# app.py
import streamlit as st
import json
import time
import os
import smtplib
import ssl
from pathlib import Path
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
import openai
from googleapiclient.errors import HttpError
from email.message import EmailMessage

from dotenv import load_dotenv
load_dotenv()  # load variables from .env into environment


# ---------- Config ----------
PROCESSED_DB = Path("processed_rows.json")

# ---------- Helpers ----------
def save_processed(data):
    PROCESSED_DB.write_text(json.dumps(data, indent=2))

def load_processed():
    if PROCESSED_DB.exists():
        return json.loads(PROCESSED_DB.read_text())
    return {"processed_row_numbers": []}

def col_letter(n):
    # 1-based column index to Excel letters (1->A, 27->AA)
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result

def update_sheet_row(sheets_service, spreadsheet_id, tab_name, row_number, row_values):
    """
    Updates a specific row in the sheet (1-based row_number).
    row_values should be a list matching the header length.
    """
    range_ = f"{tab_name}!A{row_number}:{col_letter(len(row_values))}{row_number}"
    body = {"values": [row_values]}
    try:
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_,
            valueInputOption="RAW",
            body=body
        ).execute()
    except HttpError as e:
        st.error(f"Error updating row {row_number}: {e}")
        raise

def ensure_output_headers(sheets_service, spreadsheet_id, tab_name, headers):
    """
    Ensure the first row (header) contains Score, MaxScore, Feedback, GradedAt, PerQuestionJSON, Status, EmailSent, EmailError.
    Append columns if missing.
    Returns the final header list.
    """
    range_ = f"{tab_name}!1:1"
    try:
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_, majorDimension="ROWS"
        ).execute()
        current = resp.get("values", [[""]])[0]
    except HttpError as e:
        st.error(f"Error reading header row: {e}")
        raise

    current_headers = [c.strip() for c in current]
    needed = [h for h in headers if h not in current_headers]
    if not needed:
        return current_headers

    new_header = current_headers + needed
    # write the header row back
    body = {"values": [new_header]}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_,
        valueInputOption="RAW",
        body=body
    ).execute()
    st.success(f"Added columns: {', '.join(needed)}")
    return new_header

def rows_to_objects(rows):
    if not rows or len(rows) == 0:
        return [], []
    header = [h.strip() for h in rows[0]]
    objs = []
    for i, r in enumerate(rows[1:], start=2):  # sheet row numbers start at 1; data starts at row 2
        obj = {}
        for j, h in enumerate(header):
            obj[h] = r[j] if j < len(r) else ""
        objs.append({"__row_number": i, "raw_row": r, "data": obj})
    return header, objs


# ...existing code...

def build_sheets_service_from_service_account(sa_info):
    """
    Given a service account info dict, build and return an authenticated Google Sheets API service.
    """
    credentials = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    service = build("sheets", "v4", credentials=credentials)
    return service

def read_sheet_all(sheets_service, spreadsheet_id, tab_name):
    """
    Reads all rows from the specified tab in the Google Sheet.
    Returns a list of lists (rows).
    """
    range_ = f"{tab_name}"
    try:
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_, majorDimension="ROWS"
        ).execute()
        return resp.get("values", [])
    except HttpError as e:
        st.error(f"Error reading sheet '{tab_name}': {e}")
        return []

# ...existing code...
# (Removed duplicate empty definition of parse_json_from_model)

# ---------- OpenAI grader (strict JSON output) ----------
def parse_json_from_model(text):
    # Very resilient parser for JSON anywhere in the text
    text = text.strip()
    # try direct parse
    try:
        return json.loads(text)
    except Exception:
        pass
    # find first { ... } block
    import re
    m = re.search(r"\{[\s\S]*\}", text)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            pass
    # if failed
    raise ValueError("Failed to parse JSON from model output. Output starts:\n" + text[:1000])

# ---------- Robust grader (works with old/new openai SDKs) ----------
def grade_long_answer_with_openai(openai_api_key, student_answer, reference_answer, max_score=10, question_text=""):
    """
    Robust grader:
    - Supports old openai (0.28) and new openai>=1.0.0 client styles.
    - Logs raw model output to Streamlit.
    - Returns a dict ALWAYS with keys: score, max_score, feedback, suggestions, rubric_points.
    - On failure, returns score=0 and feedback contains the error and any raw model output.
    """
    # Basic input normalization
    if student_answer is None:
        student_answer = ""
    if reference_answer is None:
        reference_answer = ""
    # assemble prompt
    system = "You are an expert teacher who grades long written answers fairly and kindly."
    user = f"""
Student answer:
\"\"\"{student_answer}\"\"\"

Reference (model) answer:
\"\"\"{reference_answer}\"\"\"

Task:
1) Compare the student's answer to the reference and assign a numeric score between 0 and {max_score}. Be strict but fair.
2) Provide a 1-2 sentence private feedback summary (what was done well / what is missing).
3) Provide up to 3 concrete, short improvement suggestions (one-liners).
4) Also return up to 3 concise rubric points that justify the score.

Important: respond ONLY with valid JSON with these fields:
{{"score": <number>, "max_score": <number>, "feedback": "<short paragraph>", "suggestions": ["..."], "rubric_points": ["..."]}}

Question context: {question_text}
    """.strip()

    raw_out = None
    last_exc = None

    # Try to detect new vs old OpenAI SDK usage
    use_new_client = False
    try:
        # new SDK exposes OpenAI in package
        from openai import OpenAI  # type: ignore
        use_new_client = True
    except Exception:
        use_new_client = False

    # choose model fallbacks
    model_candidates = ["gpt-4o", "gpt-4", "gpt-3.5-turbo"]

    for attempt in range(3):
        try:
            if use_new_client:
                # new client style
                from openai import OpenAI
                client = OpenAI(api_key=openai_api_key)
                # call chat completions
                resp = client.chat.completions.create(
                    model=model_candidates[0],
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user", "content": user}
                    ],
                    temperature=0.2,
                    max_tokens=700
                )
                # new SDK: resp.choices[0].message.content
                raw_out = None
                try:
                    raw_out = resp.choices[0].message.content
                except Exception:
                    # try alternate access
                    raw_out = getattr(resp.choices[0].message, "content", str(resp))
            else:
                # old SDK style (openai==0.28)
                import openai
                openai.api_key = openai_api_key
                resp = openai.ChatCompletion.create(
                    model=model_candidates[0],
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user", "content": user}
                    ],
                    temperature=0.2,
                    max_tokens=700
                )
                raw_out = resp["choices"][0]["message"]["content"]

            # show raw output for debugging
            try:
                st.write(f"Grader raw output (row attempt {attempt+1}, trimmed):")
                st.write((raw_out or "")[:2000])
            except Exception:
                pass

            # Try parse JSON from model output
            try:
                parsed = parse_json_from_model(raw_out or "")
                # coerce types, be defensive
                parsed_score = float(parsed.get("score", 0) or 0)
                parsed_max = float(parsed.get("max_score", max_score) or max_score)
                feedback = str(parsed.get("feedback", "") or "")
                suggestions = parsed.get("suggestions", []) or []
                rubric_points = parsed.get("rubric_points", []) or []
                return {
                    "score": parsed_score,
                    "max_score": parsed_max,
                    "feedback": feedback,
                    "suggestions": suggestions,
                    "rubric_points": rubric_points
                }
            except Exception as pe:
                # Try to ask the model to reformat only if we have raw_out
                last_exc = pe
                st.write("Parser error:", str(pe))
                if raw_out:
                    try:
                        # ask the model to reformat (old or new)
                        rescue = ("Previous output could not be parsed as JSON. "
                                  "Please reformat ONLY your previous answer as valid JSON matching the schema: "
                                  '{"score":number,"max_score":number,"feedback":"...","suggestions":[...],"rubric_points":[...]}')
                        if use_new_client:
                            resp2 = client.chat.completions.create(
                                model=model_candidates[0],
                                messages=[
                                    {"role": "system", "content": system},
                                    {"role": "user", "content": user},
                                    {"role": "assistant", "content": raw_out},
                                    {"role": "user", "content": rescue}
                                ],
                                temperature=0.1,
                                max_tokens=500
                            )
                            raw2 = resp2.choices[0].message.content
                        else:
                            resp2 = openai.ChatCompletion.create(
                                model=model_candidates[0],
                                messages=[
                                    {"role": "system", "content": system},
                                    {"role": "user", "content": user},
                                    {"role": "assistant", "content": raw_out},
                                    {"role": "user", "content": rescue}
                                ],
                                temperature=0.1,
                                max_tokens=500
                            )
                            raw2 = resp2["choices"][0]["message"]["content"]

                        st.write("Grader rescue output (trimmed):")
                        st.write((raw2 or "")[:2000])
                        parsed = parse_json_from_model(raw2 or "")
                        parsed_score = float(parsed.get("score", 0) or 0)
                        parsed_max = float(parsed.get("max_score", max_score) or max_score)
                        feedback = str(parsed.get("feedback", "") or "")
                        suggestions = parsed.get("suggestions", []) or []
                        rubric_points = parsed.get("rubric_points", []) or []
                        return {
                            "score": parsed_score,
                            "max_score": parsed_max,
                            "feedback": feedback,
                            "suggestions": suggestions,
                            "rubric_points": rubric_points
                        }
                    except Exception as e2:
                        last_exc = e2
                        st.write("Rescue parse failed:", str(e2))
                        # let loop continue to retry
                # if we cannot parse, fallthrough to retry
        except Exception as e:
            last_exc = e
            st.write(f"OpenAI call failed on attempt {attempt+1}: {e}")
            # if model name might be unsupported (e.g., gpt-4o not available), try fallback models in next loop
            # rotate model candidate to try next one
            model_candidates = model_candidates[1:] + model_candidates[:1]
            time.sleep(1 + attempt * 2)
            continue

    # after retries, return a clear failure dict (so downstream logs and email include reason)
    failure_feedback = "Auto-grading failed after retries. "
    if last_exc:
        failure_feedback += f"Error: {str(last_exc)}. "
    if raw_out:
        failure_feedback += f"Raw model output (trimmed): { (raw_out[:1000] + '...') if len(raw_out) > 1000 else raw_out }"

    st.error("Grading failed — see feedback saved into sheet for diagnostics.")
    return {
        "score": 0.0,
        "max_score": float(max_score or 10),
        "feedback": failure_feedback,
        "suggestions": [],
        "rubric_points": []
    }


# ----- Replace send_email_via_smtp with this more robust version -----
def send_email_via_smtp(smtp_config, to_email, subject, html_body, plain_body=None):
    """
    smtp_config: dict with host, port, user, pass, use_tls (bool), use_ssl (bool), from_email
    Always include a plain-text fallback to avoid mail clients showing blank messages.
    """
    if not to_email:
        raise ValueError("No recipient email provided.")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_config.get("from_email")
    msg["To"] = to_email

    # Ensure we always have a plain text body
    if not plain_body:
        # strip HTML tags lightly for fallback (simple approach)
        import re
        text_fallback = re.sub(r"<[^>]+>", "", html_body)
        text_fallback = text_fallback.strip()
        if not text_fallback:
            text_fallback = "Your feedback is available in the attached message. (No plain-text feedback generated.)"
    else:
        text_fallback = plain_body

    # Set both plain and HTML parts correctly
    msg.set_content(text_fallback)
    msg.add_alternative(html_body, subtype="html")

    host = smtp_config.get("host")
    port = int(smtp_config.get("port", 587))
    user = smtp_config.get("user")
    password = smtp_config.get("pass")
    use_ssl = smtp_config.get("use_ssl", False)
    use_tls = smtp_config.get("use_tls", True)

    # Log (Streamlit) for debug
    try:
        st.write(f"Sending email to {to_email}. Plain preview: {text_fallback[:300]}")
    except Exception:
        pass

    if use_ssl:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, context=context) as server:
            if user and password:
                server.login(user, password)
            server.send_message(msg)
    else:
        context = ssl.create_default_context()
        with smtplib.SMTP(host, port) as server:
            if use_tls:
                server.starttls(context=context)
            if user and password:
                server.login(user, password)
            server.send_message(msg)

# ---------- Streamlit UI & Main ----------
st.set_page_config(page_title="AI Long-Answer Grader (with Email & Status)", layout="wide")
st.title("AI Long-Answer Grader (Streamlit) — Status flag & Email")

st.markdown("""
This app grades long-answer questions from a Google Form (Responses in a Google Sheet),
marks `Status = Graded` in the sheet when done, and can send email feedback via SMTP.
""")

# --- Inputs ---
st.sidebar.header("Configuration")
uploaded_sa = st.sidebar.file_uploader("Upload service-account JSON", type=["json"])
sheet_id = st.sidebar.text_input("Google Sheet ID (do NOT use previous ID)", value="")
responses_tab = st.sidebar.text_input("Responses tab name", value="Form Responses 1")
answer_sheet_tab = st.sidebar.text_input("Answer Key tab name (if using sheet)", value="Answer Key")
use_local_answers = st.sidebar.checkbox("Use uploaded answers.json instead of sheet tab", value=True)
uploaded_answers = st.sidebar.file_uploader("Upload answers.json (if using local)", type=["json"])

openai_key = os.getenv("OPENAI_API_KEY")
if not openai_key:
    st.error("⚠️ OPENAI_API_KEY not found in .env file. Please set it before running.")


st.sidebar.markdown("---")
st.sidebar.header("SMTP (optional — leave blank to skip emailing)")
smtp_host = st.sidebar.text_input("SMTP Host (e.g. smtp.gmail.com)")
smtp_port = st.sidebar.text_input("SMTP Port (587/465)")
smtp_user = st.sidebar.text_input("SMTP User (email address)")
smtp_pass = st.sidebar.text_input("SMTP Password / App Password", type="password")
smtp_from = st.sidebar.text_input("From Email (e.g. Your Name <you@example.com>)")
smtp_use_tls = st.sidebar.checkbox("Use STARTTLS (True for port 587)", value=True)
smtp_use_ssl = st.sidebar.checkbox("Use SSL (True for port 465)", value=False)

model_choice = st.sidebar.selectbox("OpenAI model", options=["gpt-4o"], index=0)
run_button = st.button("Run: Grade new rows (one-shot)")

# Provide example for answers.json format
st.sidebar.markdown("**answers.json format (when uploading)**")
st.sidebar.code(json.dumps({
    "What is a decision tree?": {
        "answer": "A decision tree is a flowchart-like structure ... include Gini formula and a small example.",
        "maxScore": 10
    },
    "Explain Gini impurity": {
        "answer": "Gini impurity formula is 1 - sum(p_i^2) ...",
        "maxScore": 6
    }
}, indent=2), language="json")

if run_button:
    # Basic validation
    if not uploaded_sa:
        st.error("Please upload service-account JSON.")
    elif not sheet_id.strip():
        st.error("Please enter the new Google Sheet ID.")
    elif not openai_key:
        st.error("Please provide OpenAI API key.")
    else:
        with st.spinner("Initializing..."):
            try:
                sa_info = json.load(uploaded_sa)
                sheets_service = build_sheets_service_from_service_account(sa_info)
            except Exception as e:
                st.exception(f"Failed to build Sheets service: {e}")
                st.stop()

            # SMTP config
            smtp_config = None
            if smtp_host and smtp_port and smtp_user and smtp_pass and smtp_from:
                smtp_config = {
                    "host": smtp_host.strip(),
                    "port": smtp_port.strip(),
                    "user": smtp_user.strip(),
                    "pass": smtp_pass,
                    "from_email": smtp_from.strip(),
                    "use_tls": bool(smtp_use_tls),
                    "use_ssl": bool(smtp_use_ssl)
                }
                st.info("SMTP configured — emails will be attempted.")
            else:
                st.info("SMTP not fully configured. Emails will be skipped.")

            # load answers
            try:
                if use_local_answers:
                    if not uploaded_answers:
                        st.error("Please upload answers.json when using local answer file.")
                        st.stop()
                    answers_obj = json.load(uploaded_answers)
                    answer_key = {}
                    for k, v in answers_obj.items():
                        if isinstance(v, dict):
                            answer_key[k.strip()] = {
                                "answer": v.get("answer", ""),
                                "maxScore": int(v.get("maxScore", 10))
                            }
                        else:
                            answer_key[k.strip()] = {"answer": str(v), "maxScore": 10}
                else:
                    all_ = read_sheet_all(sheets_service, sheet_id, answer_sheet_tab)
                    if not all_ or len(all_) < 2:
                        st.error(f"Answer Key tab '{answer_sheet_tab}' must have at least header and one row.")
                        st.stop()
                    header_row = [h.strip() for h in all_[0]]
                    idx_q = header_row.index("Question") if "Question" in header_row else 0
                    idx_ans = header_row.index("ModelAnswer") if "ModelAnswer" in header_row else 1
                    idx_max = header_row.index("MaxScore") if "MaxScore" in header_row else (2 if len(header_row) > 2 else None)
                    answer_key = {}
                    for row in all_[1:]:
                        q = row[idx_q].strip() if idx_q < len(row) else ""
                        ans = row[idx_ans] if idx_ans < len(row) else ""
                        maxs = int(row[idx_max]) if (idx_max is not None and idx_max < len(row) and row[idx_max] != "") else 10
                        if q:
                            answer_key[q] = {"answer": ans, "maxScore": maxs}

                st.success(f"Loaded {len(answer_key)} answer-key items.")
            except Exception as e:
                st.exception(f"Failed to load answer key: {e}")
                st.stop()

            # Read responses sheet
            try:
                all_rows = read_sheet_all(sheets_service, sheet_id, responses_tab)
                if not all_rows or len(all_rows) < 2:
                    st.warning("Responses sheet appears empty or has only header. No rows to grade.")
                    st.stop()
                headers, row_objs = rows_to_objects(all_rows)
                st.write(f"Found {len(row_objs)} response rows (headers: {headers})")
            except Exception as e:
                st.exception(f"Failed to read responses sheet: {e}")
                st.stop()

            # Ensure output headers (including Status and EmailSent / EmailError)
            final_headers = ensure_output_headers(
                sheets_service, sheet_id, responses_tab,
                ["Score", "MaxScore", "Feedback", "GradedAt", "PerQuestionJSON", "Status", "EmailSent", "EmailError"]
            )

            processed = load_processed()
            graded_count = 0
            errors = []

            # Build a mapping header -> index for reading status quickly
            header_index = {h: i for i, h in enumerate(final_headers)}

            for row in row_objs:
                row_num = row["__row_number"]
                data = row["data"]

                # skip if Status column exists and is 'Graded'
                status_val = data.get("Status") if "Status" in data else None
                if status_val and str(status_val).strip().lower() == "graded":
                    st.info(f"Skipping row {row_num} — already graded (Status=Graded).")
                    continue

                # Also check processed_rows local backup to avoid doubles
                if row_num in processed.get("processed_row_numbers", []):
                    st.info(f"Skipping row {row_num} — already in processed_rows.json.")
                    continue

                # find student name and email (flexible)
                student_name = data.get("Name") or data.get("Student Name") or data.get("Full Name") or ""
                student_email = (data.get("Email") or data.get("Email ") or data.get("email") or "").strip()

                # Build grading queue: only grade questions that exist in answer_key and in headers
                per_question_results = []
                total_score = 0.0
                total_max = 0.0

                for qtext, qmeta in answer_key.items():
                    student_ans = data.get(qtext, "")
                    ref_ans = qmeta.get("answer", "")
                    maxs = qmeta.get("maxScore", 10)
                    st.info(f"Grading row {row_num} - question '{qtext[:40]}...'")
                    try:
                        result = grade_long_answer_with_openai(openai_key, student_ans, ref_ans, maxs, qtext)
                        score = float(result.get("score", 0))
                        max_score = float(result.get("max_score", maxs))
                        total_score += score
                        total_max += max_score
                        per_question_results.append({
                            "question": qtext,
                            "score": score,
                            "max_score": max_score,
                            "feedback": result.get("feedback", ""),
                            "suggestions": result.get("suggestions", []),
                            "rubric_points": result.get("rubric_points", [])
                        })
                        st.write(f"  → {qtext[:50]}...: {score}/{max_score}")
                    except Exception as e:
                        errors.append(f"Row {row_num} Q '{qtext}': {e}")
                        per_question_results.append({
                            "question": qtext,
                            "score": 0,
                            "max_score": maxs,
                            "feedback": "Auto-grading failed; manual review required.",
                            "suggestions": [],
                            "rubric_points": []
                        })
                        total_max += maxs

                    # small delay between calls
                    time.sleep(0.5)

                # Build combined feedback
                feedback_lines = []
                for i, pq in enumerate(per_question_results, start=1):
                    lines = [f"Q{i} ({pq['question']}) — Score: {pq['score']}/{pq['max_score']}.",
                            f"Feedback: {pq['feedback']}"]
                    if pq['suggestions']:
                        lines.append("Suggestions: " + "; ".join(pq['suggestions']))
                    feedback_lines.append("\n".join(lines))
                combined_feedback = "\n\n".join(feedback_lines)

                # If combined_feedback is empty (shouldn't happen now), set a safe fallback
                if not combined_feedback.strip():
                    combined_feedback = "No feedback generated — please review this submission manually."
                # For debugging, show feedback to Streamlit before emailing/writing:
                st.write("Combined feedback (trimmed):")
                st.write(combined_feedback[:4000])




                # Prepare row values to write back
                orig_row = row["raw_row"][:]  # may be shorter
                header_map = {}
                for idx, h in enumerate(headers):
                    header_map[h] = orig_row[idx] if idx < len(orig_row) else ""

                out_row = []
                for h in final_headers:
                    if h in headers:
                        out_row.append(header_map.get(h, ""))
                    else:
                        if h == "Score":
                            out_row.append(str(total_score))
                        elif h == "MaxScore":
                            out_row.append(str(total_max))
                        elif h == "Feedback":
                            out_row.append(combined_feedback[:50000])
                        elif h == "GradedAt":
                            out_row.append(datetime.utcnow().isoformat() + "Z")
                        elif h == "PerQuestionJSON":
                            out_row.append(json.dumps(per_question_results))
                        elif h == "Status":
                            out_row.append("Graded")
                        elif h == "EmailSent":
                            out_row.append("")  # will update after sending
                        elif h == "EmailError":
                            out_row.append("")
                        else:
                            out_row.append("")

                # Write the row back (initial write will include Status=Graded to avoid duplicate processing)
                try:
                    update_sheet_row(sheets_service, sheet_id, responses_tab, row_num, out_row)
                except Exception as e:
                    st.error(f"Failed to write graded data to row {row_num}: {e}")
                    errors.append(f"Write row {row_num} error: {e}")
                    continue

                # Attempt to send email if SMTP configured
                email_error_msg = ""
                email_sent_time = ""

                
                if smtp_config and student_email:
                    subject = "Your Graded Assignment & Feedback"
                    html_body = f"""Hi {student_name or 'Student'},<br/><br/>
Here is your feedback:<br/><br/>
<pre style="white-space:pre-wrap; font-family:inherit;">{combined_feedback}</pre>
<br/>Best regards,<br/>Your AI Learning Assistant"""
                    try:
                        send_email_via_smtp(smtp_config, student_email, subject, html_body, None)
                        email_sent_time = datetime.utcnow().isoformat() + "Z"
                        st.success(f"Email sent to {student_email}")
                    except Exception as e:
                        email_error_msg = str(e)
                        st.error(f"Failed to send email to {student_email}: {e}")

                # Update only the EmailSent and EmailError cells if needed (update the row again)
                # Build updated out_row again from scratch to ensure EmailSent / EmailError filled
                for i, h in enumerate(final_headers):
                    if h == "EmailSent":
                        out_row[i] = email_sent_time
                    if h == "EmailError":
                        out_row[i] = email_error_msg

                try:
                    update_sheet_row(sheets_service, sheet_id, responses_tab, row_num, out_row)
                    st.success(f"Row {row_num} updated with Status and Email info.")
                    graded_count += 1
                except Exception as e:
                    st.error(f"Failed to update EmailSent/EmailError for row {row_num}: {e}")
                    errors.append(f"Final update row {row_num} error: {e}")

                # mark processed locally as backup
                processed.setdefault("processed_row_numbers", []).append(row_num)
                save_processed(processed)

            st.write(f"Graded {graded_count} new rows.")
            if errors:
                st.error("Some errors occurred. See details below.")
                for e in errors:
                    st.write("- " + str(e))
            else:
                st.success("Done. No errors reported.")

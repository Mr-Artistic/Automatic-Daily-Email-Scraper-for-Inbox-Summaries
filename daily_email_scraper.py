import imaplib
import email
from email.header import decode_header
from openai import OpenAI
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import os
import datetime
import json

# =========================
# CONFIGURATION
# =========================
IMAP_SERVER = "your email server"
EMAIL_USER = "your email address"
EMAIL_PASS = "your email app password"   
OPENAI_API_KEY = "your chatgpt api key"

PDF_FILE = "email_digest.pdf"
PROCESSED_FILE = "processed_emails.json"

# Initialize OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

# =========================
# LOGGING
# =========================
def write_log(message):
    """Append timestamp + message to log file in UTF-8"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{now}] {message}\n")

# =========================
# TOKEN USAGE TRACKING
# =========================
def load_token_usage():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            return json.load(f)
    return {}

def save_token_usage(usage_data):
    with open(TOKEN_FILE, "w") as f:
        json.dump(usage_data, f, indent=2)

def log_monthly_summary(usage_data, last_month_key):
    """Write a closing summary for the previous month into the log."""
    if last_month_key in usage_data:
        prompt = usage_data[last_month_key]["prompt"]
        completion = usage_data[last_month_key]["completion"]
        total = usage_data[last_month_key]["total"]

        # Cost calculation
        prompt_cost = prompt / 1_000_000 * 0.15
        completion_cost = completion / 1_000_000 * 0.60
        monthly_cost = prompt_cost + completion_cost

        summary_message = (
            f"📊 Monthly Summary for {last_month_key}: "
            f"{total} tokens used (prompt={prompt}, completion={completion}). "
            f"Final estimated cost = ${monthly_cost:.4f} (~₹{monthly_cost*83:.2f})."
        )
        print(summary_message)
        write_log(summary_message)

# =========================
# EMAIL PROCESS TRACKING
# =========================
def load_processed_ids():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_processed_ids(ids):
    with open(PROCESSED_FILE, "w") as f:
        json.dump(list(ids), f)

# =========================
# FIRST RUN SETUP (Option 2 default)
# =========================
def setup_first_run_choice():
    """On first run, summarize unread emails now, then track new ones."""
    if not os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE, "w") as f:
            json.dump([], f)
        print("✅ First run: Summarizing unread emails, then tracking new ones going forward.")

# =========================
# FETCH UNREAD EMAILS
# =========================
def fetch_unread_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, messages = mail.search(None, 'UNSEEN')
    email_ids = messages[0].split()

    processed_ids = load_processed_ids()
    new_ids = [e_id for e_id in email_ids if e_id.decode() not in processed_ids]

    emails = []
    for e_id in new_ids:
        _, msg_data = mail.fetch(e_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        if content_type == "text/plain":
                            body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                            break
                        elif content_type == "text/html" and not body:  # fallback
                            body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                else:
                    body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

                emails.append({
                    "id": e_id.decode(),
                    "subject": subject,
                    "body": body
                })
    mail.logout()
    return emails

# =========================
# SUMMARIZE EMAILS (BATCHED, JSON OUTPUT)
# =========================
def summarize_emails_batched(emails):
    if not emails:
        return [], {"prompt": 0, "completion": 0, "total": 0}

    # Build one big prompt (ask for JSON output)
    prompt = {
        "role": "user",
        "content": f"""
        Summarize each email into JSON array. Each item must have:
        - subject
        - summary (max 2 lines)
        - tasks (list of tasks/deadlines, if any, else empty list)

        Emails:
        {json.dumps(emails, ensure_ascii=False)}
        """
    }

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[prompt],
            max_tokens=3000,
            response_format={"type": "json_object"}
        )
        result_json = json.loads(response.choices[0].message.content)

        # Token usage
        usage = {
            "prompt": response.usage.prompt_tokens,
            "completion": response.usage.completion_tokens,
            "total": response.usage.total_tokens
        }

        # Ensure mapping back to emails
        summaries = []
        for i, item in enumerate(result_json.get("emails", [])):
            if i < len(emails):
                summaries.append({
                    "id": emails[i]['id'],
                    "subject": item.get("subject", emails[i]['subject']),
                    "summary": item.get("summary", ""),
                    "tasks": item.get("tasks", [])
                })
    except Exception as e:
        summaries = [{"id": "error", "subject": "Error", "summary": str(e), "tasks": []}]
        usage = {"prompt": 0, "completion": 0, "total": 0}

    return summaries, usage

# =========================
# CREATE STYLED PDF DIGEST
# =========================
def create_pdf(summaries):
    doc = SimpleDocTemplate(PDF_FILE, pagesize=A4)
    story = []
    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle("Title", parent=styles["Heading1"], fontSize=16, textColor=colors.darkblue)
    subject_style = ParagraphStyle("Subject", parent=styles["Heading2"], fontSize=12, textColor=colors.black, spaceAfter=6)
    summary_style = ParagraphStyle("Summary", parent=styles["Normal"], fontSize=10, leading=14)
    task_style = ParagraphStyle("Task", parent=styles["Normal"], fontSize=10, textColor=colors.red, leading=14)

    # Header
    today = datetime.now().strftime("%d %B %Y")
    story.append(Paragraph(f"📅 Daily Email Digest – {today}", title_style))
    story.append(Spacer(1, 12))

    for i, mail in enumerate(summaries, 1):
        story.append(Paragraph(f"{i}. <b>{mail['subject']}</b>", subject_style))

        # Summary
        if mail['summary']:
            story.append(Paragraph(mail['summary'], summary_style))

        # Tasks (if any)
        for task in mail.get('tasks', []):
            story.append(Paragraph(f"🔴 {task}", task_style))

        story.append(Spacer(1, 8))
        story.append(Paragraph("<font color='grey'>--------------------------------------------</font>", summary_style))
        story.append(Spacer(1, 8))

    doc.build(story)

# =========================
# PRINT PDF
# =========================
def print_pdf():
    try:
        os.startfile(PDF_FILE, "print")  # Windows auto print
    except Exception as e:
        print("Printing failed:", e)

# =========================
# MAIN SCRIPT
# =========================
if __name__ == "__main__":
    setup_first_run_choice()   # Default Option 2

    emails = fetch_unread_emails()
    if not emails:
        final_message = "No new emails today."
        print(final_message)
        write_log(final_message)
    else:
        summaries, usage = summarize_emails_batched(emails)
        create_pdf(summaries)
        print_pdf()

        # Save processed IDs
        processed_ids = load_processed_ids()
        for mail in summaries:
            processed_ids.add(mail['id'])
        save_processed_ids(processed_ids)

        # Track token usage per month
        usage_data = load_token_usage()
        month_key = datetime.now().strftime("%Y-%m")

        # If new month → log last month's summary & reset
        if usage_data and month_key not in usage_data:
            last_month_key = sorted(usage_data.keys())[-1]
            log_monthly_summary(usage_data, last_month_key)

        # Ensure current month entry exists
        if month_key not in usage_data:
            usage_data[month_key] = {"prompt": 0, "completion": 0, "total": 0}

        # Update this month's tokens
        usage_data[month_key]["prompt"] += usage["prompt"]
        usage_data[month_key]["completion"] += usage["completion"]
        usage_data[month_key]["total"] += usage["total"]
        save_token_usage(usage_data)

        # Cost calculation
        prompt_cost = usage_data[month_key]["prompt"] / 1_000_000 * 0.15
        completion_cost = usage_data[month_key]["completion"] / 1_000_000 * 0.60
        monthly_cost = prompt_cost + completion_cost

        email_count = len(summaries)
        monthly_total = usage_data[month_key]["total"]

        final_message = (
            f"✅ Batched digest created, printed, and processed {email_count} emails. "
            f"Tokens this run: {usage['total']} (prompt={usage['prompt']}, completion={usage['completion']}). "
            f"Monthly total so far: {monthly_total} tokens. "
            f"Estimated monthly cost so far: ${monthly_cost:.4f} (~₹{monthly_cost*83:.2f})."
        )
        print(final_message)
        write_log(final_message)

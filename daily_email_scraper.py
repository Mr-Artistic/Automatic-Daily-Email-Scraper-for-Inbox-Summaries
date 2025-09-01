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
# LOAD/SAVE PROCESSED IDS
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
# FIRST RUN SETUP 
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
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                            break
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
# SUMMARIZE EMAILS (BATCHED)
# =========================
def summarize_emails_batched(emails):
    if not emails:
        return []

    # Build one big prompt
    prompt = "Summarize each email separately into:\n- Short summary (max 2 lines)\n- Tasks or deadlines if any\n\n"
    for i, email_data in enumerate(emails, 1):
        prompt += f"Email {i}\nSubject: {email_data['subject']}\nBody: {email_data['body']}\n\n"

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=3000
        )
        result_text = response.choices[0].message.content.strip()

        # Split output by "Email" sections
        results = [res.strip() for res in result_text.split("\n\n") if res.strip()]
        summaries = []
        for i, res in enumerate(results):
            if i < len(emails):  # ensure mapping stays aligned
                summaries.append({
                    "id": emails[i]['id'],
                    "subject": emails[i]['subject'],
                    "summary": res
                })
    except Exception as e:
        summaries = [{"id": "error", "subject": "Error", "summary": str(e)}]

    return summaries

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
    today = datetime.date.today().strftime("%d %B %Y")
    story.append(Paragraph(f"📅 Daily Email Digest – {today}", title_style))
    story.append(Spacer(1, 12))

    for i, mail in enumerate(summaries, 1):
        story.append(Paragraph(f"{i}. <b>{mail['subject']}</b>", subject_style))

        # Split summary into lines & highlight tasks
        for line in mail['summary'].split("\n"):
            if any(keyword in line.lower() for keyword in ["task", "deadline", "due", "reminder"]):
                story.append(Paragraph(line.strip(), task_style))
            else:
                story.append(Paragraph(line.strip(), summary_style))

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
    setup_first_run_choice()   

    emails = fetch_unread_emails()
    if not emails:
        print("No new emails today.")
    else:
        summaries = summarize_emails_batched(emails)
        create_pdf(summaries)
        print_pdf()

        # Save processed IDs
        processed_ids = load_processed_ids()
        for mail in summaries:
            processed_ids.add(mail['id'])
        save_processed_ids(processed_ids)

        print("✅ Batched digest created, printed, and processed emails saved.")

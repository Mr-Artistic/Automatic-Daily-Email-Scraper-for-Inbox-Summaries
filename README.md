
# 📬 Automated Email Digest with ChatGPT

1. This Python project connects to your email inbox via IMAP, fetches new/unread emails, and summarizes them using the OpenAI GPT-4o-mini model in a batched request (optimized for API cost and speed). 

2. The summaries are formatted into a clean, styled PDF digest and automatically printed through your PC. 

3. This script can be scheduled with Windows Task Scheduler to generate a daily digest at fixed times (e.g., 10 AM and 3 PM).

4. The script also tracks processed emails to avoid duplicates.

# Additional Requirements:

1. Pyhton Packages -
openai>=1.0.0
reportlab>=4.0.0

2. OpenAI API
3. Email IMAP Server
4. Email App Password

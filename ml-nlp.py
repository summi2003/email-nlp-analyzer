"""
NLP + ML Email Keyword Analysis

Description:
- Connects to Gmail via IMAP
- Fetches last 100 emails.
- Cleans text, tokenizes, removes stopwords (NLP)
- Extracts top 20 keywords using TF-IDF (ML)
- Outputs 2 Excel files:
    1) ml-nlp_top_20_keywords.xlsx
    2) ml-nlp_email_keyword_details.xlsx
"""

import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import re
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from datetime import datetime, timedelta

# ==============================
# CONFIGURATION
# ==============================
EMAIL_ACCOUNT = "abcd7@gmail.com"
APP_PASSWORD = "eurjdnjkdifsijk"

IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993
MAX_EMAILS = 250

OUTPUT_KEYWORD_FILE = "ml-nlp_top_20_keywords.xlsx"
OUTPUT_EMAIL_FILE = "ml-nlp_email_keyword_details.xlsx"

# ==============================
# HELPER FUNCTIONS
# ==============================

def clean_text(text):
    """Lowercase, remove punctuation, numbers, and extra spaces"""
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r"[^a-z\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def extract_body(msg):
    """Extract email body (plain text preferred, HTML fallback)"""
    body = ""

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if content_type == "text/plain" and "attachment" not in content_disposition:
                body = part.get_payload(decode=True)
                break

            elif content_type == "text/html" and not body:
                html = part.get_payload(decode=True)
                soup = BeautifulSoup(html, "html.parser")
                body = soup.get_text()

    else:
        body = msg.get_payload(decode=True)

    try:
        return body.decode(errors="ignore")
    except:
        return ""

# ==============================
# MAIN LOGIC
# ==============================

def main():
    print("Connecting to Gmail...")

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_ACCOUNT, APP_PASSWORD)
    mail.select("inbox")

    # -------- DATE FILTER : LAST 2 YEARS --------
    two_years_ago = (datetime.now() - timedelta(days=730)).strftime("%d-%b-%Y")

    status, messages = mail.search(None, f'SINCE {two_years_ago}')
    email_ids = messages[0].split()

    # Take only last 100 emails
    email_ids = email_ids[-MAX_EMAILS:]

    print(f"Processing {len(email_ids)} emails from last 2 years...")

    email_texts = []
    email_meta = []

    for eid in email_ids:
        status, msg_data = mail.fetch(eid, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        # Decode subject
        subject, enc = decode_header(msg.get("Subject"))[0]
        if isinstance(subject, bytes):
            subject = subject.decode(enc or "utf-8", errors="ignore")

        sender = msg.get("From")
        receiver = msg.get("To")
        date_ = msg.get("Date")

        body = extract_body(msg)

        subject_clean = clean_text(subject)
        body_clean = clean_text(body)

        full_text = subject_clean + " " + body_clean
        email_texts.append(full_text)

        email_meta.append({
            "Date": date_,
            "From": sender,
            "To": receiver,
            "Subject": subject_clean,
            "Body": body_clean
        })

    mail.logout()

    # ==============================
    # TF-IDF (ML PART)
    # ==============================

    vectorizer = TfidfVectorizer(
        
        stop_words="english",
        max_features=50
    )

    tfidf_matrix = vectorizer.fit_transform(email_texts)
    feature_names = vectorizer.get_feature_names_out()
    tfidf_scores = tfidf_matrix.sum(axis=0).A1

    keyword_scores = sorted(
        zip(feature_names, tfidf_scores),
        key=lambda x: x[1],
        reverse=True
    )[:30]

    # -------- Excel File 1: Top Keywords --------
    df_keywords = pd.DataFrame(keyword_scores, columns=["Keyword", "TF-IDF Score"])
    df_keywords.to_excel(OUTPUT_KEYWORD_FILE, index=False, engine="openpyxl")

    print(f"✅ Keyword file created: {OUTPUT_KEYWORD_FILE}")

    # ==============================
    # EMAIL KEYWORD DETAIL TABLE
    # ==============================

    rows = []

    for meta in email_meta:
        for keyword, _ in keyword_scores:
            subject_count = meta["Subject"].split().count(keyword)
            body_count = meta["Body"].split().count(keyword)

            if subject_count > 0 and body_count > 0:
                location = "Both"
            elif subject_count > 0:
                location = "Subject"
            elif body_count > 0:
                location = "Body"
            else:
                continue

            rows.append({
                "Date of Mail": meta["Date"],
                "Sender": meta["From"],
                "Receiver": meta["To"],
                "Keyword": keyword,
                "Count in Subject": subject_count,
                "Count in Body": body_count,
                "Location": location
            })

    df_email_details = pd.DataFrame(rows)
    df_email_details.to_excel(OUTPUT_EMAIL_FILE, index=False, engine="openpyxl")

    print(f"✅ Email analysis file created: {OUTPUT_EMAIL_FILE}")
    print("🎉 NLP + ML Email Analysis Completed Successfully")

# ==============================
# RUN SCRIPT
# ==============================

if __name__ == "__main__":
    main()

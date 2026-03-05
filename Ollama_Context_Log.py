import imaplib
import email
from email.header import decode_header
import os
import platform
import requests
import json
from datetime import datetime
from fpdf import FPDF

# ------ FILL OUT ------
COMPANY_NAME = ''
KEYWORDS = ['']
TARGET_YEAR = 2025  # AI will look at TARGET_YEAR, TARGET_YEAR-1, and TARGET_YEAR+1

# Login Details
MY_EMAIL = ''
APP_PASSWORD = '' # Your 16-character App Password

# Ollama Local AI Configuration
OLLAMA_URL = ""
MODEL_NAME = "llama3:8b-instruct-q8_0"

BASE_OUTPUT_PATH = '' # Set Mac folder path here

# ---------------------

# --- FOLDER SETUP ---
SAVE_FOLDER_NAME = f"{COMPANY_NAME} Context Log"
SAVE_BASE_DIR = os.path.join(BASE_OUTPUT_PATH, SAVE_FOLDER_NAME)
if not os.path.exists(SAVE_BASE_DIR): os.makedirs(SAVE_BASE_DIR)

"""
    Search string that Gmail's server understands (IMAP protocol) +- 1 TARGET_YEAR.
    It handles the 'OR' logic for zero or multiple keywords and sets the date boundaries.
    IMAP documentation: https://tools.ietf.org/html/rfc3501#section-6.4.4
"""
def build_ai_imap_query(year, keywords, company):
    start_year = year - 1
    end_year = year + 2
    date_range = f'SINCE "01-Jan-{start_year}" BEFORE "01-Jan-{end_year}"'
    # Start with the company name as a baseline keyword
    query = f'(TEXT "{company}")'
    for kw in keywords:
        query = f'OR (TEXT "{kw}") {query}'
    return f'({query} {date_range})'

"""
    A filter to prevent FPDF from crashing by deleting non-Latin-1 characters (emojis, etc).
    Maybe there's a more elegant way to do this, but it works and is essential for messy lab emails.
"""
def clean_for_pdf(text):
    if not text: return ""
    return text.encode('latin-1', 'ignore').decode('latin-1')

"""
    Consults Llama 3 to see if the metadata matches the project scope. Prompt can be adjusted
    for more nuance, but this is a simple yes/no filter to reduce noise.
"""
def ask_ollama(subject, sender, date):
    prompt = (
        f"Context: I am organizing lab records for a project named '{COMPANY_NAME}'. "
        f"The primary activity was in {TARGET_YEAR}. "
        f"Identify if this email is relevant to this specific research project:\n"
        f"Subject: {subject}\nFrom: {sender}\nDate: {date}\n"
        f"Respond with ONLY 'YES' or 'NO'."
    )
    payload = {"model": MODEL_NAME, "prompt": prompt, "stream": False}
    
    try:
        response = requests.post(OLLAMA_URL, json=payload)
        result = response.json().get('response', 'NO').strip().upper()
        return "YES" in result
    except:
        return True # Safety default: Keep if AI is unreachable

"""
    Uses everything above to connect to Gmail, download the data filtered by Ollama, and build the context log. 
"""
def create_ai_audit():
    # Setup Directory
    save_dir = os.path.join(BASE_OUTPUT_PATH, f"{COMPANY_NAME} Context Log")
    if not os.path.exists(save_dir): os.makedirs(save_dir)

    print(f"Initializing Ollama-Audit for {COMPANY_NAME}...")
    
    # ---- Connect to Gmail Server ----
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(MY_EMAIL, APP_PASSWORD)
    mail.select("inbox")
    
    # --- Read def build_ai_imap_query ---
    search_query = build_ai_imap_query(TARGET_YEAR, KEYWORDS, COMPANY_NAME)
    status, data = mail.search(None, search_query)
    mail_ids = data[0].split()
    print(f"Found {len(mail_ids)} candidates. Consulting Ollama...")

    # ----- Build PDF Transcript ------
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("helvetica", 'B', 20)    # Large Project Title
    pdf.cell(0, 20, txt=f"PROJECT TRANSCRIPT: {COMPANY_NAME.upper()}", ln=True, align='C')    
    pdf.set_font("helvetica", size=10)    # Subheader with Metadata
    pdf.cell(0, 10, txt=f"Audit Year: {TARGET_YEAR} | Model: {MODEL_NAME}", ln=True, align='C')
    pdf.ln(10) # Large gap before first email starts
    accepted_count = 0

    # ----- Process Each Email -----
    for num in mail_ids:
        res, msg_data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        
        # Extract metadata for the AI filter
        sender = msg.get('From')
        subject, encoding = decode_header(msg.get("Subject"))[0]
        if isinstance(subject, bytes): subject = subject.decode(encoding or "utf-8")
        date_raw = msg.get('Date')

        # AI Filtering Decision
        if ask_ollama(subject, sender, date_raw):
            accepted_count += 1
            print(f"   [KEEP]: {subject[:40]}...")
            
            # Timestamping for attachments
            dt = email.utils.parsedate_to_datetime(date_raw)
            date_prefix = dt.strftime('%Y%m%d')

            # Create shaded header block for the email
            pdf.set_font("helvetica", 'B', 10)
            pdf.set_fill_color(240, 240, 240)
            header_text = f"FROM: {sender}\nDATE: {date_raw}\nSUBJECT: {subject}"
            pdf.multi_cell(0, 7, txt=clean_for_pdf(header_text), border=1, fill=True)
            
            # Content & Attachment Extraction
            body_content = ""
            # "Walks" through email looking at each "part"
            for part in msg.walk():
                ctype = part.get_content_type()
                cdisp = str(part.get('Content-Disposition'))

                # If the "part" is plain text, add to body content
                if ctype == "text/plain" and 'attachment' not in cdisp:
                    payload = part.get_payload(decode=True)
                    if payload: body_content += payload.decode('utf-8', errors='ignore')

                # If the "part" is an attachment (Excel, PDF, Word), save it to the folder                
                elif 'attachment' in cdisp:
                    filename = part.get_filename()
                    # Rename attachments to preferred format (YYYYMMDD_OriginalName.file)
                    if filename:
                        final_fn = f"{date_prefix}_{filename.replace(' ', '_').replace('/', '_')}"
                        with open(os.path.join(save_dir, final_fn), 'wb') as f:
                            f.write(part.get_payload(decode=True))

            pdf.set_font("helvetica", size=9)
            pdf.multi_cell(0, 5, txt=clean_for_pdf(body_content))
            pdf.ln(10)
        else:
            print(f"   [SKIP]: {subject[:40]}...")
            
    # Save the finished PDF transcript into the folder
    pdf.output(os.path.join(save_dir, f"PROJECT_TRANSCRIPT_{COMPANY_NAME}.pdf"))
    print(f"\nSUCCESS: {accepted_count} emails archived in {save_dir}")
    mail.logout() # End the session with Gmail

if __name__ == "__main__":
    create_ai_audit()

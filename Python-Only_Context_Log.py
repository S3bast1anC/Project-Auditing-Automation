import imaplib
import email
from email.header import decode_header
import os
import platform
from datetime import datetime
from fpdf import FPDF

# ----- FILL OUT -----
COMPANY_NAME = ''
TARGET_EMAIL = ''   # Leave as '' to search keywords only
KEYWORDS = [''] # Leave as [] to search email only
TARGET_YEAR = '2025'

# Login Details
MY_EMAIL = ''
APP_PASSWORD = '' # Your 16-character App Password

# Detects if you are on Windows or Mac
if platform.system() == "Windows":
    BASE_OUTPUT_PATH = r''  # Set your Windows folder path here
    print("Operating System: Windows detected.")
else:
    BASE_OUTPUT_PATH = ''   # Set your Mac folder path here
    print("Operating System: Mac detected.")
# --------------------

# --- FOLDER SETUP ---
SAVE_FOLDER_NAME = f"{COMPANY_NAME} Context Log"
SAVE_BASE_DIR = os.path.join(BASE_OUTPUT_PATH, SAVE_FOLDER_NAME)
if not os.path.exists(SAVE_BASE_DIR):
    os.makedirs(SAVE_BASE_DIR)

""" 
    Search string that Gmail's server understands (IMAP protocol).
    It handles the 'OR' logic for zero or multiple keywords and sets the date boundaries.
    IMAP documentation: https://tools.ietf.org/html/rfc3501#section-6.4.4
""" 
def build_imap_query(email_addr, keywords, year):
    date_range = f'SINCE "01-Jan-{year}" BEFORE "01-Jan-{int(year)+1}"'
    if keywords:
        query = f'(TEXT "{keywords[0]}")'
        for kw in keywords[1:]:
            query = f'OR (TEXT "{kw}") {query}'
        if email_addr:
            return f'(FROM "{email_addr}" {query} {date_range})'
        return f'({query} {date_range})'
    return f'(FROM "{email_addr}" {date_range})'

"""
    A filter to prevent FPDF from crashing by deleting non-Latin-1 characters (emojis, etc).
    Maybe there's a more elegant way to do this, but it works and is essential for messy lab emails.
"""
def clean_for_pdf(text):
    if not text: return ""
    return text.encode('latin-1', 'ignore').decode('latin-1')

"""
    Uses everything above to connect to Gmail, download the data, and build the context log.
"""
def create_audit_package():
    print(f"Building Full Transcript for {TARGET_YEAR} {COMPANY_NAME} Project...")
    # ---- Connect to Gmail Server ----
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        mail.login(MY_EMAIL, APP_PASSWORD)
    except Exception as e:
        print(f"Connection Failed: {e}")
        return
    mail.select("inbox")

    # --- Read def build_imap_query ---
    search_query = build_imap_query(TARGET_EMAIL, KEYWORDS, TARGET_YEAR)
    status, data = mail.search(None, search_query)
    mail_ids = data[0].split() # List of emails matching the search criteria
    print(f"Found {len(mail_ids)} emails. Stitching Transcript...")

    # ----- Build PDF Transcript ------
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Title Page
    pdf.set_font("helvetica", 'B', 20)
    pdf.cell(0, 20, txt=f"PROJECT TRANSCRIPT: {COMPANY_NAME.upper()}", ln=True, align='C')
    pdf.set_font("helvetica", size=12)
    pdf.cell(0, 10, txt=f"Year: {TARGET_YEAR} | Source: {MY_EMAIL}", ln=True, align='C')
    pdf.ln(20)

    # ---- Process Each Email -----
    for num in mail_ids:
        res, msg_data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        
        # Extract Date. Converts to 'YYYY-MM-DD' format
        raw_date = msg.get('Date')
        try:
            dt = email.utils.parsedate_to_datetime(raw_date)
            date_stamp = dt.strftime('%Y-%m-%d %H:%M')
            date_prefix = dt.strftime('%Y%m%d')
        except:
            date_stamp = "Unknown Date"; date_prefix = "00000000"
        
        # Extract Sender and Subject
        sender = msg.get('From')
        subject, encoding = decode_header(msg.get("Subject"))[0]
        if isinstance(subject, bytes): subject = subject.decode(encoding or "utf-8")
        print(f"   + Adding: {subject[:40]}...")

        # ----- Writing the PDF -----
        # Header Block
        pdf.set_font("helvetica", 'B', 10)
        pdf.set_fill_color(240, 240, 240)
        safe_header = f"FROM: {sender}\nDATE: {date_stamp}\nSUBJECT: {subject}"
        pdf.multi_cell(0, 7, txt=clean_for_pdf(safe_header), border=1, fill=True)
        body_content = ""
        attachments_found = []

        # "Walks" through email looking at each "part"
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                cdisp = str(part.get('Content-Disposition'))

                # If the "part" is plain text, add to body content
                if ctype == "text/plain" and 'attachment' not in cdisp:
                    payload = part.get_payload(decode=True)
                    if payload: body_content += payload.decode('utf-8', errors='ignore')
                
                # If the "part" is an attachment (Excel, PDF, Word), save it to the folder
                if 'attachment' in cdisp:
                    filename = part.get_filename()
                    # Rename attachments to preferred format (YYYYMMDD_OriginalName.file)
                    if filename:
                        clean_fn = filename.replace(' ', '_').replace('/', '_').replace('\\', '_')
                        final_fn = f"{date_prefix}_{clean_fn}"
                        attachments_found.append(final_fn)
                        with open(os.path.join(SAVE_BASE_DIR, final_fn), 'wb') as f:
                            f.write(part.get_payload(decode=True))
        else:
            # If the email isn't multipart, just decode the payload directly
            payload = msg.get_payload(decode=True)
            if payload: body_content = payload.decode('utf-8', errors='ignore')

        # If attachments were found, print a red note in the PDF so you know where they are
        if attachments_found:
            pdf.set_font("helvetica", 'I', 9)
            pdf.set_text_color(200, 0, 0)
            pdf.multi_cell(0, 6, txt=f"ATTACHED FILES SAVED: {', '.join(attachments_found)}")
            pdf.set_text_color(0, 0, 0)

        # Body Text Layout
        pdf.set_font("helvetica", size=9)
        pdf.multi_cell(0, 5, txt=clean_for_pdf(body_content))
        pdf.ln(10)
        pdf.cell(0, 0, '', 'T')
        pdf.ln(5)

    # Save the finished PDF transcript into the folder
    final_log_path = os.path.join(SAVE_BASE_DIR, f"TRANSCRIPT_{COMPANY_NAME}.pdf")
    pdf.output(final_log_path)
    print(f"\nSUCCESS\n Output Folder: {SAVE_BASE_DIR}")
    mail.logout() # End the session with Gmail

"""
This ensures the audit only runs when YOU want it to.
If another program 'borrows' a tool from this file, 
this block prevents the audit from starting automatically.
"""
if __name__ == "__main__":
    create_audit_package()

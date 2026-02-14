import requests
import pdfplumber
from ebooklib import epub
import json
import os
import smtplib
from email.message import EmailMessage
import re
from datetime import datetime
import time

# ---------------- SETTINGS ---------------- #
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
    "Referer": "https://www.bseindia.com/",
    "Origin": "https://www.bseindia.com",
    "Accept": "application/json, text/plain, */*"
}
TIMEOUT = 30

EMAIL = os.environ.get("EMAIL_ADDRESS")
PASSWORD = os.environ.get("EMAIL_PASSWORD")
KINDLE = os.environ.get("KINDLE_EMAIL")

# ---------------- FUNCTIONS ---------------- #

def get_latest_transcript_link(bse_code):
    api_url = f"https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w?strCat=-1&strPrevDate=&strScripCode={bse_code}&strSearch=P&strToDate=&strType=C&subcategory=-1"
    try:
        r = requests.get(api_url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code != 200: return None, None, None
        
        data = r.json()
        for item in data.get("Table", []):
            headline = item.get("HEADLINE", "").lower()
            # Look for Transcript keywords
            if "transcript" in headline and ("earnings" in headline or "call" in headline):
                pdf_file = item.get("ATTACHMENTNAME", "")
                ann_date = item.get("NEWS_DT", "")
                if pdf_file:
                    full_url = f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{pdf_file}.pdf"
                    return full_url, item.get("HEADLINE"), ann_date
        return None, None, None
    except Exception as e:
        print(f"API Error for {bse_code}: {e}")
        return None, None, None

def download_pdf(url, filename):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code == 200 and "application/pdf" in r.headers.get("Content-Type", ""):
            with open(filename, "wb") as f:
                f.write(r.content)
            return True
    except Exception as e:
        print(f"Download Error: {e}")
    return False

def process_transcript(filename):
    text = ""
    with pdfplumber.open(filename) as pdf:
        for page in pdf.pages:
            content = page.extract_text()
            if content: text += content + "\n"
    
    # Cleaning and Splitting
    text = re.sub(r"Safe Harbor.*", "", text, flags=re.IGNORECASE | re.DOTALL)
    qa_pattern = r"(question\s*[-&]?\s*and\s*answer\s*session|q\s*&\s*a\s*session)"
    split = re.split(qa_pattern, text, maxsplit=1, flags=re.IGNORECASE)
    
    mgt = split[0]
    qa = split[2] if len(split) > 2 else ""
    
    # Highlights (lines with currency or percentages)
    highlights = [l.strip() for l in text.split('\n') if any(x in l for x in ["₹", "%", "Cr", "Mn", "billion"])]
    return mgt.strip(), qa.strip(), highlights[:15]

def create_epub(company, mgt, qa, highlights, title, date):
    safe_name = re.sub(r'\W+', '', company)
    filename = f"{safe_name}_Transcript.epub"
    
    book = epub.EpubBook()
    book.set_title(f"{company} - {title}")
    book.set_language('en')
    book.add_author("BSE Automated Scraper")
    
    content = f"<h1>{company}</h1><h3>{title}</h3><p>Date: {date}</p>"
    content += "<h2>Highlights</h2><ul>" + "".join(f"<li>{h}</li>" for h in highlights) + "</ul>"
    content += f"<h2>Management</h2><p>{mgt.replace(chr(10), '<br>')}</p>"
    content += f"<h2>Q&A</h2><p>{qa.replace(chr(10), '<br>')}</p>"
    
    c1 = epub.EpubHtml(title="Transcript", file_name="transcript.xhtml", content=content)
    book.add_item(c1)
    book.spine = [c1]
    epub.write_epub(filename, book)
    return filename

def send_to_kindle(file_path):
    msg = EmailMessage()
    msg["Subject"] = "Transcript Update"
    msg["From"], msg["To"] = EMAIL, KINDLE
    with open(file_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="epub+zip", filename=file_path)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL, PASSWORD)
        smtp.send_message(msg)

# ---------------- MAIN ---------------- #

if __name__ == "__main__":
    if not os.path.exists("processed.json"):
        with open("processed.json", "w") as f: json.dump([], f)
    
    with open("watchlist.json") as f: watchlist = json.load(f)["companies"]
    with open("processed.json") as f: processed = json.load(f)

    for company in watchlist:
        print(f"Checking {company['name']}...")
        link, title, date = get_latest_transcript_link(company["bse_code"])
        
        if link and link not in processed:
            pdf_path = "temp.pdf"
            if download_pdf(link, pdf_path):
                mgt, qa, highlights = process_transcript(pdf_path)
                epub_file = create_epub(company["name"], mgt, qa, highlights, title, date)
                send_to_kindle(epub_file)
                processed.append(link)
                print(f"Success: {company['name']}")
                time.sleep(5) # Avoid BSE rate limits
    
    with open("processed.json", "w") as f:
        json.dump(processed, f, indent=2)

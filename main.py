import requests
import pdfplumber
from ebooklib import epub
import json
import os
import smtplib
from email.message import EmailMessage
import re
from datetime import datetime, timedelta
import time

# ---------------- SETTINGS ---------------- #
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "https://www.bseindia.com/",
    "Origin": "https://www.bseindia.com"
}

EMAIL = os.environ.get("EMAIL_ADDRESS")
PASSWORD = os.environ.get("EMAIL_PASSWORD")
KINDLE = os.environ.get("KINDLE_EMAIL")

def get_dates():
    """Generates YYYYMMDD strings for today and 30 days ago."""
    end_dt = datetime.now()
    start_dt = end_dt - timedelta(days=30)
    return start_dt.strftime("%Y%m%d"), end_dt.strftime("%Y%m%d")

def get_transcripts_for_period(bse_code, company_name):
    from_date, to_date = get_dates()
    # strType=C for Corporate, strCat=-1 for All Categories
    api_url = (
        f"https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w"
        f"?strCat=-1&strPrevDate={from_date}&strScripCode={bse_code}"
        f"&strSearch=P&strToDate={to_date}&strType=C&subcategory=-1"
    )
    
    found_items = []
    try:
        r = requests.get(api_url, headers=HEADERS, timeout=25)
        data = r.json()
        table = data.get("Table", [])
        
        if not table:
            return []

        for item in table:
            headline = item.get("HEADLINE", "")
            # Looking for keywords in the headline
            if "transcript" in headline.lower():
                pdf_file = item.get("ATTACHMENTNAME", "")
                if pdf_file:
                    url = f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{pdf_file}.pdf"
                    found_items.append({
                        "url": url,
                        "title": headline,
                        "date": item.get("NEWS_DT")
                    })
        return found_items
    except Exception as e:
        print(f"  [!] API Error for {company_name}: {e}")
        return []

def process_pdf(url):
    r = requests.get(url, headers=HEADERS)
    with open("temp.pdf", "wb") as f:
        f.write(r.content)
    
    text = ""
    with pdfplumber.open("temp.pdf") as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    
    # Clean text and split Q&A
    parts = re.split(r"(Question-and-Answer|Q&A|Questions and Answers)", text, flags=re.IGNORECASE)
    mgt = parts[0][:40000] # Kindle limit safety
    qa = (parts[2] if len(parts) > 2 else "No Q&A section detected.")[:40000]
    highlights = [l.strip() for l in text.split('\n') if any(x in l for x in ["₹", "%", "Cr", "Mn"])]
    
    return mgt, qa, highlights[:15]

def create_and_send(company, mgt, qa, highlights, title, date):
    book = epub.EpubBook()
    book.set_title(f"{company} - {date[:10]}")
    book.add_author("BSE Scraper")
    
    content = f"<h1>{company}</h1><p><b>Date:</b> {date}</p><h3>{title}</h3>"
    content += "<h2>Highlights</h2><ul>" + "".join(f"<li>{h}</li>" for h in highlights) + "</ul>"
    content += f"<h2>Management</h2><p>{mgt.replace(chr(10), '<br>')}</p>"
    content += f"<h2>Q&A</h2><p>{qa.replace(chr(10), '<br>')}</p>"
    
    chap = epub.EpubHtml(title="Transcript", file_name="transcript.xhtml", content=content)
    book.add_item(chap)
    book.spine = [chap]
    
    filename = f"transcript.epub"
    epub.write_epub(filename, book)
    
    msg = EmailMessage()
    msg["Subject"] = f"Transcript: {company} ({date[:10]})"
    msg["From"], msg["To"] = EMAIL, KINDLE
    with open(filename, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="epub+zip", filename=f"{company}.epub")
    
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL, PASSWORD)
        smtp.send_message(msg)
    print(f"  [SUCCESS] Sent: {company} - {date[:10]}")

if __name__ == "__main__":
    with open("watchlist.json") as f: 
        watchlist = json.load(f)["companies"]
    try:
        with open("processed.json") as f: 
            processed = json.load(f)
    except: 
        processed = []

    for comp in watchlist:
        print(f"Scanning {comp['name']} (Last 30 Days)...")
        transcripts = get_transcripts_for_period(comp['bse_code'], comp['name'])
        
        for ts in transcripts:
            if ts['url'] not in processed:
                try:
                    mgt, qa, highs = process_pdf(ts['url'])
                    create_and_send(comp['name'], mgt, qa, highs, ts['title'], ts['date'])
                    processed.append(ts['url'])
                    time.sleep(3) # Be nice to BSE
                except Exception as e:
                    print(f"  [!] Error processing {comp['name']}: {e}")
            else:
                # Already processed this specific link
                pass

    with open("processed.json", "w") as f:
        json.dump(processed, f, indent=2)

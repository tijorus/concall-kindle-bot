import requests
from bs4 import BeautifulSoup
import pdfplumber
from ebooklib import epub
import json
import os
import smtplib
from email.message import EmailMessage
import re

# -------- SETTINGS --------
BSE_BASE = "https://www.bseindia.com/corporates/ann.aspx?code="
HEADERS = {"User-Agent": "Mozilla/5.0"}

EMAIL = os.environ["EMAIL_ADDRESS"]
PASSWORD = os.environ["EMAIL_PASSWORD"]
KINDLE = os.environ["KINDLE_EMAIL"]

# -------- LOAD FILES --------
with open("watchlist.json") as f:
    watchlist = json.load(f)["companies"]

with open("processed.json") as f:
    processed = json.load(f)

# -------- FUNCTIONS --------

def get_transcript_link(bse_code):
    url = BSE_BASE + bse_code
    r = requests.get(url, headers=HEADERS)
    soup = BeautifulSoup(r.text, "html.parser")

    for link in soup.find_all("a", href=True):
        if "Transcript" in link.text or "Earnings" in link.text:
            return "https://www.bseindia.com" + link["href"]
    return None


def download_pdf(url, filename):
    r = requests.get(url, headers=HEADERS)
    with open(filename, "wb") as f:
        f.write(r.content)


def extract_text_from_pdf(filename):
    text = ""
    with pdfplumber.open(filename) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text


def clean_text(text):
    text = re.sub(r"Safe Harbor.*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def structure_text(text):
    if "Question-and-Answer" in text:
        parts = text.split("Question-and-Answer")
        management = parts[0]
        qa = parts[1]
    else:
        management = text
        qa = ""
    return management, qa


def create_epub(company, content):
    book = epub.EpubBook()
    book.set_title(company)
    book.set_language("en")

    chapter = epub.EpubHtml(title="Transcript", file_name="chap_01.xhtml")
    chapter.content = f"""
    <h1>{company}</h1>
    <h2>Management Commentary</h2>
    <p>{content[0].replace("\n", "<br>")}</p>
    <h2>Q&A Session</h2>
    <p>{content[1].replace("\n", "<br>")}</p>
    """

    book.add_item(chapter)
    book.spine = ["nav", chapter]
    epub.write_epub(f"{company}.epub", book)
    return f"{company}.epub"


def send_to_kindle(file_path):
    msg = EmailMessage()
    msg["Subject"] = "Concall Transcript"
    msg["From"] = EMAIL
    msg["To"] = KINDLE

    with open(file_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="epub+zip", filename=file_path)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL, PASSWORD)
        smtp.send_message(msg)


# -------- MAIN LOOP --------

for company in watchlist:
    name = company["name"]
    code = company["bse_code"]

    link = get_transcript_link(code)

    if link and link not in processed:
        pdf_file = f"{name}.pdf"
        download_pdf(link, pdf_file)

        text = extract_text_from_pdf(pdf_file)
        cleaned = clean_text(text)
        structured = structure_text(cleaned)

        epub_file = create_epub(name, structured)
        send_to_kindle(epub_file)

        processed.append(link)

        with open("processed.json", "w") as f:
            json.dump(processed, f)

        print(f"Processed {name}")

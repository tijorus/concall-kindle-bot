import requests
from bs4 import BeautifulSoup
import pdfplumber
from ebooklib import epub
import json
import os
import smtplib
from email.message import EmailMessage
import re
from datetime import datetime

# ---------------- SETTINGS ---------------- #

BSE_BASE = "https://www.bseindia.com/corporates/ann.aspx?code="
HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = 20

EMAIL = os.environ["EMAIL_ADDRESS"]
PASSWORD = os.environ["EMAIL_PASSWORD"]
KINDLE = os.environ["KINDLE_EMAIL"]

# ---------------- LOAD FILES ---------------- #

with open("watchlist.json") as f:
    watchlist = json.load(f)["companies"]

with open("processed.json") as f:
    processed = json.load(f)

# ---------------- FUNCTIONS ---------------- #

def get_latest_transcript_link(bse_code):
    try:
        api_url = (
            "https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w"
            "?strCat=-1"
            "&strPrevDate="
            "&strScripCode=" + bse_code +
            "&strSearch=P"
            "&strToDate="
            "&strType=C"
            "&subcategory=-1"
        )

        headers = {
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://www.bseindia.com/",
            "Origin": "https://www.bseindia.com"
        }

        r = requests.get(api_url, headers=headers, timeout=TIMEOUT)

        if r.status_code != 200:
            print("API status error:", r.status_code)
            return None, None, None

        data = r.json()

        if "Table" not in data:
            print("No Table data in API response.")
            return None, None, None

        for item in data["Table"]:
            headline = item.get("HEADLINE", "").lower()
            print("Headline found:", headline)

            # Strict transcript match
            if (
                "earnings call transcript" in headline
                or "conference call transcript" in headline
                or ("earnings call" in headline and "transcript" in headline)
            ):

                pdf_file = item.get("ATTACHMENTNAME", "")
                ann_date = item.get("NEWS_DT", "")

                if pdf_file:
                    full_url = (
                        "https://www.bseindia.com/xml-data/corpfiling/AttachHis/"
                        + pdf_file + ".pdf"
                    )

                    return full_url, headline, ann_date

        return None, None, None

    except Exception as e:
        print("API error:", e)
        return None, None, None

def download_pdf(url, filename):
    try:
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://www.bseindia.com/",
            "Accept": "application/pdf",
            "Origin": "https://www.bseindia.com"
        }

        r = requests.get(url, headers=headers, timeout=TIMEOUT)

        print("Status:", r.status_code)
        print("Content-Type:", r.headers.get("Content-Type"))

        if r.status_code == 200 and "application/pdf" in r.headers.get("Content-Type", ""):
            with open(filename, "wb") as f:
                f.write(r.content)
            return True
        else:
            print("Not a valid PDF response.")
            return False

    except Exception as e:
        print(f"Error downloading PDF: {e}")
        return False


def extract_text_from_pdf(filename):
    text = ""
    try:
        with pdfplumber.open(filename) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        return text
    except Exception as e:
        print(f"Error extracting PDF text: {e}")
        return ""


def clean_text(text):
    text = re.sub(r"Safe Harbor.*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def split_sections(text):
    qa_patterns = ["question-and-answer", "q&a", "questions and answers"]

    lower_text = text.lower()
    for pattern in qa_patterns:
        if pattern in lower_text:
            idx = lower_text.index(pattern)
            return text[:idx], text[idx:]

    return text, ""


def extract_numeric_highlights(text):
    lines = text.split("\n")
    highlights = []

    for line in lines:
        if any(x in line for x in ["₹", "%", "crore", "million", "bn", "lakh"]):
            highlights.append(line.strip())

    return highlights[:15]


def extract_quarter(title_text):
    if not title_text:
        return "UnknownQ"

    match = re.search(r"(Q[1-4]\s?FY\s?\d{2,4})", title_text, re.IGNORECASE)
    if match:
        return match.group(1).replace(" ", "")
    return "UnknownQ"


def create_epub(company, management, qa, highlights, quarter, ann_date):
    safe_company = company.replace(" ", "_")

    if ann_date:
        safe_date = ann_date.replace(":", "-").replace("/", "-")
    else:
        safe_date = datetime.today().strftime("%Y-%m-%d")

    filename = f"{safe_company}{quarter}{safe_date}.epub"

    book = epub.EpubBook()
    book.set_identifier(safe_company)
    book.set_title(f"{company} {quarter}")
    book.set_language("en")
    book.add_author("Investor Relations")

    # Create chapter
    chapter = epub.EpubHtml(title="Transcript", file_name="chap_01.xhtml")

    highlight_html = "<br>".join(highlights)
    management_html = management.replace("\n", "<br>")
    qa_html = qa.replace("\n", "<br>")

    chapter.content = f"""
    <h1>{company} {quarter}</h1>
    <h3>Announcement Date: {safe_date}</h3>

    <h2>Numeric Highlights</h2>
    <p>{highlight_html}</p>

    <h2>Management Commentary</h2>
    <p>{management_html}</p>

    <h2>Q&A Session</h2>
    <p>{qa_html}</p>
    """

    book.add_item(chapter)

    # Add default navigation files (REQUIRED for Kindle)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())

    # Define Table of Contents
    book.toc = (chapter,)

    # Define spine
    book.spine = ["nav", chapter]

    epub.write_epub(filename, book)

    return filename

def send_to_kindle(file_path):
    try:
        msg = EmailMessage()
        msg["Subject"] = "Automated Concall Transcript"
        msg["From"] = EMAIL
        msg["To"] = KINDLE
        msg.set_content("Concall transcript attached.")

        with open(file_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="epub+zip",
                filename=file_path
            )

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL, PASSWORD)
            smtp.send_message(msg)

        print("Sent to Kindle successfully.")

    except Exception as e:
        print(f"Error sending email: {e}")


# ---------------- MAIN PROCESS ---------------- #

for company in watchlist:
    name = company["name"]
    code = company["bse_code"]

    print(f"Checking {name}...")

    try:
        link, title_text, ann_date = get_latest_transcript_link(code)

        if not link:
            print("No transcript found.")
            continue

        if link in processed:
            print("Already processed.")
            continue

        pdf_file = f"{name.replace(' ', '_')}.pdf"

        if not download_pdf(link, pdf_file):
            continue

        raw_text = extract_text_from_pdf(pdf_file)
        if not raw_text:
            continue

        cleaned = clean_text(raw_text)
        management, qa = split_sections(cleaned)
        highlights = extract_numeric_highlights(cleaned)
        quarter = extract_quarter(title_text)

        epub_file = create_epub(name, management, qa, highlights, quarter, ann_date)
        send_to_kindle(epub_file)

        processed.append(link)

        with open("processed.json", "w") as f:
            json.dump(processed, f, indent=2)

        print(f"Completed {name}")

    except Exception as e:
        print(f"Unexpected error for {name}: {e}")

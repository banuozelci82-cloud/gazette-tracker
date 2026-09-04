"""
Cayman Islands Gazette scraper — uses the Claude API (vision) to read
liquidation/winding-up notices instead of a traditional PDF text-extraction
library like pdfplumber, since some Gazette issues are scans without a
real text layer.

Exposes refresh_cayman(), called from app.py's /api/refresh the same way
refresh_uk() and refresh_france() are. Writes new rows into the same
`insolvencies` table, country code "KY".

Requires ANTHROPIC_API_KEY as an environment variable (set in Railway the
same way GMAIL_APP_PASSWORD and DATABASE_URL are set).
"""

import os
import re
import json
import base64
from datetime import datetime

import requests
import fitz  # PyMuPDF
import psycopg2
from bs4 import BeautifulSoup
from anthropic import Anthropic

GAZETTE_LIST_PAGE = "https://gov.ky/web/gazettes/gazettes"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

client = Anthropic()  # reads ANTHROPIC_API_KEY from the environment automatically

EXTRACTION_PROMPT = """You are reading one page of the Cayman Islands Government Gazette.

Find every notice on this page that relates to a company or partnership going
into liquidation (voluntary OR court-ordered), winding up, receivership, or
being appointed an administrator or examiner. This includes notices titled
things like "Voluntary Liquidation Notices," "Notice to Creditors from
Liquidator," and "Notices of Final Meeting of Shareholders" — these are all
in scope, not just court-ordered ones. Ignore unrelated notices (legislation,
appointments, trademarks, planning notices, partnership notices unrelated to
winding up, etc.).

For each relevant notice, extract:
- company_name
- notice_type (one of: Voluntary Liquidation, Court-Ordered Liquidation, Winding Up, Receivership, Administration, Examinership, Final Meeting of Shareholders)
- date (the date mentioned in the notice, format YYYY-MM-DD if possible, else as written)
- liquidator_or_contact (name of liquidator/firm handling it, if given)

Respond with ONLY a JSON array, nothing else. If there are no relevant notices
on this page, respond with an empty array: []

Example:
[
  {"company_name": "Example Fund Ltd", "notice_type": "Voluntary Liquidation", "date": "2026-08-20", "liquidator_or_contact": "Maples Liquidation Services Limited"}
]
"""


def get_db_connection():
    return psycopg2.connect(os.environ.get("DATABASE_URL"))


def find_latest_gazette_pdf():
    """Find the most recent regular Gazette PDF (not Legislation/Extraordinary
    Gazette, which don't carry commercial liquidation notices)."""
    r = requests.get(GAZETTE_LIST_PAGE, headers=HEADERS, timeout=20)
    soup = BeautifulSoup(r.text, "html.parser")

    for link in soup.find_all("a", href=True):
        href = link["href"]
        text = link.get_text(strip=True)
        if href.lower().endswith(".pdf") and "gazette" in text.lower() \
                and "legislation" not in text.lower() and "extraordinary" not in text.lower():
            return href if href.startswith("http") else "https://gov.ky" + href
    return None


def download_pdf(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.content


def pdf_to_page_images(pdf_bytes, dpi=150):
    """Render each PDF page to a PNG image (as base64) using PyMuPDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=matrix)
        images.append(base64.b64encode(pix.tobytes("png")).decode("utf-8"))
    doc.close()
    return images


def extract_notices_from_page(image_b64):
    """Send one page image to Claude and get back structured notices."""
    try:
        response = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=1024,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {"type": "base64", "media_type": "image/png", "data": image_b64},
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }],
        )
        text = "".join(b.text for b in response.content if b.type == "text").strip()
        text = re.sub(r"^```json\s*|\s*```$", "", text)
        return json.loads(text)
    except Exception as e:
        print("Cayman page extraction error: " + str(e))
        return []


def make_notice_id(notice):
    """Deterministic ID so re-running the same gazette issue doesn't create duplicates."""
    raw = f"{notice.get('company_name','')}-{notice.get('notice_type','')}-{notice.get('date','')}"
    slug = re.sub(r"[^A-Za-z0-9]+", "-", raw).strip("-").upper()
    return ("KY-" + slug)[:250]


def refresh_cayman():
    """Fetch the latest Cayman gazette, extract notices via Claude vision,
    insert new ones into the insolvencies table. Returns count of new rows,
    matching the pattern of refresh_uk() / refresh_france()."""
    try:
        pdf_url = find_latest_gazette_pdf()
        if not pdf_url:
            print("Cayman: could not find gazette PDF link")
            return 0

        pdf_bytes = download_pdf(pdf_url)
        page_images = pdf_to_page_images(pdf_bytes)

        conn = get_db_connection()
        cur = conn.cursor()
        new = 0

        for i, image_b64 in enumerate(page_images):
            notices = extract_notices_from_page(image_b64)
            for n in notices:
                company_name = (n.get("company_name") or "").strip()
                if not company_name:
                    continue
                notice_type = n.get("notice_type", "Liquidation")
                notice_date = n.get("date", "")
                notice_id = make_notice_id(n)
                url = pdf_url + "#page=" + str(i + 1)

                try:
                    cur.execute(
                        "INSERT INTO insolvencies VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (id) DO NOTHING",
                        (notice_id, company_name, notice_type, url,
                         datetime.now().strftime("%Y-%m-%d %H:%M"),
                         notice_date, "", "", "KY")
                    )
                    if cur.rowcount > 0:
                        new += 1
                except Exception as e:
                    print("KY insert error: " + str(e))
                    conn.rollback()

        conn.commit()
        conn.close()
        return new

    except Exception as e:
        print("Cayman refresh error: " + str(e))
        return 0

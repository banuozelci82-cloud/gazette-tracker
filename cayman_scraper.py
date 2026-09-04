"""
Cayman Islands Gazette scraper — uses the Claude API (vision) to read
liquidation/winding-up notices instead of a traditional PDF text-extraction
library like pdfplumber.

Why this approach:
- The Cayman Islands Gazette is published as a PDF, but some issues are
  scans/images rather than PDFs with a real text layer, so pdfplumber's
  extract_text() can silently return nothing.
- Instead of extracting text and writing regex to parse it, this script
  renders each PDF page as an image and asks Claude directly: "here's a
  page, give me the notices as JSON." That works whether or not the PDF
  has a text layer, and survives layout changes better than regex.

Cost: roughly 1-2 cents per page with Claude Sonnet 5. A fortnightly
gazette with a handful of relevant pages costs well under $5/year.

Setup needed (one-time):
    pip install requests beautifulsoup4 pymupdf anthropic

Environment variable needed:
    ANTHROPIC_API_KEY   (get one at console.anthropic.com, add it to
                          Railway the same way you added GMAIL_APP_PASSWORD)
"""

import os
import re
import json
import base64
import requests
import fitz  # PyMuPDF — renders PDF pages to images, no system install needed
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
  {"company_name": "Example Fund Ltd", "notice_type": "Liquidation", "date": "2026-08-20", "liquidator_or_contact": "Maples Liquidation Services Limited"}
]
"""


def find_latest_gazette_pdf():
    """Find the most recent regular Gazette PDF (not Legislation/Extraordinary Gazette,
    which don't carry commercial liquidation notices)."""
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
    zoom = dpi / 72  # PDF default is 72 dpi
    matrix = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=matrix)
        png_bytes = pix.tobytes("png")
        images.append(base64.b64encode(png_bytes).decode("utf-8"))
    doc.close()
    return images


def extract_notices_from_page(image_b64):
    """Send one page image to Claude and get back structured notices."""
    response = client.messages.create(
        model="claude-sonnet-5",
        max_tokens=1024,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": image_b64,
                        },
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
    )

    text = "".join(block.text for block in response.content if block.type == "text")
    text = text.strip()
    # Strip markdown code fences if Claude adds them despite instructions
    text = re.sub(r"^```json\s*|\s*```$", "", text.strip())

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        print(f"Could not parse JSON from response, skipping page. Raw: {text[:200]}")
        return []


def scrape_cayman_gazette():
    print("Finding latest Cayman Islands Gazette PDF...")
    pdf_url = find_latest_gazette_pdf()
    if not pdf_url:
        print("Could not find a Gazette PDF link on the page.")
        return []

    print(f"Found: {pdf_url}")
    pdf_bytes = download_pdf(pdf_url)

    print("Rendering pages to images...")
    page_images = pdf_to_page_images(pdf_bytes)
    print(f"{len(page_images)} pages to check.")

    all_notices = []
    for i, image_b64 in enumerate(page_images):
        print(f"Checking page {i + 1}/{len(page_images)}...")
        notices = extract_notices_from_page(image_b64)
        for notice in notices:
            notice["source_url"] = pdf_url
            notice["jurisdiction"] = "KY"  # Cayman Islands
            notice["page"] = i + 1
        all_notices.extend(notices)

    print(f"Found {len(all_notices)} relevant notices total.")
    return all_notices


# ---------------------------------------------------------------------------
# NOTE: this saves results to a local JSON file so you can inspect them first.
# To wire this into your existing database, replace save_to_json() below with
# the same sqlite3 insert logic your ireland_scraper.py already uses — just
# match the column names in your `notices` table (or whatever it's called).
# Ask your Claude Code session to do this integration step for you.
# ---------------------------------------------------------------------------

def save_to_json(notices, path="cayman_notices.json"):
    with open(path, "w") as f:
        json.dump(notices, f, indent=2)
    print(f"Saved to {path}")


if __name__ == "__main__":
    results = scrape_cayman_gazette()
    save_to_json(results)

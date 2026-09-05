"""
Cayman Islands Gazette notices — uses the Claude API (vision) to read
liquidation/winding-up notices instead of a traditional PDF text-extraction
library like pdfplumber, since some Gazette issues are scans without a
real text layer.

PRIMARY PATH: manual upload, same workflow as the Ireland CRO PDF.
gov.ky is a generic government CMS with thousands of unrelated documents
(credit card reports, travel expenses, bills, supplements...) mixed into
one feed with no clean way to programmatically isolate just "Gazette" and
"Extraordinary Gazette" issues. Rather than fight that, you browse to
https://gov.ky/web/gazettes yourself, click into Gazettes / Extraordinary
Gazettes, download the PDF(s) you want processed, and upload them through
the site — same motion as the Ireland upload, just pointed at a messier
source. app.py's /api/upload_cayman route calls process_cayman_upload()
below.

SECONDARY (best-effort, optional): refresh_cayman() tries to auto-find the
single latest regular Gazette issue. It can fail silently and return 0 if
gov.ky's structure shifts — that's fine, it's a bonus, not the main path.

Requires ANTHROPIC_API_KEY as an environment variable (set in Railway the
same way GMAIL_APP_PASSWORD and DATABASE_URL are set) — needed for BOTH
the manual upload and the auto-refresh, since both use Claude vision.
"""

import os
import re
import io
import json
import base64
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import fitz  # PyMuPDF
import psycopg2
from bs4 import BeautifulSoup
from anthropic import Anthropic

GAZETTE_HOME_PAGE = "https://gov.ky/web/gazettes"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

client = Anthropic()  # reads ANTHROPIC_API_KEY from the environment automatically

EXTRACTION_PROMPT = """You are reading one page of the Cayman Islands Government Gazette
(this may be a regular Gazette or an Extraordinary Gazette — both carry
commercial liquidation notices).

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
        print("Cayman page raw response: " + text[:300])
        text = re.sub(r"^```json\s*|\s*```$", "", text)
        return json.loads(text)
    except Exception as e:
        print("Cayman page extraction error: " + str(e))
        return []


def make_notice_id(notice):
    """Deterministic ID so re-uploading the same gazette issue doesn't create duplicates."""
    raw = f"{notice.get('company_name','')}-{notice.get('notice_type','')}-{notice.get('date','')}"
    slug = re.sub(r"[^A-Za-z0-9]+", "-", raw).strip("-").upper()
    return ("KY-" + slug)[:250]


def insert_notices_from_pdf(pdf_bytes, source_url):
    """Shared logic: render pages, extract via Claude vision (in parallel,
    a handful of pages at a time), insert new rows into the insolvencies
    table. Returns (new_count, total_found)."""
    page_images = pdf_to_page_images(pdf_bytes)

    # Process pages concurrently instead of one-by-one — for a 30-page issue,
    # doing them sequentially can take 2-3 minutes and gets killed by
    # gunicorn's request timeout. Running a handful at once keeps it well
    # under a minute in most cases.
    results_by_page = {}
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_page = {
            executor.submit(extract_notices_from_page, img): i
            for i, img in enumerate(page_images)
        }
        for future in as_completed(future_to_page):
            page_num = future_to_page[future]
            try:
                results_by_page[page_num] = future.result()
            except Exception as e:
                print(f"Cayman page {page_num + 1} failed: {e}")
                results_by_page[page_num] = []

    conn = get_db_connection()
    cur = conn.cursor()
    new = 0
    total_found = 0

    for i in range(len(page_images)):
        notices = results_by_page.get(i, [])
        for n in notices:
            company_name = (n.get("company_name") or "").strip()
            if not company_name:
                continue
            total_found += 1
            notice_type = n.get("notice_type", "Liquidation")
            notice_date = n.get("date", "")
            notice_id = make_notice_id(n)
            url = source_url + ("#page=" + str(i + 1) if source_url.startswith("http") else "")

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
    return new, total_found


def process_cayman_upload(pdf_bytes, filename="Manual upload"):
    """Called from app.py's /api/upload_cayman route. Returns (new_count, total_found)."""
    source_label = "Uploaded: " + filename
    return insert_notices_from_pdf(pdf_bytes, source_label)


# ---------------------------------------------------------------------------
# Best-effort automatic path (optional bonus, not the main workflow — see
# module docstring). Tries to find the single latest regular Gazette issue.
# Safe to leave wired into /api/refresh: it catches its own errors and just
# returns 0 if gov.ky's structure has shifted, same as the other refresh_*
# functions in app.py.
# ---------------------------------------------------------------------------

def find_latest_gazette_pdf():
    r = requests.get(GAZETTE_HOME_PAGE, headers=HEADERS, timeout=20)
    soup = BeautifulSoup(r.text, "html.parser")

    detail_url = None
    pattern = re.compile(r"^\d{4}[\s-]+Gazette[\s-]+\d+$", re.IGNORECASE)
    for link in soup.find_all("a", href=True):
        text = link.get_text(strip=True)
        if pattern.match(text):
            href = link["href"]
            detail_url = href if href.startswith("http") else "https://gov.ky" + href
            break

    if not detail_url:
        return None

    r2 = requests.get(detail_url, headers=HEADERS, timeout=20)
    soup2 = BeautifulSoup(r2.text, "html.parser")
    for link in soup2.find_all("a", href=True):
        href = link["href"]
        if ".pdf" in href.lower():
            return href if href.startswith("http") else "https://gov.ky" + href
    return None


def refresh_cayman():
    """Best-effort auto-refresh. Returns 0 (not an error) if it can't find
    anything — the manual upload path is the reliable one."""
    try:
        pdf_url = find_latest_gazette_pdf()
        if not pdf_url:
            print("Cayman auto-refresh: no issue found (use manual upload instead)")
            return 0
        r = requests.get(pdf_url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        new, _ = insert_notices_from_pdf(r.content, pdf_url)
        return new
    except Exception as e:
        print("Cayman refresh error: " + str(e))
        return 0

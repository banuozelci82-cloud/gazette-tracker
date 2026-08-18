import psycopg2
import requests
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

BASE_URL = "https://www.thegazette.co.uk"
SEC_HEADERS = {"User-Agent": "GazetteTracker banuozelci82@gmail.com"}
ALERT_CODES = {
    "2406": "Compulsory Liquidation",
    "2410": "Compulsory Liquidation",
    "2441": "Administration",
    "2442": "Administration",
    "2443": "Administration",
}
HEADERS = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}

def get_db():
    conn = psycopg2.connect(os.environ.get("DATABASE_URL"))
    conn.autocommit = False
    return conn

def fetch_new_uk():
    try:
        cutoff = datetime.now() - timedelta(days=1)
        new_notices = []
        page = 1
        stop = False
        while not stop and page <= 10:
            r = requests.get(
                BASE_URL + "/all-notices/notice",
                params={"category-code": "400", "results-page-size": "50", "results-page": str(page)},
                headers=HEADERS,
                timeout=10
            )
            entries = r.json().get("entry", [])
            if not entries:
                break
            for n in entries:
                nd_str = n.get("f:publish-date", "") or n.get("updated", "")
                nd = None
                if nd_str:
                    try:
                        nd = datetime.fromisoformat(nd_str[:10])
                    except:
                        pass
                if nd and nd < cutoff:
                    stop = True
                    break
                if n.get("f:notice-code") in ALERT_CODES:
                    company_name = n.get("title", "N/A").replace("&apos;", "'").replace("&amp;", "&")
                    nid = n.get("id", "").split("/")[-1]
                    new_notices.append({
                        "company": company_name,
                        "type": ALERT_CODES[n.get("f:notice-code")],
                        "date": nd_str[:10] if nd_str else "",
                        "url": BASE_URL + "/notice/" + nid,
                        "country": "UK"
                    })
            page += 1
        return new_notices
    except Exception as e:
        print("UK fetch error: " + str(e))
        return []

def fetch_new_us():
    try:
        cutoff = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        today = datetime.now().strftime("%Y-%m-%d")
        url = "https://efts.sec.gov/LATEST/search-index?q=%221.03%22&forms=8-K&dateRange=custom&startdt=" + cutoff + "&enddt=" + today
        r = requests.get(url, headers=SEC_HEADERS, timeout=15)
        hits = r.json().get("hits", {}).get("hits", [])
        new_notices = []
        for h in hits:
            src = h.get("_source", {})
            items = src.get("items", [])
            if "1.03" not in items:
                continue
            names = src.get("display_names", [])
            if not names:
                continue
            company_name = names[0].split("(")[0].strip()
            file_date = src.get("file_date", "")
            cik = src.get("ciks", [""])[0]
            filing_url = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=" + cik + "&type=8-K&owner=include&count=10"
            new_notices.append({
                "company": company_name,
                "type": "Chapter 11",
                "date": file_date,
                "url": filing_url,
                "country": "US"
            })
        return new_notices
    except Exception as e:
        print("US fetch error: " + str(e))
        return []

def send_email(uk_notices, us_notices):
    gmail = os.environ.get("GMAIL_ADDRESS")
    password = os.environ.get("GMAIL_APP_PASSWORD")
    today = datetime.now().strftime("%A %d %B %Y")
    total = len(uk_notices) + len(us_notices)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Insolvency Alert - " + str(total) + " new notice(s) - " + today
    msg["From"] = "Gazette Insolvency Tracker <" + gmail + ">"
    msg["To"] = gmail

    def make_rows(notices):
        rows = ""
        for n in notices:
            flag = "🇬🇧" if n["country"] == "UK" else "🇺🇸"
            rows += "<tr>"
            rows += "<td style='padding:10px; border-bottom:1px solid #eee;'>" + flag + " " + n["company"] + "</td>"
            rows += "<td style='padding:10px; border-bottom:1px solid #eee;'>" + n["type"] + "</td>"
            rows += "<td style='padding:10px; border-bottom:1px solid #eee;'>" + n["date"] + "</td>"
            rows += "<td style='padding:10px; border-bottom:1px solid #eee;'><a href='" + n["url"] + "'>View</a></td>"
            rows += "</tr>"
        return rows

    if total == 0:
        body_html = "<p>No new Administration, Liquidation or Chapter 11 notices today (" + today + ").</p>"
        body_text = "No new notices today."
    else:
        all_rows = make_rows(uk_notices) + make_rows(us_notices)
        body_html = """
        <html><body>
        <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
            <div style="background:#1a1a2e;color:white;padding:20px;">
                <h2 style="margin:0;">Insolvency Alert</h2>
                <p style="margin:5px 0 0 0;opacity:0.7;">""" + today + """</p>
            </div>
            <div style="padding:20px;">
                <p><strong>""" + str(total) + """ new notice(s) today:</strong></p>
                <table style="width:100%;border-collapse:collapse;margin-top:15px;">
                    <thead>
                        <tr style="background:#f5f5f5;">
                            <th style="padding:10px;text-align:left;">Company</th>
                            <th style="padding:10px;text-align:left;">Type</th>
                            <th style="padding:10px;text-align:left;">Date</th>
                            <th style="padding:10px;text-align:left;">Filing</th>
                        </tr>
                    </thead>
                    <tbody>""" + all_rows + """</tbody>
                </table>
            </div>
            <div style="padding:20px;background:#f5f5f5;font-size:12px;color:#666;">
                Gazette Insolvency Tracker — by Banu Summers
            </div>
        </div>
        </body></html>"""
        body_text = str(total) + " new notices today. Check your email client for details."

    msg.attach(MIMEText(body_text, "plain"))
    msg.attach(MIMEText(body_html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail, password)
            server.sendmail(gmail, gmail, msg.as_string())
        print("Email sent with " + str(total) + " notices")
    except Exception as e:
        print("Email error: " + str(e))

if __name__ == "__main__":
    print("Fetching new notices...")
    uk = fetch_new_uk()
    us = fetch_new_us()
    print("UK: " + str(len(uk)) + " | US: " + str(len(us)))
    send_email(uk, us)
    print("Done!")

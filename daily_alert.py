import psycopg2
import requests
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

BASE_URL = "https://www.thegazette.co.uk"
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

def fetch_new_notices():
    try:
        cutoff = datetime.now() - timedelta(days=1)
        new_notices = []
        page = 1
        stop = False

        while not stop and page <= 20:
            r = requests.get(
                f"{BASE_URL}/all-notices/notice",
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
                    nid = n.get("id", "").split("/")[-1]
                    company_name = n.get("title", "N/A").replace("&apos;", "'").replace("&amp;", "&")
                    notice_type = ALERT_CODES[n.get("f:notice-code")]
                    url = f"{BASE_URL}/notice/{nid}"
                    new_notices.append({
                        "company": company_name,
                        "type": notice_type,
                        "date": nd_str[:10] if nd_str else "",
                        "url": url
                    })

            page += 1

        return new_notices
    except Exception as e:
        print(f"Error fetching notices: {e}")
        return []

def send_email(notices):
    gmail = os.environ.get("GMAIL_ADDRESS")
    password = os.environ.get("GMAIL_APP_PASSWORD")
    today = datetime.now().strftime("%A %d %B %Y")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Insolvency Alert — {len(notices)} new notice(s) — {today}"
    msg["From"] = f"Gazette Insolvency Tracker <{gmail}>"
    msg["To"] = gmail

    if not notices:
        body_text = f"No new Administration or Compulsory Liquidation notices today ({today})."
        body_html = f"<p>No new Administration or Compulsory Liquidation notices today ({today}).</p>"
    else:
        rows = ""
        for n in notices:
            rows += f"""
            <tr>
                <td style="padding:10px; border-bottom:1px solid #eee;">{n['company']}</td>
                <td style="padding:10px; border-bottom:1px solid #eee;">{n['type']}</td>
                <td style="padding:10px; border-bottom:1px solid #eee;">{n['date']}</td>
                <td style="padding:10px; border-bottom:1px solid #eee;"><a href="{n['url']}">View</a></td>
            </tr>"""

        body_html = f"""
        <html><body>
        <div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">
            <div style="background:#1a1a2e; color:white; padding:20px;">
                <h2 style="margin:0;">Gazette Insolvency Alert</h2>
                <p style="margin:5px 0 0 0; opacity:0.7;">{today}</p>
            </div>
            <div style="padding:20px;">
                <p><strong>{len(notices)} new Administration / Compulsory Liquidation notice(s) today:</strong></p>
                <table style="width:100%; border-collapse:collapse;">
                    <thead>
                        <tr style="background:#f5f5f5;">
                            <th style="padding:10px; text-align:left;">Company</th>
                            <th style="padding:10px; text-align:left;">Type</th>
                            <th style="padding:10px; text-align:left;">Date</th>
                            <th style="padding:10px; text-align:left;">Notice</th>
                        </tr>
                    </thead>
                    <tbody>{rows}</tbody>
                </table>
            </div>
            <div style="padding:20px; background:#f5f5f5; font-size:12px; color:#666;">
                Gazette Insolvency Tracker — by Banu Summers
            </div>
        </div>
        </body></html>"""

        body_text = "\n".join([f"{n['company']} | {n['type']} | {n['date']} | {n['url']}" for n in notices])

    msg.attach(MIMEText(body_text, "plain"))
    msg.attach(MIMEText(body_html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail, password)
            server.sendmail(gmail, gmail, msg.as_string())
        print(f"Email sent successfully with {len(notices)} notices")
    except Exception as e:
        print(f"Error sending email: {e}")

if __name__ == "__main__":
    print("Fetching new insolvency notices...")
    notices = fetch_new_notices()
    print(f"Found {len(notices)} Administration/Liquidation notices")
    send_email(notices)
    print("Done!")

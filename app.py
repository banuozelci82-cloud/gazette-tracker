from flask import Flask, render_template, jsonify, send_file
import requests, csv, io, os, openpyxl
from datetime import datetime, timedelta
from collections import Counter
import psycopg2

app = Flask(__name__)
BASE_URL = "https://www.thegazette.co.uk"
SEC_HEADERS = {"User-Agent": "GazetteTracker banuozelci82@gmail.com", "Accept": "application/json"}
CODES = {
    "2406": "Compulsory Liquidation",
    "2410": "Compulsory Liquidation",
    "2431": "Creditors Voluntary Liquidation",
    "2432": "Creditors Voluntary Liquidation",
    "2433": "Creditors Voluntary Liquidation",
    "2441": "Administration",
    "2442": "Administration",
    "2443": "Administration",
    "2446": "Receivership",
    "2450": "Receivership",
    "2452": "Liquidation",
    "2454": "Winding Up",
}
HEADERS = {"Accept": "application/json", "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

def clean_name(name):
    return name.replace("&apos;", "'").replace("&amp;", "&").replace("&quot;", '"').replace("&#39;", "'")

def get_db():
    conn = psycopg2.connect(os.environ.get("DATABASE_URL"))
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS insolvencies (
        id TEXT PRIMARY KEY,
        company_name TEXT,
        notice_code TEXT,
        url TEXT,
        date_fetched TEXT,
        notice_date TEXT,
        company_number TEXT,
        sector TEXT,
        country TEXT
    )""")
    for col in ["company_number", "sector", "notice_date", "country"]:
        try:
            cur.execute(f"ALTER TABLE insolvencies ADD COLUMN {col} TEXT")
            conn.commit()
        except:
            conn.rollback()
    conn.commit()
    return conn

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/notices")
def notices():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT company_name, notice_code, date_fetched, url, notice_date, sector, country FROM insolvencies ORDER BY notice_date DESC, date_fetched DESC")
    rows = cur.fetchall()
    conn.close()
    return jsonify([{"company": r[0], "type": CODES.get(r[1], r[1]), "date": r[4] or r[2], "url": r[3], "sector": r[5] or "", "country": r[6] or "UK"} for r in rows])

@app.route("/api/refresh")
def refresh():
    new_uk = refresh_uk()
    new_us = refresh_us()
    return jsonify({"status": "ok", "new_uk": new_uk, "new_us": new_us})

def refresh_uk():
    try:
        new = 0
        page = 1
        conn = get_db()
        cur = conn.cursor()
        while page <= 10:
            r = requests.get(f"{BASE_URL}/all-notices/notice",
                params={"category-code": "400", "results-page-size": "50", "results-page": str(page)},
                headers=HEADERS, timeout=10)
            entries = r.json().get("entry", [])
            if not entries:
                break
            for n in entries:
                nd_str = n.get("f:publish-date", "") or n.get("updated", "")
                if n.get("f:notice-code") in CODES:
                    nid = n.get("id", "").split("/")[-1]
                    company_name = clean_name(n.get("title", "N/A"))
                    try:
                        cur.execute("INSERT INTO insolvencies VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (id) DO NOTHING",
                            (nid, company_name, n.get("f:notice-code", ""),
                             f"{BASE_URL}/notice/{nid}",
                             datetime.now().strftime("%Y-%m-%d %H:%M"),
                             nd_str[:10] if nd_str else "",
                             "", "", "UK"))
                        if cur.rowcount > 0:
                            new += 1
                    except Exception as e:
                        print(f"Insert error: {e}")
                        conn.rollback()
            conn.commit()
            page += 1
        conn.close()
        return new
    except Exception as e:
        print(f"UK refresh error: {e}")
        return 0

def refresh_us():
    try:
        cutoff = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
        today = datetime.now().strftime("%Y-%m-%d")
        r = requests.get(
            f"https://efts.sec.gov/LATEST/search-index?q=%22Item+1.03%22+%22bankruptcy%22&forms=8-K&dateRange=custom&startdt={cutoff}&enddt={today}",
            headers=SEC_HEADERS, timeout=15)
        hits = r.json().get("hits", {}).get("hits", [])
        conn = get_db()
        cur = conn.cursor()
        new = 0
        for h in hits:
            src = h.get("_source", {})
            names

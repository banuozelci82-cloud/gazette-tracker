from flask import Flask, render_template, jsonify, send_file
import requests, csv, io, os, openpyxl
from datetime import datetime, timedelta
from collections import Counter
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)
BASE_URL = "https://www.thegazette.co.uk"
CH_BASE = "https://api.company-information.service.gov.uk"
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
    conn = psycopg2.connect(os.environ.get("DATABASE_URL"), sslmode="require")
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS insolvencies (
        id TEXT PRIMARY KEY,
        company_name TEXT,
        notice_code TEXT,
        url TEXT,
        date_fetched TEXT,
        notice_date TEXT,
        company_number TEXT,
        sector TEXT
    )""")
    conn.commit()
    return conn

def get_sector_from_sic(sic):
    if not sic:
        return ""
    code = int(sic[:2]) if sic[:2].isdigit() else 0
    if code <= 3: return "Agriculture"
    if code <= 9: return "Mining"
    if code <= 33: return "Manufacturing"
    if code <= 35: return "Energy"
    if code <= 39: return "Water & Waste"
    if code <= 43: return "Construction"
    if code <= 47: return "Retail & Wholesale"
    if code <= 49: return "Transport"
    if code <= 53: return "Postal & Courier"
    if code <= 56: return "Hospitality"
    if code <= 63: return "IT & Media"
    if code <= 66: return "Finance"
    if code <= 68: return "Real Estate"
    if code <= 75: return "Professional Services"
    if code <= 82: return "Admin & Support"
    if code <= 84: return "Public Administration"
    if code <= 85: return "Education"
    if code <= 88: return "Health & Social"
    if code <= 93: return "Arts & Entertainment"
    if code <= 96: return "Other Services"
    return "Other"

def get_company_info(company_name):
    try:
        api_key = os.environ.get("COMPANIES_HOUSE_API_KEY", "")
        r = requests.get(
            f"{CH_BASE}/search/companies",
            params={"q": company_name, "items_per_page": 1},
            auth=(api_key, ""),
            timeout=10
        )
        items = r.json().get("items", [])
        if items:
            company_number = items[0].get("company_number", "")
            detail = requests.get(
                f"{CH_BASE}/company/{company_number}",
                auth=(api_key, ""),
                timeout=10
            ).json()
            sic_codes = detail.get("sic_codes", [])
            sector = get_sector_from_sic(sic_codes[0] if sic_codes else "")
            return company_number, sector
    except:
        pass
    return "", ""

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/notices")
def notices():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT company_name, notice_code, date_fetched, url, notice_date, sector FROM insolvencies ORDER BY notice_date DESC, date_fetched DESC")
    rows = cur.fetchall()
    conn.close()
    return jsonify([{"company": r[0], "type": CODES.get(r[1], r[1]), "date": r[4] or r[2], "url": r[3], "sector": r[5] or ""} for r in rows])

@app.route("/api/refresh")
def refresh():
    try:
        cutoff = datetime.now() - timedelta(days=7)
        new = 0
        page = 1
        stop = False
        conn = get_db()
        cur = conn.cursor()

        while not stop and page <= 50:
            r = requests.get(f"{BASE_URL}/all-notices/notice", params={"category-code": "400", "results-page-size": "50", "results-page": str(page)}, headers=HEADERS, timeout=10)
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
                if n.get("f:notice-code") in CODES:
                    nid = n.get("id", "").split("/")[-1]
                    company_name = clean_name(n.get("title", "N/A"))
                    company_number, sector = get_company_info(company_name)
                    try:
                        cur.execute(
                            "INSERT INTO insolvencies VALUES (%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (id) DO NOTHING",
                            (nid, company_name, n.get("f:notice-code", ""),
                             f"{BASE_URL}/notice/{nid}",
                             datetime.now().strftime("%Y-%m-%d %H:%M"),
                             nd_str[:10] if nd_str else "",
                             company_number, sector)
                        )
                        if cur.rowcount > 0:
                            new += 1
                    except:
                        pass
            conn.commit()
            page += 1
        conn.close()
        return jsonify({"status": "ok", "new": new})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route("/api/chart")
def chart():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT notice_code FROM insolvencies")
    rows = cur.fetchall()
    conn.close()
    return jsonify(Counter(CODES.get(r[0], r[0]) for r in rows))

@app.route("/export/csv")
def export_csv():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT company_name, notice_code, date_fetched, url, notice_date, sector FROM insolvencies ORDER BY notice_date DESC, date_fetched DESC")
    rows = cur.fetchall()
    conn.close()
    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(["Company", "Sector", "Type", "Date", "Gazette URL"])
    for row in rows:
        w.writerow([row[0], row[5] or "", CODES.get(row[1], row[1]), row[4] or row[2], row[3]])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype="text/csv", as_attachment=True, download_name="insolvencies.csv")

@app.route("/export/excel")
def export_excel():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT company_name, notice_code, date_fetched, url, notice_date, sector FROM insolvencies ORDER BY notice_date DESC, date_fetched DESC")
    rows = cur.fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Insolvencies"
    ws.append(["Company", "Sector", "Type", "Date", "Gazette URL"])
    for row in rows:
        ws.append([row[0], row[5] or "", CODES.get(row[1], row[1]), row[4] or row[2], row[3]])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 25
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="insolvencies.xlsx")

if __name__ == "__main__":
    app.run(debug=True)

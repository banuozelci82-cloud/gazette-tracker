from flask import Flask, render_template, jsonify, send_file, request
import requests, csv, io, os, openpyxl, re, pdfplumber
from datetime import datetime, timedelta
from collections import Counter
import psycopg2

app = Flask(__name__)
BASE_URL = "https://www.thegazette.co.uk"
SEC_HEADERS = {"User-Agent": "GazetteTracker banuozelci82@gmail.com"}
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
IRELAND_CODES = {
    "E2": "Liquidation",
    "E12": "Winding Up Order",
    "E19": "Administration",
    "E20": "Provisional Liquidation",
    "E21": "Liquidation",
    "E22": "Winding Up Order",
    "E35": "Examinership",
    "E8": "Receivership",
    "F15": "Insolvency Notice",
    "G1": "Voluntary Winding Up",
    "G2": "Voluntary Winding Up",
    "G4": "Creditors Voluntary Liquidation",
    "G1L": "Voluntary Winding Up",
}
HEADERS = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}

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
    all_codes = {**CODES, **IRELAND_CODES}
    return jsonify([{
        "company": r[0],
        "type": all_codes.get(r[1], r[1]),
        "date": r[4] or r[2],
        "url": r[3],
        "sector": r[5] or "",
        "country": r[6] or "UK"
    } for r in rows])

@app.route("/api/refresh")
def refresh():
    new_uk = refresh_uk()
    new_us_public = refresh_us_public()
    new_us_nonpublic = refresh_us_nonpublic()
    total_us = new_us_public + new_us_nonpublic
    return jsonify({"status": "ok", "new_uk": new_uk, "new_us": total_us, "new_us_public": new_us_public, "new_us_nonpublic": new_us_nonpublic})

@app.route("/api/upload_ireland", methods=["POST"])
def upload_ireland():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"})
    f = request.files["file"]
    if not f.filename.endswith(".pdf"):
        return jsonify({"error": "Please upload a PDF file"})
    try:
        pdf_bytes = f.read()
        notices = parse_ireland_pdf(pdf_bytes)
        conn = get_db()
        cur = conn.cursor()
        new = 0
        for n in notices:
            try:
                cur.execute(
                    "INSERT INTO insolvencies VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (id) DO NOTHING",
                    (n["id"], n["company_name"], n["notice_code"],
                     "https://cro.ie/cro-gazette-publications/",
                     datetime.now().strftime("%Y-%m-%d %H:%M"),
                     n["notice_date"], n["company_number"], "", "IE")
                )
                if cur.rowcount > 0:
                    new += 1
            except Exception as e:
                print("IE insert error: " + str(e))
                conn.rollback()
        conn.commit()
        conn.close()
        return jsonify({"status": "ok", "new": new, "total_found": len(notices)})
    except Exception as e:
        return jsonify({"error": str(e)})

def parse_ireland_pdf(pdf_bytes):
    notices = []
    try:
        pdf = pdfplumber.open(io.BytesIO(pdf_bytes))
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
        lines = full_text.split("\n")
        for line in lines:
            line = line.strip()
            if not line:
                continue
            for code in IRELAND_CODES:
                pattern = r"^(\d+)\s+(.+?)\s+" + re.escape(code) + r"\s+(\d{2}/\d{2}/\d{4})$"
                m = re.match(pattern, line)
                if m:
                    company_number = m.group(1)
                    company_name = m.group(2).strip()
                    date_str = m.group(3)
                    try:
                        d = datetime.strptime(date_str, "%d/%m/%Y")
                        notice_date = d.strftime("%Y-%m-%d")
                    except:
                        notice_date = date_str
                    notice_id = "IE-" + company_number + "-" + code + "-" + notice_date.replace("-", "")
                    notices.append({
                        "id": notice_id,
                        "company_name": company_name,
                        "company_number": company_number,
                        "notice_code": code,
                        "notice_date": notice_date
                    })
                    break
    except Exception as e:
        print("PDF parse error: " + str(e))
    return notices

@app.route("/api/clear_us")
def clear_us():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM insolvencies WHERE country = 'US'")
    conn.commit()
    deleted = cur.rowcount
    conn.close()
    return jsonify({"deleted": deleted})

def refresh_uk():
    try:
        new = 0
        page = 1
        conn = get_db()
        cur = conn.cursor()
        while page <= 20:
            params = {"category-code": "400", "results-page-size": "50", "results-page": str(page)}
            r = requests.get(BASE_URL + "/all-notices/notice", params=params, headers=HEADERS, timeout=10)
            entries = r.json().get("entry", [])
            if not entries:
                break
            for n in entries:
                nd_str = n.get("f:publish-date", "") or n.get("updated", "")
                if nd_str:
                    nd_date = nd_str[:10]
                    cutoff = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
                    if nd_date < cutoff:
                        break
                if n.get("f:notice-code") in CODES:
                    nid = n.get("id", "").split("/")[-1]
                    company_name = clean_name(n.get("title", "N/A"))
                    try:
                        cur.execute(
                            "INSERT INTO insolvencies VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (id) DO UPDATE SET notice_date = EXCLUDED.notice_date, date_fetched = EXCLUDED.date_fetched",
                            (nid, company_name, n.get("f:notice-code", ""),
                             BASE_URL + "/notice/" + nid,
                             datetime.now().strftime("%Y-%m-%d %H:%M"),
                             nd_str[:10] if nd_str else "",
                             "", "", "UK")
                        )
                        if cur.rowcount > 0:
                            new += 1
                    except Exception as e:
                        print("UK insert error: " + str(e))
                        conn.rollback()
            conn.commit()
            page += 1
        conn.close()
        return new
    except Exception as e:
        print("UK refresh error: " + str(e))
        return 0

def refresh_us_public():
    try:
        cutoff = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
        today = datetime.now().strftime("%Y-%m-%d")
        url = "https://efts.sec.gov/LATEST/search-index?q=%221.03%22&forms=8-K&dateRange=custom&startdt=" + cutoff + "&enddt=" + today
        r = requests.get(url, headers=SEC_HEADERS, timeout=15)
        hits = r.json().get("hits", {}).get("hits", [])
        conn = get_db()
        cur = conn.cursor()
        new = 0
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
            adsh = src.get("adsh", "").replace("-", "")
            notice_id = "US-SEC-" + adsh
            cik = src.get("ciks", [""])[0]
            filing_url = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=" + cik + "&type=8-K&owner=include&count=10"
            try:
                cur.execute(
                    "INSERT INTO insolvencies VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (id) DO NOTHING",
                    (notice_id, company_name, "Chapter 11",
                     filing_url,
                     datetime.now().strftime("%Y-%m-%d %H:%M"),
                     file_date, "", "", "US")
                )
                if cur.rowcount > 0:
                    new += 1
            except Exception as e:
                print("US public insert error: " + str(e))
                conn.rollback()
        conn.commit()
        conn.close()
        return new
    except Exception as e:
        print("US public refresh error: " + str(e))
        return 0

def refresh_us_nonpublic():
    try:
        api_key = os.environ.get("COURTLISTENER_API_KEY", "")
        if not api_key:
            return 0
        cutoff = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
        cl_headers = {"Authorization": "Token " + api_key}
        url = "https://www.courtlistener.com/api/rest/v4/search/?type=d&q=LLC+OR+Inc+OR+Corp+OR+Limited+chapter+11&order_by=dateFiled+desc&filed_after=" + cutoff + "&page_size=50"
        r = requests.get(url, headers=cl_headers, timeout=15)
        results = r.json().get("results", [])
        conn = get_db()
        cur = conn.cursor()
        new = 0
        for item in results:
            chapter = item.get("chapter")
            if str(chapter) not in ["11", "7"]:
                continue
            case_name = item.get("caseName", "")
            if not case_name:
                continue
            words = case_name.split()
            if any(w in ["v.", "vs.", "v"] for w in words):
                continue
            date_filed = item.get("dateFiled", "")
            # Skip if company already exists from SEC feed
            cur.execute(
                "SELECT id FROM insolvencies WHERE company_name ILIKE %s AND country = 'US'",
                (case_name,)
            )
            if cur.fetchone():
                continue
            docket_id = str(item.get("docket_id", ""))
            notice_id = "US-CL-" + docket_id
            filing_url = "https://www.courtlistener.com" + item.get("docket_absolute_url", "")
            notice_type = "Chapter " + str(chapter)
            try:
                cur.execute(
                    "INSERT INTO insolvencies VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (id) DO NOTHING",
                    (notice_id, case_name, notice_type,
                     filing_url,
                     datetime.now().strftime("%Y-%m-%d %H:%M"),
                     date_filed, "", "", "US")
                )
                if cur.rowcount > 0:
                    new += 1
            except Exception as e:
                print("US nonpublic insert error: " + str(e))
                conn.rollback()
        conn.commit()
        conn.close()
        return new
    except Exception as e:
        print("US nonpublic refresh error: " + str(e))
        return 0

@app.route("/api/chart")
def chart():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT notice_code FROM insolvencies")
    rows = cur.fetchall()
    conn.close()
    all_codes = {**CODES, **IRELAND_CODES}
    return jsonify(Counter(all_codes.get(r[0], r[0]) for r in rows))

@app.route("/export/csv")
def export_csv():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT company_name, notice_code, date_fetched, url, notice_date, sector, country FROM insolvencies ORDER BY notice_date DESC, date_fetched DESC")
    rows = cur.fetchall()
    conn.close()
    all_codes = {**CODES, **IRELAND_CODES}
    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(["Company", "Country", "Sector", "Type", "Date", "URL"])
    for row in rows:
        w.writerow([row[0], row[6] or "UK", row[5] or "", all_codes.get(row[1], row[1]), row[4] or row[2], row[3]])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype="text/csv", as_attachment=True, download_name="insolvencies.csv")

@app.route("/export/excel")
def export_excel():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT company_name, notice_code, date_fetched, url, notice_date, sector, country FROM insolvencies ORDER BY notice_date DESC, date_fetched DESC")
    rows = cur.fetchall()
    conn.close()
    all_codes = {**CODES, **IRELAND_CODES}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Insolvencies"
    ws.append(["Company", "Country", "Sector", "Type", "Date", "URL"])
    for row in rows:
        ws.append([row[0], row[6] or "UK", row[5] or "", all_codes.get(row[1], row[1]), row[4] or row[2], row[3]])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 25
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="insolvencies.xlsx")

@app.route("/api/debug_us")
def debug_us():
    api_key = os.environ.get("COURTLISTENER_API_KEY", "")
    cutoff = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    cl_headers = {"Authorization": "Token " + api_key}
    url = "https://www.courtlistener.com/api/rest/v4/search/?type=d&q=LLC+OR+Inc+OR+Corp+OR+Limited&order_by=dateFiled+desc&filed_after=" + cutoff + "&page_size=10&court=deb+OR+nysbk+OR+casbke+OR+ilnb+OR+txsb+OR+mdb+OR+njb+OR+ganb+OR+flsb+OR+ohsb"
    r = requests.get(url, headers=cl_headers, timeout=15)
    results = r.json().get("results", [])
    return jsonify([{"name": item.get("caseName"), "chapter": item.get("chapter"), "date": item.get("dateFiled"), "court": item.get("court_id")} for item in results])
    
if __name__ == "__main__":
        app.run(debug=True)

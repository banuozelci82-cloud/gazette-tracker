from flask import Flask, render_template, jsonify, send_file
import sqlite3, requests, csv, io, os, openpyxl
from datetime import datetime, timedelta
from collections import Counter
from bs4 import BeautifulSoup
import anthropic

app = Flask(__name__)
BASE_URL = "https://www.thegazette.co.uk"
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
DB = "gazette.db"

def clean_name(name):
    return name.replace("&apos;", "'").replace("&amp;", "&").replace("&quot;", '"').replace("&#39;", "'")

def get_db():
    conn = sqlite3.connect(DB)
    conn.execute("""CREATE TABLE IF NOT EXISTS insolvencies (
        id TEXT PRIMARY KEY, company_name TEXT, notice_code TEXT,
        url TEXT, date_fetched TEXT, notice_date TEXT, company_number TEXT)""")
    for col in ["company_number", "notice_date"]:
        try:
            conn.execute(f"ALTER TABLE insolvencies ADD COLUMN {col} TEXT")
        except:
            pass
    conn.commit()
    return conn

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/notices")
def notices():
    conn = get_db()
    rows = conn.execute("SELECT company_name, notice_code, date_fetched, url, notice_date, company_number FROM insolvencies ORDER BY date_fetched DESC").fetchall()
    conn.close()
    return jsonify([{"company": r[0], "type": CODES.get(r[1], r[1]), "date": r[4] or r[2], "url": r[3], "company_number": r[5] or ""} for r in rows])

@app.route("/api/refresh")
def refresh():
    try:
        cutoff = datetime.now() - timedelta(days=7)
        new = 0
        page = 1
        stop = False
        conn = get_db()
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
                    try:
                        conn.execute("INSERT INTO insolvencies VALUES (?,?,?,?,?,?,?)",
                            (nid, company_name, n.get("f:notice-code", ""),
                             f"{BASE_URL}/notice/{nid}", datetime.now().strftime("%Y-%m-%d %H:%M"),
                             nd_str[:10] if nd_str else "", ""))
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
    rows = conn.execute("SELECT notice_code FROM insolvencies").fetchall()
    conn.close()
    return jsonify(Counter(CODES.get(r[0], r[0]) for r in rows))

@app.route("/api/briefing")
def briefing():
    try:
        conn = get_db()
        rows = conn.execute("SELECT company_name, notice_code, date_fetched FROM insolvencies ORDER BY date_fetched DESC").fetchall()
        conn.close()
        total = len(rows)
        type_counts = Counter(CODES.get(r[1], r[1]) for r in rows)
        recent_names = [r[0] for r in rows[:10]]
        today = datetime.now().strftime("%A %d %B %Y")
        summary = f"Today is {today}.\nTotal: {total}\nTypes: {dict(type_counts)}\nRecent: {', '.join(recent_names)}"
        client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        msg = client.messages.create(model="claude-sonnet-4-20250514", max_tokens=300,
            messages=[{"role": "user", "content": f"You are a credit analyst assistant. Write a concise 3-4 sentence morning briefing based on this insolvency data:\n{summary}"}])
        return jsonify({"briefing": msg.content[0].text})
    except Exception as e:
        return jsonify({"briefing": f"Briefing unavailable: {str(e)}"})

@app.route("/export/csv")
def export_csv():
    conn = get_db()
    rows = conn.execute("SELECT company_name, notice_code, date_fetched, url, company_number FROM insolvencies ORDER BY date_fetched DESC").fetchall()
    conn.close()
    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(["Company", "Company Number", "Type", "Date Fetched", "Gazette URL", "Companies House URL"])
    for row in rows:
        ch = f"https://find-and-update.company-information.service.gov.uk/company/{row[4]}" if row[4] else ""
        w.writerow([row[0], row[4] or "", CODES.get(row[1], row[1]), row[2], row[3], ch])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype="text/csv", as_attachment=True, download_name="insolvencies.csv")

@app.route("/export/excel")
def export_excel():
    conn = get_db()
    rows = conn.execute("SELECT company_name, notice_code, date_fetched, url, company_number FROM insolvencies ORDER BY date_fetched DESC").fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Insolvencies"
    ws.append(["Company", "Company Number", "Type", "Date Fetched", "Gazette URL", "Companies House URL"])
    for row in rows:
        ch = f"https://find-and-update.company-information.service.gov.uk/company/{row[4]}" if row[4] else ""
        ws.append([row[0], row[4] or "", CODES.get(row[1], row[1]), row[2], row[3], ch])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 25
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="insolvencies.xlsx")

if __name__ == "__main__":
    app.run(debug=True)

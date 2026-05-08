from flask import Flask, render_template, jsonify, send_file
import sqlite3, requests, csv, io, openpyxl
from datetime import datetime
from collections import Counter

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
HEADERS = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}
DB = "gazette.db"

def get_db():
    conn = sqlite3.connect(DB)
    conn.execute("CREATE TABLE IF NOT EXISTS insolvencies (id TEXT PRIMARY KEY, company_name TEXT, notice_code TEXT, url TEXT, date_fetched TEXT)")
    conn.commit()
    return conn

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/notices")
def notices():
    conn = get_db()
    rows = conn.execute("SELECT company_name, notice_code, date_fetched, url FROM insolvencies ORDER BY date_fetched DESC").fetchall()
    conn.close()
    return jsonify([{"company": r[0], "type": CODES.get(r[1], r[1]), "date": r[2], "url": r[3]} for r in rows])

@app.route("/api/refresh")
def refresh():
    try:
        r = requests.get(f"{BASE_URL}/all-notices/notice", params={"category-code": "400", "results-page-size": "50"}, headers=HEADERS, timeout=10)
        entries = r.json().get("entry", [])
        corporate = [n for n in entries if n.get("f:notice-code") in CODES]
        conn = get_db()
        new = 0
        for n in corporate:
            nid = n.get("id", "").split("/")[-1]
            try:
                conn.execute("INSERT INTO insolvencies VALUES (?,?,?,?,?)", (nid, n.get("title", "N/A"), n.get("f:notice-code", ""), f"https://www.thegazette.co.uk/notice/{nid}", datetime.now().strftime("%Y-%m-%d %H:%M")))
                new += 1
            except:
                pass
        conn.commit()
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

@app.route("/export/csv")
def export_csv():
    conn = get_db()
    rows = conn.execute("SELECT company_name, notice_code, date_fetched, url FROM insolvencies ORDER BY date_fetched DESC").fetchall()
    conn.close()
    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(["Company", "Type", "Date Fetched", "URL"])
    for row in rows:
        w.writerow([row[0], CODES.get(row[1], row[1]), row[2], row[3]])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype="text/csv", as_attachment=True, download_name="insolvencies.csv")

@app.route("/export/excel")
def export_excel():
    conn = get_db()
    rows = conn.execute("SELECT company_name, notice_code, date_fetched, url FROM insolvencies ORDER BY date_fetched DESC").fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Insolvencies"
    ws.append(["Company", "Type", "Date Fetched", "URL"])
    for row in rows:
        ws.append([row[0], CODES.get(row[1], row[1]), row[2], row[3]])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 30
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="insolvencies.xlsx")

if __name__ == "__main__":
    app.run(debug=True)

"""
AdsTxt Checker — Python Backend
================================
Flask server that checks domains for google.com in their ads.txt files.

HOW TO RUN:
  1. pip install flask flask-cors requests openpyxl
  2. python app.py
  3. Open http://localhost:5000 in your browser
"""

import asyncio
import csv
import io
import re
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import openpyxl
from flask import Flask, jsonify, render_template, request, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ── Config ─────────────────────────────────────────────────────────────────────
TIMEOUT        = 8          # seconds per request
MAX_WORKERS    = 30         # concurrent threads
GOOGLE_PATTERN = re.compile(r'google\.com', re.IGNORECASE)

# In-memory job store  {job_id: {...}}
jobs = {}


# ── Helpers ────────────────────────────────────────────────────────────────────

def clean_domain(raw: str) -> str:
    """Strip protocol, paths, whitespace — keep bare domain."""
    raw = raw.strip()
    raw = re.sub(r'^https?://', '', raw, flags=re.IGNORECASE)
    raw = raw.split('/')[0]          # drop any path
    raw = raw.split('?')[0]          # drop query string
    return raw.lower()


def check_domain(domain: str) -> dict:
    """Fetch /ads.txt for a domain and look for google.com."""
    result = {
        "domain":      domain,
        "status":      "Failed",
        "google":      "No",
        "status_code": None,
        "error":       "",
    }

    for scheme in ("https", "http"):
        url = f"{scheme}://{domain}/ads.txt"
        try:
            resp = requests.get(
                url,
                timeout=TIMEOUT,
                headers={"User-Agent": "AdsTxtChecker/1.0"},
                allow_redirects=True,
            )
            result["status_code"] = resp.status_code

            if resp.status_code == 404:
                result["status"] = "Not Found"
                return result

            if resp.status_code == 200:
                result["status"] = "Success"
                if GOOGLE_PATTERN.search(resp.text):
                    result["google"] = "Yes"
                return result

            # other 4xx/5xx — try next scheme
            result["status"] = f"HTTP {resp.status_code}"

        except requests.exceptions.SSLError:
            continue                  # fall back to http
        except requests.exceptions.ConnectionError as e:
            result["error"] = "Connection error"
        except requests.exceptions.Timeout:
            result["error"] = "Timeout"
        except Exception as e:
            result["error"] = str(e)[:80]

    if result["status"] == "Failed" and not result["error"]:
        result["error"] = "Could not connect"

    return result


def run_job(job_id: str, domains: list):
    """Background worker — processes all domains and updates job state."""
    job = jobs[job_id]
    job["total"]   = len(domains)
    job["done"]    = 0
    job["results"] = []
    job["started"] = time.time()

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        future_map = {pool.submit(check_domain, d): d for d in domains}
        for future in as_completed(future_map):
            res = future.result()
            job["results"].append(res)
            job["done"] += 1

    job["finished"] = True
    job["elapsed"]  = round(time.time() - job["started"], 1)


def parse_domains_from_excel(file_bytes: bytes) -> list:
    """Extract all non-empty cell values from an xlsx/xls file."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    domains = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell and str(cell).strip():
                    domains.append(clean_domain(str(cell)))
    # deduplicate while preserving order
    seen = set()
    unique = []
    for d in domains:
        if d and d not in seen:
            seen.add(d)
            unique.append(d)
    return unique


def parse_domains_from_text(text: str) -> list:
    lines = text.strip().splitlines()
    seen, unique = set(), []
    for line in lines:
        d = clean_domain(line)
        if d and d not in seen:
            seen.add(d)
            unique.append(d)
    return unique


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/check", methods=["POST"])
def start_check():
    domains = []

    # --- Excel upload ---
    if "file" in request.files:
        f = request.files["file"]
        try:
            domains = parse_domains_from_excel(f.read())
        except Exception as e:
            return jsonify({"error": f"Could not read Excel file: {e}"}), 400

    # --- Plain text / paste ---
    elif request.is_json:
        data = request.get_json()
        raw  = data.get("domains", "")
        domains = parse_domains_from_text(raw)

    if not domains:
        return jsonify({"error": "No valid domains found."}), 400

    if len(domains) > 10_000:
        return jsonify({"error": "Max 10 000 domains per run."}), 400

    job_id = f"job_{int(time.time()*1000)}"
    jobs[job_id] = {"finished": False, "total": 0, "done": 0, "results": []}

    thread = threading.Thread(target=run_job, args=(job_id, domains), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id, "total": len(domains)})


@app.route("/api/status/<job_id>")
def job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "finished": job["finished"],
        "total":    job["total"],
        "done":     job["done"],
        "results":  job["results"],
        "elapsed":  job.get("elapsed"),
    })


@app.route("/api/export/<job_id>")
def export_csv(job_id):
    job = jobs.get(job_id)
    if not job or not job["finished"]:
        return jsonify({"error": "Job not ready"}), 404

    output = io.StringIO()
    writer = csv.DictWriter(
        output,
        fieldnames=["domain", "status", "google", "status_code", "error"],
        extrasaction="ignore",
    )
    writer.writeheader()
    writer.writerows(job["results"])

    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue().encode()),
        mimetype="text/csv",
        as_attachment=True,
        download_name="adstxt_results.csv",
    )


@app.route("/api/export_excel/<job_id>")
def export_excel(job_id):
    job = jobs.get(job_id)
    if not job or not job["finished"]:
        return jsonify({"error": "Job not ready"}), 404

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AdsTxt Results"

    headers = ["Domain", "Status", "Google.com Present", "HTTP Code", "Error"]
    ws.append(headers)

    # Style header row
    from openpyxl.styles import Font, PatternFill, Alignment
    header_fill = PatternFill("solid", fgColor="1A1A2E")
    for col, cell in enumerate(ws[1], 1):
        cell.font      = Font(bold=True, color="00E5A0")
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    green_fill = PatternFill("solid", fgColor="D4EDDA")
    red_fill   = PatternFill("solid", fgColor="F8D7DA")

    for row in job["results"]:
        ws.append([
            row["domain"],
            row["status"],
            row["google"],
            row["status_code"] or "",
            row["error"],
        ])
        last_row = ws.max_row
        if row["google"] == "Yes":
            ws.cell(last_row, 3).fill = green_fill
        else:
            ws.cell(last_row, 3).fill = red_fill

    # Auto-width columns
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="adstxt_results.xlsx",
    )


if __name__ == "__main__":
    print("\n🚀  AdsTxt Checker running at http://localhost:5000\n")
    app.run(debug=True, port=5000)

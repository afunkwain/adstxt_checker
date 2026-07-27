# 📡 AdsTxt Radar — AdSense-Only Domain Scanner

Check hundreds or thousands of domains and flag the ones whose `ads.txt` lists
`google.com` as its **only** demand partner — i.e. publishers monetizing with
AdSense alone, with no other ad networks alongside it.

A domain that lists `google.com` *plus* other partners (Rubicon, AppNexus,
Amazon, etc.) is **not** flagged.

---

## 🚀 Quick Start (3 steps)

### Step 1 — Install dependencies
Open a terminal in this folder and run:
```
pip install -r requirements.txt
```

### Step 2 — Start the server
```
python app.py
```

You should see:
```
🚀  AdsTxt Checker running at http://localhost:5000
```

### Step 3 — Open your browser
Go to: **http://localhost:5000**

---

## 📋 How to Use

1. **Upload an Excel file** (.xlsx / .xls / .csv) — put domains in any column, one per cell
2. **Or paste domains** directly — one per line
3. Hit **Run Scan**
4. Watch results come in live
5. **Export** to CSV or Excel when done

---

## ⚙️ Features

- ✅ Handles hundreds / thousands of domains (up to 10,000 per run)
- ✅ Concurrent requests (up to 50 parallel — configurable)
- ✅ Tries HTTPS first, falls back to HTTP
- ✅ Parses `ads.txt` properly — ignores comments, blank lines and variable
  declarations (`CONTACT=`, `OWNERDOMAIN=`, …) when deciding
- ✅ Case-insensitive exact match on the `google.com` domain field
- ✅ Live progress bar with ETA
- ✅ Filter & search results in the browser
- ✅ Export to CSV or styled Excel
- ✅ Graceful handling of timeouts, 404s, and errors

---

## 📁 File Structure

```
adstxt_checker/
├── app.py              ← Python backend (Flask)
├── requirements.txt    ← Python dependencies
├── README.md           ← This file
└── templates/
    └── index.html      ← Web frontend
```

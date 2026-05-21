#!/usr/bin/env python3
"""Run a full ads.txt scan from an Excel file and save results."""

import csv
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from app import MAX_WORKERS, check_domain, dedupe_domains, parse_domains_from_excel

BATCH_SIZE = 30
OUTPUT = Path(__file__).parent / "adstxt_results.csv"


def main():
    xlsx = Path(__file__).parent / "domains 2.xlsx"
    if not xlsx.exists():
        print(f"File not found: {xlsx}", file=sys.stderr)
        sys.exit(1)

    print(f"Parsing {xlsx.name}…")
    parsed = parse_domains_from_excel(xlsx.read_bytes())
    domains = parsed["domains"]
    total = len(domains)
    print(f"Found {total} unique domains ({parsed['submitted']} submitted)")

    results = []
    started = time.time()

    for i in range(0, total, BATCH_SIZE):
        batch = domains[i : i + BATCH_SIZE]
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = [pool.submit(check_domain, d) for d in batch]
            for fut in as_completed(futures):
                results.append(fut.result())

        done = len(results)
        pct = round(done / total * 100)
        elapsed = time.time() - started
        rate = done / elapsed if elapsed else 0
        eta = int((total - done) / rate) if rate else 0
        print(f"  {done}/{total} ({pct}%) — ~{eta}s remaining", flush=True)

    elapsed = round(time.time() - started, 1)
    google_yes = sum(1 for r in results if r["google"] == "Yes")
    success = sum(1 for r in results if r["status"] == "Success")

    with OUTPUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(
            f,
            fieldnames=["domain", "status", "google", "status_code", "error"],
        )
        w.writeheader()
        w.writerows(results)

    print(f"\nDone in {elapsed}s")
    print(f"  Google.com found: {google_yes}")
    print(f"  Ads.txt loaded:   {success}")
    print(f"  Saved to:         {OUTPUT}")


if __name__ == "__main__":
    main()

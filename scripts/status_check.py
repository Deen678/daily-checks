#!/usr/bin/env python3
"""Checks which websites in a spreadsheet are up or down.

Reads an Excel (.xlsx/.xlsm) or CSV file, works out which column holds the
website and which holds the site's status (Live, WIP, In design, In dev...),
checks every site whose status is in the "check" list of
reference/status-rules.json, and writes a colour-coded copy of the sheet with
the results added, plus a Summary sheet.

Usage:
  python3 scripts/status_check.py <sheet.xlsx|.csv> [--sheet NAME] [--site-col "Website"]
        [--status-col "Status"] [--rules reference/status-rules.json] [--out results.xlsx]
        [--only-columns]   # just show which columns it would use, then stop

Only plain page loads (GET) are sent. Nothing is submitted.
"""
import argparse, csv, json, os, re, sys, time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from urllib.parse import urlparse

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SITE_WORDS = ("website", "site", "domain", "url", "link", "web address", "address")
STATUS_WORDS = ("status", "stage", "state", "phase", "progress", "live?")
DOMAIN_RE = re.compile(r"^(https?://)?([a-z0-9-]+\.)+[a-z]{2,}(/.*)?$", re.I)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36 (site status check)"

# Page text that means the site isn't really working even though it answers.
BROKEN_SIGNS = [
    ("HTTP Server Test Page", "shows the web server's test page, not the site"),
    ("CentOS-WebPanel", "shows the hosting control panel's test page"),
    ("Account Suspended", "says the hosting account is suspended"),
    ("This Account has been suspended", "says the hosting account is suspended"),
    ("domain is for sale", "shows a 'domain for sale' page"),
    ("This domain has expired", "says the domain has expired"),
    ("domain name has expired", "says the domain has expired"),
    ("parked free", "shows a parked-domain page"),
    ("Error establishing a database connection", "shows a database error"),
    ("Index of /", "shows a bare file listing instead of the site"),
    ("An error has occurred!", "shows the site's error page"),
    ("Whoops, looks like something went wrong", "shows a server error"),
    ("Service Unavailable", "says the service is unavailable"),
    ("is under maintenance", "shows a maintenance page"),
    ("Site is down for maintenance", "shows a maintenance page"),
]


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


def load_rows(path, sheet=None):
    """Returns (sheet_name, header_row_index, headers, rows) where rows are lists of cell values."""
    if path.lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig") as f:
            data = [r for r in csv.reader(f)]
        name = os.path.basename(path)
    else:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        name = ws.title
        data = [[c for c in r] for r in ws.iter_rows(values_only=True)]
    # Header row: the first of the top 15 rows with 2+ text cells, preferring one that names a site/status column.
    best = 0
    for i, r in enumerate(data[:15]):
        texts = [norm(c) for c in r if isinstance(c, str) and c.strip()]
        if len(texts) >= 2 and any(any(w in t for w in SITE_WORDS + STATUS_WORDS) for t in texts):
            best = i; break
    headers = [str(c).strip() if c is not None else "" for c in data[best]] if data else []
    rows = [list(r) + [None] * (len(headers) - len(r)) for r in data[best + 1:] if any(c not in (None, "") for c in r)]
    return name, best, headers, rows


def pick_columns(headers, rows, rules, site_col=None, status_col=None):
    def find(name):
        for i, h in enumerate(headers):
            if norm(h) == norm(name): return i
        raise SystemExit(f'Column "{name}" not found. Columns are: {", ".join(h for h in headers if h)}')
    known = {norm(s) for s in rules["check"] + rules["skip"]}
    if site_col: si = find(site_col)
    else:
        scores = []
        for i, h in enumerate(headers):
            vals = [str(r[i]).strip() for r in rows if r[i] not in (None, "")]
            dom = sum(bool(DOMAIN_RE.match(v.replace(" ", ""))) for v in vals) / max(1, len(vals))
            head = 1 if any(w in norm(h) for w in SITE_WORDS) else 0
            scores.append((dom * 2 + head, i))
        si = max(scores)[1] if scores else 0
    if status_col: ti = find(status_col)
    else:
        scores = []
        for i, h in enumerate(headers):
            if i == si: continue
            vals = [norm(r[i]) for r in rows if r[i] not in (None, "")]
            hit = sum(v in known for v in vals) / max(1, len(vals))
            head = 1 if any(w in norm(h) for w in STATUS_WORDS) else 0
            scores.append((hit * 2 + head, i))
        ti = max(scores)[1] if scores else None
    return si, ti


def to_url(v):
    s = str(v or "").strip().replace(" ", "")
    if not s or not DOMAIN_RE.match(s): return None
    return s if s.startswith("http") else "https://" + s


def dns_problem(host):
    """Returns a reason if the domain doesn't resolve, else None. Uses DNS over https, which works through proxies."""
    try:
        for t in ("A", "AAAA"):
            j = requests.get("https://cloudflare-dns.com/dns-query", params={"name": host, "type": t}, headers={"Accept": "application/dns-json"}, timeout=15).json()
            if j.get("Status") == 3: return "Domain doesn't exist (no DNS record)"
            if any(a.get("type") in (1, 28, 5) for a in j.get("Answer", [])): return None
        return "Domain has no address (DNS isn't pointing anywhere)"
    except Exception:
        return None


def check_site(url):
    """One site. Retries once before calling it down. Returns a dict."""
    host = urlparse(url).hostname or ""
    last = None
    for attempt in (1, 2):
        t0 = time.time()
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=25, allow_redirects=True)
            secs = round(time.time() - t0, 1)
            final = r.url; fhost = (urlparse(final).hostname or "")
            body = r.text[:200000] if "text" in r.headers.get("Content-Type", "") or not r.headers.get("Content-Type") else ""
            base = dict(code=r.status_code, final=final, secs=secs)
            if r.status_code in (403, 429, 503) and ("cf-mitigated" in r.headers or "Just a moment" in body or "challenge-platform" in body):
                return dict(base, result="Couldn't check", details="Blocked by bot protection (Cloudflare). Open it in a browser to confirm.")
            if r.status_code >= 500:
                last = dict(base, result="Down", details=f"Server error {r.status_code}")
            elif r.status_code >= 400:
                last = dict(base, result="Problem", details=f"Page answers with error {r.status_code}")
            else:
                for sign, what in BROKEN_SIGNS:
                    if sign.lower() in body.lower():
                        return dict(base, result="Problem", details=f"Loads, but {what}")
                if fhost.replace("www.", "") != host.replace("www.", ""):
                    return dict(base, result="Problem", details=f"Redirects to a different site: {fhost}")
                if body and len(re.sub(r"<[^>]+>|\s+", "", body)) < 200:
                    return dict(base, result="Problem", details="Loads, but the page is almost empty")
                return dict(base, result="Up", details="" if secs < 8 else f"Slow to load ({secs}s)")
        except requests.exceptions.SSLError:
            try:
                r = requests.get("http://" + host, headers={"User-Agent": UA}, timeout=20, allow_redirects=False)
                last = dict(code=r.status_code, final="http://" + host, secs=None, result="Problem", details="Security certificate (https) error. The site answers on plain http only")
            except Exception:
                last = dict(code=None, final="", secs=None, result="Down", details="Security certificate (https) error, and no answer on http")
        except requests.exceptions.ConnectionError:
            why = dns_problem(host) or "The server didn't answer or refused the connection"
            last = dict(code=None, final="", secs=None, result="Down", details=why)
        except requests.exceptions.Timeout:
            last = dict(code=None, final="", secs=None, result="Down", details="No answer within 25 seconds")
        except Exception as e:
            last = dict(code=None, final="", secs=None, result="Down", details=f"Couldn't connect ({type(e).__name__})")
        if attempt == 1: time.sleep(4)
    return last


FILL = {"Up": "D9F2E3", "Down": "F8D7DA", "Problem": "FFF0C2", "Couldn't check": "E7E9F2", "Not checked": "F2F2F2", "Unknown status": "E8DDF5", "No website": "F2F2F2"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file"); ap.add_argument("--sheet"); ap.add_argument("--site-col"); ap.add_argument("--status-col")
    ap.add_argument("--rules", default=os.path.join(os.path.dirname(__file__), "..", "reference", "status-rules.json"))
    ap.add_argument("--out"); ap.add_argument("--only-columns", action="store_true")
    a = ap.parse_args()
    rules = json.load(open(a.rules))
    check_set = {norm(s) for s in rules["check"]}; skip_set = {norm(s) for s in rules["skip"]}
    sheet, hdr_i, headers, rows = load_rows(a.file, a.sheet)
    si, ti = pick_columns(headers, rows, rules, a.site_col, a.status_col)
    statuses = sorted({str(r[ti]).strip() for r in rows if ti is not None and r[ti] not in (None, "")})
    print(f'Sheet "{sheet}": {len(rows)} rows. Header row: {hdr_i + 1}.')
    print(f'Website column: "{headers[si]}" (column {get_column_letter(si + 1)}). Status column: "{headers[ti] if ti is not None else "(none found)"}"' + (f" (column {get_column_letter(ti + 1)})" if ti is not None else ""))
    print("Statuses found: " + ", ".join(f'{s} -> {"CHECK" if norm(s) in check_set else "skip" if norm(s) in skip_set else "UNKNOWN"}' for s in statuses))
    if a.only_columns: return

    jobs = []
    for idx, r in enumerate(rows):
        st = str(r[ti]).strip() if ti is not None and r[ti] not in (None, "") else ""
        url = to_url(r[si])
        if not url: rows[idx] = (r, dict(result="No website", details="No website address in this row")); continue
        if ti is not None and norm(st) in skip_set: rows[idx] = (r, dict(result="Not checked", details=f'Status "{st}" is not checked')); continue
        if ti is not None and norm(st) not in check_set: rows[idx] = (r, dict(result="Unknown status", details=f'Status "{st or "(blank)"}" isn\'t in the rules, so it wasn\'t checked')); continue
        jobs.append((idx, url))
    print(f"Checking {len(jobs)} sites...")
    with ThreadPoolExecutor(8) as ex:
        for (idx, url), res in zip(jobs, ex.map(lambda j: check_site(j[1]), jobs)):
            rows[idx] = (rows[idx], dict(res, checked_url=url))

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    out = a.out or re.sub(r"\.(xlsx|xlsm|csv)$", "", a.file, flags=re.I) + f"-status-{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    wb = Workbook(); ws = wb.active; ws.title = "Results"
    new_cols = ["Result", "Details", "HTTP code", "Final address", "Load time (s)", "Checked at"]
    ws.append(headers + new_cols)
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864"); c.alignment = Alignment(vertical="center", wrap_text=True)
    for r, res in rows:
        ws.append(list(r) + [res.get("result"), res.get("details", ""), res.get("code"), res.get("final", ""), res.get("secs"), stamp if res.get("result") not in ("Not checked", "No website", "Unknown status") else ""])
        fill = PatternFill("solid", fgColor=FILL.get(res.get("result"), "FFFFFF"))
        for c in ws[ws.max_row][len(headers):len(headers) + 2]: c.fill = fill
        ws[ws.max_row][len(headers)].font = Font(bold=True)
    for i, h in enumerate(headers + new_cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(60, max(12, len(str(h)) + 4, *[len(str(ws.cell(row=k, column=i).value or "")) + 2 for k in range(2, min(ws.max_row, 200) + 1)]))
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions

    # Summary first: anything that needs attention.
    sm = wb.create_sheet("Summary", 0)
    counts = {}
    for _, res in rows: counts[res["result"]] = counts.get(res["result"], 0) + 1
    sm.append([f"Site status check: {os.path.basename(a.file)} ({sheet})"]); sm["A1"].font = Font(bold=True, size=14)
    sm.append([f"Checked {stamp}. Statuses checked: {', '.join(rules['check'])}."]); sm.append([])
    for k in ["Up", "Down", "Problem", "Couldn't check", "Unknown status", "Not checked", "No website"]:
        if counts.get(k): sm.append([k, counts[k]]); sm[sm.max_row][0].fill = PatternFill("solid", fgColor=FILL[k])
    sm.append([]); sm.append(["Needs attention", "Status", "Result", "Details"])
    for c in sm[sm.max_row]: c.font = Font(bold=True)
    attention = [(r, res) for r, res in rows if res["result"] in ("Down", "Problem", "Couldn't check", "Unknown status")]
    order = {"Down": 0, "Problem": 1, "Couldn't check": 2, "Unknown status": 3}
    for r, res in sorted(attention, key=lambda x: order[x[1]["result"]]):
        sm.append([str(r[si]), str(r[ti]) if ti is not None else "", res["result"], res.get("details", "")])
        sm[sm.max_row][2].fill = PatternFill("solid", fgColor=FILL[res["result"]])
    if not attention: sm.append(["Nothing needs attention: every checked site is up."])
    for col, w in zip("ABCD", (38, 22, 16, 80)): sm.column_dimensions[col].width = w
    wb.save(out)

    print(f"\nSaved: {out}")
    print("Totals: " + ", ".join(f"{k} {v}" for k, v in counts.items()))
    for r, res in sorted(attention, key=lambda x: order[x[1]["result"]]):
        print(f'  {res["result"].upper():15} {str(r[si]):35} [{str(r[ti]) if ti is not None else ""}] {res.get("details", "")}')


if __name__ == "__main__":
    main()

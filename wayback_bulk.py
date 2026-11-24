#!/usr/bin/env python3
"""
Bulk version of wayback_upc_lookup.py, for thousands of UPCs.

Instead of searching the Wayback Machine once per UPC per site, it downloads
each site's whole list of pre-2016 captures once (a few big CDX requests),
matches every UPC against those lists on your computer, and only downloads
the archived pages that matched.

Needs Python 3.8+ and openpyxl (for reading/writing Excel):
    python3 -m pip install openpyxl truststore

Run:
    python3 wayback_bulk.py --input yogurt_upc_all.xlsx

Everything it downloads is saved in a work folder (wayback_work/ next to the
input file), so if it stops for any reason, just run the same command again
and it picks up where it left off.

Steps (all run by default, in order; pick some with --steps):
  upc-index   Download pre-2016 capture lists for sites that put the UPC in
              their URLs (upcitemdb, Open Food Facts, EWG Food Scores, ...).
  name-index  Download pre-2016 capture lists of retailer product pages
              whose URLs start with each brand name (walmart.com/ip/Dannon-...).
              One request per brand per retailer, so this is the slow step.
  match       Match the UPCs and product names against those lists.
  fetch       Download the matched archived pages and pull out nutrition text.
  report      Write wayback_results.xlsx next to the input file.
"""

import argparse
import csv
import gzip
import hashlib
import html
import http.client
import os
import random
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from html.parser import HTMLParser

try:
    # Makes Python trust the same certificates as macOS/Windows; fixes
    # CERTIFICATE_VERIFY_FAILED on networks that inspect HTTPS traffic.
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

WAYBACK = "https://web.archive.org"
USER_AGENT = "pre2016-nutrition-research/1.0 (personal research script)"
STEPS = ["upc-index", "name-index", "match", "fetch", "report"]

# Sites whose page URLs contain the UPC. Each is downloaded in full (pre-2016
# captures only). "filter" narrows big sites to URLs with a long digit run.
# Add or remove lines freely; a site with no pre-2016 captures costs one request.
UPC_SITES = [
    {"name": "upcitemdb",       "prefix": "upcitemdb.com/upc/"},
    {"name": "upcdatabase_com", "prefix": "upcdatabase.com/item/"},
    {"name": "upcdatabase_org", "prefix": "upcdatabase.org/code/"},
    {"name": "buycott",         "prefix": "buycott.com/upc/"},
    {"name": "off_world",       "prefix": "world.openfoodfacts.org/product/"},
    {"name": "off_us",          "prefix": "us.openfoodfacts.org/product/", "skip_in_trial": True},
    {"name": "off_fr",          "prefix": "fr.openfoodfacts.org/produit/", "skip_in_trial": True},
    {"name": "ewg_foodscores",  "prefix": "ewg.org/foodscores/products/"},
    {"name": "barcodelookup",   "prefix": "barcodelookup.com/", "filter": r"original:.*/[0-9]{8,14}.*",
     "upc_right_after_prefix": True},
    {"name": "searchupc",       "prefix": "searchupc.com/", "filter": r"original:.*[0-9]{10,14}.*"},
    {"name": "meijer",          "prefix": "meijer.com/product/", "filter": r"original:.*[0-9]{10,14}.*"},
    {"name": "walmart_search",  "prefix": "walmart.com/search/", "filter": r"original:.*[0-9]{10,14}.*"},
]

# Retailer URL prefixes whose product URLs start with the brand name.
# "{brand}" becomes a lowercase slug like "dannon" or "la-yogurt"; CDX
# ignores case, so "walmart.com/ip/dannon" matches "/ip/Dannon-Oikos-...".
NAME_SITES = [
    {"name": "walmart", "prefix": "walmart.com/ip/{brand}"},
    {"name": "target",  "prefix": "target.com/p/{brand}"},
    {"name": "amazon",  "prefix": "amazon.com/{brand}"},
]

# Words in product descriptions that say nothing about which product it is.
GENERIC_WORDS = {
    "the", "and", "with", "on", "of", "a", "an", "in", "for", "oz", "ounce",
    "ounces", "pack", "cup", "cups", "container", "yogurt", "yoghurt", "y",
    "spoonable", "refrigerated", "box", "tub", "bottle", "multi", "pk", "ct",
    "count", "x", "fl", "lb", "dairy", "snack", "drinkable", "drink", "style",
}

NUTRIENT_PATTERNS = {
    "ServSize": r"serving size[:\s]*([^\n|]{1,40}?)(?=\s+(?:servings|amount|calories)|\n|$)",
    "ServingsPerContainer": r"servings per container[:\s]*(?:about\s*)?(\d+(?:\.\d+)?)",
    "calories": r"(?<!from fat )calories[:\s]*(\d{1,4})\b",
    "totalfat": r"total fat[:\s]*(<?\d+(?:\.\d+)?)\s*g",
    "totalcarb": r"total carb(?:ohydrate)?s?[:\s]*(<?\d+(?:\.\d+)?)\s*g",
    "sugar": r"(?<!added )(?<!incl\. )sugars?[:\s]*(<?\d+(?:\.\d+)?)\s*g",
    "addedsugar": r"added sugars?[:\s]*(<?\d+(?:\.\d+)?)\s*g",
    "protein": r"protein[:\s]*(<?\d+(?:\.\d+)?)\s*g",
}

CANDIDATE_FIELDS = ["row", "upc", "match_type", "site", "name_score", "timestamp",
                    "original_url"]


def log(*a, end="\n"):
    print(*a, end=end, file=sys.stderr, flush=True)


# --------------------------------------------------------------------------
# HTTP: polite pacing; waits out rate limits and outages instead of quitting
# --------------------------------------------------------------------------

class Client:
    def __init__(self, delay, timeout=120, waits=(10, 30, 60, 120, 300, 600)):
        self.delay, self.timeout, self._last = delay, timeout, 0.0
        self.waits = list(waits)  # default adds up to about 18 minutes

    def get(self, url):
        waits = self.waits
        for attempt in range(len(waits) + 1):
            pause = self.delay - (time.time() - self._last)
            if pause > 0:
                time.sleep(pause)
            self._last = time.time()
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    return resp.status, resp.read()
            except urllib.error.HTTPError as e:
                if e.code not in (429, 500, 502, 503, 504):
                    return e.code, e.read() if e.fp else b""
                why = {429: "rate limited", 503: "archive busy or offline"}.get(e.code, f"HTTP {e.code}")
            # IncompleteRead etc.: the archive dropped the connection mid-response.
            except (urllib.error.URLError, http.client.HTTPException, OSError) as e:
                if "CERTIFICATE_VERIFY_FAILED" in str(e):
                    sys.exit("\nCertificate error. Run: python3 -m pip install truststore\n"
                             "then run this script again.")
                why = type(e).__name__
            if attempt < len(waits):
                log(f"    ({why}; waiting {waits[attempt]}s and trying again)")
                time.sleep(waits[attempt])
        raise RuntimeError(f"gave up after repeated failures ({why})")


def cdx_dump(client, prefix, cutoff, filters, out_path, page_size, max_pages=0):
    """Download every pre-cutoff capture under `prefix` into out_path (TSV).

    Uses CDX resume keys so big sites come down in pages. Progress is kept in
    out_path + ".resume", so an interrupted download continues where it was.
    """
    done_flag, resume_file = out_path + ".done", out_path + ".resume"
    if os.path.exists(done_flag):
        return sum(1 for _ in open(out_path, encoding="utf-8")) if os.path.exists(out_path) else 0
    resume = open(resume_file).read().strip() if os.path.exists(resume_file) else ""
    total = pages = 0
    while True:
        params = [("url", prefix), ("matchType", "prefix"), ("to", cutoff),
                  ("fl", "timestamp,original"), ("filter", "statuscode:200"),
                  ("collapse", "timestamp:4"),  # at most one capture per URL per year
                  ("limit", str(page_size)), ("showResumeKey", "true")]
        params += [("filter", f) for f in filters]
        if resume:
            params.append(("resumeKey", resume))
        status, body = client.get(f"{WAYBACK}/cdx/search/cdx?" + urllib.parse.urlencode(params))
        if status != 200:
            raise RuntimeError(f"CDX HTTP {status}: {body[:150]!r}")
        lines = body.decode("utf-8", "replace").rstrip("\n").split("\n")
        # Results, then (if more remain) a blank line and the resume key.
        resume = ""
        if len(lines) >= 2 and "" in lines:
            blank = len(lines) - 1 - lines[::-1].index("")
            tail = [l for l in lines[blank + 1:] if l.strip()]
            if tail and " " not in tail[0]:
                resume, lines = tail[0].strip(), lines[:blank]
        rows = [l.split(" ", 1) for l in lines if l.strip() and " " in l]
        with open(out_path, "a", encoding="utf-8") as f:
            for ts, orig in rows:
                f.write(f"{ts}\t{orig}\n")
        total += len(rows)
        pages += 1
        if not resume or (max_pages and pages >= max_pages):
            break
        with open(resume_file, "w") as f:
            f.write(resume)
        log(f"    ...{total} captures so far")
    open(done_flag, "w").close()
    if os.path.exists(resume_file):
        os.remove(resume_file)
    return total


def read_dump(path):
    """{original_url: latest pre-cutoff timestamp} from a TSV dump."""
    latest = {}
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            ts, _, orig = line.rstrip("\n").partition("\t")
            if orig and ts > latest.get(orig, ""):
                latest[orig] = ts
    return latest


# --------------------------------------------------------------------------
# UPC normalization
# --------------------------------------------------------------------------

def check_digit_ok(d):
    if len(d) < 8:
        return False
    s = sum(int(c) * (3 if i % 2 == 0 else 1) for i, c in enumerate(reversed(d[:-1])))
    return (10 - s % 10) % 10 == int(d[-1])


def with_check_digit(body):
    s = sum(int(c) * (3 if i % 2 == 0 else 1) for i, c in enumerate(reversed(body)))
    return body + str((10 - s % 10) % 10)


def upc_keys(digits, may_have_check):
    """Comparable forms of a code: leading zeros stripped, with and without
    a trailing check digit. Your UPCs (Nielsen style) have no check digit;
    URLs usually do. Comparing these forms makes them meet in the middle."""
    d = digits.lstrip("0")
    keys = {d} if len(d) >= 6 else set()
    if may_have_check and check_digit_ok(digits):
        k = digits[:-1].lstrip("0")
        if len(k) >= 6:
            keys.add(k)
    return keys


def full_upc(digits):
    """Best guess at the printed barcode (UPC-A or EAN-13) for display."""
    if len(digits) <= 11:
        return with_check_digit(digits.zfill(11))
    if len(digits) == 12:
        return with_check_digit(digits)  # 12-digit body -> EAN-13
    return digits


# --------------------------------------------------------------------------
# Input / names
# --------------------------------------------------------------------------

def load_rows(path, sheet, upc_col, name_col, brand_col):
    if path.lower().endswith((".xlsx", ".xlsm")):
        try:
            import openpyxl
        except ImportError:
            sys.exit("Reading Excel needs openpyxl. Run: python3 -m pip install openpyxl")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        header = [str(h).strip() if h is not None else "" for h in next(it)]
        raw = [dict(zip(header, r)) for r in it]
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            raw = list(csv.DictReader(f))
        header = list(raw[0].keys()) if raw else []
    for col in (upc_col, name_col, brand_col):
        if col not in header:
            sys.exit(f"Column '{col}' not found. Columns are: {', '.join(h for h in header if h)}")
    rows = []
    for r in raw:
        u = r.get(upc_col)
        if u is None or str(u).strip() == "":
            continue
        u = str(u).strip()
        if re.fullmatch(r"\d+\.0", u):  # number cells read as floats
            u = u[:-2]
        digits = re.sub(r"\D", "", u)
        if not digits:
            continue
        rows.append({"upc": u, "digits": digits,
                     "name": str(r.get(name_col) or "").strip(),
                     "brand": str(r.get(brand_col) or "").strip()})
    return rows


def words(text):
    return re.findall(r"[a-z0-9]+", (text or "").lower().replace("'", ""))


def brand_slug(brand, name):
    """Lowercase URL-ish prefix for a brand: 'dannon', 'la-yogurt', 'siggis'."""
    # Cut at apostrophes: "Siggi'S" -> "siggi", which covers both the
    # "Siggi-s-..." and "Siggis-..." spellings retailers use.
    split = lambda t: re.findall(r"[a-z0-9]+", (t or "").lower())
    w = split(brand) or split(name)[:2]
    if not w:
        return ""
    if len(w[0]) >= 4 or len(w) == 1:
        return w[0]
    return f"{w[0]}-{w[1]}"


def name_tokens(name, brand):
    bw = set(words(brand))
    return [t for t in dict.fromkeys(words(name))
            if t not in GENERIC_WORDS and t not in bw and len(t) > 1
            and not re.fullmatch(r"\d+|x\d+|\d+(?:p|pk|ct|oz)", t)]


# --------------------------------------------------------------------------
# Page parsing
# --------------------------------------------------------------------------

class _TextExtractor(HTMLParser):
    BLOCK = {"p", "div", "br", "tr", "li", "td", "th", "h1", "h2", "h3",
             "h4", "table", "section", "dt", "dd"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts, self._skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self._skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript") and self._skip:
            self._skip -= 1
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def page_text(raw_bytes):
    raw = raw_bytes.decode("utf-8", "replace")
    p = _TextExtractor()
    try:
        p.feed(raw)
        text = "".join(p.parts)
    except Exception:
        text = re.sub(r"<[^>]+>", " ", raw)
    text = html.unescape(text)
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    return re.sub(r"\n\s*\n+", "\n", text).strip()


def extract_nutrition(text):
    low = text.lower()
    idx = low.find("nutrition facts")
    if idx < 0:
        idx = low.find("calories")
    if idx < 0:
        return {}, ""
    window = low[idx: idx + 2500]
    found = {}
    for field, pat in NUTRIENT_PATTERNS.items():
        m = re.search(pat, window)
        if m:
            found[field] = m.group(1).strip()
    snippet = re.sub(r"\s+", " ", text[idx: idx + 600]).strip()
    return found, snippet


def page_cache_path(work, ts, orig):
    h = hashlib.sha1(f"{ts} {orig}".encode()).hexdigest()
    return os.path.join(work, "pages", h[:2], h + ".txt.gz")


# --------------------------------------------------------------------------
# Steps
# --------------------------------------------------------------------------

def step_upc_index(client, args, work, rows):
    os.makedirs(os.path.join(work, "cdx"), exist_ok=True)
    if args.sample_mode:
        return step_upc_index_sample(client, args, work, rows)
    for i, site in enumerate(UPC_SITES, 1):
        log(f"[{i}/{len(UPC_SITES)}] {site['prefix']}")
        path = os.path.join(work, "cdx", f"upc_{site['name']}.tsv")
        try:
            n = cdx_dump(client, site["prefix"], args.cutoff,
                         [site["filter"]] if site.get("filter") else [], path, args.page_size)
            log(f"    {n} captures through {args.cutoff[:4]}")
        except RuntimeError as e:
            log(f"    skipped: {e}  (run again later to retry this site)")


def step_upc_index_sample(client, args, work, rows):
    """Trial runs: search each site for just these UPCs instead of downloading
    the site's whole list. Sites where the UPC isn't right after the prefix
    (Meijer, Walmart search...) are skipped here; the full run covers them."""
    sites = [s for s in UPC_SITES if not s.get("skip_in_trial")
             and (not s.get("filter") or s.get("upc_right_after_prefix"))]
    jobs = []
    for r in rows:
        full = full_upc(r["digits"])
        for form in dict.fromkeys([full, "0" + full]):
            for site in sites:
                jobs.append((site, form))
    for n, (site, form) in enumerate(jobs, 1):
        prefix = site["prefix"] + form
        path = os.path.join(work, "cdx", f"upc_{site['name']}__{form}.tsv")
        if os.path.exists(path + ".done"):
            continue
        log(f"[{n}/{len(jobs)}] {prefix} ... ", end="")
        try:
            log(f"{cdx_dump(client, prefix, args.cutoff, [], path, args.page_size)} captures")
        except RuntimeError as e:
            log(f"skipped: {e}")


def step_name_index(client, args, work, rows):
    os.makedirs(os.path.join(work, "cdx"), exist_ok=True)
    counts = defaultdict(int)
    for r in rows:
        s = brand_slug(r["brand"], r["name"])
        if s:
            counts[s] += 1
    slugs = sorted((s for s, n in counts.items() if n >= args.min_brand_rows),
                   key=lambda s: -counts[s])
    total = len(slugs) * len(NAME_SITES)
    log(f"{len(slugs)} brands x {len(NAME_SITES)} retailers = {total} searches "
        f"(already-finished ones are skipped)")
    n = 0
    for slug in slugs:
        for site in NAME_SITES:
            n += 1
            prefix = site["prefix"].format(brand=slug)
            path = os.path.join(work, "cdx", f"name_{site['name']}_{slug}.tsv")
            if os.path.exists(path + ".done"):
                continue
            log(f"[{n}/{total}] {prefix} ... ", end="")
            try:
                got = cdx_dump(client, prefix, args.cutoff, [], path, args.page_size,
                               max_pages=args.name_max_pages)
                log(f"{got} captures")
            except RuntimeError as e:
                log(f"skipped: {e}")


def step_match(args, work, rows):
    cdx_dir = os.path.join(work, "cdx")
    if not os.path.isdir(cdx_dir):
        sys.exit("No capture lists yet. Run the upc-index step first.")

    # UPC matches: digit runs in URLs vs. your UPCs.
    index = defaultdict(set)
    for i, r in enumerate(rows):
        for k in upc_keys(r["digits"], may_have_check=len(r["digits"]) >= 12):
            index[k].add(i)
    cands = []
    for fn in sorted(os.listdir(cdx_dir)):
        if not (fn.startswith("upc_") and fn.endswith(".tsv")):
            continue
        site = fn[4:-4].split("__")[0]
        for orig, ts in read_dump(os.path.join(cdx_dir, fn)).items():
            path = urllib.parse.urlsplit(orig)
            hits = set()
            for run in re.findall(r"\d{8,14}", path.path + "?" + path.query):
                for k in upc_keys(run, may_have_check=True):
                    hits |= index.get(k, set())
            for i in hits:
                cands.append({"row": i, "upc": rows[i]["upc"], "match_type": "upc_in_url",
                              "site": site, "name_score": "", "timestamp": ts,
                              "original_url": orig})
    n_upc = len(cands)

    # Name matches: brand-prefixed retailer URLs vs. product description words.
    by_slug = defaultdict(list)
    for i, r in enumerate(rows):
        by_slug[brand_slug(r["brand"], r["name"])].append(i)
    for fn in sorted(os.listdir(cdx_dir)):
        if not (fn.startswith("name_") and fn.endswith(".tsv")):
            continue
        site, _, slug = fn[5:-4].partition("_")
        urls = read_dump(os.path.join(cdx_dir, fn))
        if not urls:
            continue
        url_words = {u: set(words(urllib.parse.unquote(u))) for u in urls}
        for i in by_slug.get(slug, []):
            toks = name_tokens(rows[i]["name"], rows[i]["brand"])
            if len(toks) < 2:
                continue  # descriptions like "Ups Nf Ft Y St&Bn 4P" can't be name-matched
            scored = []
            for u, uw in url_words.items():
                hit = sum(1 for t in toks if t in uw)
                if hit >= 2 and hit / len(toks) >= args.min_name_score:
                    scored.append((hit / len(toks), urls[u], u))
            for score, ts, u in sorted(scored, reverse=True)[:args.name_candidates]:
                cands.append({"row": i, "upc": rows[i]["upc"], "match_type": "name_in_url",
                              "site": site, "name_score": f"{score:.2f}", "timestamp": ts,
                              "original_url": u})

    with open(os.path.join(work, "candidates.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CANDIDATE_FIELDS)
        w.writeheader()
        w.writerows(cands)
    matched = len({c["row"] for c in cands})
    log(f"{n_upc} UPC-in-URL and {len(cands) - n_upc} name candidates, "
        f"covering {matched} of {len(rows)} UPCs")


def load_candidates(work):
    path = os.path.join(work, "candidates.csv")
    if not os.path.exists(path):
        sys.exit("No candidates yet. Run the match step first.")
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def step_fetch(client, args, work):
    cands = load_candidates(work)
    # Each archived page once, UPC matches first, then best name scores.
    pages = {}
    for c in sorted(cands, key=lambda c: (c["match_type"] != "upc_in_url",
                                          -float(c["name_score"] or 0))):
        pages.setdefault((c["timestamp"], c["original_url"]), None)
    todo = [p for p in pages if not os.path.exists(page_cache_path(work, *p))]
    already = len(pages) - len(todo)
    if args.max_fetch:
        todo = todo[:args.max_fetch]
    log(f"{len(pages)} pages to check, {already} already downloaded, fetching {len(todo)} "
        f"now (roughly {len(todo) * (args.delay + 2) / 3600:.1f} hours)")
    for n, (ts, orig) in enumerate(todo, 1):
        log(f"[{n}/{len(todo)}] {ts[:8]} {orig[:90]}")
        try:
            status, body = client.get(f"{WAYBACK}/web/{ts}id_/{orig}")
        except RuntimeError as e:
            log(f"    skipped: {e}")
            continue
        text = page_text(body) if status == 200 else ""
        path = page_cache_path(work, ts, orig)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with gzip.open(path, "wt", encoding="utf-8") as f:
            f.write(f"HTTP {status}\n{text}")


def step_report(args, work, rows):
    try:
        import openpyxl
    except ImportError:
        sys.exit("Writing Excel needs openpyxl. Run: python3 -m pip install openpyxl")
    cands = load_candidates(work)
    nut_cols = list(NUTRIENT_PATTERNS)
    results = []
    for c in cands:
        r = rows[int(c["row"])]
        ts, orig = c["timestamp"], c["original_url"]
        rec = {**c, "upc_descr": r["name"], "brand_descr": r["brand"],
               "full_upc": full_upc(r["digits"]),
               "capture_date": f"{ts[:4]}-{ts[4:6]}-{ts[6:8]}",
               "archive_url": f"https://web.archive.org/web/{ts}/{orig}",
               "fetched": "no", "upc_in_page": "", "nutrition_found": "", "nutrition_snippet": ""}
        path = page_cache_path(work, ts, orig)
        if os.path.exists(path):
            with gzip.open(path, "rt", encoding="utf-8") as f:
                status, _, text = f.read().partition("\n")
            rec["fetched"] = "yes" if status == "HTTP 200" else status
            if text:
                page_keys = set()
                for run in re.findall(r"\d{8,14}", re.sub(r"(?<=\d)[ -](?=\d)", "", text)):
                    page_keys |= upc_keys(run, may_have_check=True)
                mine = upc_keys(r["digits"], may_have_check=len(r["digits"]) >= 12)
                rec["upc_in_page"] = "yes" if mine & page_keys else "no"
                nut, snippet = extract_nutrition(text)
                rec.update(nut)
                rec["nutrition_found"] = "yes" if "calories" in nut else "no"
                rec["nutrition_snippet"] = snippet
        results.append(rec)

    def rank(rec):
        return ((rec["upc_in_page"] == "yes") * 4 + (rec["nutrition_found"] == "yes") * 2
                + (rec["match_type"] == "upc_in_url"), float(rec["name_score"] or 0),
                rec["timestamp"])

    best = {}
    for rec in results:
        i = int(rec["row"])
        if i not in best or rank(rec) > rank(best[i]):
            best[i] = rec

    cand_cols = ["upc", "full_upc", "upc_descr", "brand_descr", "match_type", "site",
                 "name_score", "capture_date", "archive_url", "fetched", "upc_in_page",
                 "nutrition_found", *nut_cols, "nutrition_snippet", "original_url"]
    best_cols = ["upc", "full_upc", "upc_descr", "brand_descr", "found", "confidence",
                 *cand_cols[4:]]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "best_by_upc"
    ws.append(best_cols)
    for i, r in enumerate(rows):
        rec = best.get(i)
        if rec is None:
            ws.append([r["upc"], full_upc(r["digits"]), r["name"], r["brand"], "no"]
                      + [""] * (len(best_cols) - 5))
            continue
        conf = ("high" if rec["upc_in_page"] == "yes" and rec["nutrition_found"] == "yes"
                else "medium" if rec["upc_in_page"] == "yes" or
                (rec["match_type"] == "upc_in_url" and rec["nutrition_found"] == "yes")
                else "low")
        ws.append([r["upc"], rec["full_upc"], r["name"], r["brand"], "yes", conf]
                  + [rec.get(k, "") for k in cand_cols[4:]])
    ws2 = wb.create_sheet("all_candidates")
    ws2.append(cand_cols)
    for rec in sorted(results, key=lambda x: int(x["row"])):
        ws2.append([rec.get(k, "") for k in cand_cols])
    for sheet in (ws, ws2):
        for row in sheet.iter_rows(min_row=2, max_col=2):
            for cell in row:
                cell.number_format = "@"  # keep UPCs as text (leading zeros)
        sheet.freeze_panes = "A2"
    out = os.path.join(os.path.dirname(os.path.abspath(args.input)), args.out)
    wb.save(out)
    n_high = sum(1 for i in best if ws.cell(row=i + 2, column=6).value == "high")
    log(f"Wrote {out}: {len(best)} of {len(rows)} UPCs have a candidate, "
        f"{n_high} with the UPC and a nutrition panel on the same archived page")


# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", default="yogurt_upc_all.xlsx")
    ap.add_argument("--sheet", help="sheet name (default: the first sheet)")
    ap.add_argument("--upc-col", default="upc")
    ap.add_argument("--name-col", default="upc_descr")
    ap.add_argument("--brand-col", default="brand_descr")
    ap.add_argument("--out", help="output file name (default wayback_results.xlsx, "
                    "or wayback_results_sample.xlsx for a trial run)")
    ap.add_argument("--sample", type=int, default=0,
                    help="trial run on this many randomly picked UPCs (takes a few minutes)")
    ap.add_argument("--upcs", help="trial run on these UPCs, comma-separated")
    ap.add_argument("--rows", help="full-method run on a block of UPC rows, e.g. 1-1000 "
                    "then 1001-2000 (numbered in sheet order, header not counted)")
    ap.add_argument("--seed", type=int, default=1,
                    help="change this to get a different random --sample")
    ap.add_argument("--work", help="folder for downloaded data (default: wayback_work next to input)")
    ap.add_argument("--steps", default=",".join(STEPS),
                    help=f"comma-separated subset of: {','.join(STEPS)}")
    ap.add_argument("--cutoff", default="20151231235959",
                    help="latest capture time to accept, YYYYMMDDhhmmss")
    ap.add_argument("--delay", type=float, default=1.5,
                    help="seconds between requests (the Wayback Machine limits fast clients)")
    ap.add_argument("--page-size", type=int, default=20000, help="CDX rows per request")
    ap.add_argument("--min-brand-rows", type=int, default=1,
                    help="only search retailers for brands with at least this many UPCs")
    ap.add_argument("--name-max-pages", type=int, default=3,
                    help="CDX pages to download per brand/retailer search (caps short, "
                         "common brand words like 'joy' on amazon.com)")
    ap.add_argument("--min-name-score", type=float, default=0.5,
                    help="share of description words that must appear in a URL (0-1)")
    ap.add_argument("--name-candidates", type=int, default=3,
                    help="max name-matched pages kept per UPC")
    ap.add_argument("--max-fetch", type=int, default=0,
                    help="download at most this many pages this run (0 = all)")
    ap.add_argument("--wayback-base", help=argparse.SUPPRESS)  # for testing
    args = ap.parse_args()

    global WAYBACK
    if args.wayback_base:
        WAYBACK = args.wayback_base.rstrip("/")
    if not os.path.exists(args.input):
        beside = os.path.join(os.path.dirname(os.path.abspath(__file__)), args.input)
        if os.path.exists(beside):
            args.input = beside
    if not os.path.exists(args.input):
        sys.exit(f"Can't find {args.input}. Put it in the same folder as this script, "
                 f"or give its full path with --input.")
    args.sample_mode = bool(args.sample or args.upcs)
    args.out = args.out or ("wayback_results_sample.xlsx" if args.sample_mode
                            else f"wayback_results_rows_{args.rows}.xlsx" if args.rows
                            else "wayback_results.xlsx")
    work = args.work or os.path.join(os.path.dirname(os.path.abspath(args.input)),
                                     "wayback_work_sample" if args.sample_mode else "wayback_work")
    os.makedirs(work, exist_ok=True)

    steps = [s.strip() for s in args.steps.split(",") if s.strip()]
    bad = [s for s in steps if s not in STEPS]
    if bad:
        sys.exit(f"Unknown step(s): {', '.join(bad)}. Choose from {', '.join(STEPS)}")

    rows = load_rows(args.input, args.sheet, args.upc_col, args.name_col, args.brand_col)
    if args.rows:
        m = re.fullmatch(r"\s*(\d+)\s*-\s*(\d+)\s*", args.rows)
        if not m or int(m.group(1)) < 1 or int(m.group(2)) < int(m.group(1)):
            sys.exit("--rows must look like 1-1000")
        rows = rows[int(m.group(1)) - 1:int(m.group(2))]
        log(f"Using UPC rows {args.rows}: {rows[0]['upc']} ... {rows[-1]['upc']}"
            if rows else "No UPC rows in that range.")
        if not rows:
            sys.exit(1)
    if args.upcs:
        wanted = {re.sub(r"\D", "", u) for u in args.upcs.split(",") if u.strip()}
        rows = [r for r in rows if r["digits"] in wanted or r["digits"].lstrip("0") in wanted]
        if not rows:
            sys.exit("None of the --upcs values are in the spreadsheet.")
    elif args.sample:
        # Prefer rows with a brand and a readable description, so the trial
        # exercises both UPC and name matching.
        good = [r for r in rows if r["brand"] and
                sum(len(t) >= 4 for t in name_tokens(r["name"], r["brand"])) >= 2]
        rows = random.Random(args.seed).sample(good, min(args.sample, len(good)))
    if args.sample_mode:
        log("Trial run on: " + ", ".join(f"{r['upc']} ({r['name'][:40]})" for r in rows))
    log(f"Loaded {len(rows)} UPCs from {os.path.basename(args.input)}. Work folder: {work}")
    if args.sample_mode:
        # Trial runs should finish in minutes: give up on a slow or failing
        # request quickly, and only look at the first page of brand results.
        client = Client(args.delay, timeout=45, waits=(10,))
        args.name_max_pages, args.page_size = 1, min(args.page_size, 5000)
    else:
        client = Client(args.delay)
    for step in STEPS:
        if step not in steps:
            continue
        log(f"\n=== {step} ===")
        if step == "upc-index":
            step_upc_index(client, args, work, rows)
        elif step == "name-index":
            step_name_index(client, args, work, rows)
        elif step == "match":
            step_match(args, work, rows)
        elif step == "fetch":
            step_fetch(client, args, work)
        elif step == "report":
            step_report(args, work, rows)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
PFE / Internship Scraper v2 - Casablanca & Morocco
==================================================
Drop-in replacement for the original scrape_pfe.py (same place, same imports).

What's new vs v1
----------------
 SOURCES     * Per-site error reporting (a silent 0 from a board is now visible)
             * Retry + exponential backoff + jitter on rate limits
             * Rekrute: pagination, multiple queries, per-post error isolation, runs in parallel
             * Bayt is opt-in (returned 0 rows for Morocco in v1); Google optional
 QUALITY     * Relevance filter (title / site job-type / description) - no more "Agent de saisie"
             * Honest "Job Type" (no fake "Internship / PFE" fallback)
             * Date window enforced on ALL boards (Rekrute ignored it before)
             * Location cleaning ("Casablanca, 6, Morocco" -> "Casablanca, Morocco"),
               city extraction, Greater-Casablanca scope filter
             * Accent-insensitive, scored domain classifier with 9 domains
             * Cross-board dedup that MERGES duplicates (adds an "Also On" column)
 USEFULNESS  * Run history: "NEW since last run" flag, --new-only
             * Presets: core / tech / business / engineering / all  (+ --keywords)
             * --deep: fetch descriptions to extract Duration & Paid, rescue borderline offers
             * Excel: Offers / Summary / Filtered Out sheets, real dates, colored domains
             * --from-csv: re-process an old CSV offline (no network)
 HYGIENE     * Full CLI, logging, output folder, dated files (no more *_1793365898.xlsx clutter),
               lazy imports (module imports cleanly without network libs)

Usage
-----
    python scrape_pfe.py                              # core keywords, last 30 days
    python scrape_pfe.py --preset core,tech --days 14
    python scrape_pfe.py --domain data,software --new-only
    python scrape_pfe.py --sites indeed,linkedin,bayt --deep --limit 50
    python scrape_pfe.py --from-csv old.csv --days 90   # offline re-processing
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import random
import re
import sys
import time
import unicodedata
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote_plus, urljoin

# --- Windows console UTF-8 ---------------------------------------------------
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# --- Make the JobSpy package importable when this file lives inside it -------
_here = Path(__file__).resolve().parent
_root = _here.parent if _here.name == "jobspy" else _here
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

log = logging.getLogger("pfe")

# =============================================================================
# 0. CONFIG
# =============================================================================
PRESETS: dict[str, list[str]] = {
    "core": ["stage pfe", "stage pré-embauche", "stagiaire pfe", "internship"],
    "tech": ["pfe informatique", "stage développeur", "pfe data", "pfe cybersécurité", "stage ingénieur"],
    "business": ["pfe finance", "stage marketing digital", "stage ressources humaines", "stage supply chain"],
    "engineering": ["pfe génie industriel", "pfe mécanique", "pfe électrique", "stage qualité"],
}
PRESETS["all"] = [k for v in PRESETS.values() for k in v]

DEFAULT_SITES = ["indeed", "linkedin"]          # bayt / google are opt-in via --sites
DEFAULT_REKRUTE_QUERIES = ["pfe", "stage"]

COLUMNS = [
    "Job Title", "Domain", "Company", "City", "Location", "Date Posted", "Days Old",
    "Job Board", "Also On", "Job Type", "Remote", "Relevance", "Duration", "Paid",
    "New", "Job Link",
]

# =============================================================================
# 1. TEXT HELPERS
# =============================================================================


def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def norm(s: object) -> str:
    """lowercase, accent-free, punctuation -> spaces (keeps # + .)"""
    if s is None:
        return ""
    s = strip_accents(str(s)).lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9#+.]+", " ", s)).strip()


def clean(v: object) -> str:
    """Safe string from pandas/NaN/None/list values."""
    if v is None:
        return ""
    if isinstance(v, float) and v != v:  # NaN
        return ""
    if isinstance(v, (list, tuple, set)):
        return ", ".join(str(getattr(x, "value", x)) for x in v)
    s = str(getattr(v, "value", v)).strip()
    return "" if s.lower() in {"nan", "none", "nat"} else s


# =============================================================================
# 2. DOMAIN CLASSIFIER (scored, accent-insensitive)
# =============================================================================
# Order = tie-break priority.
DOMAINS: dict[str, tuple[str, str]] = {
    # key: (label, regex on normalized text)
    "data": ("Data & AI", r"\b(?:data\w*|donnees?|analytics?|business intelligence|bi|power ?bi|tableau|etl|"
                          r"machine learning|deep learning|ml|nlp|llm|genai|ia|ai|intelligence artificielle|"
                          r"artificial intelligence|computer vision|scientist|big data)\b"),
    "software": ("Software & Web Dev", r"\b(?:dev|develop\w*|developp\w*|software|logiciels?|informatique|web|full ?stack|"
                                       r"front ?end|back ?end|react\w*|python|java|javascript|typescript|spring|angular|"
                                       r"vue|node\w*|php|laravel|symfony|mobile|ios|android|flutter|applications?|api|"
                                       r"erp|sap|hybris|salesforce|odoo|qa|testeur|testing|ui|ux|programm\w*|coding)\b"),
    "cloud": ("Cloud, Cyber & Systems", r"\b(?:cyber\w*|securite (?:informatique|des systemes|reseaux?|des si)|soc|pentest\w*|"
                                        r"cloud|aws|azure|gcp|devops|devsecops|kubernetes|docker|linux|reseaux?|networks?|"
                                        r"networking|telecoms?|telecommunications?|systemes?|systems?|infra\w*|sysadmin|"
                                        r"embarques?|embedded|iot|5g|fibre|cisco)\b"),
    "engineering": ("Engineering & Industry", r"\b(?:meca\w*|mechanic\w*|mechatron\w*|electri\w*|electron\w*|industri\w*|genie|"
                                              r"civil|btp|batiment|chantier|maintenance|automat\w*|robot\w*|qhse|hse|qualite|"
                                              r"quality|production|procedes?|menuiserie|bois|aeronaut\w*|aerospace|"
                                              r"automobile|methodes|lean|energie|energy|renouvelables?|solaire|engineer\w*|"
                                              r"ingenieur\w*)\b"),
    "finance": ("Finance & Accounting", r"\b(?:financ\w*|compta\w*|accounting|accountant|audit\w*|controle de gestion|"
                                        r"tresorerie|fiscal\w*|tax|banque|bank\w*|assurances?|actuar\w*|controlling|"
                                        r"credit|risques?|risk|invest\w*|analyst|fund|fonds)\b"),
    "marketing": ("Marketing & Communication", r"\b(?:marketing|communication|community manager|brand\w*|marque|social media|"
                                               r"reseaux sociaux|contenus?|content|seo|sea|graphi\w*|design\w*|creati\w*|"
                                               r"video|monteu\w*|montage|medias?|events?|evenement\w*|exhibitions?|salons?|"
                                               r"relations? (?:presse|publiques)|journalis\w*|redact\w*|copywrit\w*|"
                                               r"influence\w*|photograph\w*)\b"),
    "supply": ("Supply Chain & Logistics", r"\b(?:supply chain|logisti\w*|achats?|procurement|approvisionnements?|transports?|"
                                           r"douane\w*|import|export|stocks?|entrepots?|warehouse|planification|planning|"
                                           r"sourcing|transit|transitaire)\b"),
    "business": ("Business & Consulting", r"\b(?:commercial\w*|ventes?|vendeu\w*|sales|business|key account|client\w*|"
                                          r"conseill\w*|consult\w*|projets?|project|pmo|management|manager|gestion|"
                                          r"strateg\w*|partenariats?|account|e commerce|ecommerce|customer|"
                                          r"entrepreneur\w*|gestionnaire|founder|associate|retail|commerce|appels? d offres?)\b"),
    "hr": ("HR, Admin & Legal", r"\b(?:rh|hr|ressources humaines|human resources|recrut\w*|recruit\w*|talents?|"
                                r"administrati\w*|secretaire|secretariat|saisie|accueil|juridique|juriste|legal|droit|"
                                r"conformite|compliance|paie|payroll)\b"),
}
DOMAIN_RE = {k: re.compile(v[1]) for k, v in DOMAINS.items()}
DOMAIN_LABEL = {k: v[0] for k, v in DOMAINS.items()}
LABEL_TO_KEY = {v: k for k, v in DOMAIN_LABEL.items()}
OTHER_DOMAIN = "Other / General"

# "business development", "développement commercial" are NOT software dev.
DEV_NOISE = re.compile(
    r"\b(?:business|commercial)\s+(?:development|developpement)\b|"
    r"\b(?:developpement|development)\s+(?:commercial|business|des affaires|durable|rh|international|"
    r"de carriere|de marque|et partenariats?|personnel)\b"
)


def classify_domain(title: str) -> str:
    raw_text = norm(title)
    text = DEV_NOISE.sub(" ", raw_text)
    biz_bonus = len(DEV_NOISE.findall(raw_text))   # "business development" IS business
    best_label, best_score = OTHER_DOMAIN, 0
    for key, rx in DOMAIN_RE.items():          # dict order = priority
        score = len(set(rx.findall(text))) + (biz_bonus if key == "business" else 0)
        if score > best_score:
            best_label, best_score = DOMAIN_LABEL[key], score
    return best_label


# =============================================================================
# 3. RELEVANCE (is this actually an internship?)
# =============================================================================
INTERN_RE = re.compile(
    r"\b(?:stages?|stagiaires?|pfe|pfa|intern|interns|internships?|pre ?embauche|alternan\w*|apprenti\w*|"
    r"fin d etudes|end of studies|trainees?)\b"
)
REL_RANK = {"High": 3, "Medium": 2, "Low": 1}


def relevance(title: str, job_type: str, description: str) -> str:
    if INTERN_RE.search(norm(title)):
        return "High"
    if "intern" in norm(job_type):
        return "Medium"
    if description and INTERN_RE.search(norm(description[:1500])):
        return "Medium"
    return "Low"


_DUR_RE = re.compile(r"\b(\d{1,2})(?:\s*(?:a|-|to)\s*(\d{1,2}))?\s*(?:mois|months?)\b")
_PAID_RE = re.compile(r"\b(?:remunere\w*|gratification|indemnite\w*|paid|allowance|compensat\w*)\b")


def extract_extras(description: str) -> tuple[str, str]:
    if not description:
        return "", ""
    n = norm(description)
    m = _DUR_RE.search(n)
    dur = ""
    if m:
        dur = f"{m.group(1)}-{m.group(2)} months" if m.group(2) else f"{m.group(1)} months"
    return dur, ("Yes" if _PAID_RE.search(n) else "")


# =============================================================================
# 4. LOCATION CLEANING
# =============================================================================
CASA_TERMS = [
    "casablanca", "casa", "bouskoura", "nouaceur", "mohammedia", "ain sebaa", "sidi maarouf", "maarouf",
    "mechouar", "dar bouazza", "bernoussi", "sidi moumen", "anfa", "maarif", "ain diab", "californie",
    "oulfa", "ain chock", "hay hassani", "ben msik", "mediouna", "tit mellil", "ain harrouda", "berrechid",
]
OTHER_CITIES = [
    "rabat", "sale", "temara", "tanger", "tangier", "tetouan", "marrakech", "fes", "fez", "meknes",
    "agadir", "kenitra", "oujda", "el jadida", "laayoune", "dakhla", "nador", "safi", "beni mellal",
    "khouribga", "tanger med", "ouarzazate", "essaouira",
]
GENERIC_PARTS = {"morocco", "maroc", "casablanca settat", "casablanca metropolitan area",
                 "prefecture of casablanca", "region de casablanca settat", "grand casablanca"}


def _has_term(n: str, terms: list[str]) -> str | None:
    padded = f" {n} "
    for t in terms:
        if f" {t} " in padded:
            return t
    return None


def clean_location(raw: str) -> tuple[str, str, str]:
    """-> (display_location, city, zone) ; zone in {Greater Casablanca, Other city, Unknown}"""
    parts = [p.strip() for p in re.split(r"[,|]", raw or "") if p.strip()]
    parts = [p for p in parts if not re.fullmatch(r"\d+", p)]       # drop Indeed region code ("6")
    n_all = norm(" ".join(parts))
    zone = "Unknown"
    if _has_term(n_all, CASA_TERMS):
        zone = "Greater Casablanca"
    elif _has_term(n_all, OTHER_CITIES):
        zone = "Other city"

    city = ""
    for p in parts:
        if norm(p) not in GENERIC_PARTS:
            city = p
            break
    if not city and "casablanca" in n_all:
        city = "Casablanca"

    display = ", ".join(parts)
    if display and "morocco" not in norm(display) and "maroc" not in norm(display):
        display += ", Morocco"
    return display, city, zone


# =============================================================================
# 5. DATES, LINKS, DEDUP KEYS
# =============================================================================


def parse_date(v: object) -> date | None:
    s = clean(v)
    if not s:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        pass
    for fmt in ("%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def link_key(url: str) -> str:
    """Stable identity for a posting across URL variants."""
    if not url:
        return ""
    u = url.strip()
    m = re.search(r"linkedin\.[a-z.]+/jobs/view/(?:[^/?#]*?-)?(\d{6,})", u)
    if m:
        return f"li:{m.group(1)}"
    m = re.search(r"[?&]jk=([0-9a-f]+)", u)
    if m and "indeed" in u:
        return f"in:{m.group(1)}"
    return u.split("?")[0].split("#")[0].rstrip("/").lower()


def normalize_job_link(url: str) -> str:
    """Display link: keeps Indeed ?jk= (v1 behaviour), strips tracking elsewhere."""
    if not url:
        return ""
    url = url.strip()
    if "indeed.com" in url:
        m = re.search(r"([?&]jk=[^&]+)", url)
        return f"https://ma.indeed.com/viewjob{m.group(1)}" if m else url
    return url.split("?")[0]


_COMPANY_NOISE = {"maroc", "morocco", "sa", "sarl", "sas", "group", "groupe", "ltd", "inc", "international"}
_UNKNOWN_COMPANY = {"", "not specified", "confidential", "confidentiel", "anonyme", "n a"}


def company_key(company: str) -> str:
    n = norm(company)
    if n in _UNKNOWN_COMPANY:
        return ""
    return "".join(w for w in n.split() if w not in _COMPANY_NOISE)


def title_key(title: str) -> str:
    return re.sub(r"[^a-z0-9]", "", norm(title))


# =============================================================================
# 6. SCRAPERS
# =============================================================================
# Internal raw schema: title, company, location, date_posted, site, url,
#                      job_type, is_remote, description


def search_jobspy(term: str, sites: list[str], location: str, hours_old: int, limit: int,
                  deep: bool, retries: int = 2) -> tuple[list[dict], dict[str, str]]:
    """One keyword across several boards, with retry/backoff. Returns (raw rows, per-site errors)."""
    try:
        from jobspy import scrape_jobs
    except ImportError as e:  # pragma: no cover
        return [], {"import": f"JobSpy not importable: {e}"}

    errors: dict[str, str] = {}
    for attempt in range(retries + 1):
        try:
            kwargs = dict(
                site_name=sites, search_term=term, location=location, country_indeed="morocco",
                results_wanted=limit, hours_old=hours_old, fetch_description=deep, verbose=0,
            )
            if "google" in sites:
                kwargs["google_search_term"] = f"{term} jobs in {location}"
            df = scrape_jobs(**kwargs)
            errors = dict(getattr(df, "attrs", {}).get("errors", {}) or {})
            if df.empty and errors and attempt < retries:
                wait = (2 ** attempt) * 4 + random.uniform(0, 2)
                log.warning("    all boards errored (%s) - retry in %.0fs", ", ".join(errors), wait)
                time.sleep(wait)
                continue
            rows = []
            for _, r in df.iterrows():
                rows.append({
                    "title": clean(r.get("title")), "company": clean(r.get("company")),
                    "location": clean(r.get("location")), "date_posted": r.get("date_posted"),
                    "site": clean(r.get("site")), "url": clean(r.get("job_url")),
                    "job_type": clean(r.get("job_type")), "is_remote": r.get("is_remote"),
                    "description": clean(r.get("description")),
                })
            return rows, errors
        except Exception as e:  # noqa: BLE001
            errors = {"exception": str(e)[:200]}
            if attempt < retries:
                time.sleep((2 ** attempt) * 4 + random.uniform(0, 2))
    return [], errors


_REKRUTE_TAIL = re.compile(r"\s*\|\s*([^|()]+?)\s*\(Maroc\)\s*$", re.I)


def _parse_rekrute_post(post) -> dict | None:
    title_el = post.select_one("a.titreJob")
    if not title_el:
        return None
    title = title_el.get_text(" ", strip=True)
    url = urljoin("https://www.rekrute.com", title_el.get("href", ""))
    photo = post.select_one("img.photo")
    company = (photo.get("title", "").strip() if photo else "") or ""

    section = post.select_one(".section")
    section_text = section.get_text(" ") if section else post.get_text(" ")
    m = re.search(r"(\d{2}/\d{2}/\d{4})", section_text)
    posted = parse_date(m.group(1)) if m else None

    city = ""
    tail = _REKRUTE_TAIL.search(title)
    if tail:
        city = tail.group(1).strip()
        title = title[: tail.start()].strip()
    if city:
        location = f"{city}, Morocco"
    elif "casablanca" in norm(section_text):
        location = "Casablanca, Morocco"
    else:
        location = "Morocco"
    return {"title": title, "company": company, "location": location, "date_posted": posted,
            "site": "rekrute", "url": url, "job_type": "", "is_remote": None, "description": ""}


def scrape_rekrute(queries: list[str], pages: int = 3) -> list[dict]:
    """Native Rekrute.com scraper: multi-query, paginated, per-post error isolation."""
    try:
        from bs4 import BeautifulSoup
        from curl_cffi import requests as curl_requests
    except ImportError as e:
        log.warning("[!] Rekrute skipped (missing dependency: %s)", e)
        return []

    out: list[dict] = []
    seen: set[str] = set()
    for q in queries:
        for page in range(1, pages + 1):
            url = f"https://www.rekrute.com/offres.html?s=1&p={page}&o=1&query={quote_plus(q)}"
            try:
                resp = curl_requests.get(url, impersonate="chrome", timeout=20)
            except Exception as e:  # noqa: BLE001
                log.warning("[!] Rekrute '%s' p%d failed: %s", q, page, str(e)[:120])
                break
            if resp.status_code != 200:
                log.warning("[!] Rekrute '%s' p%d -> HTTP %s", q, page, resp.status_code)
                break
            posts = BeautifulSoup(resp.text, "html.parser").select("li.post-id")
            if not posts:
                break
            fresh = 0
            for post in posts:
                try:
                    rec = _parse_rekrute_post(post)
                except Exception as e:  # noqa: BLE001
                    log.debug("Rekrute post skipped: %s", e)
                    continue
                if rec and rec["url"] not in seen:
                    seen.add(rec["url"])
                    out.append(rec)
                    fresh += 1
            if fresh == 0:
                break
            time.sleep(random.uniform(1.0, 2.5))
    return out


def load_raw_from_csv(path: str) -> list[dict]:
    """Offline mode: re-process a v1/v2 CSV (utf-8 or utf-8-sig, comma or semicolon)."""
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        delim = ";" if sample.count(";") > sample.count(",") else ","
        rows = list(csv.DictReader(f, delimiter=delim))
    out = []
    for r in rows:
        jt = r.get("Job Type") or r.get("job_type") or ""
        if jt.strip().lower() == "internship / pfe":      # v1 fabricated fallback
            jt = ""
        remote = (r.get("Remote") or r.get("is_remote") or "").strip().lower()
        out.append({
            "title": r.get("Job Title") or r.get("title") or "",
            "company": r.get("Company") or r.get("company") or "",
            "location": r.get("Location") or r.get("location") or "",
            "date_posted": r.get("Date Posted") or r.get("date_posted") or "",
            "site": r.get("Job Board") or r.get("site") or "",
            "url": r.get("Job Link") or r.get("job_url") or "",
            "job_type": jt,
            "is_remote": True if remote in {"yes", "true"} else (False if remote in {"no", "false"} else None),
            "description": r.get("description") or "",
        })
    return out


# =============================================================================
# 7. PIPELINE: normalize -> filter -> dedup/merge -> history
# =============================================================================


def normalize(raw: dict, today: date) -> dict:
    title = clean(raw.get("title")) or "Not specified"
    company = clean(raw.get("company"))
    loc_display, city, zone = clean_location(clean(raw.get("location")))
    d = parse_date(raw.get("date_posted"))
    desc = clean(raw.get("description"))
    jt = clean(raw.get("job_type"))
    dur, paid = extract_extras(desc)
    remote = raw.get("is_remote")
    url = normalize_job_link(clean(raw.get("url")))
    return {
        "Job Title": title,
        "Domain": classify_domain(title),
        "Company": company or "Not specified",
        "City": city,
        "Location": loc_display,
        "Date Posted": d,
        "Days Old": (today - d).days if d else "",
        "Job Board": clean(raw.get("site")) or "unknown",
        "Also On": "",
        "Job Type": jt,
        "Remote": "Yes" if remote is True else ("No" if remote is False else ""),
        "Relevance": relevance(title, jt, desc),
        "Duration": dur,
        "Paid": paid,
        "New": "",
        "Job Link": url,
        "_zone": zone,
        "_lkey": link_key(url),
        "_tkey": (title_key(title), company_key(company)),
        "_also": set(),
    }


def merge_duplicates(rows: list[dict]) -> list[dict]:
    by_link: dict[str, dict] = {}
    by_tc: dict[tuple, dict] = {}
    unique: list[dict] = []
    fill = ("City", "Location", "Date Posted", "Days Old", "Job Type", "Remote", "Duration", "Paid")
    for r in rows:
        tk = r["_tkey"]
        has_tc = bool(tk[0] and tk[1])
        keeper = (by_link.get(r["_lkey"]) if r["_lkey"] else None) or (by_tc.get(tk) if has_tc else None)
        if keeper is None:
            unique.append(r)
            keeper = r
        else:
            if r["Job Board"] != keeper["Job Board"]:
                keeper["_also"].add(r["Job Board"])
            for f in fill:
                if keeper[f] in ("", None) and r[f] not in ("", None):
                    keeper[f] = r[f]
            if keeper["Company"] == "Not specified" and r["Company"] != "Not specified":
                keeper["Company"] = r["Company"]
            if REL_RANK[r["Relevance"]] > REL_RANK[keeper["Relevance"]]:
                keeper["Relevance"] = r["Relevance"]
        if r["_lkey"]:
            by_link[r["_lkey"]] = keeper
        if has_tc:
            by_tc[tk] = keeper
    for r in unique:
        r["Also On"] = ", ".join(sorted(r["_also"]))
    return unique


def record_key(r: dict) -> str:
    return r["_lkey"] or f"tc:{r['_tkey'][0]}|{r['_tkey'][1]}"


def load_history(path: Path) -> dict[str, str] | None:
    try:
        return json.loads(path.read_text(encoding="utf-8")).get("first_seen", {})
    except (OSError, ValueError):
        return None


def save_history(path: Path, seen: dict[str, str], today: date) -> None:
    cutoff = (today - timedelta(days=365)).isoformat()
    seen = {k: v for k, v in seen.items() if v >= cutoff}
    path.write_text(json.dumps({"first_seen": seen}, ensure_ascii=False, indent=0), encoding="utf-8")


def process(raw_rows: list[dict], *, today: date, days: int, scope: str, keep_all: bool,
            domain_keys: set[str]) -> tuple[list[dict], list[dict], dict[str, int]]:
    """-> (kept, rejected, funnel)"""
    funnel = {"collected": len(raw_rows)}
    rows = [normalize(r, today) for r in raw_rows]
    kept, rejected = [], []
    cutoff = today - timedelta(days=days)
    allowed_labels = {DOMAIN_LABEL[k] for k in domain_keys}

    for r in rows:
        reason = ""
        if r["Date Posted"] and r["Date Posted"] < cutoff:
            reason = f"Older than {days} days"
        elif scope == "casablanca" and r["_zone"] == "Other city":
            reason = "Outside Greater Casablanca"
        elif not keep_all and r["Relevance"] == "Low":
            reason = "Low relevance (no stage/PFE/intern signal)"
        elif domain_keys and r["Domain"] not in allowed_labels:
            reason = "Domain filtered out"
        (rejected if reason else kept).append({**r, "Reason": reason} if reason else r)

    funnel["after filters"] = len(kept)
    kept = merge_duplicates(kept)
    funnel["after dedup"] = len(kept)
    return kept, rejected, funnel


# =============================================================================
# 8. EXPORT
# =============================================================================
DOMAIN_COLORS = {
    "Data & AI": "E2EFDA", "Software & Web Dev": "DDEBF7", "Cloud, Cyber & Systems": "FFF2CC",
    "Engineering & Industry": "FCE4D6", "Finance & Accounting": "E4DFEC", "Marketing & Communication": "FDE9F2",
    "Supply Chain & Logistics": "DAEEF3", "Business & Consulting": "EDEDED", "HR, Admin & Legal": "F2DCDB",
}


def _csv_value(v):
    return v.isoformat() if isinstance(v, date) else v


def export_csv(rows: list[dict], path: Path, columns: list[str], sep: str = ",") -> Path:
    def write(p: Path):
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=sep)
            w.writerow(columns)
            for r in rows:
                w.writerow([_csv_value(r.get(c, "")) for c in columns])
    try:
        write(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{int(time.time())}{path.suffix}")
        write(alt)
        log.warning("[!] %s is locked (open in Excel) -> saved to %s", path.name, alt.name)
        return alt


def export_xlsx(rows: list[dict], rejected: list[dict], summary: dict, path: Path,
                columns: list[str]) -> Path:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Offers"

    head_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    body = Font(name="Calibri", size=10)
    link = Font(name="Calibri", size=10, color="0563C1", underline="single")
    thin = Side(style="thin", color="E0E0E0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_cols = {"Domain", "Date Posted", "Days Old", "Job Board", "Also On", "Job Type", "Remote",
                   "Relevance", "Duration", "Paid", "New"}

    def fill_sheet(sheet, data: list[dict], cols: list[str], with_domain_colors: bool):
        for ci, name in enumerate(cols, 1):
            c = sheet.cell(row=1, column=ci, value=name)
            c.font, c.fill = head_font, head_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        sheet.row_dimensions[1].height = 24
        for ri, r in enumerate(data, 2):
            for ci, name in enumerate(cols, 1):
                v = r.get(name, "")
                c = sheet.cell(row=ri, column=ci)
                if name == "Job Link" and str(v).startswith("http"):
                    c.value, c.hyperlink, c.font = "Open ↗", str(v), link
                elif name == "Job Title" and str(r.get("Job Link", "")).startswith("http"):
                    c.value, c.hyperlink, c.font = v, str(r["Job Link"]), link
                elif isinstance(v, date):
                    c.value, c.number_format, c.font = v, "yyyy-mm-dd", body
                elif name == "Days Old" and v != "":
                    c.value, c.font = int(v), body
                else:
                    c.value, c.font = ("" if v is None else v), body
                c.border = border
                c.alignment = Alignment(horizontal="center" if name in center_cols else "left", vertical="center")
                if with_domain_colors and name == "Domain" and v in DOMAIN_COLORS:
                    c.fill = PatternFill("solid", start_color=DOMAIN_COLORS[v], end_color=DOMAIN_COLORS[v])
                if name == "New" and v:
                    c.fill = PatternFill("solid", start_color="FFD966", end_color="FFD966")
                    c.font = Font(name="Calibri", size=10, bold=True)
        for ci, name in enumerate(cols, 1):
            longest = max([len(str(name))] + [len(str(r.get(name, ""))) for r in data[:300]])
            width = 12 if name == "Job Link" else min(max(longest + 3, 10), 50)
            sheet.column_dimensions[get_column_letter(ci)].width = width
        sheet.freeze_panes = "B2"
        sheet.auto_filter.ref = sheet.dimensions

    fill_sheet(ws, rows, columns, True)

    # ---- Summary sheet ----
    ss = wb.create_sheet("Summary")
    bold = Font(name="Calibri", size=11, bold=True)
    row = 1

    def block(title: str, pairs: list[tuple]):
        nonlocal row
        ss.cell(row=row, column=1, value=title).font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
        row += 1
        for k, v in pairs:
            ss.cell(row=row, column=1, value=k).font = body
            ss.cell(row=row, column=2, value=v).font = bold
            row += 1
        row += 1

    block("Run", summary["run"])
    block("Funnel", list(summary["funnel"].items()))
    block("Offers by domain", Counter(r["Domain"] for r in rows).most_common())
    block("Offers by board", Counter(r["Job Board"] for r in rows).most_common())
    block("Offers by city", Counter((r["City"] or "Unknown") for r in rows).most_common(15))
    block("Top companies", [x for x in Counter(r["Company"] for r in rows
                                               if r["Company"] != "Not specified").most_common(15)])
    if summary["errors"]:
        block("Board errors / warnings", list(summary["errors"].items()))
    ss.column_dimensions["A"].width = 38
    ss.column_dimensions["B"].width = 60

    # ---- Filtered out sheet ----
    if rejected:
        wr = wb.create_sheet("Filtered Out")
        rcols = ["Job Title", "Company", "Location", "Date Posted", "Job Board", "Reason", "Job Link"]
        fill_sheet(wr, rejected, rcols, False)

    try:
        wb.save(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{int(time.time())}{path.suffix}")
        wb.save(alt)
        log.warning("[!] %s is locked (open in Excel) -> saved to %s", path.name, alt.name)
        return alt


# =============================================================================
# 9. MAIN
# =============================================================================


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="PFE / internship scraper for Casablanca & Morocco",
                                formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--days", type=int, default=30, help="Keep offers posted in the last N days")
    p.add_argument("--limit", type=int, default=30, help="Results wanted per keyword per board")
    p.add_argument("--preset", default="core", help=f"Keyword presets, comma list of: {', '.join(PRESETS)}")
    p.add_argument("--keywords", default="", help="Extra keywords, comma-separated")
    p.add_argument("--location", default="Casablanca, Morocco")
    p.add_argument("--sites", default=",".join(DEFAULT_SITES),
                   help="Boards: indeed,linkedin,bayt,google,glassdoor,ziprecruiter ...")
    p.add_argument("--no-rekrute", action="store_true", help="Skip Rekrute.com")
    p.add_argument("--rekrute-queries", default=",".join(DEFAULT_REKRUTE_QUERIES))
    p.add_argument("--rekrute-pages", type=int, default=3)
    p.add_argument("--scope", choices=["casablanca", "morocco"], default="casablanca",
                   help="casablanca drops offers clearly located in other cities")
    p.add_argument("--domain", default="", help=f"Only these domains: {', '.join(DOMAINS)}")
    p.add_argument("--keep-all", action="store_true", help="Keep low-relevance rows (no stage/PFE/intern signal)")
    p.add_argument("--new-only", action="store_true", help="Export only offers not seen in previous runs")
    p.add_argument("--no-history", action="store_true", help="Do not read/write run history")
    p.add_argument("--deep", action="store_true", help="Fetch descriptions (slow): Duration, Paid, rescue borderline")
    p.add_argument("--outdir", default="pfe_output")
    p.add_argument("--base", default="stage_pfe_casablanca")
    p.add_argument("--french-csv", action="store_true", help="Also write a semicolon CSV for French-locale Excel")
    p.add_argument("--from-csv", default="", help="Re-process an existing CSV offline (no scraping)")
    p.add_argument("--today", default="", help=argparse.SUPPRESS)
    p.add_argument("-v", "--verbose", action="store_true")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(message)s")
    today = date.fromisoformat(args.today) if args.today else date.today()

    domain_keys = {d.strip().lower() for d in args.domain.split(",") if d.strip()}
    bad = domain_keys - set(DOMAINS)
    if bad:
        log.error("Unknown --domain key(s): %s (valid: %s)", ", ".join(sorted(bad)), ", ".join(DOMAINS))
        return 2

    presets = [p.strip() for p in args.preset.split(",") if p.strip()]
    unknown = [p for p in presets if p not in PRESETS]
    if unknown:
        log.error("Unknown preset(s): %s (valid: %s)", ", ".join(unknown), ", ".join(PRESETS))
        return 2
    keywords = list(dict.fromkeys(
        [k for p in presets for k in PRESETS[p]] + [k.strip() for k in args.keywords.split(",") if k.strip()]))
    sites = [s.strip().lower() for s in args.sites.split(",") if s.strip()]
    hours_old = args.days * 24

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    log.info("=" * 66)
    log.info("[*] PFE Internship Scraper v2  |  %s", today.isoformat())
    if args.from_csv:
        log.info("[*] OFFLINE mode: %s", args.from_csv)
    else:
        log.info("[*] Location: %s  |  Window: %d days  |  Scope: %s", args.location, args.days, args.scope)
        log.info("[*] Boards:   %s%s", ", ".join(sites), "" if args.no_rekrute else " + rekrute")
        log.info("[*] Keywords: %s", ", ".join(keywords))
    log.info("=" * 66)

    raw: list[dict] = []
    errors: dict[str, str] = {}
    per_board_raw: Counter = Counter()

    if args.from_csv:
        raw = load_raw_from_csv(args.from_csv)
    else:
        rek_future = None
        with ThreadPoolExecutor(max_workers=1) as pool:
            if not args.no_rekrute:
                queries = [q.strip() for q in args.rekrute_queries.split(",") if q.strip()]
                rek_future = pool.submit(scrape_rekrute, queries, args.rekrute_pages)
            for i, term in enumerate(keywords):
                if i:
                    time.sleep(random.uniform(2.5, 4.5))        # polite, jittered
                log.info("[*] '%s' ...", term)
                rows, errs = search_jobspy(term, sites, args.location, hours_old, args.limit, args.deep)
                log.info("    %d raw rows%s", len(rows), f"   [errors: {errs}]" if errs else "")
                for site, msg in errs.items():
                    errors[f"{site} ('{term}')"] = str(msg)[:160]
                raw.extend(rows)
            if rek_future:
                rek = rek_future.result()
                log.info("[*] Rekrute: %d listings", len(rek))
                raw.extend(rek)

    for r in raw:
        per_board_raw[clean(r.get("site")) or "unknown"] += 1
    for site in sites if not args.from_csv else []:
        if per_board_raw[site] == 0:
            errors.setdefault(f"{site} (zero rows)", "Board returned nothing - blocked, unsupported for Morocco, or rate-limited")

    kept, rejected, funnel = process(raw, today=today, days=args.days, scope=args.scope,
                                     keep_all=args.keep_all, domain_keys=domain_keys)

    # ---- history / NEW flag ----
    hist_path = outdir / f"{args.base}_history.json"
    history = None if args.no_history else load_history(hist_path)
    first_run = history is None and not args.no_history
    seen = dict(history or {})
    for r in kept:
        k = record_key(r)
        if history is not None and k not in history:
            r["New"] = "NEW"
        seen.setdefault(k, today.isoformat())
    new_count = sum(1 for r in kept if r["New"])
    exported = [r for r in kept if r["New"]] if args.new_only else kept

    exported.sort(key=lambda r: (r["Date Posted"] or date.min, REL_RANK[r["Relevance"]], r["Job Title"]),
                  reverse=True)

    # drop columns that are empty everywhere
    columns = [c for c in COLUMNS if c in {"Job Title", "Domain", "Company", "Job Link", "Job Board"}
               or any(r.get(c) not in ("", None) for r in exported)]

    if not exported:
        log.info("\n[-] Nothing to export (%s).", "no NEW offers" if args.new_only else "no matching offers")
        if not args.no_history and not args.new_only:
            save_history(hist_path, seen, today)
        return 0

    stamp = today.isoformat()
    summary = {
        "run": [("Date", stamp), ("Location", args.location), ("Window", f"{args.days} days"),
                ("Scope", args.scope), ("Boards", ", ".join(sites) + ("" if args.no_rekrute else ", rekrute")),
                ("Keywords", ", ".join(keywords) or "(offline)"), ("NEW since last run", new_count if history is not None else "n/a (first run)")],
        "funnel": {**funnel, **{f"raw from {b}": n for b, n in per_board_raw.most_common()}},
        "errors": errors,
    }
    xlsx = export_xlsx(exported, rejected, summary, outdir / f"{args.base}_{stamp}.xlsx", columns)
    csv_path = export_csv(exported, outdir / f"{args.base}_{stamp}.csv", columns)
    fr = export_csv(exported, outdir / f"{args.base}_{stamp}_french_excel.csv", columns, ";") if args.french_csv else None

    if not args.no_history:
        save_history(hist_path, seen, today)

    # ---- console report ----
    log.info("\n[+] Funnel: " + "  ->  ".join(f"{k}: {v}" for k, v in funnel.items()))
    if rejected:
        reasons = Counter(r["Reason"] for r in rejected)
        log.info("    Filtered out: " + "; ".join(f"{n} x {why}" for why, n in reasons.most_common())
                 + "  (see 'Filtered Out' sheet)")
    log.info("[+] Saved %d offers:", len(exported))
    log.info("    Excel: %s", xlsx.resolve())
    log.info("    CSV:   %s", csv_path.resolve())
    if fr:
        log.info("    CSV (FR): %s", fr.resolve())
    if first_run:
        log.info("[i] First run: history baseline created. Next run will flag NEW offers.")
    elif history is not None:
        log.info("[*] NEW since last run: %d", new_count)
    if errors:
        log.info("[!] Board warnings:")
        for k, v in errors.items():
            log.info("    - %s: %s", k, v)
    fresh = [r for r in exported if r["New"]] or [r for r in exported if r["Date Posted"] == today]
    if fresh:
        log.info("[*] %s:", "NEW offers" if any(r["New"] for r in fresh) else f"Posted today ({stamp})")
        for r in fresh[:8]:
            log.info("    - [%s] %s @ %s  <%s>", r["Job Board"], r["Job Title"], r["Company"], r["Domain"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
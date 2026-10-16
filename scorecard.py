"""
College Scorecard (US Dept. of Education) lookups + a transparent admissions-chance estimate.

The AI only PROPOSES school names. Everything numeric (acceptance rate, SAT range, net price,
graduation rate...) comes from the government dataset, and Safety/Target/Reach is computed here
from that data plus the student's own stats, so it can't be hallucinated.

Free API key: https://api.data.gov/signup/
"""
import difflib
import json
import math
import re
import sqlite3
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor

BASE_URL = "https://api.data.gov/ed/collegescorecard/v1/schools"
CACHE_TTL = 30 * 24 * 3600

CORE_FIELDS = [
    "id",
    "school.name",
    "school.city",
    "school.state",
    "school.school_url",
    "school.ownership",
    "latest.student.size",
    "latest.admissions.admission_rate.overall",
    "latest.cost.attendance.academic",
]
FIELD_GROUPS = {
    "sat": [
        "latest.admissions.sat_scores.25th_percentile.critical_reading",
        "latest.admissions.sat_scores.75th_percentile.critical_reading",
        "latest.admissions.sat_scores.25th_percentile.math",
        "latest.admissions.sat_scores.75th_percentile.math",
        "latest.admissions.sat_scores.average.overall",
    ],
    "act": ["latest.admissions.act_scores.midpoint.cumulative"],
    "completion": ["latest.completion.completion_rate_4yr_150nt"],
    "earnings": ["latest.earnings.10_yrs_after_entry.median"],
}
ALL_FIELDS = CORE_FIELDS + [f for g in FIELD_GROUPS.values() for f in g]

OWNERSHIP = {1: "Public", 2: "Private nonprofit", 3: "Private for-profit"}

# Official 2018 ACT -> SAT concordance
ACT_TO_SAT = {
    36: 1590, 35: 1540, 34: 1500, 33: 1460, 32: 1430, 31: 1400, 30: 1370, 29: 1340,
    28: 1310, 27: 1280, 26: 1240, 25: 1210, 24: 1180, 23: 1140, 22: 1110, 21: 1080,
    20: 1040, 19: 1010, 18: 970, 17: 930, 16: 890, 15: 850,
}


class ScorecardError(Exception):
    pass


class ScorecardAuthError(ScorecardError):
    pass


#student stats
def parse_test_score(text):
    """'SAT: 1540', 'ACT 34', '1450', 'N/A' -> SAT-equivalent int or None."""
    if not text:
        return None
    t = str(text)
    for m in re.finditer(r"(?<!\d)(\d{3,4})(?!\d)", t):
        v = int(m.group(1))
        before = t[max(0, m.start() - 6): m.start()].upper()
        if 400 <= v <= 1600 and "ACT" not in before:
            return v
    m = re.search(r"ACT\D{0,6}(\d{1,2})(?!\d)", t, re.I) or re.fullmatch(r"\s*(\d{2})\s*", t)
    if m:
        return ACT_TO_SAT.get(int(m.group(1)))
    return None


def parse_gpa(text):
    if not text:
        return None
    m = re.search(r"\d(?:\.\d+)?", str(text))
    if not m:
        return None
    v = float(m.group(0))
    return v if 0 < v <= 5.0 else None


# API plumbing
def _get_json(params, timeout=15):
    url = BASE_URL + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def pick(row, path):
    """Scorecard returns flat dotted keys when you pass `fields`; handle nested too."""
    if path in row:
        return row[path]
    cur = row
    for part in path.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return None
    return cur


def _num(row, path):
    v = pick(row, path)
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def normalize_school(row):
    r25 = _num(row, "latest.admissions.sat_scores.25th_percentile.critical_reading")
    r75 = _num(row, "latest.admissions.sat_scores.75th_percentile.critical_reading")
    m25 = _num(row, "latest.admissions.sat_scores.25th_percentile.math")
    m75 = _num(row, "latest.admissions.sat_scores.75th_percentile.math")
    sat_25 = round(r25 + m25) if r25 is not None and m25 is not None else None
    sat_75 = round(r75 + m75) if r75 is not None and m75 is not None else None
    sat_mid = _num(row, "latest.admissions.sat_scores.average.overall")
    if sat_mid is None and sat_25 and sat_75:
        sat_mid = (sat_25 + sat_75) / 2
    if sat_mid is None:
        act_mid = _num(row, "latest.admissions.act_scores.midpoint.cumulative")
        if act_mid:
            sat_mid = ACT_TO_SAT.get(int(round(act_mid)))

    url = pick(row, "school.school_url") or ""
    if url and not url.startswith("http"):
        url = "https://" + url.strip("/ ")

    return {
        "id": pick(row, "id"),
        "name": pick(row, "school.name") or "",
        "city": pick(row, "school.city"),
        "state": pick(row, "school.state"),
        "ownership": OWNERSHIP.get(pick(row, "school.ownership")),
        "url": url or None,
        "size": _num(row, "latest.student.size"),
        "admit_rate": _num(row, "latest.admissions.admission_rate.overall"),
        "sat_25": sat_25,
        "sat_75": sat_75,
        "sat_mid": round(sat_mid) if sat_mid else None,
        "net_price": _num(row, "latest.cost.attendance.academic"),  # Mapped to sticker price
        "grad_rate": _num(row, "latest.completion.completion_rate_4yr_150nt"),
        "earnings": _num(row, "latest.earnings.10_yrs_after_entry.median"),
    }


# name matching
def _norm(s):
    s = (s or "").lower().replace("&", " and ").replace("-", " ")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\b(the|of|at)\b", " ", s)
    return " ".join(s.split())


def _best_match(wanted, rows):
    target = _norm(wanted)
    best, best_score = None, 0.0
    for row in rows:
        cand = _norm(pick(row, "school.name") or "")
        if not cand:
            continue
        score = difflib.SequenceMatcher(None, target, cand).ratio()
        if cand.startswith(target) or target.startswith(cand):
            score = max(score, 0.93)
        size = _num(row, "latest.student.size") or 0
        score += min(size, 60000) / 60000 * 0.02  # tie-break toward the main campus
        if score > best_score:
            best, best_score = row, score
    return best if best_score >= 0.80 else None


def _variants(name):
    v = [name, name.replace(", ", "-"), re.sub(r"^The\s+", "", name)]
    return list(dict.fromkeys(x for x in v if x))


# cache
def _cache_conn(db_path):
    conn = sqlite3.connect(db_path, timeout=10)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS school_cache (name TEXT PRIMARY KEY, data TEXT, fetched REAL)"
    )
    return conn


def _cache_get(db_path, name):
    try:
        conn = _cache_conn(db_path)
        row = conn.execute("SELECT data, fetched FROM school_cache WHERE name = ?", (name.lower(),)).fetchone()
        conn.close()
        if row and time.time() - row[1] < CACHE_TTL:
            return json.loads(row[0])
    except Exception:
        pass
    return None


def _cache_put(db_path, name, data):
    try:
        conn = _cache_conn(db_path)
        conn.execute(
            "INSERT OR REPLACE INTO school_cache (name, data, fetched) VALUES (?, ?, ?)",
            (name.lower(), json.dumps(data), time.time()),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


# lookup
_ALL_FIELDS_OK = None


def _search(variant, api_key, fields):
    data = _get_json(
        {"api_key": api_key, "school.name": variant, "fields": ",".join(fields), "per_page": 10}
    )
    return data.get("results", [])


def _fetch_rows(variant, api_key):
    """Try the full field list, then the core fields. A 5xx or 400 only means that request failed."""
    global _ALL_FIELDS_OK
    attempts = []
    if _ALL_FIELDS_OK is not False:
        attempts.append((ALL_FIELDS, False))
    attempts.append((CORE_FIELDS, True))
    last_code = None
    for fields, needs_groups in attempts:
        for try_num in range(2):
            try:
                rows = _search(variant, api_key, fields)
                if fields is ALL_FIELDS:
                    _ALL_FIELDS_OK = True
                return rows, needs_groups
            except urllib.error.HTTPError as e:
                last_code = e.code
                if e.code in (401, 403):
                    raise ScorecardAuthError("Scorecard rejected the API key (check SCORECARD_API_KEY).")
                if e.code == 400 and fields is ALL_FIELDS:
                    _ALL_FIELDS_OK = False 
                    break
                if e.code >= 500 and try_num == 0:
                    time.sleep(2)
                    continue
                break
            except (urllib.error.URLError, TimeoutError) as e:
                last_code = str(e)
                break
    raise ScorecardError(f"Scorecard request failed ({last_code})")


def lookup_school(name, api_key, db_path):
    """Return a normalized dict of real stats for `name`, or None if not found."""
    cached = _cache_get(db_path, name)
    if cached:
        return cached

    best, needs_groups = None, False
    for variant in _variants(name):
        try:
            rows, needs_groups = _fetch_rows(variant, api_key)
        except ScorecardAuthError:
            raise
        except ScorecardError as e:
            print(f"[Scorecard] {variant}: {e}")
            continue
        best = _best_match(name, rows)
        if best:
            break
    if not best:
        return None

    merged = dict(best)
    if needs_groups:
        for fields in FIELD_GROUPS.values():
            try:
                extra = _get_json({"api_key": api_key, "id": pick(best, "id"), "fields": ",".join(["id"] + fields)})
                for r in extra.get("results", []):
                    merged.update(r)
            except Exception:
                continue

    info = normalize_school(merged)
    _cache_put(db_path, name, info)
    return info


# chance estimate
def _phi(z):
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))


def estimate_chance(school, sat=None, gpa=None):
    """
    Rough, transparent estimate: start from the school's overall acceptance rate, then scale it by
    where the student's test score sits among admitted students (or by GPA if no score yet).
    NOT a prediction: real admissions depend on essays, activities, and context we can't see.
    """
    ar = school.get("admit_rate")
    if ar is None:
        return None, "no admissions data"

    mult, basis = 1.0, "selectivity only (add test scores or GPA for a better estimate)"
    if sat and school.get("sat_mid"):
        if school.get("sat_25") and school.get("sat_75"):
            sd = max((school["sat_75"] - school["sat_25"]) / 1.35, 40)
        else:
            sd = 120
        pct = _phi((sat - school["sat_mid"]) / sd)
        mult = 0.25 + 1.5 * pct
        basis = "your test score vs. admitted students"
    elif gpa:
        mult = 1.2 if gpa >= 3.9 else 1.0 if gpa >= 3.7 else 0.8 if gpa >= 3.3 else 0.6
        basis = "your GPA and the school's selectivity"

    chance = ar * mult
    if ar < 0.15:
        chance = min(chance, 0.30)
    elif ar < 0.30:
        chance = min(chance, 0.65)
    return max(0.02, min(0.95, chance)), basis


def categorize(chance):
    if chance >= 0.70:
        return "Safety"
    if chance >= 0.35:
        return "Target"
    return "Reach"


# enrichment
COPY_KEYS = ["city", "state", "ownership", "url", "size", "admit_rate", "sat_25", "sat_75", "sat_mid",
             "net_price", "grad_rate", "earnings"]


def enrich_colleges(colleges, api_key, db_path, sat=None, gpa=None, workers=5):
    def work(c):
        try:
            return lookup_school(c["name"], api_key, db_path)
        except ScorecardAuthError:
            raise
        except Exception as e:
            print(f"[Scorecard] lookup failed for {c['name']}: {e}")
            return None

    with ThreadPoolExecutor(max_workers=workers) as ex:
        infos = list(ex.map(work, colleges))

    out, seen = [], set()
    for c, info in zip(colleges, infos):
        item = dict(c)
        item.update({"data_source": "ai_only", "match_score": None, "basis": "unverified (not found in College Scorecard)"})
        if info:
            if info["id"] in seen:
                continue
            seen.add(info["id"])
            item.update({k: info.get(k) for k in COPY_KEYS})
            item["scorecard_id"] = info["id"]
            item["scorecard_name"] = info["name"]
            item["data_source"] = "scorecard"
            chance, basis = estimate_chance(info, sat, gpa)
            item["basis"] = basis
            if chance is not None:
                item["match_score"] = round(chance * 100)
                item["category"] = categorize(item["match_score"] / 100)  # category matches the number shown
        out.append(item)

    order = {"Reach": 0, "Target": 1, "Safety": 2}
    out.sort(key=lambda x: (order.get(x.get("category"), 1), x["match_score"] if x["match_score"] is not None else 50))
    return out

import csv
import os
import re
import threading
import time
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore")
requests.packages.urllib3.disable_warnings()

INPUT = r"C:\Users\vinutshreedhar.n\Downloads\company careers\new_pages.xlsx"
OUTPUT = r"C:\Users\vinutshreedhar.n\Downloads\company careers\output_careers.xlsx"
TEST_OUTPUT = r"C:\Users\vinutshreedhar.n\Downloads\company careers\output_careers_test50.xlsx"
CSV_FIELDS = ["DUNS_NUMBER", "Web Domain", "Company URL", "Careers URL", "Status"]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# English + major EU language paths
CAREER_PATHS = [
    "/careers", "/careers/", "/career", "/jobs", "/jobs/", "/job",
    "/open-positions", "/open-positions/", "/vacancies", "/vacancies/",
    "/opportunities", "/opportunities/", "/join-us", "/join-us/",
    "/join-our-team", "/join-our-team/", "/work-with-us", "/work-with-us/",
    "/work-for-us", "/work-for-us/", "/we-are-hiring", "/hiring",
    "/about/careers", "/about/careers/", "/company/careers",
    "/company/careers/", "/about/jobs", "/company/jobs",
    "/en/careers", "/en/jobs", "/us/careers", "/careers.html",
    "/careers.php", "/jobs.html", "/employment", "/life-at",
    # Multi-language
    "/karriere", "/karriere/", "/stellenangebote", "/offene-stellen",
    "/carrieres", "/carrieres/", "/recrutement", "/nos-offres",
    "/empleo", "/empleo/", "/trabaja-con-nosotros", "/ofertas-de-empleo",
    "/vacatures", "/vacatures/", "/werken-bij",
    "/vagas", "/vagas/", "/trabalhe-conosco",
    "/carriera", "/lavora-con-noi", "/posizioni-aperte",
    "/lediga-jobb", "/jobba-hos-oss",
    "/kariera", "/oferty-pracy",
]

CAREER_LINK_RE = re.compile(
    r"\b(career|careers|job|jobs|job-opening|job-posting|vacan|open-position|"
    r"join[-_ ]?us|join[-_ ]?our[-_ ]?team|work[-_ ]?(with|for|at)[-_ ]?us|"
    r"we[-_ ]?are[-_ ]?hiring|hiring|opportunit|talent|recruit|"
    r"internship[-_ ]?(program|job|opportunit|position|career)|"
    r"student[-_ ]?(job|employ|opportunit|position|career)|"
    # EU languages
    r"karriere|stellenangebote|offene[-_ ]?stellen|"
    r"carri[eè]res?|recrutement|offres?[-_ ]?d[-_ ]?emploi|"
    r"empleo|vacatures|werken[-_ ]?bij|vagas|carriera|"
    r"lavora[-_ ]?con[-_ ]?noi|lediga[-_ ]?jobb|oferty[-_ ]?pracy)\b",
    re.I,
)

BAD_EXT = re.compile(
    r"\.(pdf|jpg|jpeg|png|gif|svg|css|js|zip|docx?|xlsx?|mp4|webp|ico|woff2?)(\?|$)",
    re.I,
)

# External job boards / social — never accepted as careers URL (same-domain rule)
EXTERNAL_HOSTS = (
    "linkedin.com", "facebook.com", "x.com", "twitter.com", "instagram.com",
    "youtube.com", "glassdoor.", "indeed.", "ziprecruiter.", "monster.com",
    "greenhouse.io", "lever.co", "ashbyhq.com", "workable.com", "smartrecruiters.com",
    "bamboohr.com", "workday.com", "myworkdayjobs.com", "icims.com", "jobvite.com",
    "recruitee.com", "personio.", "breezy.hr", "teamtailor.com", "jazz.co",
    "google.com", "bing.com", "yahoo.com",
)

# Subdomains to try on company domain
SUBDOMAIN_PREFIXES = ["careers", "jobs"]

MAX_PATH_PROBES = 20
HOMEPAGE_LINK_LIMIT = 10
OVERALL_TIMEOUT = 50  # hard per-company budget (seconds)
DEFAULT_WORKERS = 40

# Reject event/news-style pages that merely mention "career" (false positives)
EVENT_RE = re.compile(
    r"(career[-_]?(connect|fair|day|week|month|night|expo|event|forum|series|talk|chat|webinar)|"
    r"job[-_]?(fair|day|event|workshop)|hiring[-_]?event|college[-_]?(career|recruiting)|"
    r"(spring|fall|winter|summer)[-_ ]?(career|job)[-_ ]?(fair|day|event)|"
    r"20\d\d[-_ ]?(career|job)|"
    # Housing / conference / student-life pages that mention internships but aren't job boards
    r"(internship|student)[-_]?(housing|dorm|residence|life|services?|portal|housing|accommodation)|"
    r"conference[-_]?services?|event[-_]?services?|"
    r"career[-_]?services?(?!\s*(job|position|opening|employ))|"
    r"/about/conference|/housing|/residence[-_]?life|/student[-_]?life)",
    re.I,
)

# Paths that typically hold the actual job listings (deeper than careers landing)
JOBS_LIST_PATH_RE = re.compile(
    r"(open[-_]?(position|role|job|vacanc)|"
    r"job[-_]?(opportunit|opening|vacanc|listing|search|board|available)|"
    r"current[-_]?(opening|vacanc|job)|"
    r"search[-_]?job|view[-_]?job|see[-_]?(our[-_]?)?(job|position|opening)|"
    r"available[-_]?(position|job|role)|"
    r"browse[-_]?job|explore[-_]?(job|role|position|opportunit)|"
    r"/jobs?/|/vacanc|/offre|/stelle|/empleo|/vaga|/posizion|"
    r"(career|job)s?/(search|list|openings?|opportunities?|positions?)|"
    r"job[-_]?opportunit|careers?/open|careers?/job|"
    r"role[s]?/|position[s]?/| openings?/|"
    # Bare openings path, or /search on a careers-ish host/path context handled by caller
    r"openings?(/|$)|"
    r"/search(/|$)|"
    r"/find[-_]?job|/job[-_]?search|/job[-_]?list)",
    re.I,
)

# Anchor/button text that points to the real listings page.
# Covers wide real-world CTA labels (View Jobs, Explore Careers, Join Our Team,
# Life at X, Talent Community, Apply Now, department/location filters, etc.)
JOBS_LIST_TEXT_RE = re.compile(
    r"(?:"
    # --- Openings / positions / roles ---
    r"open[-_ ]?(position|role|job|vacanc|ing)s?|"
    r"view[-_ ]?(job|position|role|opening)s?|"
    r"view[-_ ]?(our|all|current|available|latest|new|featured)[-_ ]?(opening|job|position|role|vacanc|opportunit)s?|"
    r"search[-_ ]?(job|opening|position|role|vacanc|opportunit|career)s?|"
    r"see[-_ ]?(all[-_ ]?)?(job|position|opening|available|role|vacanc|opportunit)s?|"
    r"see[-_ ]?(our|all|current|available|open)[-_ ]?(opening|job|position|role|vacanc|opportunit)s?|"
    r"current[-_ ]?(opening|vacanc|job|opportunit|position|role)s?|"
    r"available[-_ ]?(job|role|position|opening|vacanc|opportunit)s?|"
    r"latest[-_ ]?(job|opening|role|position|vacanc)s?|"
    r"new[-_ ]?(job|opening|role|position|opportunit|vacanc)s?|"
    r"featured[-_ ]?(job|role|position|opening)s?|"
    r"recently[-_ ]?posted[-_ ]?(job|role|position|opening)s?|"
    r"all[-_ ]?(job|opening|role|position|vacanc|opportunit)s?|"
    r"browse[-_ ]?(job|opening|position|role|vacanc|opportunit|career)s?|"
    r"explore[-_ ]?(job|role|position|opening|opportunit|career|available|open)s?|"
    r"discover[-_ ]?(job|role|position|opening|opportunit|career|vacanc)s?|"
    r"view[-_ ]?opportunit|discover[-_ ]?opportunit|"
    r"find[-_ ]?(a[-_ ]?)?(job|role|position|opening|next[-_ ]?role)|"
    r"find[-_ ]?(your[-_ ]?)?(role|future|position|job)|"
    r"job[-_ ]?(opportunit|opening|vacanc|listing|search|available|portal|board|detail)s?|"
    r"career[-_ ]?(opportunit|search|portal|detail)s?|"
    r"employment[-_ ]?(opportunit|s)?|"
    r"work[-_ ]?(opportunit|s)?|"
    r"recruitment([-_ ]?opportunit)?s?|"
    r"vacanc(y|ies)|"
    r"position[s]?|"
    r"\broles?\b|"
    # --- Careers exploration / company pages ---
    r"(explore|discover|view|browse|search|go[-_ ]?to|visit)[-_ ]?careers?|"
    r"careers?\s+at\b|"
    r"life\s+at\b|"
    r"careers?|"
    r"career[-_ ]?portal|job[-_ ]?portal|applicant[-_ ]?portal|candidate[-_ ]?portal|"
    r"external[-_ ]?careers?|visit[-_ ]?(careers?|job)[-_ ]?site|"
    r"go[-_ ]?to[-_ ]?(careers?|jobs?)|"
    # --- Join / work with us ---
    r"join[-_ ]?(our[-_ ]?)?(team|us|talent|network|community)|"
    r"work[-_ ]?(with|for|at)[-_ ]?us|"
    r"build[-_ ]?your[-_ ]?career|start[-_ ]?your[-_ ]?career|grow[-_ ]?with[-_ ]?us|"
    r"make[-_ ]?an[-_ ]?impact|your[-_ ]?career[-_ ]?starts[-_ ]?here|"
    r"meet[-_ ]?our[-_ ]?team|"
    # --- Hiring ---
    r"we['’]?re[-_ ]?hiring|we[-_ ]?are[-_ ]?hiring|now[-_ ]?hiring|hiring[-_ ]?now|"
    r"teams?[-_ ]?hiring|departments?[-_ ]?hiring|"
    # --- Apply ---
    r"apply[-_ ]?(now|here|today|online|for[-_ ]?(this[-_ ]?)?job|for[-_ ]?position)|"
    r"submit[-_ ]?(application|your[-_ ]?resume|application)|"
    r"view[-_ ]?(job[-_ ]?)?details?|read[-_ ]?more|learn[-_ ]?more|"
    r"i[-_ ]?m[-_ ]?interested|"
    r"upload[-_ ]?resume|register[-_ ]?your[-_ ]?interest|"
    r"(sign[-_ ]?up|create|get|start)[-_ ]?for[-_ ]?job[-_ ]?alerts?|"
    r"job[-_ ]?alerts?|"
    # --- Talent community / network ---
    r"talent[-_ ]?community|talent[-_ ]?network|"
    r"connect[-_ ]?with[-_ ]?recruiters?|"
    # --- Filters / by location / department / team ---
    r"jobs?[-_ ]?by[-_ ]?(location|department|team)|"
    r"locations?\\b|"
    r"(engineering|software[-_ ]?engineering|technology|tech|backend|developer|it)[-_ ]?(jobs?|careers?|opportunit|roles?)|"
    r"(internship|graduate|early[-_ ]?careers?|university[-_ ]?recruiting|campus[-_ ]?recruiting|student)[-_ ]?(opportunit|job|program|s)?|"
    r"entry[-_ ]?level[-_ ]?jobs?|"
    r"experienced[-_ ]?professionals?|professional[-_ ]?careers?|"
    r"remote[-_ ]?jobs?|work[-_ ]?from[-_ ]?home[-_ ]?jobs?|"
    r"contract[-_ ]?opportunit|full[-_ ]?time[-_ ]?jobs?|part[-_ ]?time[-_ ]?jobs?|"
    r"join[-_ ]?our[-_ ]?(engineering|tech)[-_ ]?team|"
    # --- Bare common labels ---
    r"^\s*(jobs?|careers?|positions?|roles?|openings?|vacanc\w*|opportunit\w*|internships?|graduates?)\s*$|"
    # --- EU languages ---
    r"offene[-_ ]?stellen|stellenangebot|"
    r"carri[eè]res?|emploi|offres?[-_ ]?d[-_ ]?emploi|recrutement|nos[-_ ]?offres|"
    r"empleo|ofertas?[-_ ]?de[-_ ]?empleo|trabaja[-_ ]?con[-_ ]?nosotros|"
    r"vacatures|werken[-_ ]?bij|"
    r"vagas|trabalhe[-_ ]?conosco|"
    r"carriera|lavora[-_ ]?con[-_ ]?noi|posizioni[-_ ]?aperte|"
    r"lediga[-_ ]?jobb|jobba[-_ ]?hos[-_ ]?oss|"
    r"kariera|oferty[-_ ]?pracy"
    r")",
    re.I,
)

# Generic CTA text — weaker signal, only used when path also looks jobs-related
GENERIC_CTA_TEXT_RE = re.compile(
    r"(check[-_ ]?(it|this|them)?[-_ ]?(out)?|"
    r"learn[-_ ]?more|read[-_ ]?more|see[-_ ]?more|"
    r"click[-_ ]?here|more[-_ ]?info|details?|"
    r"explore|discover|get[-_ ]?started|go|"
    r"continue|proceed|next|view[-_ ]?more)",
    re.I,
)

# Signals that a page actually lists jobs (not just careers marketing copy)
JOB_LISTING_SIGNALS_RE = re.compile(
    r"(apply now|apply online|job description|full[-_ ]?time|part[-_ ]?time|"
    r"remote|on[-_ ]?site|hybrid|department|location|requisition|"
    r"\b\d+\s+(open|job|position)|openings?\b|we.?re hiring|"
    r"equal opportunity|benefits?\b|salary|compensation)",
    re.I,
)

JOB_POSTING_JSONLD_RE = re.compile(r'"@type"\s*:\s*"JobPosting"', re.I)


def normalize_domain(d):
    d = str(d).strip()
    if not d or d.lower() in ("nan", "none", ""):
        return None
    d = re.sub(r"^https?://", "", d, flags=re.I).strip("/")
    d = d.split("/")[0].split("?")[0]
    if not d or "." not in d:
        return None
    return d.lower()


def ensure_url(base, path):
    if not base.startswith(("http://", "https://")):
        base = "https://" + base
    return base.rstrip("/") + path


def domain_matches(url, base_domain):
    """True if url host is base_domain or its subdomain (same site only)."""
    host = urlparse(url).netloc.lower().replace("www.", "")
    base = base_domain.lower().replace("www.", "")
    if not host or not base:
        return False
    return host == base or host.endswith("." + base)


def is_external(url, base_domain):
    if any(e in urlparse(url).netloc.lower() for e in EXTERNAL_HOSTS):
        return True
    return not domain_matches(url, base_domain)


# Thread-local Session for connection keep-alive (big speed win)
_thread_local = threading.local()


def get_session():
    s = getattr(_thread_local, "session", None)
    if s is None:
        s = requests.Session()
        s.headers.update(HEADERS)
        s.verify = False
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=10,
            pool_maxsize=20,
            max_retries=0,
        )
        s.mount("https://", adapter)
        s.mount("http://", adapter)
        _thread_local.session = s
    return s


def fetch(url, timeout=10):
    return get_session().get(url, timeout=timeout, allow_redirects=True)


def fetch_head(url, timeout=5):
    """Cheap existence check — only used to skip hopeless URLs fast."""
    try:
        return get_session().head(url, timeout=timeout, allow_redirects=True)
    except Exception:
        return None


def looks_like_careers(url, text=""):
    if BAD_EXT.search(url):
        return False
    if any(e in urlparse(url).netloc.lower() for e in EXTERNAL_HOSTS):
        return False
    path_q = urlparse(url).path + "?" + (urlparse(url).query or "")
    if EVENT_RE.search(path_q) or EVENT_RE.search(text):
        return False
    return bool(CAREER_LINK_RE.search(url + " " + text))


def page_is_careers(html, url):
    """Validate fetched page really is a careers/jobs page (reduce false positives)."""
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ").lower() if soup.title else ""
    h1s = " ".join(h.get_text(" ").lower() for h in soup.find_all("h1")[:3])
    meta = ""
    m = soup.find("meta", attrs={"name": "description"})
    if m and m.get("content"):
        meta = m["content"].lower()
    path = urlparse(url).path.lower()
    hay = title + " " + h1s + " " + meta + " " + path
    if EVENT_RE.search(hay):
        return False
    strong = re.search(
        r"(career|job|we are hiring|join (us|our team)|open (roles|positions)|"
        r"current (openings|vacancies)|karriere|carri[eè]res?|empleo|vacatures|vagas)",
        hay,
    )
    if strong:
        return True
    # Weak: body has multiple job-listing signals
    body = soup.get_text(" ").lower()
    signals = sum(k in body for k in ("apply now", "open position", "job description", "requirements", "what you'll do", "location", "full-time", "department"))
    return signals >= 3


def extract_career_links(page_url, html, base_domain):
    """Same-domain career links from a page's anchors."""
    soup = BeautifulSoup(html, "html.parser")
    found = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript:")):
            continue
        full = urljoin(page_url, href).split("#")[0]
        if not full.startswith(("http://", "https://")):
            continue
        if BAD_EXT.search(full):
            continue
        text = " ".join(a.get_text(" ").split())
        if not looks_like_careers(full, text):
            continue
        if is_external(full, base_domain):
            continue
        if full.rstrip("/") == page_url.rstrip("/"):
            continue
        found.append(full)
    seen, uniq = set(), []
    for u in found:
        k = u.rstrip("/").lower()
        if k not in seen:
            seen.add(k)
            uniq.append(u)
    return uniq


def probe_one_path(url, base_domain):
    """Probe a single URL. Returns (priority, result_url, how) or (999, None, None).

    Priority: lower = better (common paths first).
    """
    if is_external(url, base_domain):
        return (999, None, None)
    try:
        r = fetch(url, timeout=4)
    except Exception:
        return (998, None, None)
    if r.status_code == 200 and "html" in r.headers.get("Content-Type", "text/html"):
        if is_external(r.url, base_domain):
            return (999, None, None)
        if page_is_careers(r.text, r.url):
            return (0, r.url, "direct path")
        # 200 but not careers-like — still a soft signal, low priority
        return (50, None, None)
    if r.status_code in (403, 429):
        path_part = urlparse(url).path.lower()
        if re.search(r"(career|job|hiring|vacanc|karriere|carri[eè]re|empleo|vacature|vaga)", path_part):
            return (10, url, f"blocked (HTTP {r.status_code})")
    return (998, None, None)


def probe_direct_paths(home, base_domain, deadline=None):
    """Probe common careers paths + subdomains in parallel (company domain only)."""
    candidates = []
    # Path probes — most common first (priority = index)
    for i, path in enumerate(CAREER_PATHS[:MAX_PATH_PROBES]):
        candidates.append((i, ensure_url(home, path)))
    # Subdomain probes
    base_no_www = base_domain.replace("www.", "")
    idx = len(candidates)
    for prefix in SUBDOMAIN_PREFIXES:
        candidates.append((idx, f"https://{prefix}.{base_no_www}/careers"))
        candidates.append((idx + 1, f"https://{prefix}.{base_no_www}/"))
        idx += 2

    if not candidates:
        return None, None, []

    # Probe all in parallel — connection-pooled session per thread
    # Cap inner workers: too many concurrent sockets causes timeouts under load
    results = []
    with ThreadPoolExecutor(max_workers=min(8, len(candidates))) as pex:
        futs = {
            pex.submit(probe_one_path, url, base_domain): (prio, url)
            for prio, url in candidates
        }
        for fut in as_completed(futs):
            prio_base, url = futs[fut]
            try:
                _prio, rurl, how = fut.result()
            except Exception:
                continue
            if rurl:
                # Rank by original candidate order so /careers beats /careers.php etc.
                results.append((prio_base, rurl, how))

    if results:
        results.sort(key=lambda x: x[0])
        _p, rurl, how = results[0]
        return rurl, how, [f"{u}={h}" for _, u, h in results[:5]]
    return None, None, []


def scan_homepage_links(home, base_domain):
    """Scan homepage for same-domain career links; validate candidates in parallel."""
    home_page = None
    for scheme in ("https://", "http://"):
        try:
            r = fetch(scheme + base_domain if not scheme in home else home, timeout=6)
            if r.status_code < 400:
                home_page = r
                break
        except Exception:
            continue
    if not home_page:
        return None, "site unreachable"

    links = extract_career_links(home_page.url, home_page.text, base_domain)
    if not links:
        return None, "not found"

    # Validate candidates in parallel
    def _check(cand):
        try:
            r = fetch(cand, timeout=5)
            if r.status_code < 400 and not is_external(r.url, base_domain):
                if page_is_careers(r.text, r.url):
                    return r.url
        except Exception:
            pass
        return None

    with ThreadPoolExecutor(max_workers=min(8, len(links))) as pex:
        futs = [pex.submit(_check, c) for c in links[:HOMEPAGE_LINK_LIMIT]]
        for fut in as_completed(futs):
            hit = fut.result()
            if hit:
                # Cancel remaining? Keep simple — first completed hit wins
                return hit, "homepage link"

    # Fallback: return first candidate without deep validation
    return links[0], "homepage link (unverified)"


def looks_like_jobs_list_page(html, url):
    """Heuristic: does this page actually list individual job openings?"""
    if JOB_POSTING_JSONLD_RE.search(html):
        return True
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(" ").lower()
    signals = len(JOB_LISTING_SIGNALS_RE.findall(text))
    if signals >= 2:
        return True
    # Many links under /jobs/... or containing apply = likely listings
    job_links = 0
    for a in soup.find_all("a", href=True):
        href = a["href"].lower()
        atext = a.get_text(" ").lower()
        if re.search(r"(/job[s]?/|job[-_]?id=|/position/|/opening/|requisition)", href) or \
           (re.search(r"apply", atext) and len(atext) < 60):
            job_links += 1
    return job_links >= 3


def drill_to_jobs_page(careers_url, base_domain, max_hops=2, deadline=None):
    """From a careers landing page, follow links to the actual job-listings page.

    Returns (jobs_url, how) where how describes the path taken.
    Falls back to the original careers_url if no deeper listings page is found.
    """
    if not careers_url:
        return None, None

    best = careers_url
    best_how = "careers landing"
    visited = set()
    frontier = [careers_url]
    # Score candidates: higher = more likely the real job-listings page
    candidates = []  # (score, url, how)

    def score_candidate(url, text=""):
        """0 = not a listings CTA; higher = stronger match."""
        path_q = urlparse(url).path + "?" + (urlparse(url).query or "")
        if EVENT_RE.search(path_q) or EVENT_RE.search(text):
            return 0
        s = 0
        # Any button/link text that literally contains job-ish words = strong hit
        # e.g. "View 25 Jobs", "See Open Jobs", "XXXX Careers Jobs", "Job Openings"
        if re.search(r"\b(jobs?|openings?|vacanc\w*|positions?|roles?)\b", text, re.I):
            s += 4
        if JOBS_LIST_PATH_RE.search(path_q):
            s += 3
        if JOBS_LIST_TEXT_RE.search(text):
            s += 3
        # Generic CTA ("check it", "explore", "learn more") only counts with jobs-ish path
        elif GENERIC_CTA_TEXT_RE.search(text) and re.search(r"(job|position|role|opening|vacanc|opportunit|career|hire)", path_q, re.I):
            s += 1
        # Bare short labels like "Jobs", "Careers", "Openings" on nav-ish links
        elif re.fullmatch(r"(jobs?|careers?|positions?|roles?|openings?|vacanc\w*|opportunit\w*|hiring|join us)", text.strip(), re.I):
            s += 2
        if s and looks_like_url_jobsish(url):
            s += 1
        return s

    for hop in range(max_hops):
        if deadline and time.time() > deadline:
            break
        next_frontier = []
        for page_url in frontier:
            if deadline and time.time() > deadline:
                break
            try:
                r = fetch(page_url, timeout=6)
            except Exception:
                continue
            if r.status_code >= 400 or is_external(r.url, base_domain):
                continue
            final = r.url
            key = final.rstrip("/").lower()
            if key in visited:
                continue
            visited.add(key)

            # If this page itself looks like a real listings page and is deeper, prefer it
            path = urlparse(final).path
            if path.strip("/") and looks_like_jobs_list_page(r.text, final):
                base_score = 5 + len([p for p in path.split("/") if p])
                if JOBS_LIST_PATH_RE.search(path):
                    base_score += 3
                candidates.append((base_score, final, f"jobs list (hop {hop + 1})"))

            # Collect same-domain child links that look like the listings page
            soup = BeautifulSoup(r.text, "html.parser")

            # 1) <a href> links
            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                if not href or href.startswith(("#", "mailto:", "tel:", "javascript:")):
                    continue
                full = urljoin(final, href).split("#")[0]
                if not full.startswith(("http://", "https://")):
                    continue
                if BAD_EXT.search(full) or is_external(full, base_domain):
                    continue
                if full.rstrip("/").lower() in visited:
                    continue
                text = " ".join(a.get_text(" ").split())
                s = score_candidate(full, text)
                if s > 0:
                    depth = len([p for p in urlparse(full).path.split("/") if p])
                    candidates.append((s + depth * 0.1, full, f"link CTA hop {hop + 1}: {text[:40]!r}"))
                    next_frontier.append(full)
                    next_frontier = list(dict.fromkeys(next_frontier))[:12]

            # 2) <button> and other elements with data-href / onclick navigation containing job words
            for el in soup.find_all(["button", "a", "span", "div", "li"]):
                text = " ".join(el.get_text(" ").split())
                if not text or len(text) > 80:
                    continue
                if not re.search(r"\b(jobs?|openings?|vacanc\w*|positions?|roles?|careers?)\b", text, re.I):
                    continue
                # data-href / data-url / data-link attributes
                for attr in ("data-href", "data-url", "data-link", "data-target", "data-action-url"):
                    href = (el.get(attr) or "").strip()
                    if not href or not href.startswith(("http://", "https://", "/")):
                        continue
                    full = urljoin(final, href).split("#")[0]
                    if not full.startswith(("http://", "https://")):
                        continue
                    if BAD_EXT.search(full) or is_external(full, base_domain):
                        continue
                    if full.rstrip("/").lower() in visited:
                        continue
                    s = score_candidate(full, text)
                    if s > 0:
                        depth = len([p for p in urlparse(full).path.split("/") if p])
                        candidates.append((s + depth * 0.1, full, f"button attr hop {hop + 1}: {text[:40]!r}"))
                        next_frontier.append(full)
                # onclick="window.location='...'" patterns
                onclick = el.get("onclick") or ""
                m = re.search(r"(?:location(?:\.href)?|window\.open)\s*[=(]\s*['\"]([^'\"]+)['\"]", onclick)
                if m:
                    href = m.group(1).strip()
                    if href.startswith(("/", "http://", "https://")):
                        full = urljoin(final, href).split("#")[0]
                        if full.startswith(("http://", "https://")) and not is_external(full, base_domain) \
                           and not BAD_EXT.search(full) and full.rstrip("/").lower() not in visited:
                            s = score_candidate(full, text)
                            if s > 0:
                                candidates.append((s + 1, full, f"onclick hop {hop + 1}: {text[:40]!r}"))
                                next_frontier.append(full)

            next_frontier = list(dict.fromkeys(next_frontier))[:12]

        if not next_frontier:
            break
        frontier = next_frontier

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        top_score, top_url, top_how = candidates[0]
        # Prefer cleaner URL among near-equal scores (strip noisy query strings)
        clean = re.sub(r"[?&](createNewAlert|q|locationsearch|geolocation|optionsFacetsDD_\w+|searchby)=[^&]*", "", top_url)
        clean = re.sub(r"[?&]+$", "", clean).rstrip("?&")
        if clean.startswith("http") and clean.rstrip("/").lower() != top_url.rstrip("/").lower():
            # Only use cleaned URL if path still looks jobs-related
            if JOBS_LIST_PATH_RE.search(urlparse(clean).path) or looks_like_url_jobsish(clean):
                top_url = clean
        # Only replace careers landing if we found a meaningfully better candidate
        if top_url.rstrip("/").lower() != careers_url.rstrip("/").lower() and top_score >= 3:
            # Resolve redirects so we return the real final URL
            try:
                rr = fetch(top_url, timeout=6)
                if rr.status_code < 400 and not is_external(rr.url, base_domain):
                    top_url = rr.url
            except Exception:
                pass
            return top_url, top_how

    return best, best_how


def looks_like_url_jobsish(url):
    parsed = urlparse(url)
    # Host like careers.company.com or jobs.company.com
    host = parsed.netloc.lower()
    if re.search(r"^(careers?|jobs?|talent|hr)\.", host) or re.search(r"\.(careers?|jobs)\.", host):
        return True
    return bool(re.search(
        r"(job|position|role|opening|vacanc|opportunit|career|hire|apply|talent|search)",
        parsed.path,
        re.I,
    ))


def check_company(domain, timeout_budget=OVERALL_TIMEOUT):
    d = normalize_domain(domain)
    if not d:
        return None, "invalid domain"

    base_domain = d
    home = "https://" + d
    start = time.time()
    deadline = start + timeout_budget

    careers_url = None
    how = None

    # 1) Direct path + subdomain probes
    url, how1, _tried = probe_direct_paths(home, base_domain, deadline=deadline)
    if url:
        careers_url, how = url, how1

    # 2) Homepage link scan
    if not careers_url and time.time() < deadline:
        url, how2 = scan_homepage_links(home, base_domain)
        if url and not is_external(url, base_domain):
            careers_url, how = url, how2

    if not careers_url:
        return None, "not found"

    # 3) Drill down to the actual job-listings page
    if time.time() < deadline:
        jobs_url, jobs_how = drill_to_jobs_page(careers_url, base_domain, deadline=deadline)
        if jobs_url and not is_external(jobs_url, base_domain):
            # Prefer deeper jobs-list page; keep original how if drill found nothing better
            if jobs_url.rstrip("/").lower() != careers_url.rstrip("/").lower():
                return jobs_url, f"{how} -> {jobs_how}"
            return careers_url, how

    return careers_url, how


def load_done(csv_path):
    """Load already-processed domains from a previous run's CSV (for resume)."""
    done = {}
    if not os.path.isfile(csv_path):
        return done
    try:
        with open(csv_path, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                dom = normalize_domain(row.get("Web Domain", ""))
                if dom:
                    done[dom] = row
        print(f"Resume: found {len(done)} already-processed domains in {os.path.basename(csv_path)}")
    except Exception as e:
        print(f"Resume: could not read existing CSV ({e}) — starting fresh")
    return done


def append_csv_row(csv_path, row_dict, write_header=False):
    """Append one result row to CSV immediately (crash-safe)."""
    file_exists = os.path.isfile(csv_path) and os.path.getsize(csv_path) > 0
    with open(csv_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row_dict)
        f.flush()
        os.fsync(f.fileno())


def run(df, output_path, label, workers=DEFAULT_WORKERS, csv_path=None):
    if csv_path is None:
        csv_path = os.path.splitext(output_path)[0] + ".csv"

    print(f"\n=== {label}: {len(df)} rows | workers={workers} | live CSV: {os.path.basename(csv_path)} ===")

    # Resume: skip domains already in CSV
    done = load_done(csv_path)
    remaining = []
    resumed_rows = []
    for idx, row in df.iterrows():
        dom = normalize_domain(row["Web Domain"])
        if dom and dom in done:
            resumed_rows.append({
                "DUNS_NUMBER": row["DUNS_NUMBER"],
                "Web Domain": row["Web Domain"],
                "Company URL": done[dom].get("Company URL", ""),
                "Careers URL": done[dom].get("Careers URL", ""),
                "Status": done[dom].get("Status", "resumed"),
            })
        else:
            remaining.append((idx, row))

    print(f"Resume: {len(resumed_rows)} skipped (already done) | {len(remaining)} to process")
    if not remaining:
        out = pd.DataFrame(resumed_rows, columns=CSV_FIELDS)
        try:
            out.to_excel(output_path, index=False)
            print(f"Nothing to do. Excel saved: {output_path}")
        except PermissionError:
            print(f"Excel save skipped (file open?). CSV is complete: {csv_path}")
        return out

    results = {}
    start = time.time()
    # Live CSV writer: append each result the moment it finishes
    # Ensure header exists before any appends
    if not os.path.isfile(csv_path) or os.path.getsize(csv_path) == 0:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            csv.DictWriter(f, fieldnames=CSV_FIELDS).writeheader()

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(check_company, row["Web Domain"]): (idx, row) for idx, row in remaining}
        done_count = 0
        total = len(remaining)
        for fut in as_completed(futures):
            idx, row = futures[fut]
            try:
                careers_url, status = fut.result()
            except Exception as e:
                careers_url, status = None, f"error: {e}"
            # Final safety: same-domain only
            if careers_url:
                dom = normalize_domain(row["Web Domain"])
                if dom and is_external(careers_url, dom):
                    careers_url, status = None, "external rejected"
            careers_url = careers_url or ""
            company_url = "https://" + str(row["Web Domain"]).strip() if pd.notna(row["Web Domain"]) and str(row["Web Domain"]).strip() else ""
            record = {
                "DUNS_NUMBER": row["DUNS_NUMBER"],
                "Web Domain": row["Web Domain"],
                "Company URL": company_url,
                "Careers URL": careers_url,
                "Status": status,
            }
            results[idx] = record
            # Write immediately — do not wait for the whole run
            append_csv_row(csv_path, record)
            done_count += 1
            if done_count % 5 == 0 or done_count == total:
                found_so_far = sum(1 for r in results.values() if r["Careers URL"])
                print(f"Progress: {done_count}/{total}  found={found_so_far}  ({time.time()-start:.0f}s)", flush=True)

    # Merge resumed + new for final dataframe / excel
    all_rows = resumed_rows + [results[i] for i in sorted(results.keys())]
    out = pd.DataFrame(all_rows, columns=CSV_FIELDS)
    try:
        out.to_excel(output_path, index=False)
        print(f"\nDone in {time.time()-start:.0f}s | Excel: {output_path}")
    except PermissionError:
        print(f"\nDone in {time.time()-start:.0f}s | Excel save skipped (file open in Excel?)")
    found = (out["Careers URL"] != "").sum()
    print(f"Careers found: {found}/{len(out)} | Live CSV: {csv_path}")
    return out


def main():
    import sys
    limit = None
    workers = DEFAULT_WORKERS
    fresh = False
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    if "--workers" in sys.argv:
        workers = int(sys.argv[sys.argv.index("--workers") + 1])
    if "--fresh" in sys.argv:
        fresh = True

    df = pd.read_excel(INPUT)
    print(f"Loaded {len(df)} rows | columns: {list(df.columns)}")

    if limit:
        test_df = df.head(limit).copy()
        csv_path = os.path.splitext(TEST_OUTPUT)[0] + ".csv"
        if fresh and os.path.isfile(csv_path):
            os.remove(csv_path)
            print(f"--fresh: deleted {csv_path}")
        out = run(test_df, TEST_OUTPUT, f"TEST first {limit}", workers=workers, csv_path=csv_path)
        print("\n--- Sample results ---")
        print(out.head(min(limit, 20)).to_string(index=False))
    else:
        csv_path = os.path.splitext(OUTPUT)[0] + ".csv"
        if fresh and os.path.isfile(csv_path):
            os.remove(csv_path)
            print(f"--fresh: deleted {csv_path}")
        out = run(df, OUTPUT, "FULL 1000", workers=workers, csv_path=csv_path)
        print(out["Status"].value_counts().to_string())


if __name__ == "__main__":
    main()

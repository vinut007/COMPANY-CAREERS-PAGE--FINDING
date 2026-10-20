# Company Careers Page Finder

Finds the **actual job-listings page** for a list of company websites (same domain only) and writes results to Excel + CSV.

Input: `new_pages.xlsx` with columns `DUNS_NUMBER`, `Web Domain`
Output: `output_careers.xlsx` + `output_careers.csv` (live-appended) with columns `DUNS_NUMBER`, `Web Domain`, `Company URL`, `Careers URL`, `Status`

## Install

```
pip install -r requirements.txt
```

## Usage

```
# Full run (1000 rows) — resume-safe, live CSV
python find_careers.py --workers 40

# Test first 50 rows
python find_careers.py --limit 50 --workers 40

# Start over (delete existing CSV first)
python find_careers.py --limit 50 --workers 40 --fresh
```

| Flag | Default | Meaning |
|------|---------|---------|
| `--limit N` | all rows | only process first N rows (writes to `output_careers_test50.xlsx/.csv`) |
| `--workers N` | 40 | parallel company workers |
| `--fresh` | off | delete the CSV before starting (ignore resume data) |

## How it works

1. **Direct path probe (parallel)** — probes `/careers`, `/jobs`, `/open-positions`, `/vacancies`, `/join-us`, `/karriere`, `/carrieres`, `/empleo`… plus `careers.domain.com` and `jobs.domain.com` subdomains, all concurrently per company.

2. **Homepage link scan (parallel)** — fetches the homepage, extracts same-domain links whose URL or anchor text looks careers-related, validates candidates concurrently.

3. **Drill-down to real listings page** — from the careers landing page, follows CTA buttons/links to the page that actually lists jobs. Recognizes 140+ real-world CTA labels, e.g.:
   - View Jobs, Search Jobs, Browse Open Positions, Current Openings
   - Explore Careers, Join Our Team, Work With Us, Life at [Company]
   - Talent Community, Job Alerts, Apply Now, Jobs by Location/Department
   - Also reads `data-href`, `data-url`, and `onclick="location='...'"` button patterns

4. **Validation** — checks page title/H1/meta/JSON-LD (`JobPosting`) to confirm it is a real listings page, not an event/fair/housing page.

5. **Same-domain rule only** — external ATS boards (greenhouse.io, lever.co, workday, indeed, linkedin…) are **rejected**. Subdomains of the company (e.g. `careers.hmhco.com`) are accepted.

## Resume + live CSV

- Every result is **appended to the CSV immediately** (flushed to disk) — no waiting for program end.
- If the script crashes or is stopped, **re-run the same command** — already-processed domains are skipped automatically.
- Use `--fresh` to discard resume data and start over.
- CSV can be opened while the script is running.

## Output columns

| Column | Description |
|--------|-------------|
| `DUNS_NUMBER` | from input |
| `Web Domain` | from input |
| `Company URL` | `https://` + domain |
| `Careers URL` | resolved job-listings URL (empty if not found) |
| `Status` | how it was found / why not |

### Status values

| Status | Meaning |
|--------|---------|
| `direct path` | matched a common careers path directly |
| `homepage link` | found via homepage anchor scan |
| `jobs list (hop N)` | drilled down from landing page to real listings |
| `link CTA hop N: '...'` | followed a specific button/link text |
| `blocked (HTTP 403/429)` | path exists but bot-blocks scripts (URL is best guess) |
| `not found` | no same-domain careers page (often careers live on external ATS) |
| `external rejected` | careers found but on external domain — rejected by rule |

## Performance

- Connection keep-alive (`requests.Session` per thread)
- Parallel path probing within each company (8 inner threads)
- 40 outer workers by default
- 50-record benchmark: ~35–90s depending on network
- Full 1000 estimate: **~25–35 min** (varies with site responsiveness)

## Files

| File | Purpose |
|------|---------|
| `find_careers.py` | main script |
| `requirements.txt` | Python dependencies |
| `new_pages.xlsx` | input (1000 companies) |
| `output_careers.csv` / `.xlsx` | full-run output |
| `output_careers_test50.csv` / `.xlsx` | `--limit` test output |

# jenkins-fortify-scan-reporter

A Python script that pulls pipeline stage data from Jenkins and exports it as a formatted Excel report. Built primarily to track **Fortify SAST scan durations** across all CI/CD pipelines — but it works for any stage you want to monitor.

---

## Why I built this

When you're running Fortify scans across 50+ pipelines, you quickly lose visibility into which ones are running, how long they're taking, and which ones silently stopped including the scan stage. Clicking through Jenkins job-by-job is tedious and doesn't scale.

This script automates that. Point it at your Jenkins server, give it a date range, and it comes back with a clean Excel report showing every build, whether it had a Fortify stage, how long the scan took, and the failure reason if something went wrong — all in one place.

---

## Adapting it for a different stage

The script isn't hardcoded to Fortify. If you want to track a different stage — SonarQube, DAST scans, deployments, anything — just change one line in the config:

```python
# Default (Fortify)
STAGE_KEYWORDS = ["fortify scan"]

# For SonarQube
STAGE_KEYWORDS = ["sonarqube", "sonar scan"]

# For deployments
STAGE_KEYWORDS = ["deploy to prod", "deploy"]
```

The keyword match is case-insensitive and partial, so `"sonar"` would match a stage named `"Run SonarQube Analysis"`. Everything else in the script — the Excel output, the duration chart, the failure extraction — stays the same.

---

## What it produces

An `.xlsx` file with three sheets:

**All Builds** — every build that ran in the date range, with a Yes/No column showing whether the target stage was detected. Good for spotting pipelines where the stage was accidentally removed or skipped.

**Fortify Stages Only** — filtered to only the builds where a Fortify (or your target) stage ran. This is the main view most people will use.

**Summary** — key counts plus a doughnut chart showing how scan durations are spread across three bands: under 10 minutes, 10–15 minutes, and over 15 minutes. Useful for catching scans that are creeping up in duration over time.

The report also colour-codes rows — green where the stage ran, red for failed builds — and pulls the actual error message from the console log for anything that failed, so you don't have to go back to Jenkins to figure out what went wrong.

---

## Setup

```bash
pip install requests openpyxl
```

That's all you need. No other dependencies.

---

## Configuration

Open the script and update the `CONFIG` block at the top:

```python
JENKINS_URL   = "https://your-jenkins-server"   # No trailing slash
JENKINS_USER  = "your_username"
JENKINS_TOKEN = "your_api_token"

# Date range — defaults to yesterday. Change to a fixed string like "2024-06-01" if needed.
FROM_DATE = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
TO_DATE   = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

# How many jobs to fetch in parallel
MAX_WORKERS = 5

# The stage name keywords to look for
STAGE_KEYWORDS = ["fortify scan"]
```

For `JENKINS_TOKEN`: go to **Jenkins → Your Profile → Configure → API Token** and generate one. Don't use your password here.

**Timezone:** Build timestamps are displayed in UTC+3 by default. Change the `hours=3` offset in `fetch_build_data()` to match your timezone.

---

## Running it

```bash
python3 jenkins_fortify_report.py
```

Output file will be named `fortify_report_YYYY-MM-DD_to_YYYY-MM-DD.xlsx` in the current directory.

**Using cached data:** The script saves the raw Jenkins API response to `jenkins_raw_cache.json` after every run. If you want to change the Excel layout and regenerate without re-fetching, flip this at the bottom of the script:

```python
rows = collect_all_data(use_cache=True)
```

Delete the cache file to force a fresh fetch.

---

## A note on SSL

By default SSL verification is disabled (`verify=False`) since most internal Jenkins instances use self-signed certificates. If yours has a valid cert, set `s.verify = True` in the `make_session()` function.

For production use, move credentials out of the script and into environment variables:

```python
import os
JENKINS_USER  = os.environ["JENKINS_USER"]
JENKINS_TOKEN = os.environ["JENKINS_TOKEN"]
```

---

## Adjusting the duration thresholds

The Summary chart splits scans into three bands. The defaults are 10 and 15 minutes. If your scans typically run longer or shorter, change these constants in `write_excel()`:

```python
TEN_MIN_MS     = 10 * 60 * 1000   # 10 minutes
FIFTEEN_MIN_MS = 15 * 60 * 1000   # 15 minutes
```

---

## Requirements

- Python 3.8+
- `requests`
- `openpyxl`
- Jenkins with the Pipeline Stage View plugin (needed for the `/wfapi/describe` endpoint)

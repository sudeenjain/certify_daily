# Daily Certification Update Monitor

Checks 55 learning/certification platforms once a day (evening IST) and
emails you a digest of what changed. Runs entirely on GitHub Actions —
free, no server needed, and it keeps running even if your laptop is off.

## What it actually does (read this first)

- Most of these platforms **don't publish** a structured "new certification"
  feed. So for HTML sources, the script hashes the page text and flags it
  in the email if the page **changed since yesterday** — new cert, price
  change, or a banner ad, it can't tell the difference. Treat the email as
  a shortlist to go check, not a guaranteed cert announcement.
- A handful of sources (GitHub, freeCodeCamp) have real RSS feeds — those
  entries are shown as-is, newest post first, regardless of topic.
- Some sites may block automated requests (403 errors, especially
  Cloudflare-protected ones). Those will show up under "failed to check"
  in the email — check `sources.json` and fix/remove them as needed.

## Setup (10 minutes)

### 1. Create a new GitHub repo
Push this whole folder to a new **private** GitHub repo (private, since
state.json will accumulate over time — not sensitive, but no reason to
make it public).

```bash
cd cert-monitor
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/sudeenjain/cert-monitor.git
git branch -M main
git push -u origin main
```

### 2. Create a Gmail App Password
You cannot use your normal Gmail password for this — Google blocks it.

1. Go to https://myaccount.google.com/security
2. Turn on **2-Step Verification** if it isn't already on.
3. Go to https://myaccount.google.com/apppasswords
4. Create an app password (name it e.g. "cert-monitor"), copy the
   16-character code it gives you.

### 3. Add GitHub Secrets
In your repo: **Settings → Secrets and variables → Actions → New repository secret**

Add three secrets:
| Name | Value |
|---|---|
| `SMTP_USER` | your Gmail address (e.g. `you@gmail.com`) |
| `SMTP_PASS` | the 16-character app password from step 2 |
| `TO_EMAIL` | the email address you want the digest sent to |

### 4. Test it manually
Go to the **Actions** tab → "Daily Certification Update Check" →
**Run workflow** (this uses the `workflow_dispatch` trigger). Check the
run logs, and check your inbox.

### 5. Let it run
It's scheduled for **13:30 UTC = 7:00 PM IST** every day
(`.github/workflows/daily-cert-check.yml`). Change the cron line there if
you want a different time — cron times are always in UTC.

## Files

- `sources.json` — the 55 companies and the page/feed the script checks.
  Edit freely: fix broken URLs, add/remove companies.
- `monitor.py` — the script that does the checking and emails the digest.
- `state.json` — auto-created/updated by the script; stores yesterday's
  page hashes so it can detect changes. The workflow commits this back
  to the repo after every run — don't edit it by hand.
- `.github/workflows/daily-cert-check.yml` — the schedule.

## Note on GitHub Actions free tier
Public repos: unlimited free minutes. Private repos: 2,000 free minutes/month
on the free plan — this job takes well under a minute a day, so you're fine.

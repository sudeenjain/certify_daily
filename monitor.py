#!/usr/bin/env python3
"""
Daily certification/course update monitor.

For each source in sources.json:
  - RSS sources: pulls the latest entry title + link.
  - HTML sources: hashes the page content and compares to the last
    saved hash. If it changed since yesterday, it's flagged.

Sends one email digest at the end listing:
  - RSS: newest item found today
  - HTML: which pages changed since the last run (best-effort signal
    that "something updated" - not a guarantee it's a new certification,
    since most of these sites don't expose that distinction publicly)

State (previous hashes / seen RSS items) is persisted in state.json,
which this script rewrites each run. In GitHub Actions, the workflow
commits state.json back to the repo so history carries over between runs.
"""

import json
import os
import re
import smtplib
import sys
import hashlib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

import requests
import feedparser
from bs4 import BeautifulSoup

SOURCES_FILE = "sources.json"
STATE_FILE = "state.json"
REQUEST_TIMEOUT = 20
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; CertMonitorBot/1.0; +https://github.com/)"
}


def load_json(path, default):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def normalize_html_text(html):
    """Strip tags/scripts and collapse whitespace so minor markup churn
    (ads, timestamps, nonces) doesn't trigger false-positive diffs."""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "footer", "nav"]):
        tag.decompose()
    text = soup.get_text(separator=" ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def check_rss(source):
    try:
        feed = feedparser.parse(source["url"])
        if not feed.entries:
            return {"status": "error", "detail": "No entries found in feed"}
        latest = feed.entries[0]
        return {
            "status": "ok",
            "latest_title": latest.get("title", "(no title)"),
            "latest_link": latest.get("link", source["url"]),
            "published": latest.get("published", ""),
        }
    except Exception as e:
        return {"status": "error", "detail": str(e)}


def check_html(source, prev_hash):
    try:
        resp = requests.get(source["url"], headers=HEADERS, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        text = normalize_html_text(resp.text)
        digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
        changed = prev_hash is not None and digest != prev_hash
        first_check = prev_hash is None
        return {
            "status": "ok",
            "hash": digest,
            "changed": changed,
            "first_check": first_check,
        }
    except Exception as e:
        return {"status": "error", "detail": str(e)}


def build_digest(results, run_date):
    changed_html = [r for r in results if r["type"] == "html" and r["result"].get("changed")]
    first_check_html = [r for r in results if r["type"] == "html" and r["result"].get("first_check")]
    rss_updates = [r for r in results if r["type"] == "rss" and r["result"].get("status") == "ok"]
    errors = [r for r in results if r["result"].get("status") == "error"]

    lines = []
    lines.append(f"Certification / course update digest - {run_date}\n")

    lines.append("=== Pages that changed since yesterday ({}): ===".format(len(changed_html)))
    if changed_html:
        for r in changed_html:
            lines.append(f"  - {r['name']}: {r['url']}")
    else:
        lines.append("  (none detected)")
    lines.append("")

    lines.append("=== Latest RSS items ({}): ===".format(len(rss_updates)))
    for r in rss_updates:
        res = r["result"]
        lines.append(f"  - {r['name']}: {res['latest_title']}")
        lines.append(f"      {res['latest_link']}")
    lines.append("")

    if first_check_html:
        lines.append(f"=== First-time baseline saved for {len(first_check_html)} page(s) (no diff yet, will report from tomorrow): ===")
        for r in first_check_html:
            lines.append(f"  - {r['name']}")
        lines.append("")

    if errors:
        lines.append(f"=== Sources that failed to check ({len(errors)}): ===")
        for r in errors:
            lines.append(f"  - {r['name']}: {r['result'].get('detail','unknown error')}")
        lines.append("")

    lines.append(
        "Note: HTML-page checks flag ANY content change on that page (news, "
        "prices, banners, etc.) as a best-effort signal - most of these sites "
        "don't publish a structured 'new certification' feed, so this isn't a "
        "guarantee of a new cert specifically. RSS sources show the newest post "
        "regardless of topic. Treat this as a shortlist to go check, not a final answer."
    )
    return "\n".join(lines)


def send_email(subject, body):
    def clean(val):
        if val is None:
            return val
        # Strip regular whitespace and non-breaking spaces (\xa0), which
        # commonly get copied in from Google's app-password display page.
        return val.replace("\xa0", "").replace(" ", "").strip()

    smtp_user = os.environ.get("SMTP_USER", "").strip() or None
    smtp_pass = clean(os.environ.get("SMTP_PASS")) or None
    to_addr = os.environ.get("TO_EMAIL", smtp_user)
    if to_addr:
        to_addr = to_addr.strip()

    if not smtp_user or not smtp_pass:
        print("SMTP_USER / SMTP_PASS not set - skipping email send, printing digest instead:\n")
        print(body)
        return

    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, [to_addr], msg.as_string())
    print(f"Email sent to {to_addr}")


def main():
    sources = load_json(SOURCES_FILE, [])
    state = load_json(STATE_FILE, {})
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M")

    results = []
    for source in sources:
        name = source["name"]
        stype = source["type"]
        print(f"Checking: {name} ({stype})...")

        if stype == "rss":
            res = check_rss(source)
        else:
            prev_hash = state.get(name, {}).get("hash")
            res = check_html(source, prev_hash)
            if res.get("status") == "ok":
                state[name] = {"hash": res["hash"], "last_checked": run_date}

        results.append({"name": name, "url": source["url"], "type": stype, "result": res})

    save_json(STATE_FILE, state)

    digest = build_digest(results, run_date)
    subject = f"Daily Certification Update Digest - {datetime.now().strftime('%d %b %Y')}"
    send_email(subject, digest)


if __name__ == "__main__":
    main()

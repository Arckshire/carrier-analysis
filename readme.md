# Carrier Analysis Tool

A Streamlit app for analyzing carrier tracking and milestone completeness from shipment exports.

## What it does

Walks you through a structured analysis of one carrier's shipment data:

1. **Upload** an Excel or CSV export
2. **Phase 1 — Tracking Analysis:** see the tracked/untracked split, surface location and tracking-method patterns in untracked shipments, and log issues with evidence BOLs and screenshots
3. **Phase 2 — Milestone Completeness:** filtered to tracked shipments only. Shows individual milestone presence rates plus all 15 missing-pattern buckets, ranked by count. Click into any bucket for examples and pattern hints. Log issues as you go.
4. **Export:** download a `.docx` one-pager with all logged issues (BOLs and screenshots embedded), plus a reusable email-drafting prompt to feed Claude in a separate chat.

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push `app.py` and `requirements.txt` to a GitHub repo (public or private).
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click "New app", select the repo, set main file to `app.py`, and deploy.

## Expected columns

The app fuzzy-matches column names (case-insensitive, handles different spacings/punctuation), so column order doesn't matter. It looks for:

- **Required:** `tracked`, `bill of lading`
- **For pattern detection:** `pickup name`, `drop-off name`, `tracking method`, `active equipment id`, `pickup city state`, `drop-off city state`
- **For milestone analysis:** `pickup arrival utc timestamp raw`, `pickup departure utc timestamp raw`, `drop-off arrival utc timestamp raw`, `drop-off departure utc timestamp raw`
- **For metadata:** `tenant name`, `carrier name`

When you upload a file, the app shows a "Detected column mapping" panel so you can verify what was matched.

## Notes

- One tenant + one carrier per file is the expected input. If multiple carriers are detected, you'll be prompted to pick one.
- **Session state caveat:** Streamlit loses everything if you refresh the tab. Use the "Save progress" button in the sidebar to download a JSON snapshot of your logged issues at any point. Screenshots are not included in the snapshot — only text — so re-attach them if you restore.
- Pattern detection uses conservative thresholds (≥30% concentration of failures + meaningfully higher failure rate than baseline). Adjust the constants at the top of `app.py` if the defaults are too strict or too loose for your data.

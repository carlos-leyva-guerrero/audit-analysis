# audit-analysis

Processes Microsoft 365 audit log CSV exports (from audit.microsoft.com) and generates per-user PDF reports for HR security review.

## Quick Start

```powershell
py -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python generate_reports.py
```

Or use the Jupyter notebook (`audit_report.ipynb`) for interactive exploration.

## Input

Place Microsoft 365 Unified Audit Log CSV exports in `input/`. Any number of CSVs are supported and will be combined automatically.

## Output

One PDF per user in `output/`, containing:

- **Executive Summary** with composite risk score (0–100) and severity badge
- **Operation Category Counts** — bar chart + table (Delete, Download, Share, Modify, Upload, etc.)
- **Per-Site Breakdown** — heatmap + table showing operations by SharePoint site (top 12)
- **Daily Activity Timeline** — bar chart with spike detection threshold
- **Hourly Activity Distribution** — highlights off-hours (before 07:00 / after 20:00 UTC)
- **Last 7 Days vs Prior Period** — comparison table to detect end-of-tenure anomalies
- **Top Files in Sensitive Operations** — files involved in delete/download/share
- **Access Context** — client IPs, platforms, geolocations
- **Risk Indicators** — heuristic flags for bulk deletions, mass downloads, excessive sharing, off-hours activity, activity spikes/drops

## Risk Indicators

| Indicator | Trigger |
|---|---|
| Bulk deletions | >50 delete ops (HIGH) or >10 (MEDIUM) |
| Mass manual downloads | >50 manual download ops (HIGH) or >20 (MEDIUM) |
| Very high sync volume | >500 auto-sync ops/day |
| Excessive sharing | >30 share ops (HIGH) or >10 (MEDIUM) |
| Off-hours activity | >15% outside 07:00–20:00 UTC (MEDIUM) |
| Last-week activity spike | >2× daily average vs prior (HIGH) |
| Sudden activity drop | <20% of prior daily average in final week |
| Spike days | Days with >3× overall daily average |
| Login failures | ≥5 failed login attempts |
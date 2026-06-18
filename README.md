# Home Leads Funnel — Automated Dashboard

Upload an updated **Home Leads Funnel** workbook and instantly get a refined visual
funnel summary plus time-bucketed insights (weekly → annual) — no more manual rollup.

## What it does
- Reads **every channel sheet** automatically (ignores the manual `Summary` sheet, and
  is robust to channels being added/removed — any sheet with `Deal Status` + `Deal Date`
  is treated as channel data).
- Recreates the per-channel funnel **automatically** (Total Leads → Deals Done / Rejected /
  Pending / Awaiting Confirmation + conversion %), fixing the off-by-one counting in the
  hand-built Summary.
- Charts: overall funnel, outcome split, **funnel by channel**, **leads & conversion over
  time** with a time-bucket selector, top rejection reasons, category & brand mix, and the
  inspection **regrade** economics (online quote vs. physical inspection).
- Filters: channel, deal-date range, and time bucket
  (**Weekly · Bi-weekly · Monthly · Quarterly · Semi-annual · Annual**).
- One-click **Excel export** of the computed summary.

## Run it
```bash
pip install -r requirements.txt
streamlit run app.py
```
or just double-click **`run.bat`** on Windows. Your browser opens at
`http://localhost:8501`; drag in the `.xlsx` and the dashboard rebuilds itself.

## Files
| File | Purpose |
|------|---------|
| `app.py` | Streamlit UI (charts, KPIs, filters, export) |
| `data.py` | Data engine: load, clean, funnel mapping, time-bucketing (no Streamlit deps) |
| `requirements.txt` | Dependencies |
| `run.bat` | Windows launcher |

## Assumptions (from the source file)
- Time axis = **`Deal Date`** (parsed as `DD/MM/YYYY`).
- Funnel status from **`Deal Status`** (`Done` / `Rejected` / `Not Completed` /
  `Waiting For Customer`); `Not Completed` is split into **Pending** vs
  **Awaiting Confirmation** using `Terminal Status`.
- `Reject Reason` / `Rejection Reason` (duplicate columns) are coalesced.

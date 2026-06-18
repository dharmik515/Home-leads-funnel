"""
Data engine for the Home Leads Funnel dashboard.

Responsibilities:
  - Read every channel sheet from an uploaded/loaded workbook (any sheet that
    looks like a deal sheet -- robust to channels being added or removed).
  - Harmonise columns, parse dates and money fields.
  - Derive a single clean funnel status per deal.
  - Provide time-bucketing (weekly -> annual) and summary aggregations.

Nothing here is Streamlit-specific so it can be reused / unit-tested.
"""
from __future__ import annotations

import io
import pandas as pd
import numpy as np

# --- Configuration -------------------------------------------------------

# Sheets that are NOT channel data and must be ignored on load.
# A sheet is treated as a channel/deal sheet *purely by structure*: it must
# carry these columns. This is why a "Summary"/rollup sheet is ignored
# automatically whether or not it is present -- nothing depends on its name or
# even its existence. (Exact match, so a sheet with off-spec headers fails loudly
# rather than being mis-counted downstream.)
REQUIRED_COLUMNS = {"Deal Status", "Deal Date"}

# Money / numeric fields that arrive as text in the source file.
NUMERIC_COLUMNS = [
    "Deal Value (AED)", "Initial Value", "Final Value", "Difference",
    "Top-up Amount", "Voucher Amount",
]

# Friendly channel names (match the wording used in the manual Summary sheet).
CHANNEL_RENAME = {
    "SharafDG Onlinehelp": "Sharaf DG Online Help",
    "Samsung Web": "Samsung Online (Web)",
    "Etisalat": "Etisalat Online",
    "Istyle Online": "iStyle Online",
    "Sharafdg Tradein Later": "Sharaf DG Trade-in Later",
}

# Canonical funnel buckets, in funnel order.
# Not Completed / Waiting For Customer -> Pending; Terminal Status "Awaiting
# confirmation" is kept as its own bucket.
FUNNEL_ORDER = ["Total Leads", "Deals Done", "Rejected", "Pending", "Awaiting Confirmation"]
STATUS_ORDER = ["Done", "Rejected", "Pending", "Awaiting Confirmation"]

STATUS_COLORS = {
    "Done": "#16a34a",
    "Rejected": "#dc2626",
    "Pending": "#f59e0b",
    "Awaiting Confirmation": "#3b82f6",
}

# Time-bucket -> pandas resample/Grouper frequency (start-anchored for clean labels).
BUCKET_FREQ = {
    "Weekly": "W-MON",
    "Bi-weekly": "2W-MON",
    "Monthly": "MS",
    "Quarterly": "QS",
    "Semi-annual": "2QS",
    "Annual": "YS",
}
BUCKET_OPTIONS = list(BUCKET_FREQ.keys())


# --- Loading -------------------------------------------------------------

def _clean_channel(sheet_name: str) -> str:
    name = sheet_name.strip()
    return CHANNEL_RENAME.get(name, name)


def load_workbook(source, fero_source=None) -> pd.DataFrame:
    """Read the Home Leads Funnel `source` (path, bytes, or file-like) into one tidy DataFrame.

    Works for both layouts: one channel per sheet, OR a single flat sheet where the
    channel lives in the `Store Name` column. If `fero_source` is given, the deals are
    enriched with Fero's rejection reason + appraisal remark (see apply_fero).
    """
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)

    xl = pd.ExcelFile(source)
    frames, skipped = [], []
    for sheet in xl.sheet_names:
        raw = xl.parse(sheet, dtype=str)
        raw.columns = [str(c).strip() for c in raw.columns]
        if not REQUIRED_COLUMNS.issubset(set(raw.columns)):
            # Not a deal sheet (Summary, notes, blank tab, ...) -- skip.
            skipped.append(sheet)
            continue
        # Channel = Store Name column (authoritative; works for the flat file too),
        # falling back to the sheet name where Store Name is blank.
        fallback = _clean_channel(sheet)
        if "Store Name" in raw.columns:
            sn = raw["Store Name"].fillna("").astype(str).str.strip()
            raw["Channel"] = sn.where(sn.ne("") & sn.str.lower().ne("nan"), fallback)
        else:
            raw["Channel"] = fallback
        frames.append(raw)

    if not frames:
        raise ValueError(
            "No deal sheets found in this workbook. A deal/channel sheet must have "
            f"both {sorted(REQUIRED_COLUMNS)} columns. Sheets seen: {list(xl.sheet_names)}."
        )

    df = _normalise(pd.concat(frames, ignore_index=True))
    if fero_source is not None:
        df = apply_fero(df, fero_source)
    return df


def apply_fero(df: pd.DataFrame, fero_source) -> pd.DataFrame:
    """Overlay Fero appraisal data onto the funnel deals.

    Joins Fero `Appraisal ID` to the funnel `Deal ID` and brings over ONLY the
    rejection reason and appraisal remark. Fero's reason becomes the primary
    rejection reason (the funnel's own is the fallback where there's no match).
    Adds helper columns: _fero_reason, _fero_remark, _fero_matched.
    """
    if isinstance(fero_source, (bytes, bytearray)):
        fero_source = io.BytesIO(fero_source)
    fero = pd.read_excel(fero_source, dtype=str)
    fero.columns = [str(c).strip() for c in fero.columns]
    if "Appraisal ID" not in fero.columns:
        raise ValueError(
            "Fero file must contain an 'Appraisal ID' column to map to the funnel's "
            f"Deal ID. Columns seen: {list(fero.columns)[:10]}..."
        )

    def cl(series):
        # normalise text (Fero values contain non-breaking spaces)
        return series.fillna("").astype(str).str.replace("\xa0", " ", regex=False).str.strip()

    fero["_aid"] = cl(fero["Appraisal ID"])
    fero["_reason"] = cl(fero["Reason"]) if "Reason" in fero.columns else ""
    fero["_remark"] = cl(fero["Appraisal Remarks"]) if "Appraisal Remarks" in fero.columns else ""
    # One row per appraisal id; prefer the row that actually has a reason.
    fero = fero[fero["_aid"] != ""].sort_values("_reason", key=lambda s: s.eq(""))
    fmap = fero.drop_duplicates("_aid").set_index("_aid")

    key = (df["Deal ID"].fillna("").astype(str).str.strip()
           if "Deal ID" in df.columns else pd.Series("", index=df.index))
    df["_fero_reason"] = key.map(fmap["_reason"]).fillna("")
    df["_fero_remark"] = key.map(fmap["_remark"]).fillna("")
    df["_fero_matched"] = key.ne("") & key.isin(fmap.index)

    # Rejection reason comes ONLY from Fero (matched by Deal ID = Appraisal ID).
    # (When no Fero file is supplied, the funnel's own reason set in _normalise stands.)
    df["_reject_reason"] = df["_fero_reason"].replace("", pd.NA)
    return df


def _normalise(df: pd.DataFrame) -> pd.DataFrame:
    # --- dates ---
    df["_date"] = pd.to_datetime(df.get("Deal Date"), format="%d/%m/%Y", errors="coerce")
    # Fallback for any other date style in a future file.
    if df["_date"].isna().any():
        mask = df["_date"].isna()
        df.loc[mask, "_date"] = pd.to_datetime(
            df.loc[mask, "Deal Date"], errors="coerce", dayfirst=True
        )

    # --- money / numerics ---
    def to_num(series):
        return pd.to_numeric(
            series.astype(str).str.replace(",", "", regex=False).str.strip(),
            errors="coerce",
        )

    def num_col(col):
        """Numeric version of `col` if present, else an all-NaN column."""
        if col in df.columns:
            return to_num(df[col])
        return pd.Series(np.nan, index=df.index)

    df["_deal_value"] = num_col("Deal Value (AED)")
    df["_initial_value"] = num_col("Initial Value")
    df["_final_value"] = num_col("Final Value")
    df["_top_up"] = num_col("Top-up Amount")
    # Deal value occasionally blank -> fall back to the initial appraisal.
    df["_deal_value"] = df["_deal_value"].fillna(df["_initial_value"])
    df["_regrade"] = df["_final_value"] - df["_initial_value"]

    # --- coalesce duplicate reject reason columns ---
    reasons = [c for c in ["Reject Reason", "Rejection Reason"] if c in df.columns]
    if reasons:
        df["_reject_reason"] = df[reasons[0]]
        for c in reasons[1:]:
            df["_reject_reason"] = df["_reject_reason"].fillna(df[c])
    else:
        df["_reject_reason"] = np.nan

    # --- Fero enrichment placeholders (filled by apply_fero if a Fero file is given) ---
    df["_fero_reason"] = ""
    df["_fero_remark"] = ""
    df["_fero_matched"] = False

    # --- canonical funnel status ---
    df["Status"] = df.apply(_derive_status, axis=1)
    return df


def _derive_status(row) -> str:
    """Map raw Deal Status (+ Terminal Status) onto the canonical funnel buckets.

    Done -> Done, Rejected -> Rejected. For the rest: Terminal Status
    "Awaiting confirmation" -> Awaiting Confirmation; everything else
    (Not Completed / Waiting For Customer / blank) -> Pending.
    """
    raw = str(row.get("Deal Status", "")).strip().lower()
    if raw == "done":
        return "Done"
    if raw == "rejected":
        return "Rejected"
    terminal = str(row.get("Terminal Status", "") or "").strip().lower()
    if "awaiting confirmation" in terminal:
        return "Awaiting Confirmation"
    return "Pending"


# --- Aggregations --------------------------------------------------------

def funnel_counts(df: pd.DataFrame) -> dict:
    """Return the headline funnel counts for the given (already-filtered) frame."""
    vc = df["Status"].value_counts()
    total = len(df)
    out = {"Total Leads": total}
    for s in STATUS_ORDER:
        out[{"Done": "Deals Done"}.get(s, s)] = int(vc.get(s, 0))
    return out


def channel_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Recreate (automatically) the per-channel Summary table."""
    g = df.groupby("Channel")["Status"].value_counts().unstack(fill_value=0)
    for s in STATUS_ORDER:
        if s not in g.columns:
            g[s] = 0
    g = g[STATUS_ORDER]
    g.insert(0, "Total Leads", g.sum(axis=1))
    g["Conversion %"] = (g["Done"] / g["Total Leads"].replace(0, np.nan) * 100).round(1)
    g = g.sort_values("Total Leads", ascending=False)
    # Grand total row (status columns keep their raw names: Done, Rejected, ...).
    total = g.drop(columns="Conversion %").sum(numeric_only=True)
    total["Conversion %"] = round(total["Done"] / total["Total Leads"] * 100, 1)
    total.name = "ALL CHANNELS"
    g = pd.concat([g, total.to_frame().T])
    return g.reset_index().rename(columns={"index": "Channel"})


def trend(df: pd.DataFrame, bucket: str) -> pd.DataFrame:
    """Counts per status per time bucket, with a clean period label."""
    freq = BUCKET_FREQ[bucket]
    d = df.dropna(subset=["_date"]).copy()
    if d.empty:
        return pd.DataFrame(columns=["period", "label"] + STATUS_ORDER + ["Total Leads", "Conversion %"])
    grp = d.groupby([pd.Grouper(key="_date", freq=freq), "Status"]).size().unstack(fill_value=0)
    for s in STATUS_ORDER:
        if s not in grp.columns:
            grp[s] = 0
    grp = grp[STATUS_ORDER].reset_index().rename(columns={"_date": "period"})
    grp["Total Leads"] = grp[STATUS_ORDER].sum(axis=1)
    grp["Conversion %"] = (grp["Done"] / grp["Total Leads"].replace(0, np.nan) * 100).round(1)
    grp["label"] = grp["period"].map(lambda p: period_label(p, bucket))
    return grp


def period_partner(df: pd.DataFrame, bucket: str) -> pd.DataFrame:
    """Lead volume per **time bucket × partner**.

    Rows = period labels (chronological), columns = channels ordered by total
    volume (busiest first). Values = number of leads. No totals appended here;
    the report adds Total row/column.
    """
    freq = BUCKET_FREQ[bucket]
    d = df.dropna(subset=["_date"]).copy()
    if d.empty:
        return pd.DataFrame()
    g = (d.groupby([pd.Grouper(key="_date", freq=freq), "Channel"])
           .size().unstack(fill_value=0))
    g = g.sort_index()                       # chronological
    order = g.sum(axis=0).sort_values(ascending=False).index.tolist()
    g = g[order]
    g.index = [period_label(ts, bucket) for ts in g.index]
    return g


def period_label(ts: pd.Timestamp, bucket: str) -> str:
    if pd.isna(ts):
        return "?"
    if bucket in ("Weekly", "Bi-weekly"):
        return "w/c " + ts.strftime("%d %b %Y")
    if bucket == "Monthly":
        return ts.strftime("%b %Y")
    if bucket == "Quarterly":
        return f"Q{(ts.month - 1) // 3 + 1} {ts.year}"
    if bucket == "Semi-annual":
        return f"H{1 if ts.month <= 6 else 2} {ts.year}"
    if bucket == "Annual":
        return str(ts.year)
    return ts.strftime("%Y-%m-%d")


def reject_reasons(df: pd.DataFrame, top: int = 10) -> pd.DataFrame:
    r = df.loc[df["Status"] == "Rejected", "_reject_reason"].dropna()
    r = r[r.astype(str).str.strip() != ""]
    out = r.value_counts().head(top).rename_axis("Reason").reset_index(name="Count")
    return out


def reject_pivot(df: pd.DataFrame, top: int = 8) -> pd.DataFrame:
    """Rejection reasons broken down **by partner/channel**.

    Returns a matrix: rows = top reason (+ 'Other reasons'), columns = channels
    ordered by their rejection volume, plus a 'Total' column; a final
    'ALL REASONS' total row. Counts of rejected deals.
    """
    r = df[df["Status"] == "Rejected"].copy()
    r["_reason"] = r["_reject_reason"].fillna("").astype(str).str.strip()
    r = r[r["_reason"] != ""]
    if r.empty:
        return pd.DataFrame()

    piv = pd.crosstab(r["_reason"], r["Channel"])
    piv["Total"] = piv.sum(axis=1)
    piv = piv.sort_values("Total", ascending=False)

    # Keep the top reasons; roll the rest into a single "Other reasons" row.
    if len(piv) > top:
        head = piv.iloc[:top].copy()
        other = piv.iloc[top:].sum(axis=0)
        other.name = "Other reasons"
        piv = pd.concat([head, other.to_frame().T])

    # Order channel columns by their rejection volume (busiest first).
    chan_cols = [c for c in piv.columns if c != "Total"]
    order = piv.loc[:, chan_cols].sum(axis=0).sort_values(ascending=False).index.tolist()
    piv = piv[order + ["Total"]]

    total_row = piv.sum(axis=0)
    total_row.name = "ALL REASONS"
    piv = pd.concat([piv, total_row.to_frame().T])
    return piv.astype(int)


def rejection_detail(df: pd.DataFrame) -> pd.DataFrame:
    """Per-device rejection reasons (Channel, Status, Reason) for rejected deals
    that have a reason recorded."""
    r = df[df["Status"] == "Rejected"].copy()
    r["Reason"] = (r["_reject_reason"].fillna("").astype(str)
                   .str.replace("\xa0", " ", regex=False).str.strip())
    r = r[r["Reason"] != ""]
    if r.empty:
        return pd.DataFrame(columns=["Channel", "Status", "Reason"])
    return r[["Channel", "Status", "Reason"]].reset_index(drop=True)


def appraisal_remarks(df: pd.DataFrame) -> pd.DataFrame:
    """Deals that carry an appraiser remark from Fero (free-text notes)."""
    if "_fero_remark" not in df.columns:
        return pd.DataFrame(columns=["Deal ID", "Channel", "Status", "Reason", "Remark"])
    rem = df["_fero_remark"].fillna("").astype(str).str.strip()
    sub = df[rem != ""].copy()
    if sub.empty:
        return pd.DataFrame(columns=["Deal ID", "Channel", "Status", "Reason", "Remark"])
    out = pd.DataFrame({
        "Deal ID": sub.get("Deal ID", pd.Series("", index=sub.index)).astype(str),
        "Channel": sub["Channel"],
        "Status": sub["Status"],
        "Reason": sub["_fero_reason"].replace("", "—"),
        "Remark": sub["_fero_remark"],
    })
    return out.reset_index(drop=True)


def fero_stats(df: pd.DataFrame) -> dict:
    """Summary of how much Fero enrichment was applied."""
    matched = bool(df.get("_fero_matched", pd.Series(dtype=bool)).any())
    n_match = int(df["_fero_matched"].sum()) if "_fero_matched" in df else 0
    n_reason = int((df.get("_fero_reason", pd.Series([], dtype=str)).fillna("") != "").sum())
    n_remark = int((df.get("_fero_remark", pd.Series([], dtype=str)).fillna("") != "").sum())
    return {"applied": matched, "matched": n_match, "reasons": n_reason, "remarks": n_remark}


def breakdown(df: pd.DataFrame, col: str, top: int = 10) -> pd.DataFrame:
    if col not in df.columns:
        return pd.DataFrame(columns=[col, "Count"])
    s = df[col].dropna()
    s = s[s.astype(str).str.strip() != ""]
    return s.value_counts().head(top).rename_axis(col).reset_index(name="Count")


def regrade_stats(df: pd.DataFrame) -> dict:
    """Inspection regrade economics on completed (Done) deals."""
    d = df[(df["Status"] == "Done") & df["_regrade"].notna()]
    if d.empty:
        return {"n": 0}
    diff = d["_regrade"]
    return {
        "n": int(len(d)),
        "unchanged": int((diff == 0).sum()),
        "devalued": int((diff < 0).sum()),
        "uplift": int((diff > 0).sum()),
        "avg_drop": float(diff[diff < 0].mean()) if (diff < 0).any() else 0.0,
        "total_initial": float(d["_initial_value"].sum()),
        "total_final": float(d["_final_value"].sum()),
    }

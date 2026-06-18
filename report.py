"""
Builds the downloadable Excel **summary report** — a colour-coded, automated
replacement for the hand-built Summary sheet.

Sheets:
  - Summary: dated header, colourful KPI cards, and a channel-wise distribution
    table with **count + % side by side** for every funnel stage. Each status owns
    a colour (green/red/amber/blue) carried through headers, column tints, KPI cards
    and charts. Data bars on Total Leads, a red→green colour scale on Conversion %,
    plus a stacked bar chart (funnel per channel) and an outcome pie.
  - By Period: the chosen time bucket (weekly … annual) with counts, % and a chart.
  - Rejection Reasons: ranked reasons with share + bar chart.

The layout is driven by data.STATUS_ORDER, so adding/removing a funnel bucket
(e.g. Awaiting Confirmation) flows through automatically.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import data as D

# --- palette -------------------------------------------------------------
SLATE = "1F2937"
INK = "0F172A"
MUTE = "64748B"

# Status colour system: strong (headers/cards) + tint (column bands).
STATUS_HEX = {"Done": "16A34A", "Rejected": "DC2626",
              "Pending": "F59E0B", "Awaiting Confirmation": "2563EB"}
STATUS_TINT = {"Done": "DCFCE7", "Rejected": "FEE2E2",
               "Pending": "FEF3C7", "Awaiting Confirmation": "DBEAFE"}
DISPLAY = {"Done": "Deals Done"}   # status -> column label override

# --- reusable styles -----------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor=SLATE)
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(size=18, bold=True, color=INK)
DATE_FONT = Font(size=10, bold=True, color="FFFFFF")
SUB_FONT = Font(size=10, italic=True, color=MUTE)
TOTAL_FILL = PatternFill("solid", fgColor="E2E8F0")
TOTAL_FONT = Font(bold=True, color=INK)
BAND_FILL = PatternFill("solid", fgColor="EFF3F8")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
CANVAS_FILL = PatternFill("solid", fgColor="E7ECF3")   # light dashboard background
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
_thin = Side(style="thin", color="CBD5E1")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
CNT_FMT = "#,##0"
PCT_FMT = "0.0%"


def build_report_xlsx(df, bucket: str = "Monthly") -> bytes:
    funnel = D.funnel_counts(df)
    summ = D.channel_summary(df)
    ppartner = D.period_partner(df, bucket)
    rpivot = D.reject_pivot(df, top=8)
    dmin, dmax = df["_date"].min(), df["_date"].max()
    period = (f"{dmin:%d %b %Y} – {dmax:%d %b %Y}"
              if pd.notna(dmin) and pd.notna(dmax) else "all dates")

    wb = Workbook()
    _summary_sheet(wb.active, df, funnel, summ, period)
    _period_sheet(wb.create_sheet("By Period"), ppartner, bucket, period)
    _reasons_sheet(wb.create_sheet("Rejection Reasons"), rpivot, funnel, period)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# -------------------------------------------------------------------------
def _dated_header(ws, title, period):
    """Title + a coloured dated ribbon (the 'date in the output')."""
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"  Data period: {period}      Report generated: {date.today():%d %b %Y}"
    ws["A2"].font = DATE_FONT
    ws["A2"].fill = PatternFill("solid", fgColor=SLATE)
    ws["A2"].alignment = LEFT
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 20


def _summary_sheet(ws, df, funnel, summ, period):
    ws.title = "Summary"
    # Light dashboard canvas behind everything; cards/table/charts sit on top.
    for rr in range(1, 64):
        for cc in range(1, 16):
            ws.cell(row=rr, column=cc).fill = CANVAS_FILL
    _dated_header(ws, "Home Leads Funnel — Summary", period)
    ws["G1"] = f"{funnel['Total Leads']:,} leads · {df['Channel'].nunique()} channels"
    ws["G1"].font = SUB_FONT

    statuses = D.STATUS_ORDER
    total = funnel["Total Leads"] or 1

    # ---- colourful KPI cards (row 4 label / row 5 value), 2 cols each ----
    cards = [("Total Leads", funnel["Total Leads"], None, "4F46E5")]
    for s in statuses:
        key = "Deals Done" if s == "Done" else s
        cards.append((key, funnel.get("Deals Done" if s == "Done" else s, 0),
                      funnel.get("Deals Done" if s == "Done" else s, 0) / total, STATUS_HEX[s]))
    for i, (lbl, val, pct, hexcol) in enumerate(cards):
        c0 = 1 + i * 2
        ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 1)
        ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 1)
        lc = ws.cell(row=4, column=c0, value=lbl.upper())
        lc.fill = PatternFill("solid", fgColor=hexcol)
        lc.font = Font(bold=True, color="FFFFFF", size=9)
        lc.alignment = CENTER
        txt = f"{val:,}" + (f"   {pct:.0%}" if pct is not None else "")
        vc = ws.cell(row=5, column=c0, value=txt)
        vc.fill = PatternFill("solid", fgColor=hexcol)
        vc.font = Font(bold=True, color="FFFFFF", size=15)
        vc.alignment = CENTER
        ws.row_dimensions[5].height = 26

    # ---- channel distribution table (count + % side by side) ----
    # Dynamic columns: Channel | Total | [status count | %]... | Conversion %
    spec = [("Channel", "text", None), ("Total Leads", "count", None)]
    for s in statuses:
        spec.append((DISPLAY.get(s, s), "count", s))
        spec.append(("%", "pct", s))
    spec.append(("Conversion %", "pct", None))
    ncol = len(spec)
    total_col = 2
    conv_col = ncol

    top = 8
    for j, (label, _kind, s) in enumerate(spec, 1):
        cell = ws.cell(row=top, column=j, value=label)
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
        cell.fill = PatternFill("solid", fgColor=STATUS_HEX[s]) if s else HEADER_FILL

    def status_count(row_or_funnel, s, total_leads):
        col = "Done" if s == "Done" else s
        return row_or_funnel[col]

    body = summ[summ["Channel"] != "ALL CHANNELS"]
    r = top
    for _, row in body.iterrows():
        r += 1
        tot = row["Total Leads"] or 1
        for j, (label, kind, s) in enumerate(spec, 1):
            if kind == "text":
                v = row["Channel"]
            elif label == "Total Leads":
                v = int(row["Total Leads"])
            elif label == "Conversion %":
                v = row["Done"] / tot
            elif kind == "count":
                v = int(status_count(row, s, tot))
            else:  # pct beside its status count
                v = status_count(row, s, tot) / tot
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = BORDER
            cell.alignment = LEFT if kind == "text" else CENTER
            if kind == "pct":
                cell.number_format = PCT_FMT
            elif kind == "count" and label != "Channel":
                cell.number_format = CNT_FMT
            # colour: status columns get their tint; others get crisp banding
            if s:
                cell.fill = PatternFill("solid", fgColor=STATUS_TINT[s])
            else:
                cell.fill = BAND_FILL if (r - top) % 2 == 0 else WHITE_FILL

    # total row
    r += 1
    tot_all = funnel["Total Leads"] or 1
    for j, (label, kind, s) in enumerate(spec, 1):
        if kind == "text":
            v = "ALL CHANNELS"
        elif label == "Total Leads":
            v = funnel["Total Leads"]
        elif label == "Conversion %":
            v = funnel["Deals Done"] / tot_all
        elif kind == "count":
            v = funnel["Deals Done" if s == "Done" else s]
        else:
            v = funnel["Deals Done" if s == "Done" else s] / tot_all
        cell = ws.cell(row=r, column=j, value=v)
        cell.fill = TOTAL_FILL
        cell.font = TOTAL_FONT
        cell.border = BORDER
        cell.alignment = LEFT if kind == "text" else CENTER
        if kind == "pct":
            cell.number_format = PCT_FMT
        elif kind == "count" and label != "Channel":
            cell.number_format = CNT_FMT

    last_data_row = r - 1
    L = get_column_letter
    # data bars on Total Leads + colour scale on Conversion %  (engagement)
    ws.conditional_formatting.add(
        f"{L(total_col)}{top+1}:{L(total_col)}{last_data_row}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="6366F1"))
    ws.conditional_formatting.add(
        f"{L(conv_col)}{top+1}:{L(conv_col)}{last_data_row}",
        ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                       mid_type="percentile", mid_value=50, mid_color="FFEB84",
                       end_type="max", end_color="63BE7B"))

    # widths + freeze
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 12
    for j in range(3, ncol):
        ws.column_dimensions[L(j)].width = 11 if (j % 2 == 1) else 8
    ws.column_dimensions[L(conv_col)].width = 13
    ws.freeze_panes = f"A{top+1}"

    _charts(ws, statuses, top, last_data_row, conv_col, anchor_row=r + 3)


def _charts(ws, statuses, top, last_data_row, conv_col, anchor_row):
    """Two charts, both sourced from the VISIBLE table so they always render
    (Excel does not plot hidden cells): funnel-by-channel + conversion-by-channel.
    Series colours match the table/KPI status colours."""
    cats = Reference(ws, min_col=1, min_row=top + 1, max_row=last_data_row)

    # 1) Funnel by channel (stacked). Count columns are visible: 3, 5, 7, ...
    bar = BarChart()
    bar.type = "bar"; bar.grouping = "stacked"; bar.overlap = 100
    bar.title = "Funnel by channel"; bar.height = 9.5; bar.width = 24
    for k, s in enumerate(statuses):
        ref = Reference(ws, min_col=3 + 2 * k, min_row=top, max_row=last_data_row)
        bar.add_data(ref, titles_from_data=True)
    bar.set_categories(cats)
    for srs, s in zip(bar.series, statuses):
        srs.graphicalProperties.solidFill = STATUS_HEX[s]
        srs.graphicalProperties.line.solidFill = STATUS_HEX[s]
    bar.legend.position = "b"
    ws.add_chart(bar, f"A{anchor_row}")

    # 2) Conversion % by channel (visible Conversion % column).
    conv = BarChart()
    conv.type = "col"; conv.title = "Conversion % by channel"
    conv.height = 8; conv.width = 24
    conv.add_data(Reference(ws, min_col=conv_col, min_row=top, max_row=last_data_row),
                  titles_from_data=True)
    conv.set_categories(cats)
    conv.series[0].graphicalProperties.solidFill = "4F46E5"
    conv.y_axis.numFmt = "0%"
    conv.y_axis.majorGridlines = None
    conv.legend = None
    ws.add_chart(conv, f"A{anchor_row + 20}")


# -------------------------------------------------------------------------
LINE_PALETTE = ["2563EB", "16A34A", "DC2626", "D97706", "7C3AED", "0891B2"]


def _period_sheet(ws, pp, bucket, period):
    """Partner-wise: leads per time bucket × partner (heatmap matrix + trend lines)."""
    _dated_header(ws, f"Leads by {bucket} × Partner", period)
    if pp is None or pp.empty:
        ws["A4"] = "No dated deals in this selection."
        ws["A4"].font = SUB_FONT
        return

    partners = list(pp.columns)               # already ordered busiest-first
    periods = list(pp.index)
    ncol = 1 + len(partners) + 1              # Period | partners... | Total
    total_col = ncol
    top = 4

    # ---- header (Period | rotated partner names | Total) ----
    hp = ws.cell(row=top, column=1, value="Period")
    hp.font = HEADER_FONT; hp.alignment = LEFT; hp.border = BORDER; hp.fill = HEADER_FILL
    for k, ch in enumerate(partners, start=2):
        cell = ws.cell(row=top, column=k, value=ch)
        cell.font = HEADER_FONT; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="bottom", text_rotation=90)
        cell.fill = PatternFill("solid", fgColor="4F46E5")
    tc = ws.cell(row=top, column=total_col, value="Total")
    tc.font = HEADER_FONT; tc.alignment = CENTER; tc.border = BORDER
    tc.fill = PatternFill("solid", fgColor="312E81")
    ws.row_dimensions[top].height = 96

    # ---- body: one row per period ----
    r = top
    body_first = top + 1
    for prd in periods:
        r += 1
        a = ws.cell(row=r, column=1, value=prd); a.border = BORDER; a.alignment = LEFT
        a.fill = BAND_FILL if (r - top) % 2 == 0 else WHITE_FILL
        row_total = 0
        for k, ch in enumerate(partners, start=2):
            v = int(pp.loc[prd, ch]); row_total += v
            cell = ws.cell(row=r, column=k, value=v)
            cell.border = BORDER; cell.alignment = CENTER; cell.number_format = CNT_FMT
        tcell = ws.cell(row=r, column=total_col, value=row_total)
        tcell.border = BORDER; tcell.alignment = CENTER; tcell.number_format = CNT_FMT
    body_last = r

    # ---- total row (per partner) ----
    r += 1
    grand = ws.cell(row=r, column=1, value="TOTAL")
    grand.fill = TOTAL_FILL; grand.font = TOTAL_FONT; grand.border = BORDER; grand.alignment = LEFT
    for k, ch in enumerate(partners, start=2):
        v = int(pp[ch].sum())
        cell = ws.cell(row=r, column=k, value=v)
        cell.fill = TOTAL_FILL; cell.font = TOTAL_FONT; cell.border = BORDER
        cell.alignment = CENTER; cell.number_format = CNT_FMT
    gt = ws.cell(row=r, column=total_col, value=int(pp.values.sum()))
    gt.fill = TOTAL_FILL; gt.font = TOTAL_FONT; gt.border = BORDER; gt.alignment = CENTER; gt.number_format = CNT_FMT

    # ---- heatmap over matrix + data bars on Total column ----
    L = get_column_letter
    ws.conditional_formatting.add(
        f"{L(2)}{body_first}:{L(total_col-1)}{body_last}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                       mid_type="percentile", mid_value=60, mid_color="C7D2FE",
                       end_type="max", end_color="4F46E5"))
    ws.conditional_formatting.add(
        f"{L(total_col)}{body_first}:{L(total_col)}{body_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="312E81"))

    ws.column_dimensions["A"].width = 16
    for k in range(2, total_col):
        ws.column_dimensions[L(k)].width = 6
    ws.column_dimensions[L(total_col)].width = 9
    ws.freeze_panes = "B5"

    # ---- easy trend chart: top partners as lines over time ----
    n_lines = min(6, len(partners))
    line = LineChart()
    line.title = f"Lead trend — top {n_lines} partners"
    line.height = 9; line.width = 24
    line.y_axis.title = "Leads"
    for i in range(n_lines):
        col = 2 + i
        ref = Reference(ws, min_col=col, min_row=top, max_row=body_last)
        line.add_data(ref, titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=body_first, max_row=body_last))
    for i, srs in enumerate(line.series):
        srs.graphicalProperties.line.solidFill = LINE_PALETTE[i % len(LINE_PALETTE)]
        srs.graphicalProperties.line.width = 28000   # ~2.2pt
        srs.smooth = False
    line.legend.position = "b"
    ws.add_chart(line, f"A{r + 3}")


def _reasons_sheet(ws, pivot, funnel, period):
    _dated_header(ws, "Rejection Reasons by Partner", period)
    if pivot is None or pivot.empty:
        ws["A4"] = "No rejections with a recorded reason in this selection."
        ws["A4"].font = SUB_FONT
        return

    channels = [c for c in pivot.columns if c != "Total"]
    reasons = list(pivot.index)                 # includes 'ALL REASONS' total row last
    ncol = 1 + len(channels) + 1                # Reason | channels... | Total
    total_col = ncol
    top = 4

    # ---- header (reason | rotated channel names | total) ----
    h = ws.cell(row=top, column=1, value="Rejection reason")
    h.font = HEADER_FONT; h.alignment = LEFT; h.border = BORDER
    h.fill = PatternFill("solid", fgColor="DC2626")
    for k, ch in enumerate(channels, start=2):
        cell = ws.cell(row=top, column=k, value=ch)
        cell.font = HEADER_FONT; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="bottom", text_rotation=90)
        cell.fill = HEADER_FILL
    tc = ws.cell(row=top, column=total_col, value="Total")
    tc.font = HEADER_FONT; tc.alignment = CENTER; tc.border = BORDER; tc.fill = PatternFill("solid", fgColor="991B1B")
    ws.row_dimensions[top].height = 96

    # ---- body ----
    r = top
    body_first = top + 1
    for reason in reasons:
        r += 1
        is_total = reason == "ALL REASONS"
        a = ws.cell(row=r, column=1, value=reason)
        a.border = BORDER; a.alignment = LEFT
        for k, ch in enumerate(channels, start=2):
            cell = ws.cell(row=r, column=k, value=int(pivot.loc[reason, ch]))
            cell.border = BORDER; cell.alignment = CENTER; cell.number_format = CNT_FMT
        t = ws.cell(row=r, column=total_col, value=int(pivot.loc[reason, "Total"]))
        t.border = BORDER; t.alignment = CENTER; t.number_format = CNT_FMT
        if is_total:
            for cc in range(1, ncol + 1):
                ws.cell(row=r, column=cc).fill = TOTAL_FILL
                ws.cell(row=r, column=cc).font = TOTAL_FONT
        else:
            a.fill = BAND_FILL if (r - top) % 2 == 0 else WHITE_FILL
    body_last = r - 1            # excludes ALL REASONS row
    grand_total_row = r

    # ---- heatmap over the matrix body (channel cells) + data bars on Total ----
    L = get_column_letter
    ws.conditional_formatting.add(
        f"{L(2)}{body_first}:{L(total_col-1)}{body_last}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                       mid_type="percentile", mid_value=60, mid_color="FCA5A5",
                       end_type="max", end_color="DC2626"))
    ws.conditional_formatting.add(
        f"{L(total_col)}{body_first}:{L(total_col)}{body_last}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="991B1B"))

    # ---- widths + freeze (keep reason column + header visible) ----
    ws.column_dimensions["A"].width = 30
    for k in range(2, total_col):
        ws.column_dimensions[L(k)].width = 5.5
    ws.column_dimensions[L(total_col)].width = 8
    ws.freeze_panes = "B5"

    # ---- two simple, single-message charts (easier to read than a stack) ----
    # A) Top rejection reasons overall (horizontal bar) — uses the Total column.
    rbar = BarChart()
    rbar.type = "bar"; rbar.title = "Top rejection reasons (all partners)"
    rbar.height = 0.55 * (body_last - body_first + 1) + 2.5; rbar.width = 16
    rbar.add_data(Reference(ws, min_col=total_col, min_row=body_first, max_row=body_last))
    rbar.set_categories(Reference(ws, min_col=1, min_row=body_first, max_row=body_last))
    rbar.series[0].graphicalProperties.solidFill = "DC2626"
    rbar.legend = None
    ws.add_chart(rbar, f"A{grand_total_row + 3}")

    # B) Total rejections by partner (column) — uses the ALL REASONS total row.
    pbar = BarChart()
    pbar.type = "col"; pbar.title = "Total rejections by partner"
    pbar.height = 9; pbar.width = 16
    data = Reference(ws, min_col=2, max_col=1 + len(channels), min_row=grand_total_row, max_row=grand_total_row)
    pbar.add_data(data, from_rows=True)
    pbar.set_categories(Reference(ws, min_col=2, max_col=1 + len(channels), min_row=top, max_row=top))
    pbar.series[0].graphicalProperties.solidFill = "991B1B"
    pbar.legend = None
    ws.add_chart(pbar, f"K{grand_total_row + 3}")

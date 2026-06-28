import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, PageBreak, Spacer
from reportlab.pdfbase import pdfmetrics
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

df = pd.read_excel("Bank Submit History (26).xlsx")
df.columns = df.columns.str.strip()
df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]
df = df.dropna(axis=1, how="all")
df = df.dropna(subset=["Firm Name"])
df["_date"] = pd.to_datetime(df["Bank Submition Date"], errors="coerce")
for col in ["Maturity Date", "Payment. Rcv Dt", "Bank Accept Date"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

today = pd.Timestamp.today().normalize()
df_show = df[(df["Payment. Rcv Dt"].isna()) & (df["Maturity Date"].notna())].copy()
df_show["Days Until Maturity"] = (today - df_show["Maturity Date"]).dt.days
df_show["Due Status"] = df_show["Days Until Maturity"].apply(
    lambda d: "Overdue" if d > 0 else "Due in 7d" if d >= -7
    else "Due in 15d" if d >= -15 else "Due in 30d" if d >= -30
    else "Due in 60d+" if d >= -60 else "Safe")

overdue_df = df_show[df_show["Days Until Maturity"] > 0]
due7_df = df_show[(df_show["Days Until Maturity"] <= 0) & (df_show["Days Until Maturity"] >= -7)]

firm_overdue = overdue_df.groupby("Firm Name").agg(
    count=("Invoice Value", "size"),
    total_value=("Invoice Value", "sum"),
    max_overdue_days=("Days Until Maturity", "max"),
).sort_values("total_value", ascending=False).head(15)

bank_overdue = overdue_df.groupby("Our Bank").agg(
    count=("Invoice Value", "size"),
    total_value=("Invoice Value", "sum"),
    max_overdue_days=("Days Until Maturity", "max"),
    avg_overdue_days=("Days Until Maturity", "mean"),
).sort_values("total_value", ascending=False).head(15)

def usd(v):
    try:
        v = float(v)
    except:
        return "$0.00"
    if v >= 1e6:
        return "$" + f"{v/1e6:.2f}" + "M"
    if v >= 1e3:
        return "$" + f"{v/1e3:.1f}" + "K"
    return "$" + f"{v:.2f}"

summary_data = {
    "total": len(df_show), "total_val": usd(df_show["Invoice Value"].sum()),
    "overdue": len(overdue_df), "overdue_val": usd(overdue_df["Invoice Value"].sum()),
    "due7": len(due7_df), "due7_val": usd(due7_df["Invoice Value"].sum()),
    "due15": len(df_show[(df_show["Days Until Maturity"] < -7) & (df_show["Days Until Maturity"] >= -15)]),
    "due15_val": usd(df_show[(df_show["Days Until Maturity"] < -7) & (df_show["Days Until Maturity"] >= -15)]["Invoice Value"].sum()),
    "due30": len(df_show[(df_show["Days Until Maturity"] < -15) & (df_show["Days Until Maturity"] >= -30)]),
    "due30_val": usd(df_show[(df_show["Days Until Maturity"] < -15) & (df_show["Days Until Maturity"] >= -30)]["Invoice Value"].sum()),
    "due60": len(df_show[(df_show["Days Until Maturity"] < -30) & (df_show["Days Until Maturity"] >= -60)]),
    "due60_val": usd(df_show[(df_show["Days Until Maturity"] < -30) & (df_show["Days Until Maturity"] >= -60)]["Invoice Value"].sum()),
}

pdf_cols = ["Firm Name", "LC No", "Bank Refno", "Party Name", "Our Bank", "Bank Name",
            "Invoice Value", "Maturity Date", "Days Until Maturity", "Due Status",
            "Sales Person", "Payment. Rcv Dt"]
pdf_cols = [c for c in pdf_cols if c in df_show.columns]
pdf_df_raw = df_show[pdf_cols].copy()
pdf_df_raw["Maturity Date"] = pdf_df_raw["Maturity Date"].dt.strftime("%d %b %Y")
if "Payment. Rcv Dt" in pdf_df_raw.columns:
    pdf_df_raw["Payment. Rcv Dt"] = pd.to_datetime(pdf_df_raw["Payment. Rcv Dt"], errors="coerce").dt.strftime("%d %b %Y").fillna("")
pdf_df_raw["Days Until Maturity"] = pdf_df_raw["Days Until Maturity"].apply(lambda x: str(int(x)))
inv_raw = df_show["Invoice Value"].copy()
pdf_df_raw["Invoice Value"] = inv_raw.apply(lambda x: "$" + f"{x:,.2f}" if pd.notna(x) else "$0.00")
for _em in ["\U0001f534 ", "\U0001f7e1 ", "\U0001f7e0 ", "\U0001f535 ", "\U0001f7e2 ", "\u2705 "]:
    if "Due Status" in pdf_df_raw.columns:
        pdf_df_raw["Due Status"] = pdf_df_raw["Due Status"].str.replace(_em, "", regex=False)
pdf_df_raw = pdf_df_raw.fillna("")

pw, ph = landscape(A3)
lm = rm = 40; tm = 50; bm = 45; uw = pw - lm - rm
hf = "Helvetica-Bold"; cf = "Helvetica"
title_clr = colors.HexColor("#0d3f47"); accent = colors.HexColor("#00c9a7")
dark_bg = colors.HexColor("#0b1220"); mid_bg = colors.HexColor("#1a2235")
header_bg = colors.HexColor("#0d3f47"); header_fg = colors.white
text_dark = colors.HexColor("#222222"); text_mut = colors.HexColor("#666666")
overdue_c = colors.HexColor("#ff3b30"); warn_c = colors.HexColor("#ff9500")
orange_c = colors.HexColor("#ff6b35"); info_c = colors.HexColor("#1a8fff"); safe_c = colors.HexColor("#00c9a7")
row_even = colors.HexColor("#f5f7fa"); row_odd = colors.white; border_c = colors.HexColor("#d0d5dd")
period_label = "All Unpaid"
generated_at = datetime.now().strftime("%d %b %Y %H:%M:%S")

def draw_cover(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(dark_bg); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
    canvas.setFillColor(accent); canvas.roundRect(pw/2 - 200, ph - 180, 400, 8, 4, fill=1, stroke=0)
    canvas.setFillColor(header_fg); canvas.setFont(hf, 32)
    canvas.drawCentredString(pw/2, ph - 250, "DUE DATE TRACKER")
    canvas.setFillColor(accent); canvas.setFont(hf, 18)
    canvas.drawCentredString(pw/2, ph - 285, "Bank Submit History Report")
    canvas.setFillColor(colors.HexColor("#8899aa")); canvas.setFont(cf, 12)
    canvas.drawCentredString(pw/2, ph - 320, "Report Period: " + period_label)
    canvas.drawCentredString(pw/2, ph - 340, "Generated: " + generated_at)
    canvas.setFillColor(accent); canvas.roundRect(pw/2 - 200, ph - 370, 400, 2, 1, fill=1, stroke=0)
    box_y = ph - 530; box_h = 110; box_w = (uw - 40) / 4
    kpi_boxes = [
        ("TOTAL RECORDS", str(summary_data["total"]), summary_data["total_val"], info_c),
        ("TOTAL VALUE", summary_data["total_val"], str(summary_data["total"]) + " Bills", accent),
        ("OVERDUE", str(summary_data["overdue"]) + " Bills", summary_data["overdue_val"], overdue_c),
        ("DUE SOON (7d)", str(summary_data["due7"]) + " Bills", summary_data["due7_val"], warn_c),
    ]
    for i, (lbl, val, sub, clr) in enumerate(kpi_boxes):
        bx = lm + 20 + i * (box_w + 10)
        canvas.setFillColor(mid_bg); canvas.roundRect(bx, box_y, box_w, box_h, 8, fill=1, stroke=0)
        canvas.setStrokeColor(clr); canvas.setLineWidth(3)
        canvas.line(bx, box_y + box_h, bx + box_w, box_y + box_h)
        canvas.setFillColor(clr); canvas.setFont(hf, 10); canvas.drawString(bx + 15, box_y + box_h - 25, lbl)
        canvas.setFillColor(header_fg); canvas.setFont(hf, 26); canvas.drawString(bx + 15, box_y + 40, val)
        canvas.setFillColor(clr); canvas.setFont(hf, 14); canvas.drawString(bx + 15, box_y + 12, sub)
    canvas.setFillColor(colors.HexColor("#556677")); canvas.setFont(cf, 9)
    canvas.drawCentredString(pw/2, 100, "Prepared by: Smart Dashboard v2.0")
    canvas.drawCentredString(pw/2, 82, "Confidential - For Internal Use Only")
    canvas.setStrokeColor(accent); canvas.setLineWidth(0.5); canvas.line(lm, 60, pw - rm, 60)
    canvas.restoreState()

def draw_summary(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
    canvas.setFillColor(title_clr); canvas.setFont(hf, 20)
    canvas.drawString(lm, ph - tm - 10, "EXECUTIVE SUMMARY")
    canvas.setStrokeColor(accent); canvas.setLineWidth(2)
    canvas.line(lm, ph - tm - 18, lm + 220, ph - tm - 18)
    canvas.setFillColor(text_mut); canvas.setFont(cf, 10)
    canvas.drawString(lm, ph - tm - 38, "Report Period: " + period_label + "  |  Generated: " + generated_at)

    sec_y = ph - tm - 70
    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
    canvas.drawString(lm, sec_y, "1. Maturity Status Overview")
    canvas.setStrokeColor(border_c); canvas.setLineWidth(0.5)
    canvas.line(lm, sec_y - 5, pw - rm, sec_y - 5)

    card_y = sec_y - 95; card_h = 75; card_w = (uw - 30) / 5
    cards = [
        ("Overdue", summary_data["overdue"], summary_data["overdue_val"], overdue_c),
        ("Due in 7 Days", summary_data["due7"], summary_data["due7_val"], warn_c),
        ("Due in 15 Days", summary_data["due15"], summary_data["due15_val"], orange_c),
        ("Due in 30 Days", summary_data["due30"], summary_data["due30_val"], info_c),
        ("Due in 60 Days", summary_data["due60"], summary_data["due60_val"], safe_c),
    ]
    for i, (lbl, cnt, val, clr) in enumerate(cards):
        cx = lm + i * (card_w + 6)
        canvas.setFillColor(colors.HexColor("#f8f9fa")); canvas.roundRect(cx, card_y, card_w, card_h, 6, fill=1, stroke=0)
        canvas.setFillColor(clr); canvas.roundRect(cx, card_y + card_h - 5, card_w, 5, 2, fill=1, stroke=0)
        canvas.setFont(hf, 9); canvas.drawString(cx + 10, card_y + card_h - 22, lbl)
        canvas.setFillColor(text_dark); canvas.setFont(hf, 22); canvas.drawString(cx + 10, card_y + 30, str(cnt) + " Bills")
        canvas.setFillColor(clr); canvas.setFont(hf, 11); canvas.drawString(cx + 10, card_y + 8, val)

    # ── Firm-wise Overdue (left side) ──
    sec2_y = card_y - 40
    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
    canvas.drawString(lm, sec2_y, "2. Firm-wise Overdue Summary (Top 15)")
    canvas.setStrokeColor(border_c); canvas.line(lm, sec2_y - 5, pw - rm, sec2_y - 5)

    firm_header = ["#", "Firm Name", "Bill Count", "Total Bill Value", "Max Overdue", "Risk Level"]
    firm_rows = [firm_header]
    for rank, (firm, row) in enumerate(firm_overdue.iterrows(), 1):
        max_d = int(row["max_overdue_days"]); cnt = int(row["count"]); val = usd(row["total_value"])
        risk = "CRITICAL" if max_d > 60 else "HIGH" if max_d > 30 else "MEDIUM" if max_d > 15 else "LOW"
        firm_rows.append([str(rank), str(firm)[:30], str(cnt), val, str(max_d) + "d", risk])

    fcol_w = [25, 180, 60, 100, 65, 70]
    ftbl = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
        ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (1, 0), (1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
    ])
    for ri in range(1, len(firm_rows)):
        rv2 = firm_rows[ri][-1]
        if rv2 == "CRITICAL": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); ftbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        elif rv2 == "HIGH": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); ftbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        elif rv2 == "MEDIUM": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), orange_c)
    ft = LongTable(firm_rows, colWidths=fcol_w, hAlign="LEFT")
    ft.setStyle(ftbl)
    tw, th = ft.wrap(520, 200)
    ft.drawOn(canvas, lm, sec2_y - 25 - th)

    # ── Our Bank-wise Overdue (right side) ──
    bank_x = lm + 560
    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
    canvas.drawString(bank_x, sec2_y, "3. Our Bank-wise Overdue Summary")
    canvas.setStrokeColor(border_c); canvas.line(bank_x, sec2_y - 5, bank_x + 520, sec2_y - 5)

    bheader = ["#", "Our Bank", "Bill Count", "Total Value", "Max Days", "Risk Level"]
    brows = [bheader]
    for rank, (bname, brow) in enumerate(bank_overdue.iterrows(), 1):
        max_d = int(brow["max_overdue_days"]); cnt = int(brow["count"]); val = usd(brow["total_value"])
        risk = "CRITICAL" if max_d > 60 else "HIGH" if max_d > 30 else "MEDIUM" if max_d > 15 else "LOW"
        brows.append([str(rank), str(bname)[:25], str(cnt), val, str(max_d) + "d", risk])

    bcol_w = [25, 120, 55, 85, 55, 70]
    btbl = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
        ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (1, 0), (1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
    ])
    for ri in range(1, len(brows)):
        rv2 = brows[ri][-1]
        if rv2 == "CRITICAL": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); btbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        elif rv2 == "HIGH": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); btbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        elif rv2 == "MEDIUM": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), orange_c)
    bt = LongTable(brows, colWidths=bcol_w, hAlign="LEFT")
    bt.setStyle(btbl)
    btw, bth = bt.wrap(420, 200)
    bt.drawOn(canvas, bank_x, sec2_y - 25 - bth)

    # ── Key Observations ──
    sec3_y = sec2_y - 25 - max(th, bth) - 35
    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
    canvas.drawString(lm, sec3_y, "4. Key Observations")
    canvas.setStrokeColor(border_c); canvas.line(lm, sec3_y - 5, pw - rm, sec3_y - 5)

    total_overdue_val = overdue_df["Invoice Value"].sum()
    total_all_val = df_show["Invoice Value"].sum()
    pct_overdue = (total_overdue_val / total_all_val * 100) if total_all_val > 0 else 0
    top_firm = firm_overdue.index[0] if len(firm_overdue) > 0 else "N/A"
    top_firm_val = usd(firm_overdue.iloc[0]["total_value"]) if len(firm_overdue) > 0 else "$0"
    top_bank = bank_overdue.index[0] if len(bank_overdue) > 0 else "N/A"
    top_bank_val = usd(bank_overdue.iloc[0]["total_value"]) if len(bank_overdue) > 0 else "$0"
    critical_firms = len(firm_overdue[firm_overdue["max_overdue_days"] > 60]) if len(firm_overdue) > 0 else 0

    obs = [
        "OVERALL RISK: " + str(summary_data["overdue"]) + " bills (" + str(round(pct_overdue, 1)) + "% of total value) are overdue, totaling " + summary_data["overdue_val"] + ".",
        "TOP FIRM: " + str(top_firm) + " leads with " + top_firm_val + " overdue. Immediate collection follow-up recommended.",
        "TOP BANK: " + str(top_bank) + " exposure is " + top_bank_val + ". Coordinate with relationship manager.",
        "CRITICAL: " + str(critical_firms) + " firms overdue >60 days. Escalate to senior management for recovery.",
    ]
    oy = sec3_y - 20
    for i, o in enumerate(obs):
        canvas.setFillColor(overdue_c if i == 0 else warn_c if i < 3 else info_c)
        canvas.setFont(hf if i < 2 else cf, 8)
        canvas.drawString(lm + 5, oy - i * 15, o)

    canvas.setFillColor(text_mut); canvas.setFont(cf, 8)
    canvas.drawString(lm, bm - 15, "Page " + str(doc.page) + "  |  Bank Submit History Report")
    canvas.drawRightString(pw - rm, bm - 15, "Confidential")
    canvas.restoreState()

def draw_suggestions(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
    canvas.setFillColor(title_clr); canvas.setFont(hf, 20)
    canvas.drawString(lm, ph - tm - 10, "OVERDUE RECOVERY SUGGESTIONS")
    canvas.setStrokeColor(accent); canvas.setLineWidth(2)
    canvas.line(lm, ph - tm - 18, lm + 300, ph - tm - 18)
    canvas.setFillColor(text_mut); canvas.setFont(cf, 10)
    canvas.drawString(lm, ph - tm - 38, "Firm-wise & Bank-wise action plan for CO/ED review  |  " + generated_at)

    # ── Firm-wise Recovery Roadmap ──
    sec_y = ph - tm - 75
    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
    canvas.drawString(lm, sec_y, "1. Firm-wise Overdue Recovery Roadmap")
    canvas.setStrokeColor(border_c); canvas.setLineWidth(0.5)
    canvas.line(lm, sec_y - 5, pw - rm, sec_y - 5)

    sug_header = ["#", "Firm Name", "Bills", "Bill Value", "Max Days", "Suggested Action", "Priority"]
    sug_rows = [sug_header]
    for rank, (firm, row) in enumerate(firm_overdue.head(12).iterrows(), 1):
        max_d = int(row["max_overdue_days"]); cnt = int(row["count"]); val = usd(row["total_value"])
        if max_d > 60: action = "Immediate escalation to CO. Legal notice."; priority = "URGENT"
        elif max_d > 30 and cnt > 3: action = "Schedule meeting with firm accounts dept."; priority = "HIGH"
        elif max_d > 30: action = "Send formal reminder. Follow up weekly."; priority = "MEDIUM"
        elif cnt > 5: action = "Bulk collection drive via sales team."; priority = "MEDIUM"
        else: action = "Standard follow-up via sales team."; priority = "LOW"
        sug_rows.append([str(rank), str(firm)[:30], str(cnt), val, str(max_d) + "d", action, priority])

    scol_w = [22, 150, 40, 85, 50, 310, 60]
    stbl = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
        ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (4, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
    ])
    for ri in range(1, len(sug_rows)):
        pv = sug_rows[ri][-1]
        if pv == "URGENT": stbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); stbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        elif pv == "HIGH": stbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); stbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
    st = LongTable(sug_rows, colWidths=scol_w, hAlign="LEFT")
    st.setStyle(stbl)
    stw, sth = st.wrap(uw, 200)
    st.drawOn(canvas, lm, sec_y - 25 - sth)

    # ── Bank-wise Recovery Roadmap ──
    sec_b_y = sec_y - 25 - sth - 30
    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
    canvas.drawString(lm, sec_b_y, "2. Our Bank-wise Overdue Recovery Roadmap")
    canvas.setStrokeColor(border_c); canvas.line(lm, sec_b_y - 5, pw - rm, sec_b_y - 5)

    bsug_header = ["#", "Our Bank", "Bills", "Overdue Value", "Max Days", "Bank Action Required", "Priority"]
    bsug_rows = [bsug_header]
    for rank, (bname, brow) in enumerate(bank_overdue.head(10).iterrows(), 1):
        max_d = int(brow["max_overdue_days"]); cnt = int(brow["count"]); val = usd(brow["total_value"])
        if max_d > 60: action = "Escalate to bank management. Request written update."; priority = "URGENT"
        elif max_d > 30: action = "Meet relationship manager. Push for clearance."; priority = "HIGH"
        else: action = "Regular follow-up with bank contact."; priority = "MEDIUM"
        bsug_rows.append([str(rank), str(bname)[:25], str(cnt), val, str(max_d) + "d", action, priority])

    bsc_w = [22, 100, 40, 85, 50, 355, 60]
    bstbl = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
        ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (4, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
    ])
    for ri in range(1, len(bsug_rows)):
        pv = bsug_rows[ri][-1]
        if pv == "URGENT": bstbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); bstbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        elif pv == "HIGH": bstbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); bstbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
    bst = LongTable(bsug_rows, colWidths=bsc_w, hAlign="LEFT")
    bst.setStyle(bstbl)
    bstw, bsth = bst.wrap(uw, 200)
    bst.drawOn(canvas, lm, sec_b_y - 25 - bsth)

    # ── General Strategies ──
    box_top = sec_b_y - 25 - bsth - 25
    canvas.setFillColor(colors.HexColor("#f0f9ff")); canvas.roundRect(lm, box_top - 160, uw, 160, 8, fill=1, stroke=0)
    canvas.setStrokeColor(info_c); canvas.setLineWidth(1.5); canvas.roundRect(lm, box_top - 160, uw, 160, 8, fill=0, stroke=1)
    canvas.setFillColor(title_clr); canvas.setFont(hf, 12)
    canvas.drawString(lm + 15, box_top - 20, "3. General Recovery Strategy")
    suggestions = [
        "1. PRIORITY MATRIX: Classify overdue bills into Critical (>60d), High (>30d), Medium (>15d) and Low risk.",
        "2. FIRM ENGAGEMENT: Schedule direct meetings with top overdue firms. Prepare firm-wise outstanding statements.",
        "3. BANK COORDINATION: Work with relationship managers at CBP, SEBPLC, DBBL for faster clearance.",
        "4. SALES TEAM ALERT: Assign specific sales persons to follow up on their respective firm overdue.",
        "5. WEEKLY REVIEW: Establish weekly overdue review meeting with CO to track recovery progress.",
    ]
    sy = box_top - 38
    for i, s in enumerate(suggestions):
        canvas.setFillColor(text_dark); canvas.setFont(cf, 8)
        canvas.drawString(lm + 20, sy - i * 18, s)

    canvas.setFillColor(text_mut); canvas.setFont(cf, 8)
    canvas.drawString(lm, bm - 15, "Page " + str(doc.page) + "  |  Bank Submit History Report")
    canvas.drawRightString(pw - rm, bm - 15, "Confidential")
    canvas.restoreState()

def draw_data_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
    canvas.setFillColor(title_clr); canvas.setFont(hf, 16)
    canvas.drawString(lm, ph - tm - 5, "DETAILED BILL DATA")
    canvas.setStrokeColor(accent); canvas.setLineWidth(1.5)
    canvas.line(lm, ph - tm - 12, lm + 180, ph - tm - 12)
    canvas.setFillColor(text_mut); canvas.setFont(cf, 9)
    canvas.drawRightString(pw - rm, ph - tm - 8, "Page " + str(doc.page))
    canvas.setStrokeColor(border_c); canvas.setLineWidth(0.3)
    canvas.line(lm, bm + 5, pw - rm, bm + 5)
    canvas.setFillColor(text_mut); canvas.setFont(cf, 7)
    canvas.drawString(lm, bm - 8, "Smart Dashboard v2.0  |  " + generated_at)
    canvas.drawRightString(pw - rm, bm - 8, "Confidential")
    canvas.restoreState()

def mw2(txt, font, sz):
    return pdfmetrics.stringWidth(str(txt), font, sz)

elements = []
elements.append(Spacer(1, 20)); elements.append(PageBreak())
elements.append(Spacer(1, 20)); elements.append(PageBreak())
elements.append(Spacer(1, 20)); elements.append(PageBreak())

dt = pdf_df_raw.fillna("").astype(str)
col_list = list(pdf_df_raw.columns)
col_widths_pdf = []
for col in col_list:
    vals = dt[col].tolist()
    if len(vals) > 100: vals = vals[::max(1, len(vals) // 100)]
    measured = [mw2(v, cf, 7.5) for v in vals if v]
    mx = max([mw2(col, hf, 9)] + measured) if measured else mw2(col, hf, 9)
    col_widths_pdf.append(max(55, min(200, mx + 12)))
tot_w = sum(col_widths_pdf)
if tot_w > uw: col_widths_pdf = [w * uw / tot_w for w in col_widths_pdf]

def chunk(cols, widths, max_w):
    groups, cur, cw3 = [], [], 0.0
    for c, w in zip(cols, widths):
        if cur and cw3 + w > max_w: groups.append(cur); cur = [c]; cw3 = w
        else: cur.append(c); cw3 += w
    if cur: groups.append(cur)
    return groups

ss2 = getSampleStyleSheet()
hs = ParagraphStyle("RH", parent=ss2["Normal"], fontName=hf, fontSize=8, leading=10, textColor=header_fg, alignment=TA_LEFT)
cs = ParagraphStyle("RC", parent=ss2["Normal"], fontName=cf, fontSize=7, leading=9, alignment=TA_LEFT)

_page_callbacks = {2: draw_summary, 3: draw_suggestions}
def draw_later(canvas, doc):
    fn = _page_callbacks.get(doc.page, draw_data_page)
    fn(canvas, doc)

groups = chunk(col_list, col_widths_pdf, uw)
for g_i, g_cols in enumerate(groups):
    g_w = [col_widths_pdf[col_list.index(c)] for c in g_cols]
    rows_data = [[Paragraph(str(c), hs) for c in g_cols]]
    for rv in pdf_df_raw.fillna("").astype(str).values.tolist():
        row_cells = []
        for c in g_cols:
            idx = col_list.index(c); val = str(rv[idx]); style = cs
            if c == "Due Status":
                if "Overdue" in val: style = ParagraphStyle("OVR" + str(g_i), parent=cs, textColor=overdue_c, fontName=hf)
                elif "7d" in val: style = ParagraphStyle("D7" + str(g_i), parent=cs, textColor=warn_c, fontName=hf)
                elif "15d" in val: style = ParagraphStyle("D15" + str(g_i), parent=cs, textColor=orange_c)
                elif "30d" in val: style = ParagraphStyle("D30" + str(g_i), parent=cs, textColor=info_c)
                elif "60d" in val: style = ParagraphStyle("D60" + str(g_i), parent=cs, textColor=safe_c)
            elif c == "Days Until Maturity":
                try:
                    dv = int(val)
                    if dv > 0: style = ParagraphStyle("DVp" + str(g_i), parent=cs, textColor=overdue_c, fontName=hf)
                    elif dv >= -7: style = ParagraphStyle("DV7" + str(g_i), parent=cs, textColor=warn_c, fontName=hf)
                    elif dv >= -15: style = ParagraphStyle("DV15" + str(g_i), parent=cs, textColor=orange_c)
                    elif dv >= -30: style = ParagraphStyle("DV30" + str(g_i), parent=cs, textColor=info_c)
                    else: style = ParagraphStyle("DV60" + str(g_i), parent=cs, textColor=safe_c)
                except: pass
            row_cells.append(Paragraph(val, style))
        rows_data.append(row_cells)
    tbl = LongTable(rows_data, repeatRows=1, colWidths=g_w, hAlign="LEFT", splitByRow=1, spaceBefore=6, spaceAfter=6)
    ts = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"), ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
    ])
    tbl.setStyle(ts)
    elements.append(tbl)
    if g_i < len(groups) - 1: elements.append(PageBreak())

doc = SimpleDocTemplate("sample_due_date_report.pdf", pagesize=landscape(A3),
                        leftMargin=lm, rightMargin=rm, topMargin=tm + 20, bottomMargin=bm)
doc.build(elements, onFirstPage=draw_cover, onLaterPages=draw_later)
print("PDF generated: sample_due_date_report.pdf")

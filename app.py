"""
SAM Factsheet Updater — Web App
================================
Run with:  python app.py
Then open: http://localhost:5000
"""

from flask import Flask, request, send_file, jsonify, render_template
import zipfile, io, shutil, os, tempfile
import numpy as np
import pandas as pd
import openpyxl
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


# ─────────────────────────────────────────────────────────────────────────────
# SECTOR MAP
# ─────────────────────────────────────────────────────────────────────────────
SECTOR_MAP = {
    "GOOGL": "Communications",
    "AVGO":  "Technology",
    "JPM":   "Financials",
    "CEG":   "Utilities",
    "MSFT":  "Technology",
    "LLY":   "Healthcare",
    "CSCO":  "Technology",
    "BRKB":  "Financials",
    "GD":    "Industrials",
    "LMT":   "Industrials",
    "HON":   "Industrials",
    "ICE":   "Financials",
    "JNJ":   "Healthcare",
    "EMR":   "Industrials",
    "ABBNY": "Industrials",
    "AON":   "Financials",
    "BHP":   "Materials",
    "CVX":   "Energy",
    "DIS":   "Communications",
    "KMI":   "Energy",
    "MCHP":  "Technology",
    "MELI":  "Cons. Discretionary",
    "NSRGY": "Consumer Staples",
    "PM":    "Consumer Staples",
}


# ─────────────────────────────────────────────────────────────────────────────
# DATE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _to_short_date(long_date):
    month_map = {
        "January": "1", "February": "2", "March": "3",
        "April": "4", "May": "5", "June": "6",
        "July": "7", "August": "8", "September": "9",
        "October": "10", "November": "11", "December": "12"
    }
    for month, num in month_map.items():
        if long_date.startswith(month + " "):
            rest = long_date[len(month)+1:].replace(",", "").split()
            if len(rest) == 2:
                return f"{num}/{rest[0]}/{rest[1]}"
    return long_date

def _to_upper_date(long_date):
    parts = long_date.split(" ", 1)
    return parts[0].upper() + (" " + parts[1] if len(parts) > 1 else "")


# ─────────────────────────────────────────────────────────────────────────────
# PROCESSING
# ─────────────────────────────────────────────────────────────────────────────
def process_holdings(holdings_bytes, model_bytes):
    holdings = pd.read_excel(io.BytesIO(holdings_bytes))
    model    = pd.read_excel(io.BytesIO(model_bytes))

    model_tickers = set(model["Ticker"].astype(str).str.strip().str.upper())
    holdings["Ticker_clean"] = holdings["Ticker"].astype(str).str.strip().str.upper()

    filtered = holdings[
        holdings["Ticker_clean"].isin(model_tickers) &
        (holdings["Asset Category"] == "Equity")
    ].copy()

    if filtered.empty:
        raise ValueError(
            "No matching stocks found after filtering. "
            "Check that the holdings file has Asset Category and Ticker columns "
            "and the model file has a Ticker column."
        )

    stocks = filtered.groupby(["Ticker_clean", "ProductName"], as_index=False)["Total Value"].sum()
    total = stocks["Total Value"].sum()
    stocks["Weight_Pct"] = (stocks["Total Value"] / total * 100).round(2)
    top10 = stocks.sort_values("Weight_Pct", ascending=False).head(10).reset_index(drop=True)

    stocks["Sector"] = stocks["Ticker_clean"].map(SECTOR_MAP).fillna("Other")
    sector_weights = (
        stocks.groupby("Sector")["Weight_Pct"].sum()
        .reset_index().rename(columns={"Weight_Pct": "Sector_Pct"})
        .sort_values("Sector_Pct", ascending=False).reset_index(drop=True)
    )
    sector_weights["Sector_Pct"] = sector_weights["Sector_Pct"].round(1)
    return top10, sector_weights


def update_pie_chart_bytes(pptx_bytes, sector_weights_df):
    xlsx_key = "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"

    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zin:
        names = zin.namelist()
        if xlsx_key not in names:
            raise ValueError("Could not find the embedded chart workbook in the PPTX.")
        xlsx_data = zin.read(xlsx_key)
        all_files = {name: zin.read(name) for name in names}

    wb = openpyxl.load_workbook(io.BytesIO(xlsx_data))
    if "Sectors" not in wb.sheetnames:
        raise ValueError(f"Embedded workbook has no 'Sectors' sheet. Found: {wb.sheetnames}")

    ws = wb["Sectors"]
    sectors = sector_weights_df["Sector"].tolist()
    pcts    = sector_weights_df["Sector_Pct"].tolist()

    for col in range(1, 20):
        ws.cell(row=1, column=col).value = None
        ws.cell(row=2, column=col).value = None
    for i, (sector, pct) in enumerate(zip(sectors, pcts)):
        ws.cell(row=1, column=i + 1).value = f"{sector}  {pct}%"
        ws.cell(row=2, column=i + 1).value = round(pct / 100, 4)

    buf = io.BytesIO()
    wb.save(buf)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, buf.getvalue() if name == xlsx_key else data)

    pptx_bytes = out_buf.getvalue()

    chart_key = "ppt/charts/chart1.xml"
    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zin:
        if chart_key not in zin.namelist():
            return pptx_bytes
        chart_xml = zin.read(chart_key)
        all_files = {name: zin.read(name) for name in zin.namelist()}

    root = etree.fromstring(chart_xml)
    for strCache in root.iter(qn("c:strCache")):
        for pt in strCache.findall(qn("c:pt")): strCache.remove(pt)
        ptc = strCache.find(qn("c:ptCount"))
        if ptc is not None: ptc.set("val", str(len(sectors)))
        for i, (s, p) in enumerate(zip(sectors, pcts)):
            pt = etree.SubElement(strCache, qn("c:pt"))
            pt.set("idx", str(i))
            etree.SubElement(pt, qn("c:v")).text = f"{s}  {p}%"

    for numCache in root.iter(qn("c:numCache")):
        for pt in numCache.findall(qn("c:pt")): numCache.remove(pt)
        ptc = numCache.find(qn("c:ptCount"))
        if ptc is not None: ptc.set("val", str(len(pcts)))
        for i, p in enumerate(pcts):
            pt = etree.SubElement(numCache, qn("c:pt"))
            pt.set("idx", str(i))
            etree.SubElement(pt, qn("c:v")).text = str(round(p / 100, 4))

    all_files[chart_key] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items(): zout.writestr(name, data)
    return out_buf.getvalue()


def _set_para_text(para, text):
    if para.runs:
        para.runs[0].text = str(text)
        for run in para.runs[1:]: run.text = ""
    else:
        para.text = str(text)


def replace_text(slide, old, new):
    count = 0
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
                    count += 1
    return count


def run_update(pptx_bytes, holdings_bytes, model_bytes, old_date, new_date):
    log = []

    log.append("Processing holdings...")
    top10, sector_weights = process_holdings(holdings_bytes, model_bytes)

    log.append("Top 10 holdings:")
    for _, row in top10.iterrows():
        log.append(f"  {row['ProductName']:<42} {row['Weight_Pct']:.2f}%")

    log.append("\nSector weights:")
    for _, row in sector_weights.iterrows():
        log.append(f"  {row['Sector']:<28} {row['Sector_Pct']:.1f}%")

    log.append("\nUpdating pie chart...")
    pptx_bytes = update_pie_chart_bytes(pptx_bytes, sector_weights)
    log.append("Pie chart updated.")

    log.append("\nUpdating Top 10 and dates...")
    prs = Presentation(io.BytesIO(pptx_bytes))
    slide = prs.slides[0]

    shapes_4 = sorted(
        [s for s in slide.shapes if s.name == "object 4" and s.has_text_frame],
        key=lambda s: s.left
    )
    if len(shapes_4) < 2:
        raise RuntimeError(f"Expected 2 Top 10 text boxes, found {len(shapes_4)}.")

    left_paras  = shapes_4[0].text_frame.paragraphs
    right_paras = shapes_4[1].text_frame.paragraphs

    for i in range(min(5, len(left_paras) - 1)):
        _set_para_text(left_paras[i + 1], top10.iloc[i]["ProductName"])
    top5_pct = int(round(top10.head(5)["Weight_Pct"].sum(), 0))
    _set_para_text(right_paras[0], f"MV {top5_pct}%")
    for i in range(min(5, len(right_paras) - 1)):
        _set_para_text(right_paras[i + 1], top10.iloc[5 + i]["ProductName"])
    log.append(f"Top 10 updated. MV of top 5 = {top5_pct}%")

    old_short = _to_short_date(old_date)
    new_short = _to_short_date(new_date)
    old_upper = _to_upper_date(old_date)
    new_upper = _to_upper_date(new_date)

    n1 = replace_text(slide, old_short, new_short)
    n2 = replace_text(slide, old_upper, new_upper)
    n3 = replace_text(slide, old_date,  new_date)
    log.append(f"Dates updated ({n1+n2+n3} replacements)")

    out_buf = io.BytesIO()
    prs.save(out_buf)

    return out_buf.getvalue(), log


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/update", methods=["POST"])
def update():
    try:
        pptx_file     = request.files.get("pptx")
        holdings_file = request.files.get("holdings")
        model_file    = request.files.get("model")
        old_date      = request.form.get("old_date", "").strip()
        new_date      = request.form.get("new_date", "").strip()

        missing = []
        if not pptx_file:     missing.append("Factsheet (.pptx)")
        if not holdings_file: missing.append("Holdings (.xlsx)")
        if not model_file:    missing.append("Model (.xlsx)")
        if not old_date:      missing.append("Current date")
        if not new_date:      missing.append("New date")
        if missing:
            return jsonify({"error": f"Missing: {', '.join(missing)}"}), 400

        pptx_bytes     = pptx_file.read()
        holdings_bytes = holdings_file.read()
        model_bytes    = model_file.read()

        result_bytes, log = run_update(
            pptx_bytes, holdings_bytes, model_bytes, old_date, new_date
        )

        original_name = pptx_file.filename or "factsheet.pptx"
        base, ext = os.path.splitext(original_name)
        output_name = base.replace(old_date.split(",")[0].split(" ")[-1],
                                   new_date.split(",")[0].split(" ")[-1]) + ext
        if output_name == original_name:
            output_name = base + "_updated" + ext

        return jsonify({
            "log": log,
            "filename": output_name,
            "data": result_bytes.hex()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  SAM Factsheet Updater")
    print(f"  Open your browser to: http://localhost:{port}\n")
    app.run(debug=False, host="0.0.0.0", port=port)

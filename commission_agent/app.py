import streamlit as st
import pdfplumber
import openpyxl
import json
import io
import re
import pandas as pd
from pathlib import Path
from groq import Groq

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Commission Agent — MAR1",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #181c26; }
[data-testid="stSidebar"] * { color: #e8eaf0 !important; }
.main-header {
    background: linear-gradient(135deg, #4f8ef7, #7c3aed);
    padding: 18px 24px; border-radius: 12px; margin-bottom: 24px;
}
.main-header h1 { color: white; margin: 0; font-size: 22px; }
.main-header p  { color: rgba(255,255,255,0.75); margin: 0; font-size: 13px; }
.ok-box   { background: rgba(34,197,94,0.08);  border: 1px solid rgba(34,197,94,0.2);  border-radius:10px; padding:12px 16px; margin-bottom:8px; }
.err-box  { background: rgba(239,68,68,0.08);  border: 1px solid rgba(239,68,68,0.25); border-radius:10px; padding:14px 16px; margin-bottom:10px; }
.warn-box { background: rgba(245,158,11,0.08); border: 1px solid rgba(245,158,11,0.25);border-radius:10px; padding:14px 16px; margin-bottom:10px; }
</style>
""", unsafe_allow_html=True)

# ─── Session State ─────────────────────────────────────────────────────────────
if "analysis"     not in st.session_state: st.session_state.analysis     = None
if "chat_history" not in st.session_state: st.session_state.chat_history = []
if "api_key"      not in st.session_state:
    try:    st.session_state.api_key = st.secrets["GROQ_API_KEY"]
    except: st.session_state.api_key = ""

GROQ_MODEL = "llama-3.3-70b-versatile"

# ─── Helpers ──────────────────────────────────────────────────────────────────
def fmt(n):
    try:    return f"${float(n):,.2f}"
    except: return "$0.00"

def fmtpct(n):
    try:    return f"{float(n)*100:.1f}%"
    except: return "—"

# ─── PDF Parser ───────────────────────────────────────────────────────────────
def parse_pdf(file_bytes: bytes) -> dict:
    """
    Extract from PDF:
      Invoice#, Inv. Dt, PO#, Customer#, Name, Sales Amt (UnitCost), State, Commission
    Returns: {"month_label": str, "rows": [...]}
    """
    text = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text.append(t)
    full_text = "\n".join(text)
    return {"raw_text": full_text}


# ─── Excel Parser ─────────────────────────────────────────────────────────────
def parse_xlsx(file_bytes: bytes) -> dict:
    """
    Extract:
      - COMM REPORT sheet  → summary (Part I amount/commission, Part II totals)
      - DISTY SALES sheet  → commission rate per AB Inv No
      - INTEGRA/distributor sheets → line items
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    sheets = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
        sheets[sheet] = rows
    return sheets


def sheets_to_text(sheets: dict) -> str:
    lines = []
    for name, rows in sheets.items():
        lines.append(f"=== Sheet: {name} ===")
        for row in rows:
            vals = [str(v) if v is not None else "" for v in row]
            if any(v.strip() for v in vals):
                lines.append("\t".join(vals))
    return "\n".join(lines)


# ─── Groq Analysis ────────────────────────────────────────────────────────────
def run_analysis(pdf_text: str, xlsx_sheets: dict, api_key: str) -> dict:
    client = Groq(api_key=api_key)
    xlsx_text = sheets_to_text(xlsx_sheets)

    prompt = f"""
=== PDF FILE: Sales Commission Report from Payment ===
{pdf_text}

=== EXCEL FILE (all sheets) ===
{xlsx_text}

---
You are a commission data extraction agent for MARCTECH2, INC. (Sales Rep MAR1).

TASK: Extract structured data from the above files using these EXACT rules:

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PART I — REGULAR SALES (extract from PDF)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
For each transaction row in the PDF, extract:
  - invoice_number     ← "Invoice#" column
  - invoice_date       ← "Inv. Dt" column
  - po_number          ← "PO #" column
  - customer_number    ← "Customer#" column
  - customer_name      ← "Name" column
  - unit_cost          ← "Sales Amt" column (the dollar amount, NOT the percentage)
  - state              ← "State" column (2-letter state code)
  - commission         ← "Commission" column (the dollar amount, NOT the 5.00% figure)

Also extract:
  - part1_total_amount     ← "Sales Rep Total" payment amount
  - part1_total_commission ← "Sales Rep Total" commission amount
  - month_label            ← e.g. "December 2025" (from the date range in PDF header)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PART II — DISTRIBUTOR SALES (extract from Excel)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Step 1 — Get commission RATE from the DISTY SALES sheet:
  - Find the "RATE" column value for each AB INV NO.
  - This is the commission percentage (e.g. 0.05 = 5%)

Step 2 — Get line items from the INTEGRA sheet (or any distributor sheet):
For each row, extract:
  - state          ← "ST" column
  - zip            ← "Zip" column
  - part_number    ← "Item No." column
  - quantity       ← "Qty" column (integer)
  - unit_sale      ← "AB PRICE" column (this is the price per unit)
  - invoice_number ← "AB. Inv. No." column
  - customer_name  ← "Company" column
  - city           ← "City" column

Step 3 — Calculate commission per row:
  - total_amount = quantity × unit_sale
  - commission   = total_amount × rate   (rate from DISTY SALES for that invoice number)

Also extract:
  - part2_month_label        ← month label from DISTY SALES sheet header (e.g. "November 2025")
  - part2_total_amount       ← sum of all total_amounts
  - part2_total_commission   ← sum of all commissions

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COMM REPORT SUMMARY (from COMM REPORT sheet)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  - comm_report_part1_amount      ← Part I AMOUNT value
  - comm_report_part1_commission  ← Part I COMMISSION value
  - comm_report_part2_amount      ← Part II AMOUNT value
  - comm_report_part2_commission  ← Part II COMMISSION value
  - comm_report_total_commission  ← TOTAL COMMISSION value

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RECONCILIATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Check and report:
  1. Does PDF Part I commission match COMM REPORT Part I commission? (tolerance $0.01)
  2. Does calculated Part II commission match COMM REPORT Part II commission? (tolerance $0.01)
  3. Any calculation errors in the distributor rows?

Return this EXACT JSON (no markdown, no explanation):

{{
  "month_label": "December 2025",
  "part1": {{
    "month": "December 2025",
    "total_amount": 6090.00,
    "total_commission": 304.50,
    "rows": [
      {{
        "invoice_number": "451014081",
        "invoice_date": "11/12/25",
        "po_number": "5034774",
        "customer_number": "DEL003",
        "customer_name": "DELTA CONTROLS, INC.",
        "unit_cost": 6090.00,
        "state": "BC",
        "commission": 304.50
      }}
    ]
  }},
  "part2": {{
    "month": "November 2025",
    "rate": 0.05,
    "total_amount": 572.00,
    "total_commission": 28.60,
    "rows": [
      {{
        "invoice_number": "5034813",
        "customer_name": "SYMETRIX, INC.",
        "city": "MUKILTEO",
        "state": "WA",
        "zip": "98275",
        "part_number": "BA-10G1UD",
        "quantity": 1000,
        "unit_sale": 0.572,
        "total_amount": 572.00,
        "commission": 28.60
      }}
    ]
  }},
  "comm_report": {{
    "part1_amount": 6090.00,
    "part1_commission": 304.50,
    "part2_amount": 572.00,
    "part2_commission": 28.60,
    "total_commission": 333.10
  }},
  "reconciliation": {{
    "part1_match": true,
    "part2_match": true,
    "issues": [],
    "status": "ok"
  }}
}}
"""

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=4000,
        temperature=0.1,
    )

    raw = response.choices[0].message.content
    raw = re.sub(r"```json|```", "", raw).strip()
    match = re.search(r'\{.*\}', raw, re.DOTALL)
    if match:
        raw = match.group()
    return json.loads(raw)


def chat_with_agent(question: str, context: dict, history: list, api_key: str) -> str:
    client = Groq(api_key=api_key)
    system = (
        "You are a commission analysis assistant for MARCTECH2, INC. (Sales Rep MAR1). "
        "Answer questions about the commission data clearly. Use dollar amounts and be specific.\n\n"
        f"DATA:\n{json.dumps(context, indent=2)}"
    ) if context else "No data analyzed yet. Ask the user to upload files first."

    messages = [{"role": "system", "content": system}]
    for h in history[-6:]:
        messages.append({"role": "user",      "content": h["user"]})
        messages.append({"role": "assistant", "content": h["assistant"]})
    messages.append({"role": "user", "content": question})

    response = client.chat.completions.create(
        model=GROQ_MODEL, messages=messages, max_tokens=800, temperature=0.3,
    )
    return response.choices[0].message.content


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    st.markdown("Free API key → [console.groq.com](https://console.groq.com)")
    api_key = st.text_input("Groq API Key", type="password",
                             value=st.session_state.api_key, placeholder="gsk_…")
    if api_key:
        st.session_state.api_key = api_key

    st.markdown("---")
    st.markdown("## 📂 Upload Files")

    pdf_file  = st.file_uploader("PDF — Comm_from_Payment",  type=["pdf"],         key="pdf_up")
    xlsx_file = st.file_uploader("Excel — Comm_Report",       type=["xlsx", "xls"], key="xl_up")

    st.markdown("---")
    run = st.button("⚡ Analyze & Reconcile", use_container_width=True,
                    disabled=not (pdf_file and xlsx_file and st.session_state.api_key))
    if not st.session_state.api_key:
        st.caption("⚠️ Enter Groq API key above.")
    elif not (pdf_file and xlsx_file):
        st.caption("⚠️ Upload both PDF and Excel files.")

# ─── Run Analysis ─────────────────────────────────────────────────────────────
if run and pdf_file and xlsx_file:
    with st.spinner("Reading files…"):
        pdf_bytes  = pdf_file.read()
        xlsx_bytes = xlsx_file.read()
        pdf_data   = parse_pdf(pdf_bytes)
        xlsx_data  = parse_xlsx(xlsx_bytes)

    with st.spinner("Extracting & reconciling with AI…"):
        try:
            st.session_state.analysis = run_analysis(
                pdf_data["raw_text"], xlsx_data, st.session_state.api_key
            )
            st.success("Done!")
        except Exception as e:
            st.error(f"Failed: {e}")

# ─── Main UI ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📊 Commission Intelligence Agent</h1>
  <p>MARCTECH2, INC. · American Bright Optoelectronics Corp. · Sales Rep: MAR1 · Powered by Groq (Free)</p>
</div>
""", unsafe_allow_html=True)

data = st.session_state.analysis

if not data:
    st.info("👈 Upload the PDF and Excel files in the sidebar, then click **Analyze & Reconcile**.")

    with st.expander("📋 Field Mapping Reference"):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**From PDF → Part I (Regular Sales)**")
            st.markdown("""
| PDF Column | Field |
|---|---|
| Invoice# | Invoice Number |
| Inv. Dt | Invoice Date |
| PO # | PO Number |
| Customer# | Customer Number |
| Name | Customer Name |
| Sales Amt | Unit Cost |
| State | State |
| Commission | Commission ($) |
""")
        with col2:
            st.markdown("**From Excel INTEGRA sheet → Part II (Distributor Sales)**")
            st.markdown("""
| Excel Column | Field |
|---|---|
| ST | State |
| Zip | Zip |
| Item No. | Part Number |
| Qty | Quantity |
| AB PRICE | Unit Sale |
| AB. Inv. No. | Invoice Number |
| Company | Customer Name |
| City | City |
| DISTY SALES → RATE | Commission % |
""")
        st.info("**Commission formula:** Qty × AB PRICE = Total Amount → × Rate = Commission")
else:
    p1   = data.get("part1", {})
    p2   = data.get("part2", {})
    cr   = data.get("comm_report", {})
    rec  = data.get("reconciliation", {})

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Summary", "📄 Part I — Regular Sales",
        "🏪 Part II — Distributor Sales", "💬 Ask AI"
    ])

    # ── Tab 1: Summary ──────────────────────────────────────────────────────
    with tab1:
        st.markdown(f"### Commission Statement — {data.get('month_label','')}")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Part I — Regular Sales",
                      fmt(cr.get("part1_amount", 0)),
                      f"Commission: {fmt(cr.get('part1_commission', 0))}")
        with c2:
            st.metric("Part II — Distributor Sales",
                      fmt(cr.get("part2_amount", 0)),
                      f"Commission: {fmt(cr.get('part2_commission', 0))}")
        with c3:
            st.metric("💰 Total Commission",
                      fmt(cr.get("total_commission", 0)))

        st.markdown("---")
        st.markdown("### Reconciliation")

        status = rec.get("status", "ok")
        issues = rec.get("issues", [])

        if status == "ok":
            st.markdown('<div class="ok-box">✅ <strong>All figures reconcile perfectly</strong> — PDF, Excel, and calculated values all match.</div>', unsafe_allow_html=True)
        else:
            for issue in issues:
                st.markdown(f'<div class="err-box">⚠️ {issue}</div>', unsafe_allow_html=True)

        recon_rows = [
            {"Check": "Part I PDF Commission",     "Value": fmt(p1.get("total_commission")),          "COMM REPORT", fmt(cr.get("part1_commission")),  "Match": "✅" if rec.get("part1_match") else "❌"},
            {"Check": "Part II Calculated Commission", "Value": fmt(p2.get("total_commission")),       "COMM REPORT": fmt(cr.get("part2_commission")), "Match": "✅" if rec.get("part2_match") else "❌"},
        ]
        # Simple two-row table
        col_a, col_b, col_c, col_d = st.columns([3,2,2,1])
        col_a.markdown("**Check**"); col_b.markdown("**Calculated**"); col_c.markdown("**COMM REPORT**"); col_d.markdown("**Match**")
        col_a.write("Part I — PDF Commission");         col_b.write(fmt(p1.get("total_commission"))); col_c.write(fmt(cr.get("part1_commission"))); col_d.write("✅" if rec.get("part1_match") else "❌")
        col_a.write("Part II — Distributor Commission");col_b.write(fmt(p2.get("total_commission"))); col_c.write(fmt(cr.get("part2_commission"))); col_d.write("✅" if rec.get("part2_match") else "❌")

    # ── Tab 2: Part I Regular Sales ─────────────────────────────────────────
    with tab2:
        st.markdown(f"### Part I — Regular Sales &nbsp; `{p1.get('month','')}`")
        rows = p1.get("rows", [])
        if rows:
            df1 = pd.DataFrame(rows)
            # Rename columns for display
            rename_map = {
                "invoice_number":  "Invoice #",
                "invoice_date":    "Invoice Date",
                "po_number":       "PO #",
                "customer_number": "Customer #",
                "customer_name":   "Customer Name",
                "unit_cost":       "Unit Cost",
                "state":           "State",
                "commission":      "Commission",
            }
            df1 = df1.rename(columns=rename_map)
            # Format money columns
            for col in ["Unit Cost", "Commission"]:
                if col in df1.columns:
                    df1[col] = df1[col].apply(lambda x: f"${float(x):,.2f}" if x else "—")
            st.dataframe(df1, use_container_width=True, hide_index=True)
        else:
            st.warning("No Part I rows extracted.")

        st.markdown("---")
        c1, c2 = st.columns(2)
        c1.metric("Total Sales Amount",     fmt(p1.get("total_amount")))
        c2.metric("Total Commission (5%)",  fmt(p1.get("total_commission")))

    # ── Tab 3: Part II Distributor Sales ────────────────────────────────────
    with tab3:
        st.markdown(f"### Part II — Distributor Sales &nbsp; `{p2.get('month','')}`")
        st.caption(f"Commission rate from DISTY SALES sheet: **{fmtpct(p2.get('rate', 0.05))}**")
        st.caption("Formula: **Qty × AB PRICE = Total Amount → × Rate = Commission**")

        rows2 = p2.get("rows", [])
        if rows2:
            df2 = pd.DataFrame(rows2)
            rename2 = {
                "invoice_number": "Invoice #",
                "customer_name":  "Customer Name",
                "city":           "City",
                "state":          "State",
                "zip":            "Zip",
                "part_number":    "Part Number",
                "quantity":       "Qty",
                "unit_sale":      "AB PRICE (Unit Sale)",
                "total_amount":   "Total Amount",
                "commission":     "Commission",
            }
            df2 = df2.rename(columns=rename2)
            for col in ["AB PRICE (Unit Sale)", "Total Amount", "Commission"]:
                if col in df2.columns:
                    df2[col] = df2[col].apply(lambda x: f"${float(x):,.4f}" if col == "AB PRICE (Unit Sale)" else f"${float(x):,.2f}" if x else "—")
            st.dataframe(df2, use_container_width=True, hide_index=True)
        else:
            st.warning("No Part II rows extracted.")

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("Rate",                  fmtpct(p2.get("rate", 0.05)))
        c2.metric("Total Distributor Amount", fmt(p2.get("total_amount")))
        c3.metric("Total Commission",      fmt(p2.get("total_commission")))

    # ── Tab 4: Ask AI ───────────────────────────────────────────────────────
    with tab4:
        st.markdown("Ask anything about this commission statement.")
        for h in st.session_state.chat_history:
            with st.chat_message("user"):      st.write(h["user"])
            with st.chat_message("assistant"): st.write(h["assistant"])

        question = st.chat_input("e.g. What is the total commission? Any discrepancies?")
        if question:
            if not st.session_state.api_key:
                st.error("Enter Groq API key in the sidebar.")
            else:
                with st.chat_message("user"): st.write(question)
                with st.chat_message("assistant"):
                    with st.spinner("Thinking…"):
                        reply = chat_with_agent(question, data, st.session_state.chat_history, st.session_state.api_key)
                    st.write(reply)
                st.session_state.chat_history.append({"user": question, "assistant": reply})

import io
import re
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(
    page_title="Outstanding Tracker Dashboard",
    page_icon="📊",
    layout="wide",
)

# ---------- HELPER FUNCTIONS ----------

def find_header_row(file_bytes: bytes, required_cols=("Type", "Serial")) -> int:
    """Excel sheet mein automatically header row ka index dhundhta hai."""
    preview = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=20, dtype=str)
    for idx, row in preview.iterrows():
        vals = {str(v).strip().lower() for v in row.dropna().values}
        if all(c.lower() in vals for c in required_cols):
            return int(idx)
    return 6


def read_register(file_bytes: bytes) -> pd.DataFrame:
    header_idx = find_header_row(file_bytes)
    df = pd.read_excel(io.BytesIO(file_bytes), header=header_idx, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def normalize_bill_key(doc_type: str, number: str) -> str:
    t = str(doc_type).strip().upper() if pd.notna(doc_type) else ""
    n = str(number).strip().upper() if pd.notna(number) else ""
    if not t or not n or t == "NAN" or n == "NAN":
        return ""
    n_clean = n.lstrip("0") or "0"
    return f"{t}-{n_clean}"


def extract_paid_bills_from_ledger(ml_df: pd.DataFrame) -> set:
    """Multi Ledger ke ICICI rows ki Narration se saare paid Bill No(s) nikalta hai."""
    paid_keys = set()
    if "Particulars" not in ml_df.columns or "Narration" not in ml_df.columns:
        return paid_keys

    icici_mask = ml_df["Particulars"].astype(str).str.contains("ICICI", case=False, na=False)
    for narr in ml_df.loc[icici_mask, "Narration"].dropna():
        clean = str(narr).replace("_x000D_", " ").replace("\r", " ").replace("\n", " ")
        match = re.search(r"Bill\s*No\(s\)\s*:\s*(.*)", clean, re.IGNORECASE)
        if not match:
            continue
        raw_part = match.group(1)
        tokens = [tok.strip().upper() for tok in re.split(r"[,;\s]+", raw_part) if tok.strip()]
        for tok in tokens:
            paid_keys.add(tok)
            if "-" in tok:
                p1, p2 = tok.split("-", 1)
                paid_keys.add(f"{p1}-{p2.lstrip('0') or '0'}")
    return paid_keys


def extract_parent_party(party_name: str) -> str:
    """Branch/unit naam hata kar Main Parent Company (Group) naam nikalta hai (e.g. CIPLA)."""
    if not isinstance(party_name, str) or not party_name.strip():
        return "UNKNOWN"
    s = party_name.strip().upper()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.split(r"[-/,]", s)[0]
    s = re.sub(
        r"\b(PVT\.?|PRIVATE|LTD\.?|LIMITED|LLP|INC\.?|CORP\.?|CORPORATION|"
        r"UNIT[\s\-]*[IVX0-9]+|PHARMA(CEUTICALS?)?|LABORATORIES|LABS?|"
        r"INDUSTRIES|INDIA|SOLUTIONS|HEALTHCARE|LIFESCIENCES|BIOSCIENCES|"
        r"FORMULATION|R&D|CWH)\b",
        " ",
        s,
    )
    s = re.sub(r"[^A-Z0-9\s&]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    words = s.split()
    if not words:
        return party_name.strip().upper()
    return " ".join(words[:2]) if len(words) > 1 and len(words[0]) <= 3 else words[0]


def get_overdue_bucket(days: int) -> str:
    if pd.isna(days):
        return "0-30 Days"
    if days <= 30:
        return "0-30 Days"
    if days <= 45:
        return "31-45 Days"
    if days <= 90:
        return "46-90 Days"
    return "90+ Days"


def process_register_outstanding(
    reg_df: pd.DataFrame,
    paid_keys: set,
    as_on_date: date,
    is_purchase: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    main_df = reg_df.copy()

    party_col = "Party Name"
    has_party = main_df[party_col].notna() & (main_df[party_col].astype(str).str.strip() != "")
    main_df["_Invoice_Group"] = has_party.cumsum()

    valid_df = main_df[main_df["_Invoice_Group"] > 0].copy()
    inv_headers = valid_df[has_party[valid_df.index]].copy()

    inv_headers["Bill No"] = (
        inv_headers["Type"].fillna("").str.strip()
        + "-"
        + inv_headers["Serial"].fillna("").str.strip()
    )

    key_serial = [
        normalize_bill_key(t, s)
        for t, s in zip(inv_headers["Type"].fillna(""), inv_headers["Serial"].fillna(""))
    ]

    if is_purchase and "Bill.No" in inv_headers.columns:
        key_billno = [
            normalize_bill_key(t, b)
            for t, b in zip(inv_headers["Type"].fillna(""), inv_headers["Bill.No"].fillna(""))
        ]
        is_paid = [(ks in paid_keys) or (kb in paid_keys) for ks, kb in zip(key_serial, key_billno)]
    else:
        is_paid = [(ks in paid_keys) for ks in key_serial]

    inv_headers["Payment Status"] = ["Paid" if p else "Outstanding" for p in is_paid]
    out_df = inv_headers[inv_headers["Payment Status"] == "Outstanding"].copy()

    date_col = "Bill.Date" if (is_purchase and "Bill.Date" in out_df.columns) else "Doc.Date"
    out_df["Invoice Date"] = pd.to_datetime(out_df[date_col], errors="coerce")
    out_df["Invoice Value"] = pd.to_numeric(out_df.get("Invoice Value", 0), errors="coerce").fillna(0.0)

    if (out_df["Invoice Value"] == 0).any() and "Item Value" in valid_df.columns:
        valid_df["_item_val_num"] = pd.to_numeric(valid_df["Item Value"], errors="coerce").fillna(0.0)
        grp_sums = valid_df.groupby("_Invoice_Group")["_item_val_num"].sum()
        zero_mask = out_df["Invoice Value"] == 0
        out_df.loc[zero_mask, "Invoice Value"] = out_df.loc[zero_mask, "_Invoice_Group"].map(grp_sums).fillna(0.0)

    as_on_ts = pd.Timestamp(as_on_date)
    out_df["Overdue Days"] = (as_on_ts - out_df["Invoice Date"]).dt.days.fillna(0).astype(int)
    out_df["Overdue Bucket"] = out_df["Overdue Days"].apply(get_overdue_bucket)

    out_df["Party Name (Branch)"] = out_df["Party Name"].astype(str).str.strip()
    out_df["Party Group"] = out_df["Party Name (Branch)"].apply(extract_parent_party)

    keep_cols = [
        "Party Group",
        "Party Name (Branch)",
        "Type",
        "Serial",
        "Bill No",
    ]
    if is_purchase and "Bill.No" in out_df.columns:
        keep_cols.append("Bill.No")
    keep_cols += [
        "Invoice Date",
        "City",
        "State",
        "Invoice Value",
        "Overdue Days",
        "Overdue Bucket",
    ]
    keep_cols = [c for c in keep_cols if c in out_df.columns]
    out_df = out_df[keep_cols].sort_values(
        by=["Party Group", "Party Name (Branch)", "Overdue Days"],
        ascending=[True, True, False],
    ).reset_index(drop=True)

    main_clean = main_df.drop(columns=["_Invoice_Group"])
    return main_clean, out_df


# ---------- CLEAN EXCEL CHART HELPER ----------

def make_clean_bar_chart(
    chart_type: str,
    title: str,
    data_ref: Reference,
    cats_ref: Reference,
    bar_hex: str,
    width: float,
    height: float,
) -> BarChart:
    ch = BarChart()
    ch.type = chart_type
    ch.grouping = "clustered"
    ch.title = title
    ch.title.overlay = False
    ch.legend = None
    ch.plot_vis_only = False
    ch.varyColors = False

    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)

    if ch.series:
        ch.series[0].graphicalProperties.solidFill = bar_hex
        ch.series[0].graphicalProperties.line.solidFill = bar_hex

    dl = DataLabelList()
    dl.showVal = True
    dl.showSerName = False
    dl.showCatName = False
    dl.showPercent = False
    dl.showLegendKey = False
    ch.dataLabels = dl

    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.tickLblPos = "low"
    ch.y_axis.majorGridlines = None
    ch.x_axis.majorGridlines = None

    ch.width = width
    ch.height = height
    return ch


# ---------- EXCEL DASHBOARD WITH NATIVE CHECKBOX MULTI-SELECT FILTERS ----------

def write_raw_sheet(wb: Workbook, title: str, df: pd.DataFrame):
    ws = wb.create_sheet(title=title)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=10)

    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row in enumerate(df.itertuples(index=False, name=None), start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=None if pd.isna(val) else val)

    ws.freeze_panes = "A2"


def write_outstanding_dashboard_sheet(
    wb: Workbook,
    title: str,
    df: pd.DataFrame,
    table_name: str,
    calc_sheet_name: str,
):
    ws = wb.create_sheet(title=title)
    ws.views.sheetView[0].showGridLines = False

    # Hidden calc sheet sirf 3 charts ke summary data ke liye
    cs = wb.create_sheet(title=calc_sheet_name)
    cs.sheet_state = "hidden"

    # Styles
    banner_fill = PatternFill("solid", fgColor="1F4E78")
    banner_font = Font(bold=True, size=15, color="FFFFFF")
    kpi_label_font = Font(bold=True, size=9, color="595959")
    kpi_val_font = Font(bold=True, size=13, color="1F4E78")
    kpi_amber_font = Font(bold=True, size=13, color="B45309")
    kpi_red_font = Font(bold=True, size=13, color="C00000")
    kpi_fill = PatternFill("solid", fgColor="EDF2F8")
    kpi_amber_fill = PatternFill("solid", fgColor="FEF3C7")
    kpi_red_fill = PatternFill("solid", fgColor="FCE4D6")
    tip_fill = PatternFill("solid", fgColor="E8F4FD")
    tip_font = Font(bold=True, size=9, color="1F4E78")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    amber_fill = PatternFill("solid", fgColor="FEF3C7")
    amber_font = Font(bold=True, color="B45309")
    red_fill = PatternFill("solid", fgColor="FCE4D6")
    red_font = Font(bold=True, color="C00000")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    n_cols = len(df.columns)
    last_tbl_ltr = get_column_letter(max(n_cols, 11))

    # Top Title Banner (Row 1-2)
    ws.merge_cells(f"A1:{last_tbl_ltr}2")
    ws["A1"] = f"  {title} Dashboard"
    ws["A1"].fill = banner_fill
    ws["A1"].font = banner_font
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18

    if len(df) == 0:
        ws["A4"] = "Koi pending invoice nahi hai."
        return

    # Table Layout: Row 26 = Header (with Excel Checkbox AutoFilter!), Row 27+ = Data
    start_row = 26
    r1 = start_row + 1
    r2 = start_row + len(df)

    # Helper Visible column (Table ke bahar, hidden) — AutoFilter lagte hi 1 ya 0 hota hai!
    vis_col_idx = n_cols + 2
    vis_ltr = get_column_letter(vis_col_idx)

    grp_ltr = get_column_letter(1 + df.columns.get_loc("Party Group"))
    br_ltr = get_column_letter(1 + df.columns.get_loc("Party Name (Branch)"))
    amt_idx = 1 + df.columns.get_loc("Invoice Value")
    amt_ltr = get_column_letter(amt_idx)
    days_ltr = get_column_letter(1 + df.columns.get_loc("Overdue Days"))
    bkt_ltr = get_column_letter(1 + df.columns.get_loc("Overdue Bucket"))

    vis_rng = f"${vis_ltr}${r1}:${vis_ltr}${r2}"
    amt_rng = f"${amt_ltr}${r1}:${amt_ltr}${r2}"
    day_rng = f"${days_ltr}${r1}:${days_ltr}${r2}"
    bkt_rng = f"${bkt_ltr}${r1}:${bkt_ltr}${r2}"

    # ---- 5 DYNAMIC KPI CARDS (Row 4-5) ----
    kpi_specs = [
        ("A", "B", "Total Outstanding (₹)",
         f"=SUMIF({vis_rng},1,{amt_rng})",
         "₹#,##0.00", kpi_fill, kpi_val_font),
        ("C", "D", "Pending Invoices",
         f"=COUNTIF({vis_rng},1)",
         "#,##0", kpi_fill, kpi_val_font),
        ("E", "F", "46-90 Days Overdue (₹)",
         f'=SUMIFS({amt_rng},{vis_rng},1,{bkt_rng},"46-90 Days")',
         "₹#,##0.00", kpi_amber_fill, kpi_amber_font),
        ("G", "I", "90+ Days Overdue (₹)",
         f'=SUMIFS({amt_rng},{vis_rng},1,{bkt_rng},"90+ Days")',
         "₹#,##0.00", kpi_red_fill, kpi_red_font),
        ("J", last_tbl_ltr, "Avg Overdue Days",
         f"=IFERROR(AVERAGEIF({vis_rng},1,{day_rng}),0)",
         '0" Days"', kpi_fill, kpi_val_font),
    ]
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 24
    for c1, c2, lbl, formula, num_fmt, fill_st, val_font in kpi_specs:
        ws.merge_cells(f"{c1}4:{c2}4")
        ws.merge_cells(f"{c1}5:{c2}5")
        ws[f"{c1}4"] = lbl
        ws[f"{c1}4"].font = kpi_label_font
        ws[f"{c1}4"].fill = fill_st
        ws[f"{c1}4"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{c1}5"] = formula
        ws[f"{c1}5"].font = val_font
        ws[f"{c1}5"].fill = fill_st
        ws[f"{c1}5"].number_format = num_fmt
        ws[f"{c1}5"].alignment = Alignment(horizontal="center", vertical="center")

    # ---- FILTER INSTRUCTION BAR (Row 25, right above Table headers) ----
    ws.merge_cells(f"A25:{last_tbl_ltr}25")
    ws["A25"] = (
        "  FILTER TIP: Niche table ke kisi bhi column header arrow (▼) par click karke "
        "Party Group, Branch, ya Overdue Bucket mein multiple select / deselect karein — "
        "upar ke KPIs aur teeno Graphs turant update honge!"
    )
    ws["A25"].fill = tip_fill
    ws["A25"].font = tip_font
    ws["A25"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[25].height = 20

    # ---- TABLE HEADERS (Row 26) ----
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ---- TABLE DATA (Row 27 onwards) ----
    bkt_offset = df.columns.get_loc("Overdue Bucket")
    for r_offset, row in enumerate(df.itertuples(index=False, name=None), start=1):
        r_num = start_row + r_offset
        bkt_val = row[bkt_offset]
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_num, column=c_idx)
            if isinstance(val, pd.Timestamp):
                cell.value = val.to_pydatetime() if pd.notna(val) else None
                cell.number_format = "dd-mmm-yyyy"
            else:
                cell.value = None if pd.isna(val) else val

            if c_idx == amt_idx:
                cell.number_format = "#,##0.00"
            cell.border = thin_border

            if bkt_val == "90+ Days":
                cell.fill = red_fill
                cell.font = red_font
            elif bkt_val == "46-90 Days":
                cell.fill = amber_fill
                cell.font = amber_font

        # Helper Visible Flag (1 when row is visible, 0 when filtered out by checkbox filter)
        ws.cell(row=r_num, column=vis_col_idx, value=f"=SUBTOTAL(103,A{r_num})")

    ws.column_dimensions[vis_ltr].hidden = True

    # ---- EXCEL TABLE (Enables Search + Select All + Checkboxes on EVERY column!) ----
    tbl_last_ltr = get_column_letter(n_cols)
    tab = Table(displayName=table_name, ref=f"A{start_row}:{tbl_last_ltr}{r2}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # ---- LIVE CHART DATA ON HIDDEN CALC SHEET (cs) ----
    sh_ref = f"'{title}'"
    vis_rng_cs = f"{sh_ref}!${vis_ltr}${r1}:${vis_ltr}${r2}"
    amt_rng_cs = f"{sh_ref}!${amt_ltr}${r1}:${amt_ltr}${r2}"
    bkt_rng_cs = f"{sh_ref}!${bkt_ltr}${r1}:${bkt_ltr}${r2}"
    grp_rng_cs = f"{sh_ref}!${grp_ltr}${r1}:${grp_ltr}${r2}"
    br_rng_cs = f"{sh_ref}!${br_ltr}${r1}:${br_ltr}${r2}"

    # Chart 1 Data (A1:B5): Overdue Bucket-wise Outstanding
    cs["A1"] = "Overdue Bucket"
    cs["B1"] = "Outstanding"
    for i, b_name in enumerate(["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"], start=2):
        cs[f"A{i}"] = b_name
        cs[f"B{i}"] = f"=SUMIFS({amt_rng_cs},{vis_rng_cs},1,{bkt_rng_cs},A{i})"
        cs[f"B{i}"].number_format = "#,##0"

    # Chart 2 Data (D1:E11): Top 10 Parties (Party Group)
    cs["D1"] = "Party Group"
    cs["E1"] = "Outstanding"
    top10_groups = (
        df.groupby("Party Group")["Invoice Value"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    for i, grp_name in enumerate(top10_groups, start=2):
        cs[f"D{i}"] = grp_name
        cs[f"E{i}"] = f"=SUMIFS({amt_rng_cs},{vis_rng_cs},1,{grp_rng_cs},D{i})"
        cs[f"E{i}"].number_format = "#,##0"

    # Chart 3 Data (G1:H11): Top 10 Pending Branches (Full name in Col I for exact match)
    cs["G1"] = "Branch"
    cs["H1"] = "Outstanding"
    top10_branches = (
        df.groupby("Party Name (Branch)")["Invoice Value"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    for i, br_name in enumerate(top10_branches, start=2):
        cs[f"G{i}"] = br_name[:28] + ("..." if len(br_name) > 28 else "")
        cs[f"I{i}"] = br_name
        cs[f"H{i}"] = f"=SUMIFS({amt_rng_cs},{vis_rng_cs},1,{br_rng_cs},I{i})"
        cs[f"H{i}"].number_format = "#,##0"

    # ---- 3 CLEAN, LIVE CHARTS (Rows 7 to 23 — ABOVE the Table!) ----
    ch1 = make_clean_bar_chart(
        chart_type="col",
        title="Overdue Bucket-wise Outstanding (₹)",
        data_ref=Reference(cs, min_col=2, min_row=1, max_row=5),
        cats_ref=Reference(cs, min_col=1, min_row=2, max_row=5),
        bar_hex="1F4E78",
        width=11.5,
        height=8.6,
    )
    ws.add_chart(ch1, "A7")

    if top10_groups:
        max_r2 = 1 + len(top10_groups)
        ch2 = make_clean_bar_chart(
            chart_type="bar",
            title="Top 10 Parties by Outstanding (₹)",
            data_ref=Reference(cs, min_col=5, min_row=1, max_row=max_r2),
            cats_ref=Reference(cs, min_col=4, min_row=2, max_row=max_r2),
            bar_hex="0F766E",
            width=13.0,
            height=8.6,
        )
        ws.add_chart(ch2, "E7")

    if top10_branches:
        max_r3 = 1 + len(top10_branches)
        ch3 = make_clean_bar_chart(
            chart_type="bar",
            title="Top 10 Pending Branches (₹)",
            data_ref=Reference(cs, min_col=8, min_row=1, max_row=max_r3),
            cats_ref=Reference(cs, min_col=7, min_row=2, max_row=max_r3),
            bar_hex="4338CA",
            width=14.0,
            height=8.6,
        )
        ws.add_chart(ch3, "I7")

    # Auto-fit column widths
    for c_idx, col_name in enumerate(df.columns, start=1):
        max_len = max(len(str(col_name)), 12)
        sample_max = df[col_name].astype(str).str.len().max()
        max_len = min(max(max_len, int(sample_max) + 3), 38)
        ws.column_dimensions[get_column_letter(c_idx)].width = max_len


def build_excel_workbook(
    sales_main: pd.DataFrame,
    purchase_main: pd.DataFrame,
    sales_out: pd.DataFrame,
    purchase_out: pd.DataFrame,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    write_raw_sheet(wb, "Sales Main Sheet", sales_main)
    write_raw_sheet(wb, "Purchase Main Sheet", purchase_main)
    write_outstanding_dashboard_sheet(
        wb, "Sales Outstanding", sales_out, "SalesOutstandingTbl", "_Sales_Calc"
    )
    write_outstanding_dashboard_sheet(
        wb, "Purchase Outstanding", purchase_out, "PurchaseOutstandingTbl", "_Purchase_Calc"
    )

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


# ---------- INTERACTIVE STREAMLIT DASHBOARD RENDERER (MULTI-SELECT) ----------

def render_outstanding_tab(df: pd.DataFrame, key_prefix: str, label: str):
    if df.empty:
        st.success(f"Koi {label} pending nahi hai!")
        return

    slicer_col, main_col = st.columns([1.2, 3.8], gap="large")

    with slicer_col:
        st.markdown("### 🎛️ Multi-Select Slicers")

        all_groups = sorted(df["Party Group"].dropna().unique().tolist())
        selected_groups = st.multiselect(
            "1. Party / Company Group (Khali = All)",
            options=all_groups,
            default=[],
            key=f"{key_prefix}_groups",
            help="Ek ya zyada companies select karein. Khali chhodne par sab dikhega.",
        )

        group_df = df if not selected_groups else df[df["Party Group"].isin(selected_groups)]

        available_branches = sorted(group_df["Party Name (Branch)"].dropna().unique().tolist())
        default_branches = available_branches if selected_groups else []
        selected_branches = st.multiselect(
            "2. Branches / Sub-Parties (Khali = All)",
            options=available_branches,
            default=default_branches,
            key=f"{key_prefix}_branches_{'_'.join(selected_groups)}",
            help="Company select karte hi uski branches aa jayengi; aap kisi ko bhi deselect kar sakte hain.",
        )

        bucket_options = ["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
        selected_buckets = st.multiselect(
            "3. Overdue Days Bucket (Khali = All)",
            options=bucket_options,
            default=[],
            key=f"{key_prefix}_buckets",
            help="Ek ya zyada overdue ranges select karein (e.g. 46-90 Days + 90+ Days).",
        )

    filtered = group_df.copy()
    if selected_branches:
        filtered = filtered[filtered["Party Name (Branch)"].isin(selected_branches)]
    if selected_buckets:
        filtered = filtered[filtered["Overdue Bucket"].isin(selected_buckets)]

    with main_col:
        total_amt = filtered["Invoice Value"].sum()
        total_bills = len(filtered)
        amt_46_90 = filtered.loc[filtered["Overdue Bucket"] == "46-90 Days", "Invoice Value"].sum()
        amt_90_plus = filtered.loc[filtered["Overdue Bucket"] == "90+ Days", "Invoice Value"].sum()
        avg_days = int(filtered["Overdue Days"].mean()) if total_bills > 0 else 0

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total Outstanding", f"₹{total_amt:,.2f}")
        k2.metric("Pending Invoices", f"{total_bills:,}")
        k3.metric("46-90 Days Overdue", f"₹{amt_46_90:,.2f}")
        k4.metric("90+ Days Overdue", f"₹{amt_90_plus:,.2f}")
        k5.metric("Avg Overdue Days", f"{avg_days} Days")

        c1, c2, c3 = st.columns(3)
        with c1:
            bucket_order = ["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
            b_sum = filtered.groupby("Overdue Bucket", as_index=False)["Invoice Value"].sum()
            b_sum["Overdue Bucket"] = pd.Categorical(
                b_sum["Overdue Bucket"], categories=bucket_order, ordered=True
            )
            b_sum = b_sum.sort_values("Overdue Bucket")
            fig_b = px.bar(
                b_sum,
                x="Overdue Bucket",
                y="Invoice Value",
                title="Overdue Bucket-wise (₹)",
                text_auto=".2s",
                color="Overdue Bucket",
                color_discrete_map={
                    "0-30 Days": "#2E7D32",
                    "31-45 Days": "#F9A825",
                    "46-90 Days": "#EF6C00",
                    "90+ Days": "#C62828",
                },
            )
            fig_b.update_layout(showlegend=False, height=320)
            st.plotly_chart(fig_b, use_container_width=True)

        with c2:
            top_grp = (
                filtered.groupby("Party Group", as_index=False)["Invoice Value"]
                .sum()
                .sort_values("Invoice Value", ascending=False)
                .head(10)
            )
            fig_g = px.bar(
                top_grp,
                x="Invoice Value",
                y="Party Group",
                orientation="h",
                title="Top 10 Parties by Outstanding (₹)",
                text_auto=".2s",
            )
            fig_g.update_layout(yaxis={"categoryorder": "total ascending"}, height=320)
            st.plotly_chart(fig_g, use_container_width=True)

        with c3:
            top_br = (
                filtered.groupby("Party Name (Branch)", as_index=False)["Invoice Value"]
                .sum()
                .sort_values("Invoice Value", ascending=False)
                .head(10)
            )
            fig_p = px.bar(
                top_br,
                x="Invoice Value",
                y="Party Name (Branch)",
                orientation="h",
                title="Top 10 Pending Branches (₹)",
                text_auto=".2s",
            )
            fig_p.update_layout(yaxis={"categoryorder": "total ascending"}, height=320)
            st.plotly_chart(fig_p, use_container_width=True)

        def highlight_overdue(row):
            if row["Overdue Bucket"] == "90+ Days":
                return ["background-color: #FCE4D6; color: #C00000; font-weight: bold"] * len(row)
            if row["Overdue Bucket"] == "46-90 Days":
                return ["background-color: #FEF3C7; color: #B45309; font-weight: bold"] * len(row)
            return [""] * len(row)

        display_df = filtered.copy()
        display_df["Invoice Date"] = display_df["Invoice Date"].dt.strftime("%d-%b-%Y")

        st.dataframe(
            display_df.style.apply(highlight_overdue, axis=1).format({"Invoice Value": "₹{:,.2f}"}),
            use_container_width=True,
            height=420,
        )


# ---------- MAIN APP UI ----------

st.title("📊 Sales & Purchase Outstanding Dashboard")

with st.sidebar:
    st.header("1. Upload 3 Excel Files")
    ml_file = st.file_uploader("Multi Ledger Sheet (.xlsx)", type=["xlsx", "xls"])
    pur_file = st.file_uploader("Purchase Record Sheet (.xlsx)", type=["xlsx", "xls"])
    sal_file = st.file_uploader("Sales Record Sheet (.xlsx)", type=["xlsx", "xls"])

    st.divider()
    as_on_date = st.date_input("As-On Date (Overdue calculate karne ke liye)", value=date.today())

if not (ml_file and pur_file and sal_file):
    st.info("Left sidebar mein teeno Excel files (Multi Ledger, Purchase, aur Sales) upload karein.")
    st.stop()

with st.spinner("Files process ho rahi hain aur Excel Dashboards generate ho rahe hain..."):
    ml_df = read_register(ml_file.getvalue())
    pur_df = read_register(pur_file.getvalue())
    sal_df = read_register(sal_file.getvalue())

    paid_keys = extract_paid_bills_from_ledger(ml_df)

    sales_main, sales_out = process_register_outstanding(sal_df, paid_keys, as_on_date, is_purchase=False)
    purchase_main, purchase_out = process_register_outstanding(pur_df, paid_keys, as_on_date, is_purchase=True)

    excel_bytes = build_excel_workbook(sales_main, purchase_main, sales_out, purchase_out)

st.download_button(
    label="⬇️ Download 4-Sheet Outstanding Dashboard (.xlsx)",
    data=excel_bytes,
    file_name=f"Outstanding_Dashboard_{as_on_date.strftime('%d-%b-%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

tab1, tab2 = st.tabs([
    f"📈 Sales Outstanding ({len(sales_out)} Bills)",
    f"📉 Purchase Outstanding ({len(purchase_out)} Bills)",
])

with tab1:
    render_outstanding_tab(sales_out, "sales", "Sales Outstanding")

with tab2:
    render_outstanding_tab(purchase_out, "purchase", "Purchase Outstanding")

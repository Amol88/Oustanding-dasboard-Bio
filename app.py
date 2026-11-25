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
from openpyxl.worksheet.datavalidation import DataValidation

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

    # Clean single corporate color
    if ch.series:
        ch.series[0].graphicalProperties.solidFill = bar_hex
        ch.series[0].graphicalProperties.line.solidFill = bar_hex

    # Show ONLY numeric value on bars (no series name, no category name, no legend box)
    dl = DataLabelList()
    dl.showVal = True
    dl.showSerName = False
    dl.showCatName = False
    dl.showPercent = False
    dl.showLegendKey = False
    ch.dataLabels = dl

    # Clean axes without clutter
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.tickLblPos = "low"
    ch.y_axis.majorGridlines = None
    ch.x_axis.majorGridlines = None

    ch.width = width
    ch.height = height
    return ch


# ---------- 100% NATIVE, ZERO-ERROR EXCEL DASHBOARD GENERATOR ----------

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


def write_outstanding_dashboard_sheet(wb: Workbook, title: str, df: pd.DataFrame, data_sheet_name: str):
    ws = wb.create_sheet(title=title)
    ws.views.sheetView[0].showGridLines = False

    ds = wb.create_sheet(title=data_sheet_name)
    ds.sheet_state = "hidden"

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 3
    ws.column_dimensions["D"].width = 3

    # Styles
    banner_fill = PatternFill("solid", fgColor="1F4E78")
    banner_font = Font(bold=True, size=15, color="FFFFFF")
    panel_fill = PatternFill("solid", fgColor="F2F5F9")
    filter_hdr_fill = PatternFill("solid", fgColor="1F4E78")
    filter_hdr_font = Font(bold=True, size=11, color="FFFFFF")
    filter_lbl_font = Font(bold=True, size=9, color="1F4E78")
    dropdown_fill = PatternFill("solid", fgColor="FFFFFF")
    dropdown_font = Font(bold=True, size=10, color="000000")
    kpi_label_font = Font(bold=True, size=9, color="595959")
    kpi_val_font = Font(bold=True, size=13, color="1F4E78")
    kpi_amber_font = Font(bold=True, size=13, color="B45309")
    kpi_red_font = Font(bold=True, size=13, color="C00000")
    kpi_fill = PatternFill("solid", fgColor="EDF2F8")
    kpi_amber_fill = PatternFill("solid", fgColor="FEF3C7")
    kpi_red_fill = PatternFill("solid", fgColor="FCE4D6")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    box_border = Border(
        left=Side(style="medium", color="1F4E78"),
        right=Side(style="medium", color="1F4E78"),
        top=Side(style="medium", color="1F4E78"),
        bottom=Side(style="medium", color="1F4E78"),
    )

    # Top Title Banner (E1:R2)
    ws.merge_cells("E1:R2")
    ws["E1"] = f"  {title} Dashboard"
    ws["E1"].fill = banner_fill
    ws["E1"].font = banner_font
    ws["E1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18

    # Left Filter Panel (A1:C24)
    for r in range(1, 25):
        for c in range(1, 4):
            ws.cell(row=r, column=c).fill = panel_fill

    ws["B2"] = "DASHBOARD FILTERS"
    ws["B2"].fill = filter_hdr_fill
    ws["B2"].font = filter_hdr_font
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    ws["B4"] = "1. Party / Company Group"
    ws["B4"].font = filter_lbl_font
    ws["B5"] = "All"
    ws["B5"].fill = dropdown_fill
    ws["B5"].font = dropdown_font
    ws["B5"].border = box_border
    ws["B5"].alignment = Alignment(horizontal="left", vertical="center")

    ws["B7"] = "2. Branch / Sub-Party"
    ws["B7"].font = filter_lbl_font
    ws["B8"] = "All"
    ws["B8"].fill = dropdown_fill
    ws["B8"].font = dropdown_font
    ws["B8"].border = box_border
    ws["B8"].alignment = Alignment(horizontal="left", vertical="center")

    ws["B10"] = "3. Overdue Days Bucket"
    ws["B10"].font = filter_lbl_font
    ws["B11"] = "All"
    ws["B11"].fill = dropdown_fill
    ws["B11"].font = dropdown_font
    ws["B11"].border = box_border
    ws["B11"].alignment = Alignment(horizontal="left", vertical="center")

    ws["B13"] = "Dropdown change karte hi"
    ws["B13"].font = Font(italic=True, size=8, color="595959")
    ws["B14"] = "KPIs, Charts & Niche Table"
    ws["B14"].font = Font(italic=True, size=8, color="595959")
    ws["B15"] = "sab live update honge!"
    ws["B15"].font = Font(italic=True, size=8, color="595959")

    if len(df) == 0:
        ws["E4"] = "Koi pending invoice nahi hai."
        return

    # ---- POPULATE BACKEND DATA SHEET (ds) ----
    ds["A1"] = 0
    for idx, col_name in enumerate(df.columns, start=2):
        ds.cell(row=1, column=idx, value=col_name)

    grp_ds_ltr = get_column_letter(2 + df.columns.get_loc("Party Group"))
    br_ds_ltr = get_column_letter(2 + df.columns.get_loc("Party Name (Branch)"))
    amt_ds_idx = 2 + df.columns.get_loc("Invoice Value")
    amt_ds_ltr = get_column_letter(amt_ds_idx)
    days_ds_ltr = get_column_letter(2 + df.columns.get_loc("Overdue Days"))
    bkt_ds_ltr = get_column_letter(2 + df.columns.get_loc("Overdue Bucket"))

    sh_ref = f"'{title}'"
    ds_ref = f"'{data_sheet_name}'"

    n_rows = len(df)
    r_start = 2
    r_end = n_rows + 1

    for i, row in enumerate(df.itertuples(index=False, name=None), start=2):
        ds.cell(
            row=i,
            column=1,
            value=(
                f'=IF(AND('
                f'OR({sh_ref}!$B$5="All",{grp_ds_ltr}{i}={sh_ref}!$B$5),'
                f'OR({sh_ref}!$B$8="All",{br_ds_ltr}{i}={sh_ref}!$B$8),'
                f'OR({sh_ref}!$B$11="All",{bkt_ds_ltr}{i}={sh_ref}!$B$11)'
                f'),A{i-1}+1,A{i-1})'
            ),
        )
        for offset, val in enumerate(row):
            c = ds.cell(row=i, column=2 + offset)
            if isinstance(val, pd.Timestamp):
                c.value = val.to_pydatetime() if pd.notna(val) else None
            else:
                c.value = None if pd.isna(val) else val

    grp_rng = f"{ds_ref}!${grp_ds_ltr}${r_start}:${grp_ds_ltr}${r_end}"
    br_rng = f"{ds_ref}!${br_ds_ltr}${r_start}:${br_ds_ltr}${r_end}"
    amt_rng = f"{ds_ref}!${amt_ds_ltr}${r_start}:${amt_ds_ltr}${r_end}"
    day_rng = f"{ds_ref}!${days_ds_ltr}${r_start}:${days_ds_ltr}${r_end}"
    bkt_rng = f"{ds_ref}!${bkt_ds_ltr}${r_start}:${bkt_ds_ltr}${r_end}"
    cnt_rng = f"{ds_ref}!$A${r_start}:$A${r_end}"

    c_grp = 'IF($B$5="All","*",$B$5)'
    c_br = 'IF($B$8="All","*",$B$8)'
    c_bkt = 'IF($B$11="All","*",$B$11)'

    # ---- 5 DYNAMIC KPI CARDS (Row 4-5) ----
    kpi_specs = [
        ("E", "F", "Total Outstanding (₹)",
         f"=SUMIFS({amt_rng},{grp_rng},{c_grp},{br_rng},{c_br},{bkt_rng},{c_bkt})",
         "₹#,##0.00", kpi_fill, kpi_val_font),
        ("H", "I", "Pending Invoices",
         f"=COUNTIFS({grp_rng},{c_grp},{br_rng},{c_br},{bkt_rng},{c_bkt})",
         "#,##0", kpi_fill, kpi_val_font),
        ("K", "L", "46-90 Days Overdue (₹)",
         f'=IF(OR($B$11="All",$B$11="46-90 Days"),SUMIFS({amt_rng},{grp_rng},{c_grp},{br_rng},{c_br},{bkt_rng},"46-90 Days"),0)',
         "₹#,##0.00", kpi_amber_fill, kpi_amber_font),
        ("N", "O", "90+ Days Overdue (₹)",
         f'=IF(OR($B$11="All",$B$11="90+ Days"),SUMIFS({amt_rng},{grp_rng},{c_grp},{br_rng},{c_br},{bkt_rng},"90+ Days"),0)',
         "₹#,##0.00", kpi_red_fill, kpi_red_font),
        ("Q", "R", "Avg Overdue Days",
         f"=IFERROR(AVERAGEIFS({day_rng},{grp_rng},{c_grp},{br_rng},{c_br},{bkt_rng},{c_bkt}),0)",
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

    # ---- DROPDOWN LISTS & CHART SUMMARIES ON ds ----
    groups_list = ["All"] + sorted(df["Party Group"].dropna().unique().tolist())
    for idx, val in enumerate(groups_list, start=1):
        ds[f"T{idx}"] = val
    dv_grp = DataValidation(type="list", formula1=f"={ds_ref}!$T$1:$T${len(groups_list)}", allow_blank=False)
    ws.add_data_validation(dv_grp)
    dv_grp.add("B5")

    branches_list = ["All"] + sorted(df["Party Name (Branch)"].dropna().unique().tolist())
    for idx, val in enumerate(branches_list, start=1):
        ds[f"U{idx}"] = val
    dv_br = DataValidation(type="list", formula1=f"={ds_ref}!$U$1:$U${len(branches_list)}", allow_blank=False)
    ws.add_data_validation(dv_br)
    dv_br.add("B8")

    buckets_dd = ["All", "0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
    for idx, val in enumerate(buckets_dd, start=1):
        ds[f"V{idx}"] = val
    dv_bkt = DataValidation(type="list", formula1=f"={ds_ref}!$V$1:$V${len(buckets_dd)}", allow_blank=False)
    ws.add_data_validation(dv_bkt)
    dv_bkt.add("B11")

    c_grp_ds = f'IF({sh_ref}!$B$5="All","*",{sh_ref}!$B$5)'
    c_br_ds = f'IF({sh_ref}!$B$8="All","*",{sh_ref}!$B$8)'
    c_bkt_ds = f'IF({sh_ref}!$B$11="All","*",{sh_ref}!$B$11)'

    # Chart 1 Data (W1:X5): Overdue Bucket-wise Outstanding
    ds["W1"] = "Overdue Bucket"
    ds["X1"] = "Outstanding"
    for i, b_name in enumerate(["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"], start=2):
        ds[f"W{i}"] = b_name
        ds[f"X{i}"] = (
            f'=IF(OR({sh_ref}!$B$11="All",{sh_ref}!$B$11=W{i}),'
            f'SUMIFS({amt_rng},{grp_rng},{c_grp_ds},{br_rng},{c_br_ds},{bkt_rng},W{i}),0)'
        )
        ds[f"X{i}"].number_format = "#,##0"

    # Chart 2 Data (Y1:Z11): Top 10 Parties (Party Group)
    ds["Y1"] = "Party Group"
    ds["Z1"] = "Outstanding"
    top10_groups = (
        df.groupby("Party Group")["Invoice Value"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    for i, grp_name in enumerate(top10_groups, start=2):
        ds[f"Y{i}"] = grp_name
        ds[f"Z{i}"] = (
            f'=IF(OR({sh_ref}!$B$5="All",{sh_ref}!$B$5=Y{i}),'
            f'SUMIFS({amt_rng},{grp_rng},Y{i},{br_rng},{c_br_ds},{bkt_rng},{c_bkt_ds}),0)'
        )
        ds[f"Z{i}"].number_format = "#,##0"

    # Chart 3 Data (AA1:AB11): Top 10 Pending Branches
    ds["AA1"] = "Branch"
    ds["AB1"] = "Outstanding"
    top10_branches = (
        df.groupby("Party Name (Branch)")["Invoice Value"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    for i, br_name in enumerate(top10_branches, start=2):
        ds[f"AA{i}"] = br_name[:28] + ("..." if len(br_name) > 28 else "")
        ds[f"AC{i}"] = br_name  # Full name for exact SUMIFS matching
        ds[f"AB{i}"] = (
            f'=IF(OR({sh_ref}!$B$8="All",{sh_ref}!$B$8=AC{i}),'
            f'SUMIFS({amt_rng},{grp_rng},{c_grp_ds},{br_rng},AC{i},{bkt_rng},{c_bkt_ds}),0)'
        )
        ds[f"AB{i}"].number_format = "#,##0"

    # ---- 3 CLEAN, NON-OVERLAPPING LIVE CHARTS ----
    ch1 = make_clean_bar_chart(
        chart_type="col",
        title="Overdue Bucket-wise Outstanding (₹)",
        data_ref=Reference(ds, min_col=24, min_row=1, max_row=5),
        cats_ref=Reference(ds, min_col=23, min_row=2, max_row=5),
        bar_hex="1F4E78",
        width=12.0,
        height=9.2,
    )
    ws.add_chart(ch1, "E7")

    if top10_groups:
        max_r2 = 1 + len(top10_groups)
        ch2 = make_clean_bar_chart(
            chart_type="bar",
            title="Top 10 Parties by Outstanding (₹)",
            data_ref=Reference(ds, min_col=26, min_row=1, max_row=max_r2),
            cats_ref=Reference(ds, min_col=25, min_row=2, max_row=max_r2),
            bar_hex="0F766E",
            width=13.5,
            height=9.2,
        )
        ws.add_chart(ch2, "J7")

    if top10_branches:
        max_r3 = 1 + len(top10_branches)
        ch3 = make_clean_bar_chart(
            chart_type="bar",
            title="Top 10 Pending Branches (₹)",
            data_ref=Reference(ds, min_col=28, min_row=1, max_row=max_r3),
            cats_ref=Reference(ds, min_col=27, min_row=2, max_row=max_r3),
            bar_hex="4338CA",
            width=14.5,
            height=9.2,
        )
        ws.add_chart(ch3, "O7")

    # ---- DYNAMIC FILTERED INVOICES TABLE (Row 26 onwards, Col E onwards) ----
    col_start = 5   # Column E
    start_row = 26  # Header row (clean gap below taller charts)

    for offset, col_name in enumerate(df.columns):
        c_idx = col_start + offset
        cell = ws.cell(row=start_row, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    amt_col_offset = df.columns.get_loc("Invoice Value")
    date_col_offset = df.columns.get_loc("Invoice Date")

    for r_idx in range(1, n_rows + 1):
        excel_row = start_row + r_idx
        for offset, col_name in enumerate(df.columns):
            ds_col_ltr = get_column_letter(2 + offset)
            c_idx = col_start + offset
            cell = ws.cell(
                row=excel_row,
                column=c_idx,
                value=f'=IFERROR(INDEX({ds_ref}!${ds_col_ltr}${r_start}:${ds_col_ltr}${r_end},MATCH({r_idx},{cnt_rng},0)),"")',
            )
            if offset == amt_col_offset:
                cell.number_format = "#,##0.00"
            elif offset == date_col_offset:
                cell.number_format = "dd-mmm-yyyy"
            cell.border = thin_border

    for offset, col_name in enumerate(df.columns):
        c_idx = col_start + offset
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
    write_outstanding_dashboard_sheet(wb, "Sales Outstanding", sales_out, "_Sales_Data")
    write_outstanding_dashboard_sheet(wb, "Purchase Outstanding", purchase_out, "_Purchase_Data")

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


# ---------- INTERACTIVE STREAMLIT DASHBOARD RENDERER ----------

def render_outstanding_tab(df: pd.DataFrame, key_prefix: str, label: str):
    if df.empty:
        st.success(f"Koi {label} pending nahi hai!")
        return

    slicer_col, main_col = st.columns([1.2, 3.8], gap="large")

    with slicer_col:
        st.markdown("### 🎛️ Slicers")

        all_groups = sorted(df["Party Group"].dropna().unique().tolist())
        selected_group = st.selectbox(
            "1. Party / Company Group",
            options=["All"] + all_groups,
            key=f"{key_prefix}_group",
            help="Company select karte hi uski saari branches select ho jayengi.",
        )

        group_df = df if selected_group == "All" else df[df["Party Group"] == selected_group]

        available_branches = sorted(group_df["Party Name (Branch)"].dropna().unique().tolist())
        selected_branches = st.multiselect(
            "2. Branches / Sub-Parties (Deselect allowed)",
            options=available_branches,
            default=available_branches,
            key=f"{key_prefix}_branches_{selected_group}",
        )

        bucket_options = ["All", "0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
        selected_bucket = st.radio(
            "3. Overdue Days",
            options=bucket_options,
            key=f"{key_prefix}_bucket",
        )

    filtered = group_df[group_df["Party Name (Branch)"].isin(selected_branches)].copy()
    if selected_bucket != "All":
        filtered = filtered[filtered["Overdue Bucket"] == selected_bucket]

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

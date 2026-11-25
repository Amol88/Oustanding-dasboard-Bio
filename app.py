import io
import re
from datetime import date
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
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
    return 6  # Default BioPharm export header row


def read_register(file_bytes: bytes) -> pd.DataFrame:
    header_idx = find_header_row(file_bytes)
    df = pd.read_excel(io.BytesIO(file_bytes), header=header_idx, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    # Completely empty rows hata dein
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def normalize_bill_key(doc_type: str, number: str) -> str:
    """Type aur Serial/Bill No ko combine karke leading zeros normalize karta hai."""
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
    """
    1) Main sheet ko as-is (multi-product blank Party Name rows ke saath) preserve karta hai.
    2) Unpaid invoices ka clean Outstanding DataFrame banata hai.
    """
    main_df = reg_df.copy()

    # Check header rows jahan Party Name non-empty hai (multi-product invoice grouping)
    party_col = "Party Name"
    has_party = main_df[party_col].notna() & (main_df[party_col].astype(str).str.strip() != "")
    main_df["_Invoice_Group"] = has_party.cumsum()

    # Sirf valid invoice groups rakhein (jahan kam se kam 1 Party Name mila ho)
    valid_df = main_df[main_df["_Invoice_Group"] > 0].copy()

    # Har invoice ki first row (header row) nikal kar outstanding calculate karein
    inv_headers = valid_df[has_party[valid_df.index]].copy()

    # Bill No = Type - Serial
    inv_headers["Bill No"] = (
        inv_headers["Type"].fillna("").str.strip()
        + "-"
        + inv_headers["Serial"].fillna("").str.strip()
    )

    # Matching keys
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

    # Sirf unpaid (Outstanding) rows
    out_df = inv_headers[inv_headers["Payment Status"] == "Outstanding"].copy()

    # Dates & Amounts parse karein
    date_col = "Bill.Date" if (is_purchase and "Bill.Date" in out_df.columns) else "Doc.Date"
    out_df["Invoice Date"] = pd.to_datetime(out_df[date_col], errors="coerce")
    out_df["Invoice Value"] = pd.to_numeric(out_df.get("Invoice Value", 0), errors="coerce").fillna(0.0)

    # Multi-product invoices mein agar Invoice Value 0 ho toh Item Value ka sum fallback lein
    if (out_df["Invoice Value"] == 0).any() and "Item Value" in valid_df.columns:
        valid_df["_item_val_num"] = pd.to_numeric(valid_df["Item Value"], errors="coerce").fillna(0.0)
        grp_sums = valid_df.groupby("_Invoice_Group")["_item_val_num"].sum()
        zero_mask = out_df["Invoice Value"] == 0
        out_df.loc[zero_mask, "Invoice Value"] = out_df.loc[zero_mask, "_Invoice_Group"].map(grp_sums).fillna(0.0)

    # Overdue days calculation
    as_on_ts = pd.Timestamp(as_on_date)
    out_df["Overdue Days"] = (as_on_ts - out_df["Invoice Date"]).dt.days.fillna(0).astype(int)
    out_df["Overdue Bucket"] = out_df["Overdue Days"].apply(get_overdue_bucket)

    # Parent Company & Branch
    out_df["Party Name (Branch)"] = out_df["Party Name"].astype(str).str.strip()
    out_df["Party Group"] = out_df["Party Name (Branch)"].apply(extract_parent_party)

    # Clean output columns
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

    # Main sheet se helper column hata dein (bilkul as-is rakhne ke liye)
    main_clean = main_df.drop(columns=["_Invoice_Group"])
    return main_clean, out_df


# ---------- EXCEL GENERATOR (4 SHEETS) ----------

def write_raw_sheet(wb: Workbook, title: str, df: pd.DataFrame):
    ws = wb.create_sheet(title=title)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=10)

    # Headers
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data rows as-is
    for r_idx, row in enumerate(df.itertuples(index=False, name=None), start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=None if pd.isna(val) else val)

    ws.freeze_panes = "A2"


def write_outstanding_dashboard_sheet(wb: Workbook, title: str, df: pd.DataFrame, table_name: str):
    ws = wb.create_sheet(title=title)
    ws.views.sheetView[0].showGridLines = True

    # Styling
    title_font = Font(bold=True, size=14, color="1F4E78")
    kpi_label_font = Font(bold=True, size=10, color="595959")
    kpi_val_font = Font(bold=True, size=13, color="1F4E78")
    kpi_fill = PatternFill("solid", fgColor="EDF2F8")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    red_fill = PatternFill("solid", fgColor="FCE4D6")
    red_font = Font(bold=True, color="C00000")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Title Row
    ws["A1"] = f"{title} Dashboard (Tip: Table mein click karke Insert -> Slicer use karein)"
    ws["A1"].font = title_font

    start_row = 6
    end_row = max(start_row + len(df), start_row + 1)

    # Column letters for dynamic SUBTOTAL KPIs
    amt_col_idx = df.columns.get_loc("Invoice Value") + 1
    amt_col_ltr = get_column_letter(amt_col_idx)
    days_col_idx = df.columns.get_loc("Overdue Days") + 1
    days_col_ltr = get_column_letter(days_col_idx)

    # Dynamic KPI Cards (Row 3-4) — Excel filter/slicer lagane par auto-update honge!
    kpis = [
        ("A", "Total Outstanding (₹)", f"=SUBTOTAL(9,{amt_col_ltr}{start_row+1}:{amt_col_ltr}{end_row})", "₹#,##0.00"),
        ("C", "Pending Invoices", f"=SUBTOTAL(3,A{start_row+1}:A{end_row})", "#,##0"),
        ("E", "Avg Overdue Days", f"=IFERROR(SUBTOTAL(1,{days_col_ltr}{start_row+1}:{days_col_ltr}{end_row}),0)", "0"),
        ("G", "Max Overdue Days", f"=IFERROR(SUBTOTAL(4,{days_col_ltr}{start_row+1}:{days_col_ltr}{end_row}),0)", "0"),
    ]
    for col_ltr, lbl, formula, num_fmt in kpis:
        ws[f"{col_ltr}3"] = lbl
        ws[f"{col_ltr}3"].font = kpi_label_font
        ws[f"{col_ltr}3"].fill = kpi_fill
        ws[f"{col_ltr}4"] = formula
        ws[f"{col_ltr}4"].font = kpi_val_font
        ws[f"{col_ltr}4"].fill = kpi_fill
        ws[f"{col_ltr}4"].number_format = num_fmt

    # Table Headers
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Table Data
    bucket_col_idx = df.columns.get_loc("Overdue Bucket") + 1
    for r_offset, row in enumerate(df.itertuples(index=False, name=None), start=1):
        r_num = start_row + r_offset
        is_90_plus = row[bucket_col_idx - 1] == "90+ Days"
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_num, column=c_idx)
            if isinstance(val, pd.Timestamp):
                cell.value = val.to_pydatetime() if pd.notna(val) else None
                cell.number_format = "dd-mmm-yyyy"
            else:
                cell.value = None if pd.isna(val) else val

            if c_idx == amt_col_idx:
                cell.number_format = "#,##0.00"
            cell.border = thin_border

            if is_90_plus:
                cell.fill = red_fill
                cell.font = red_font

    # Add Excel Table (for instant AutoFilter & 1-click Slicers in Excel)
    if len(df) > 0:
        last_col_ltr = get_column_letter(len(df.columns))
        tab = Table(displayName=table_name, ref=f"A{start_row}:{last_col_ltr}{end_row}")
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(tab)

    # Auto-fit column widths
    for c_idx, col_name in enumerate(df.columns, start=1):
        max_len = max(len(str(col_name)), 12)
        if len(df) > 0:
            sample_max = df[col_name].astype(str).str.len().max()
            max_len = min(max(max_len, int(sample_max) + 3), 42)
        ws.column_dimensions[get_column_letter(c_idx)].width = max_len

    ws.freeze_panes = f"A{start_row + 1}"


def build_excel_workbook(
    sales_main: pd.DataFrame,
    purchase_main: pd.DataFrame,
    sales_out: pd.DataFrame,
    purchase_out: pd.DataFrame,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)  # Default empty sheet hatayein

    write_raw_sheet(wb, "Sales Main Sheet", sales_main)
    write_raw_sheet(wb, "Purchase Main Sheet", purchase_main)
    write_outstanding_dashboard_sheet(wb, "Sales Outstanding", sales_out, "SalesOutstandingTbl")
    write_outstanding_dashboard_sheet(wb, "Purchase Outstanding", purchase_out, "PurchaseOutstandingTbl")

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


# ---------- INTERACTIVE DASHBOARD RENDERER ----------

def render_outstanding_tab(df: pd.DataFrame, key_prefix: str, label: str):
    if df.empty:
        st.success(f"Koi {label} pending nahi hai!")
        return

    # Left-side Slicers + Right-side Dashboard layout
    slicer_col, main_col = st.columns([1.2, 3.8], gap="large")

    with slicer_col:
        st.markdown("### 🎛️ Slicers")

        # 1. Parent Party / Company Slicer (with All option)
        all_groups = sorted(df["Party Group"].dropna().unique().tolist())
        selected_group = st.selectbox(
            "1. Party / Company Group",
            options=["All"] + all_groups,
            key=f"{key_prefix}_group",
            help="Company select karte hi uski saari branches select ho jayengi.",
        )

        if selected_group == "All":
            group_df = df
        else:
            group_df = df[df["Party Group"] == selected_group]

        # 2. Branch / Sub-Party Slicer (deselect karne ka option)
        available_branches = sorted(group_df["Party Name (Branch)"].dropna().unique().tolist())
        selected_branches = st.multiselect(
            "2. Branches / Sub-Parties (Deselect allowed)",
            options=available_branches,
            default=available_branches,
            key=f"{key_prefix}_branches_{selected_group}",
        )

        # 3. Overdue Days Bucket Slicer
        bucket_options = ["All", "0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
        selected_bucket = st.radio(
            "3. Overdue Days",
            options=bucket_options,
            key=f"{key_prefix}_bucket",
        )

    # Apply Filters
    filtered = group_df[group_df["Party Name (Branch)"].isin(selected_branches)].copy()
    if selected_bucket != "All":
        filtered = filtered[filtered["Overdue Bucket"] == selected_bucket]

    with main_col:
        # KPIs
        total_amt = filtered["Invoice Value"].sum()
        total_bills = len(filtered)
        critical_amt = filtered.loc[filtered["Overdue Bucket"] == "90+ Days", "Invoice Value"].sum()
        avg_days = int(filtered["Overdue Days"].mean()) if total_bills > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Outstanding", f"₹{total_amt:,.2f}")
        k2.metric("Pending Invoices", f"{total_bills:,}")
        k3.metric("90+ Days Overdue", f"₹{critical_amt:,.2f}")
        k4.metric("Avg Overdue Days", f"{avg_days} Days")

        # Charts
        c1, c2 = st.columns(2)
        with c1:
            bucket_order = ["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
            b_sum = (
                filtered.groupby("Overdue Bucket", as_index=False)["Invoice Value"]
                .sum()
            )
            b_sum["Overdue Bucket"] = pd.Categorical(b_sum["Overdue Bucket"], categories=bucket_order, ordered=True)
            b_sum = b_sum.sort_values("Overdue Bucket")
            fig_b = px.bar(
                b_sum,
                x="Overdue Bucket",
                y="Invoice Value",
                title="Overdue Bucket-wise Outstanding (₹)",
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
            top_p = (
                filtered.groupby("Party Name (Branch)", as_index=False)["Invoice Value"]
                .sum()
                .sort_values("Invoice Value", ascending=False)
                .head(10)
            )
            fig_p = px.bar(
                top_p,
                x="Invoice Value",
                y="Party Name (Branch)",
                orientation="h",
                title="Top 10 Pending Branches (₹)",
                text_auto=".2s",
            )
            fig_p.update_layout(yaxis={"categoryorder": "total ascending"}, height=320)
            st.plotly_chart(fig_p, use_container_width=True)

        # Highlight 90+ days in table
        def highlight_90(row):
            if row["Overdue Bucket"] == "90+ Days":
                return ["background-color: #FCE4D6; color: #C00000; font-weight: bold"] * len(row)
            return [""] * len(row)

        display_df = filtered.copy()
        display_df["Invoice Date"] = display_df["Invoice Date"].dt.strftime("%d-%b-%Y")

        st.dataframe(
            display_df.style.apply(highlight_90, axis=1).format({"Invoice Value": "₹{:,.2f}"}),
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

with st.spinner("Files process ho rahi hain aur ICICI Bank payments match ho rahi hain..."):
    ml_df = read_register(ml_file.getvalue())
    pur_df = read_register(pur_file.getvalue())
    sal_df = read_register(sal_file.getvalue())

    paid_keys = extract_paid_bills_from_ledger(ml_df)

    sales_main, sales_out = process_register_outstanding(sal_df, paid_keys, as_on_date, is_purchase=False)
    purchase_main, purchase_out = process_register_outstanding(pur_df, paid_keys, as_on_date, is_purchase=True)

    excel_bytes = build_excel_workbook(sales_main, purchase_main, sales_out, purchase_out)

# Download Button (Top right)
st.download_button(
    label="⬇️ Download Complete 4-Sheet Outstanding Report (.xlsx)",
    data=excel_bytes,
    file_name=f"Outstanding_Report_{as_on_date.strftime('%d-%b-%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# Tabs for Sales & Purchase Outstanding Dashboards
tab1, tab2 = st.tabs([
    f"📈 Sales Outstanding ({len(sales_out)} Bills)",
    f"📉 Purchase Outstanding ({len(purchase_out)} Bills)",
])

with tab1:
    render_outstanding_tab(sales_out, "sales", "Sales Outstanding")

with tab2:
    render_outstanding_tab(purchase_out, "purchase", "Purchase Outstanding")

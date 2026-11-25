import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
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


# ---------- NATIVE EXCEL SLICER XML INJECTOR (CHART-SAFE) ----------

def inject_excel_table_slicers(xlsx_bytes: bytes, sheet_slicer_configs: list) -> bytes:
    """
    Existing Excel Charts (drawing.xml) ko preserve karte hue Sales & Purchase
    Outstanding sheets mein native clickable Excel Slicers inject karta hai.
    """
    in_buf = io.BytesIO(xlsx_bytes)
    with zipfile.ZipFile(in_buf, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"

    wb_root = ET.fromstring(files["xl/workbook.xml"])
    wb_rels_root = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    ct_root = ET.fromstring(files["[Content_Types].xml"])

    rid_to_target = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in wb_rels_root.findall(f"{{{ns_rel}}}Relationship")
    }
    sheet_to_path = {}
    for sh in wb_root.findall(f".//{{{ns_main}}}sheet"):
        name = sh.attrib["name"]
        rid = sh.attrib[f"{{{ns_r}}}id"]
        target = rid_to_target[rid].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        sheet_to_path[name] = target

    table_info = {}
    for fname, content in files.items():
        if fname.startswith("xl/tables/table") and fname.endswith(".xml"):
            t_root = ET.fromstring(content)
            t_name = t_root.attrib.get("displayName") or t_root.attrib.get("name")
            t_id = t_root.attrib["id"]
            cols = {
                tc.attrib["name"]: tc.attrib["id"]
                for tc in t_root.findall(f".//{{{ns_main}}}tableColumn")
            }
            table_info[t_name] = (t_id, cols)

    def next_rid(rels_root):
        max_id = 0
        for rel in rels_root.findall(f"{{{ns_rel}}}Relationship"):
            m = re.match(r"rId(\d+)", rel.attrib.get("Id", ""))
            if m:
                max_id = max(max_id, int(m.group(1)))
        return f"rId{max_id + 1}"

    def serialize_with_default_ns(root, default_ns):
        ET.register_namespace("", default_ns)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    cache_counter = 0
    slicer_part_counter = 0
    drawing_counter = sum(
        1 for k in files if k.startswith("xl/drawings/drawing") and k.endswith(".xml")
    )
    wb_cache_rids = []
    ct_overrides = []

    for cfg in sheet_slicer_configs:
        sh_name = cfg["sheet_name"]
        t_name = cfg["table_name"]
        if sh_name not in sheet_to_path or t_name not in table_info:
            continue

        t_id, col_map = table_info[t_name]
        sh_path = sheet_to_path[sh_name]
        sh_dir, sh_file = sh_path.rsplit("/", 1)
        sh_rels_path = f"{sh_dir}/_rels/{sh_file}.rels"

        if sh_rels_path in files:
            sh_rels_root = ET.fromstring(files[sh_rels_path])
        else:
            sh_rels_root = ET.Element(f"{{{ns_rel}}}Relationships")

        existing_draw_target = None
        for rel in sh_rels_root.findall(f"{{{ns_rel}}}Relationship"):
            if rel.attrib.get("Type", "").endswith("/drawing"):
                existing_draw_target = rel.attrib["Target"]
                break

        slicer_part_counter += 1
        slicer_xml_items = []
        drawing_xml_items = []
        shape_id = 100

        for sl in cfg["slicers"]:
            col_name = sl["col_name"]
            if col_name not in col_map:
                continue
            col_id = col_map[col_name]
            cache_counter += 1
            shape_id += 1

            safe_slug = re.sub(r"[^A-Za-z0-9_]", "_", f"{t_name}_{col_name}")[:36]
            cache_name = f"Slicer_{safe_slug}_{cache_counter}"
            slicer_name = xml_escape(f"{sl['caption']} {slicer_part_counter}_{shape_id}")
            caption_esc = xml_escape(sl["caption"])
            col_esc = xml_escape(col

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

        # Check if sheet already has a drawing (from openpyxl charts)
        existing_draw_target = None
        for rel in sh_rels_root.findall(f"{{{ns_rel}}}Relationship"):
            if rel.attrib.get("Type", "").endswith("/drawing"):
                existing_draw_target = rel.attrib["Target"]
                break

        slicer_part_counter += 1
        slicer_xml_items = []
        drawing_xml_items = []
        shape_id = 100  # Avoid collision with openpyxl chart shape IDs

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
            col_esc = xml_escape(col_name)

            cache_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
                'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" '
                'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                f'name="{cache_name}" sourceName="{col_esc}">'
                '<extLst>'
                '<ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" '
                'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
                f'<x15:tableSlicerCache tableId="{t_id}" column="{col_id}"/>'
                '</ext>'
                '</extLst>'
                '</slicerCacheDefinition>'
            )
            cache_path = f"xl/slicerCaches/slicerCache{cache_counter}.xml"
            files[cache_path] = cache_xml.encode("utf-8")
            ct_overrides.append((f"/{cache_path}", "application/vnd.ms-excel.slicerCache+xml"))

            c_rid = next_rid(wb_rels_root)
            ET.SubElement(
                wb_rels_root,
                f"{{{ns_rel}}}Relationship",
                {
                    "Id": c_rid,
                    "Type": "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                    "Target": f"slicerCaches/slicerCache{cache_counter}.xml",
                },
            )
            wb_cache_rids.append(c_rid)

            slicer_xml_items.append(
                f'<slicer name="{slicer_name}" cache="{cache_name}" caption="{caption_esc}" '
                'startItem="0" columnCount="1" showCaption="1" style="SlicerStyleLight1"/>'
            )

            fc, fr, tc, tr = sl["from_col"], sl["from_row"], sl["to_col"], sl["to_row"]
            drawing_xml_items.append(
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                '<mc:Choice xmlns:sle14="http://schemas.microsoft.com/office/drawing/2010/slicer" Requires="sle14">'
                '<xdr:twoCellAnchor>'
                f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                '<xdr:graphicFrame macro="">'
                '<xdr:nvGraphicFramePr>'
                f'<xdr:cNvPr id="{shape_id}" name="{slicer_name}"/>'
                '<xdr:cNvGraphicFramePr/>'
                '</xdr:nvGraphicFramePr>'
                '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">'
                f'<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="{slicer_name}"/>'
                '</a:graphicData></a:graphic>'
                '</xdr:graphicFrame><xdr:clientData/>'
                '</xdr:twoCellAnchor>'
                '</mc:Choice>'
                '<mc:Fallback>'
                '<xdr:twoCellAnchor>'
                f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                '<xdr:sp macro="" textlink="">'
                '<xdr:nvSpPr>'
                f'<xdr:cNvPr id="{shape_id}" name="{slicer_name}"/>'
                '<xdr:cNvSpPr/>'
                '</xdr:nvSpPr>'
                '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
                '</xdr:sp><xdr:clientData/>'
                '</xdr:twoCellAnchor>'
                '</mc:Fallback>'
                '</mc:AlternateContent>'
            )

        if not slicer_xml_items:
            continue

        slicer_path = f"xl/slicers/slicer{slicer_part_counter}.xml"
        files[slicer_path] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" '
            'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            + "".join(slicer_xml_items)
            + "</slicers>"
        ).encode("utf-8")
        ct_overrides.append((f"/{slicer_path}", "application/vnd.ms-excel.slicer+xml"))

        # Append slicers into existing chart drawing.xml OR create a new drawing.xml
        if existing_draw_target:
            draw_path = existing_draw_target.replace("../", "xl/").lstrip("/")
            if not draw_path.startswith("xl/"):
                draw_path = "xl/" + draw_path
            existing_draw_xml = files[draw_path].decode("utf-8")
            # Ensure xmlns:a is declared on wsDr root
            if 'xmlns:a=' not in existing_draw_xml.split(">", 1)[0]:
                existing_draw_xml = existing_draw_xml.replace(
                    "<xdr:wsDr",
                    '<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
                    1,
                )
            updated_draw_xml = existing_draw_xml.replace(
                "</xdr:wsDr>", "".join(drawing_xml_items) + "</xdr:wsDr>", 1
            )
            files[draw_path] = updated_draw_xml.encode("utf-8")
            need_drawing_tag = False
        else:
            drawing_counter += 1
            draw_path = f"xl/drawings/drawing{drawing_counter}.xml"
            files[draw_path] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                + "".join(drawing_xml_items)
                + "</xdr:wsDr>"
            ).encode("utf-8")
            ct_overrides.append(
                (f"/{draw_path}", "application/vnd.openxmlformats-officedocument.drawing+xml")
            )
            draw_rid = next_rid(sh_rels_root)
            ET.SubElement(
                sh_rels_root,
                f"{{{ns_rel}}}Relationship",
                {
                    "Id": draw_rid,
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                    "Target": f"../drawings/drawing{drawing_counter}.xml",
                },
            )
            need_drawing_tag = True

        slicer_rid = next_rid(sh_rels_root)
        ET.SubElement(
            sh_rels_root,
            f"{{{ns_rel}}}Relationship",
            {
                "Id": slicer_rid,
                "Type": "http://schemas.microsoft.com/office/2007/relationships/slicer",
                "Target": f"../slicers/slicer{slicer_part_counter}.xml",
            },
        )
        files[sh_rels_path] = serialize_with_default_ns(sh_rels_root, ns_rel)

        sh_xml_str = files[sh_path].decode("utf-8")
        if need_drawing_tag:
            drawing_tag = (
                f'<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                f'r:id="{draw_rid}"/>'
            )
            if "<tableParts" in sh_xml_str:
                sh_xml_str = sh_xml_str.replace("<tableParts", drawing_tag + "<tableParts", 1)
            else:
                sh_xml_str = sh_xml_str.replace("</worksheet>", drawing_tag + "</worksheet>", 1)

        ext_block = (
            '<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" '
            'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            '<x14:slicerList>'
            f'<x14:slicer xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'r:id="{slicer_rid}"/>'
            '</x14:slicerList>'
            '</ext>'
        )
        if "</extLst>" in sh_xml_str:
            sh_xml_str = sh_xml_str.replace("</extLst>", ext_block + "</extLst>", 1)
        else:
            sh_xml_str = sh_xml_str.replace("</worksheet>", f"<extLst>{ext_block}</extLst></worksheet>", 1)

        files[sh_path] = sh_xml_str.encode("utf-8")

    if wb_cache_rids:
        wb_xml_str = files["xl/workbook.xml"].decode("utf-8")
        cache_refs = "".join(
            f'<x14:slicerCache xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/>'
            for rid in wb_cache_rids
        )
        wb_ext = (
            '<ext uri="{7523E5D3-25F3-45E0-1632-64F254C22452}" '
            'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
            f'<x15:slicerCaches>{cache_refs}</x15:slicerCaches>'
            '</ext>'
        )
        if "</extLst>" in wb_xml_str:
            wb_xml_str = wb_xml_str.replace("</extLst>", wb_ext + "</extLst>", 1)
        else:
            wb_xml_str = wb_xml_str.replace("</workbook>", f"<extLst>{wb_ext}</extLst></workbook>", 1)
        files["xl/workbook.xml"] = wb_xml_str.encode("utf-8")

    files["xl/_rels/workbook.xml.rels"] = serialize_with_default_ns(wb_rels_root, ns_rel)
    for part_name, content_type in ct_overrides:
        ET.SubElement(ct_root, f"{{{ns_ct}}}Override", {"PartName": part_name, "ContentType": content_type})
    files["[Content_Types].xml"] = serialize_with_default_ns(ct_root, ns_ct)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    out_buf.seek(0)
    return out_buf.getvalue()


# ---------- EXCEL GENERATOR (4 SHEETS WITH FULL APP-STYLE DASHBOARDS) ----------

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


def write_outstanding_dashboard_sheet(wb: Workbook, title: str, df: pd.DataFrame, table_name: str):
    ws = wb.create_sheet(title=title)
    ws.views.sheetView[0].showGridLines = False

    # Layout:
    # Cols A-C : Left Vertical Slicers
    # Col D    : Spacer
    # Col E+   : Row 1-2 Title | Row 4-5 KPI Cards | Row 7-21 Two Dynamic Charts | Row 23+ Data Table
    col_start = 5   # Column E
    start_row = 23  # Table Header Row
    end_row = max(start_row + len(df), start_row + 1)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 3

    # Styles
    banner_fill = PatternFill("solid", fgColor="1F4E78")
    banner_font = Font(bold=True, size=15, color="FFFFFF")
    slicer_hdr_font = Font(bold=True, size=11, color="1F4E78")
    kpi_label_font = Font(bold=True, size=9, color="595959")
    kpi_val_font = Font(bold=True, size=13, color="1F4E78")
    kpi_red_font = Font(bold=True, size=13, color="C00000")
    kpi_fill = PatternFill("solid", fgColor="EDF2F8")
    kpi_red_fill = PatternFill("solid", fgColor="FCE4D6")
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

    # Left Slicers Header
    ws["A2"] = "FILTERS / SLICERS"
    ws["A2"].font = slicer_hdr_font

    # Top Title Banner (E1:O2)
    ws.merge_cells("E1:O2")
    ws["E1"] = f"  {title} Dashboard"
    ws["E1"].fill = banner_fill
    ws["E1"].font = banner_font
    ws["E1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18

    # Add helper column 'Visible' (1 if row is visible after slicer filter, 0 if hidden)
    table_cols = list(df.columns) + ["Visible"]
    first_tbl_ltr = get_column_letter(col_start)
    party_branch_ltr = get_column_letter(col_start + df.columns.get_loc("Party Name (Branch)"))
    amt_col_idx = col_start + df.columns.get_loc("Invoice Value")
    amt_col_ltr = get_column_letter(amt_col_idx)
    days_col_idx = col_start + df.columns.get_loc("Overdue Days")
    days_col_ltr = get_column_letter(days_col_idx)
    bucket_col_idx = col_start + df.columns.get_loc("Overdue Bucket")
    bucket_col_ltr = get_column_letter(bucket_col_idx)
    vis_col_idx = col_start + len(df.columns)
    vis_col_ltr = get_column_letter(vis_col_idx)

    # 4 Dynamic KPI Cards (Row 4-5)
    kpi_specs = [
        ("E", "F", "Total Outstanding (₹)",
         f"=SUBTOTAL(9,{amt_col_ltr}{start_row+1}:{amt_col_ltr}{end_row})", "₹#,##0.00", kpi_fill, kpi_val_font),
        ("H", "I", "Pending Invoices",
         f"=SUBTOTAL(3,{first_tbl_ltr}{start_row+1}:{first_tbl_ltr}{end_row})", "#,##0", kpi_fill, kpi_val_font),
        ("K", "L", "90+ Days Overdue (₹)",
         f'=SUMIFS({amt_col_ltr}{start_row+1}:{amt_col_ltr}{end_row},{bucket_col_ltr}{start_row+1}:{bucket_col_ltr}{end_row},"90+ Days",{vis_col_ltr}{start_row+1}:{vis_col_ltr}{end_row},1)',
         "₹#,##0.00", kpi_red_fill, kpi_red_font),
        ("N", "O", "Avg Overdue Days",
         f"=IFERROR(SUBTOTAL(1,{days_col_ltr}{start_row+1}:{days_col_ltr}{end_row}),0)", '0" Days"', kpi_fill, kpi_val_font),
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

    # Table Headers (Row 23)
    for offset, col_name in enumerate(table_cols):
        c_idx = col_start + offset
        cell = ws.cell(row=start_row, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Table Data (Row 24 onwards)
    bucket_offset = df.columns.get_loc("Overdue Bucket")
    for r_offset, row in enumerate(df.itertuples(index=False, name=None), start=1):
        r_num = start_row + r_offset
        is_90_plus = row[bucket_offset] == "90+ Days"
        for offset, val in enumerate(row):
            c_idx = col_start + offset
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

        # Helper Visible formula (1 when visible, 0 when filtered out by Slicer)
        vis_cell = ws.cell(row=r_num, column=vis_col_idx, value=f"=SUBTOTAL(103,{first_tbl_ltr}{r_num})")
        vis_cell.border = thin_border

    # Hide Visible helper column so sheet stays super clean
    ws.column_dimensions[vis_col_ltr].hidden = True

    # Add Excel Table
    if len(df) > 0:
        last_col_ltr = get_column_letter(col_start + len(table_cols) - 1)
        tab = Table(displayName=table_name, ref=f"{first_tbl_ltr}{start_row}:{last_col_ltr}{end_row}")
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(tab)

        # ---- DYNAMIC CHART DATA SUMMARY (Hidden Columns Z:AC) ----
        # 1. Overdue Bucket Summary (Z1:AA5)
        ws["Z1"] = "Overdue Bucket"
        ws["AA1"] = "Outstanding (₹)"
        buckets = ["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
        for i, b_name in enumerate(buckets, start=2):
            ws[f"Z{i}"] = b_name
            ws[f"AA{i}"] = (
                f'=SUMIFS({amt_col_ltr}{start_row+1}:{amt_col_ltr}{end_row},'
                f'{bucket_col_ltr}{start_row+1}:{bucket_col_ltr}{end_row},Z{i},'
                f'{vis_col_ltr}{start_row+1}:{vis_col_ltr}{end_row},1)'
            )
            ws[f"AA{i}"].number_format = "#,##0"

        # 2. Top 10 Pending Branches Summary (AB1:AC11)
        ws["AB1"] = "Branch"
        ws["AC1"] = "Outstanding (₹)"
        top10_branches = (
            df.groupby("Party Name (Branch)")["Invoice Value"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .index.tolist()
        )
        for i, br_name in enumerate(top10_branches, start=2):
            ws[f"AB{i}"] = br_name
            ws[f"AC{i}"] = (
                f'=SUMIFS({amt_col_ltr}{start_row+1}:{amt_col_ltr}{end_row},'
                f'{party_branch_ltr}{start_row+1}:{party_branch_ltr}{end_row},AB{i},'
                f'{vis_col_ltr}{start_row+1}:{vis_col_ltr}{end_row},1)'
            )
            ws[f"AC{i}"].number_format = "#,##0"

        for h_col in ["Z", "AA", "AB", "AC"]:
            ws.column_dimensions[h_col].hidden = True

        # ---- CHART 1: Overdue Bucket-wise Outstanding (Anchored at E7) ----
        ch1 = BarChart()
        ch1.type = "col"
        ch1.style = 10
        ch1.title = "Overdue Bucket-wise Outstanding (₹)"
        ch1.y_axis.title = "Amount (₹)"
        ch1.legend = None
        data1 = Reference(ws, min_col=27, min_row=1, max_row=5)
        cats1 = Reference(ws, min_col=26, min_row=2, max_row=5)
        ch1.add_data(data1, titles_from_data=True)
        ch1.set_categories(cats1)
        ch1.width = 14.5
        ch1.height = 7.5
        ws.add_chart(ch1, "E7")

        # ---- CHART 2: Top 10 Pending Branches (Anchored at K7) ----
        if top10_branches:
            ch2 = BarChart()
            ch2.type = "bar"
            ch2.style = 11
            ch2.title = "Top 10 Pending Branches (₹)"
            ch2.x_axis.title = "Amount (₹)"
            ch2.legend = None
            max_r2 = 1 + len(top10_branches)
            data2 = Reference(ws, min_col=29, min_row=1, max_row=max_r2)
            cats2 = Reference(ws, min_col=28, min_row=2, max_row=max_r2)
            ch2.add_data(data2, titles_from_data=True)
            ch2.set_categories(cats2)
            ch2.width = 16.5
            ch2.height = 7.5
            ws.add_chart(ch2, "K7")

    # Auto-fit Table columns
    for offset, col_name in enumerate(df.columns):
        c_idx = col_start + offset
        max_len = max(len(str(col_name)), 12)
        if len(df) > 0:
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
    write_outstanding_dashboard_sheet(wb, "Sales Outstanding", sales_out, "SalesOutstandingTbl")
    write_outstanding_dashboard_sheet(wb, "Purchase Outstanding", purchase_out, "PurchaseOutstandingTbl")

    raw_buf = io.BytesIO()
    wb.save(raw_buf)

    # Left-aligned vertical slicers (Cols A-C) matching the Streamlit App
    default_slicers = [
        {
            "col_name": "Party Group",
            "caption": "1. Party / Company Group",
            "from_col": 0, "from_row": 3, "to_col": 3, "to_row": 16,
        },
        {
            "col_name": "Party Name (Branch)",
            "caption": "2. Branches / Sub-Parties",
            "from_col": 0, "from_row": 17, "to_col": 3, "to_row": 31,
        },
        {
            "col_name": "Overdue Bucket",
            "caption": "3. Overdue Days",
            "from_col": 0, "from_row": 32, "to_col": 3, "to_row": 41,
        },
    ]

    slicer_configs = []
    if len(sales_out) > 0:
        slicer_configs.append({
            "sheet_name": "Sales Outstanding",
            "table_name": "SalesOutstandingTbl",
            "slicers": default_slicers,
        })
    if len(purchase_out) > 0:
        slicer_configs.append({
            "sheet_name": "Purchase Outstanding",
            "table_name": "PurchaseOutstandingTbl",
            "slicers": default_slicers,
        })

    if slicer_configs:
        return inject_excel_table_slicers(raw_buf.getvalue(), slicer_configs)
    return raw_buf.getvalue()


# ---------- INTERACTIVE DASHBOARD RENDERER ----------

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
        critical_amt = filtered.loc[filtered["Overdue Bucket"] == "90+ Days", "Invoice Value"].sum()
        avg_days = int(filtered["Overdue Days"].mean()) if total_bills > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Outstanding", f"₹{total_amt:,.2f}")
        k2.metric("Pending Invoices", f"{total_bills:,}")
        k3.metric("90+ Days Overdue", f"₹{critical_amt:,.2f}")
        k4.metric("Avg Overdue Days", f"{avg_days} Days")

        c1, c2 = st.columns(2)
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

with st.spinner("Files process ho rahi hain aur Excel Dashboards + Slicers generate ho rahe hain..."):
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

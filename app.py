import difflib
import io
import re
import sqlite3
import openpyxl
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. DATABASE & PERSISTENCE LAYER (SQLite)
# ==========================================
DB_FILE = "quotation_master.db"


def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS master_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            main_material TEXT NOT NULL,
            sub_material TEXT NOT NULL UNIQUE,
            purchase_price REAL DEFAULT 0.0,
            discount REAL DEFAULT 0.0
        )
    """)
    conn.commit()
    return conn


def load_master_data():
    conn = get_db_connection()
    df = pd.read_sql_query(
        "SELECT id, main_material, sub_material, purchase_price, discount FROM master_items ORDER BY main_material, sub_material",
        conn,
    )
    conn.close()
    return df


def save_master_data(df):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM master_items")
    for _, row in df.iterrows():
        main_mat = str(row.get("main_material", "")).strip()
        sub_mat = str(row.get("sub_material", "")).strip()
        if not sub_mat:
            continue
        try:
            p_price = float(row.get("purchase_price", 0.0) or 0.0)
        except (ValueError, TypeError):
            p_price = 0.0
        try:
            disc = float(row.get("discount", 0.0) or 0.0)
        except (ValueError, TypeError):
            disc = 0.0

        cursor.execute(
            """
            INSERT OR REPLACE INTO master_items (main_material, sub_material, purchase_price, discount)
            VALUES (?, ?, ?, ?)
        """,
            (main_mat, sub_mat, p_price, disc),
        )
    conn.commit()
    conn.close()


def clear_master_db():
    conn = get_db_connection()
    conn.execute("DELETE FROM master_items")
    conn.commit()
    conn.close()


# ==========================================
# 2. FILE PARSING (EXCEL / XLSM / PDF)
# ==========================================
KNOWN_MAIN_CATEGORIES = [
    "CABLE GLANDS - POWER WITH SUITABLE RUBBER HOOD",
    "CABLE GLANDS -POWER WITH SUITABLE RUBBER HOOD",
    "CABLE GLANDS -POWER WITH RUBBER HOOD",
    "CABLE GLANDS - POWER",
    "CABLE GLANDS - SIGNAL",
    "CABLE GLAND - SIGNAL",
    "CABLE GLANDS  SIGNAL",
    "CABLE GLANDS SIGNAL",
    "REDUCER- ADAPTOR POWER",
    "REDUCER-ADAPTOR POWER",
    "REDUCER",
    "BLINDS POWER & SIGNAL",
    "BLIND - SIGNAL",
    "BLINDSIGNAL",
    "POWER CABLE LUGS & EARTHING LUGS (TINNED COPPER)",
    "LUGS POWER + EARTHING (TINNED COPPER)",
    "SIGNAL CABLE LUGS",
    "LUGS SIGNALS",
    "LUGS SIGNAL",
    "SLEEVES",
    "CABLE LABELS",
    "CABLE LABLES",
    "FERRULES",
    "CABLE TIE",
    "MISCELLANEOUS",
    "MISCELENIOUS",
]


def clean_text(val):
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def parse_master_from_excel(file_bytes):
    extracted_data = []
    wb_dict = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    for sheet_name, df in wb_dict.items():
        if "metadata" in sheet_name.lower():
            continue

        current_main = "MISCELLANEOUS"
        for _, row in df.iterrows():
            row_vals = [clean_text(v) for v in row if clean_text(v)]
            if not row_vals:
                continue

            first_col = row_vals[0]
            upper_first = first_col.upper()

            is_main = any(cat in upper_first for cat in KNOWN_MAIN_CATEGORIES)
            if is_main and (
                len(row_vals) <= 2
                or "POWER" in upper_first
                or "GLAND" in upper_first
            ):
                current_main = first_col
                continue

            desc = ""
            if (
                len(row_vals) >= 2
                and (
                    row_vals[0].isdigit()
                    or (len(row_vals[0]) <= 4 and "." in row_vals[0])
                )
            ):
                desc = row_vals[1]
            else:
                desc = row_vals[0]

            if (
                desc.lower()
                in [
                    "description",
                    "item",
                    "sr.no",
                    "particulars",
                    "part number",
                    "qty",
                    "remark",
                    "date",
                ]
                or desc.isdigit()
                or desc.startswith("ALFA LAVAL")
                or desc.startswith("PROJECT :")
                or desc.startswith("BILL OF MATERIAL")
            ):
                continue

            if len(desc) > 3:
                extracted_data.append(
                    {
                        "main_material": current_main,
                        "sub_material": desc,
                        "purchase_price": 0.0,
                        "discount": 0.0,
                    }
                )

    return pd.DataFrame(extracted_data)


def parse_master_from_pdf(file_bytes):
    extracted_data = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        current_main = "MISCELLANEOUS"
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                line_clean = clean_text(line)
                if not line_clean:
                    continue

                upper_line = line_clean.upper()
                is_main = any(cat in upper_line for cat in KNOWN_MAIN_CATEGORIES)
                if is_main and len(line_clean) < 60:
                    current_main = line_clean
                    continue

                sub_mat = re.sub(r"^\d+[\.\)\s]+", "", line_clean).strip()
                if (
                    sub_mat
                    and len(sub_mat) > 3
                    and not sub_mat.lower().startswith("sr.no")
                ):
                    extracted_data.append(
                        {
                            "main_material": current_main,
                            "sub_material": sub_mat,
                            "purchase_price": 0.0,
                            "discount": 0.0,
                        }
                    )

    return pd.DataFrame(extracted_data)


def parse_bom_file(uploaded_file):
    items = []
    fname = uploaded_file.name.lower()

    if fname.endswith((".xlsx", ".xls", ".xlsm")):
        wb_dict = pd.read_excel(
            uploaded_file,
            sheet_name=None,
            header=None,
            engine="openpyxl" if not fname.endswith(".xls") else None,
        )

        target_sheets = [s for s in wb_dict.keys() if "bom" in s.lower()]
        sheets_to_process = target_sheets if target_sheets else list(wb_dict.keys())

        for s_name in sheets_to_process:
            df = wb_dict[s_name]
            desc_col_idx = None
            qty_col_idx = None

            for r_idx in range(min(20, len(df))):
                row_str = [str(x).lower() for x in df.iloc[r_idx].values]
                for c_idx, val in enumerate(row_str):
                    if any(
                        k in val for k in ["description", "sub material", "item"]
                    ):
                        desc_col_idx = c_idx
                    if any(
                        k in val
                        for k in ["qty", "qwt", "quantity", "total qty"]
                    ):
                        qty_col_idx = c_idx
                if desc_col_idx is not None and qty_col_idx is not None:
                    df = df.iloc[r_idx + 1 :]
                    break

            if desc_col_idx is None:
                desc_col_idx = 1 if len(df.columns) > 1 else 0
            if qty_col_idx is None:
                qty_col_idx = 3 if len(df.columns) > 3 else desc_col_idx

            for _, row in df.iterrows():
                desc = clean_text(row.iloc[desc_col_idx])
                qty_val = str(row.iloc[qty_col_idx])
                qty_match = re.search(r"(\d+(\.\d+)?)", qty_val)
                qty = float(qty_match.group(1)) if qty_match else 0.0

                if (
                    desc
                    and len(desc) > 2
                    and not desc.isdigit()
                    and desc.lower() not in ["description", "part number"]
                ):
                    items.append({"bom_description": desc, "qty": qty})

    elif fname.endswith(".pdf"):
        with pdfplumber.open(io.BytesIO(uploaded_file.getvalue())) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row or len(row) < 2:
                            continue
                        row_clean = [clean_text(c) for c in row if c]
                        if len(row_clean) >= 2:
                            desc = row_clean[1] if len(row_clean) > 2 else row_clean[0]
                            qty_str = (
                                row_clean[2]
                                if len(row_clean) > 2
                                else row_clean[-1]
                            )
                            qty_match = re.search(r"(\d+(\.\d+)?)", qty_str)
                            qty = float(qty_match.group(1)) if qty_match else 0.0
                            if desc and len(desc) > 2:
                                items.append(
                                    {"bom_description": desc, "qty": qty}
                                )

    return pd.DataFrame(items)


# ==========================================
# 3. EXCEL EXPORT (EXACT 6500389 FORMAT)
# ==========================================
def build_quotation_excel_exact_6500389(
    quotation_df, profit_pct, annexure_title="Annexure Quotation"
):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Row 1: Title
    ws["B1"] = annexure_title
    ws["B1"].font = Font(name="Calibri", size=12, bold=True)

    # Row 2: Standard Headers & Profit
    ws["D2"] = "QTY"
    ws["E2"] = "Amnt"
    ws["F2"] = "Discount"
    ws["G2"] = "Total Amnt"
    ws["I2"] = "Purchase"
    ws["K2"] = "Profit"
    ws["L2"] = float(profit_pct)  # e.g., 12.0 for 12%

    header_font = Font(name="Calibri", size=11, bold=True)
    for col_ref in ["D2", "E2", "F2", "G2", "I2", "K2", "L2"]:
        ws[col_ref].font = header_font
        if col_ref not in ["K2", "L2"]:
            ws[col_ref].alignment = Alignment(
                horizontal="center", vertical="center"
            )

    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    current_r = 2
    first_category = True
    data_row_indices = []

    grouped = quotation_df.groupby("main_material", sort=False)

    for main_mat, group in grouped:
        if first_category:
            ws["A2"] = str(main_mat)
            ws["A2"].font = header_font
            current_r = 3
            first_category = False
        else:
            ws.cell(
                row=current_r, column=1, value=str(main_mat)
            ).font = header_font
            current_r += 1

        sr_no = 1
        for _, item in group.iterrows():
            r = current_r
            data_row_indices.append(r)

            ws.cell(row=r, column=1, value=sr_no).alignment = Alignment(
                horizontal="center"
            )
            ws.cell(row=r, column=2, value=str(item["sub_material"]))
            ws.cell(row=r, column=3, value="Reputed")
            ws.cell(row=r, column=4, value=float(item.get("qty", 0) or 0))

            # Exact quotation 6500389 formula: =((I{r}+J{r})+$L$2%*(I{r}+J{r}))
            ws.cell(row=r, column=5, value=f"=((I{r}+J{r})+$L$2%*(I{r}+J{r}))")

            disc = float(item.get("discount", 0) or 0)
            ws.cell(row=r, column=6, value=disc)

            # Exact quotation 6500389 formula: =(E{r}-(F{r}%*E{r})) *D{r}
            ws.cell(row=r, column=7, value=f"=(E{r}-(F{r}%*E{r})) *D{r}")

            p_price = float(item.get("purchase_price", 0) or 0)
            ws.cell(row=r, column=9, value=p_price)
            ws.cell(row=r, column=10, value=0)  # Column J spacer (0)

            ws[f"D{r}"].number_format = "#,##0"
            ws[f"E{r}"].number_format = "#,##0.00"
            ws[f"F{r}"].number_format = "0.00"
            ws[f"G{r}"].number_format = "#,##0.00"
            ws[f"I{r}"].number_format = "#,##0.00"

            for c in range(1, 11):
                ws.cell(row=r, column=c).border = thin_border

            sr_no += 1
            current_r += 1

    # Total Row
    total_r = current_r
    ws.cell(row=total_r, column=5, value="TOTAL").font = header_font
    ws.cell(row=total_r, column=5).alignment = Alignment(horizontal="center")

    if data_row_indices:
        sum_formula = f"=SUM(G{data_row_indices[0]}:G{data_row_indices[-1]})"
    else:
        sum_formula = "=0.00"

    ws.cell(row=total_r, column=7, value=sum_formula).font = header_font
    ws.cell(row=total_r, column=7).number_format = "#,##0.00"

    # Set practical column widths
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 4
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["J"].width = 6
    ws.column_dimensions["K"].width = 10
    ws.column_dimensions["L"].width = 10

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ==========================================
# 4. STREAMLIT APPLICATION INTERFACE
# ==========================================
st.set_page_config(page_title="Panel Quotation Generator", layout="wide")
st.title("Industrial Quotation Generator")

if "quotation_records" not in st.session_state:
    st.session_state.quotation_records = []

tabs = st.tabs(
    ["1. Quotation Section", "2. Master Sheet Section", "3. File Upload Section"]
)

# ------------------------------------------
# TAB 1: QUOTATION GENERATION
# ------------------------------------------
with tabs[0]:
    st.subheader("Generate Quotation from BOM Files")

    col_cfg1, col_cfg2 = st.columns([3, 3])
    with col_cfg1:
        profit_input = st.number_input(
            "Profit % (Cell L2 in quotation 6500389)",
            min_value=0.0,
            max_value=200.0,
            value=12.0,
            step=0.5,
        )
    with col_cfg2:
        if st.button("Delete / Clear Quotation Data", type="secondary"):
            st.session_state.quotation_records = []
            st.success("Quotation table wiped clean.")

    uploaded_boms = st.file_uploader(
        "Upload 1 or Multiple BOM Files (.xlsx, .xls, .xlsm, .pdf)",
        type=["xlsx", "xls", "xlsm", "pdf"],
        accept_multiple_files=True,
        key="bom_uploader",
    )

    master_df = load_master_data()

    if uploaded_boms:
        if master_df.empty:
            st.warning(
                "Master Sheet is currently empty. Please add or upload materials in Section 2 or 3 first."
            )
        else:
            all_bom_items = []
            for bom_file in uploaded_boms:
                parsed_bom = parse_bom_file(bom_file)
                if not parsed_bom.empty:
                    all_bom_items.append(parsed_bom)

            if all_bom_items:
                combined_bom = (
                    pd.concat(all_bom_items, ignore_index=True)
                    .groupby("bom_description", as_index=False)["qty"]
                    .sum()
                )

                master_sub_list = master_df["sub_material"].tolist()
                matched_rows = []
                unmatched_items = []

                for _, item in combined_bom.iterrows():
                    bom_desc = item["bom_description"]
                    qty = item["qty"]

                    exact = master_df[
                        master_df["sub_material"].str.lower()
                        == bom_desc.lower()
                    ]
                    if not exact.empty:
                        m_row = exact.iloc[0]
                        matched_rows.append(
                            {
                                "main_material": m_row["main_material"],
                                "sub_material": m_row["sub_material"],
                                "qty": qty,
                                "purchase_price": m_row["purchase_price"],
                                "discount": m_row["discount"],
                            }
                        )
                    else:
                        unmatched_items.append({"desc": bom_desc, "qty": qty})

                if unmatched_items:
                    st.info(
                        "Review Unmatched Items: Select matching item from Master Sheet."
                    )
                    for idx, un_item in enumerate(unmatched_items):
                        u_name = un_item["desc"]
                        u_qty = un_item["qty"]

                        close_matches = difflib.get_close_matches(
                            u_name, master_sub_list, n=3, cutoff=0.3
                        )
                        options = (
                            ["-- Exclude / Reject --"]
                            + close_matches
                            + ["-- Keep as Custom New Item --"]
                        )

                        col_u1, col_u2 = st.columns([3, 2])
                        with col_u1:
                            st.write(
                                f"**BOM Item:** `{u_name}` (Qty: {u_qty})"
                            )
                        with col_u2:
                            choice = st.selectbox(
                                f"Match candidate #{idx+1}:",
                                options,
                                key=f"fuzzy_{idx}",
                            )

                        if (
                            choice
                            and choice != "-- Exclude / Reject --"
                            and choice != "-- Keep as Custom New Item --"
                        ):
                            m_row = master_df[
                                master_df["sub_material"] == choice
                            ].iloc[0]
                            matched_rows.append(
                                {
                                    "main_material": m_row["main_material"],
                                    "sub_material": m_row["sub_material"],
                                    "qty": u_qty,
                                    "purchase_price": m_row["purchase_price"],
                                    "discount": m_row["discount"],
                                }
                            )
                        elif choice == "-- Keep as Custom New Item --":
                            matched_rows.append(
                                {
                                    "main_material": "MISCELLANEOUS",
                                    "sub_material": u_name,
                                    "qty": u_qty,
                                    "purchase_price": 0.0,
                                    "discount": 0.0,
                                }
                            )

                st.session_state.quotation_records = matched_rows

    if st.session_state.quotation_records:
        q_df = pd.DataFrame(st.session_state.quotation_records)

        # UI Preview calculations matching quotation 6500389
        q_df["Amnt"] = q_df["purchase_price"] * (1 + (profit_input / 100.0))
        q_df["Total Amnt"] = (
            q_df["qty"] * q_df["Amnt"] * (1 - (q_df["discount"] / 100.0))
        )

        st.markdown("### Quotation Summary Table")
        st.dataframe(
            q_df[
                [
                    "main_material",
                    "sub_material",
                    "qty",
                    "purchase_price",
                    "Amnt",
                    "discount",
                    "Total Amnt",
                ]
            ].style.format(
                {
                    "qty": "{:,.0f}",
                    "purchase_price": "₹{:,.2f}",
                    "Amnt": "₹{:,.2f}",
                    "discount": "{:.1f}%",
                    "Total Amnt": "₹{:,.2f}",
                }
            ),
            width="stretch",
        )

        grand_total = q_df["Total Amnt"].sum()
        st.markdown(
            f"### **Grand Total Amount: ₹ {grand_total:,.2f}** (Includes {profit_input}% profit)"
        )

        excel_buffer = build_quotation_excel_exact_6500389(q_df, profit_input)
        st.download_button(
            label="Download Complete Quotation (.xlsx)",
            data=excel_buffer,
            file_name="Quotation_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ------------------------------------------
# TAB 2: MASTER SHEET
# ------------------------------------------
with tabs[1]:
    st.subheader("Master Sheet Management")
    st.caption(
        "Edits saved here remain persistent in the database across sessions and restarts."
    )

    current_master = load_master_data()

    col_btn1, col_btn2 = st.columns([2, 4])
    with col_btn1:
        if st.button("Delete Entire Master Sheet", type="secondary"):
            clear_master_db()
            st.success("Master sheet completely wiped.")
            st.rerun()

    edited_df = st.data_editor(
        current_master,
        num_rows="dynamic",
        column_config={
            "id": None,
            "main_material": st.column_config.TextColumn(
                "Main Material", required=True
            ),
            "sub_material": st.column_config.TextColumn(
                "Sub Material", required=True
            ),
            "purchase_price": st.column_config.NumberColumn(
                "Purchase Price", min_value=0.0, format="₹%.2f"
            ),
            "discount": st.column_config.NumberColumn(
                "Discount %", min_value=0.0, max_value=100.0, format="%.1f%%"
            ),
        },
        width="stretch",
        key="master_editor",
    )

    if st.button("Save Master Sheet Changes", type="primary"):
        save_master_data(edited_df)
        st.success(f"Saved {len(edited_df)} items to the Master Sheet database.")
        st.rerun()

# ------------------------------------------
# TAB 3: FILE UPLOAD SECTION
# ------------------------------------------
with tabs[2]:
    st.subheader("Populate Master Sheet from Raw Files")
    st.write(
        "Upload reference quotation sheets, BOMs, or annexures (.xlsx, .xls, .xlsm, .pdf)."
    )

    ref_files = st.file_uploader(
        "Upload Quotation / Annexure / BOM Files",
        type=["xlsx", "xls", "xlsm", "pdf"],
        accept_multiple_files=True,
        key="ref_uploader",
    )

    if ref_files and st.button("Extract and Append to Master Sheet"):
        total_extracted = []
        for f in ref_files:
            bytes_data = f.read()
            fname = f.name.lower()
            if fname.endswith((".xlsx", ".xls", ".xlsm")):
                extracted = parse_master_from_excel(bytes_data)
            else:
                extracted = parse_master_from_pdf(bytes_data)
            total_extracted.append(extracted)

        if total_extracted:
            combined_extracted = pd.concat(
                total_extracted, ignore_index=True
            ).drop_duplicates(subset=["sub_material"])
            existing_db = load_master_data()

            final_df = (
                pd.concat([existing_db, combined_extracted], ignore_index=True)
                .drop_duplicates(subset=["sub_material"], keep="first")
                .fillna(0.0)
            )

            save_master_data(final_df)
            st.success(
                f"Successfully extracted {len(combined_extracted)} items and synchronized with Master Sheet!"
            )
            st.dataframe(combined_extracted.head(20), width="stretch")

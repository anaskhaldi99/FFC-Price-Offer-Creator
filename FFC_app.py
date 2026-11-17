from datetime import datetime
import io
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.shared import Inches, Pt
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Price Offer Generator",
    layout="centered",
    page_icon="📄",
    initial_sidebar_state="collapsed",
)


def fetch_base_products(file_buffer, codes_str):
    """Fetches base product details and list prices from the 'Old to New' sheet."""
    code_list = [
        code.strip() for code in codes_str.split(",") if code.strip()
    ]

    df = pd.read_excel(file_buffer, sheet_name="Old to New")
    df.columns = df.columns.str.strip()

    base_results = []

    for idx, code in enumerate(code_list, start=1):
        row = df[df["PRODUCT CODE"].astype(str).str.strip() == str(code)]

        if not row.empty:
            row_data = row.iloc[0]

            product_name = str(row_data.get("PRODUCT NAME", "")).strip()
            if product_name.lower() == "nan":
                product_name = ""

            designer_name = str(row_data.get("DESIGNER NAME", "")).strip()
            designer_variant = str(row_data.get("DESIGNER VARIANT", "")).strip()

            if designer_name.lower() == "nan":
                designer_name = ""
            if designer_variant.lower() == "nan":
                designer_variant = ""

            if designer_name and designer_variant:
                product_details = f"{designer_name} - {designer_variant}"
            else:
                product_details = designer_name or designer_variant or "-"

            gender = str(row_data.get("GENDER", "")).strip()
            if gender.lower() == "nan":
                gender = ""

            try:
                list_price = float(row_data.get("LIST PRICE USD", 0.0))
            except (ValueError, TypeError):
                list_price = 0.0

            base_results.append(
                {
                    "SR NO.": idx,
                    "PRODUCT CODE": int(code) if code.isdigit() else code,
                    "PRODUCT NAME": product_name,
                    "PRODUCT DETAILS": product_details,
                    "GENDER": gender,
                    "LIST PRICE USD": list_price,
                    "FOUND": True,
                }
            )
        else:
            base_results.append(
                {
                    "SR NO.": idx,
                    "PRODUCT CODE": code,
                    "PRODUCT NAME": "NOT FOUND",
                    "PRODUCT DETAILS": "-",
                    "GENDER": "-",
                    "LIST PRICE USD": 0.0,
                    "FOUND": False,
                }
            )

    return pd.DataFrame(base_results)


def apply_discounts_and_build_dfs(base_df, discounts_list):
    """Applies per-row discounts and builds final DataFrames for Word and Excel."""
    word_results = []
    excel_results = []

    for idx, row in base_df.iterrows():
        disc_pct = discounts_list[idx]
        discount_str = (
            f"{int(disc_pct)}%" if disc_pct.is_integer() else f"{disc_pct}%"
        )

        list_price = row["LIST PRICE USD"]
        if disc_pct > 0 and row["FOUND"]:
            customer_price = list_price * (1.0 - (disc_pct / 100.0))
        else:
            customer_price = list_price

        # Word Data (Truncated price to whole integer)
        word_results.append(
            {
                "SR NO.": row["SR NO."],
                "PRODUCT CODE": row["PRODUCT CODE"],
                "PRODUCT NAME": row["PRODUCT NAME"],
                "PRODUCT DETAILS": row["PRODUCT DETAILS"],
                "GENDER": row["GENDER"],
                "PRICE/KG ($)": int(customer_price) if row["FOUND"] else 0,
            }
        )

        # Excel Data (Exact unrounded two-decimal floats)
        excel_results.append(
            {
                "serial no.": row["SR NO."],
                "product code": row["PRODUCT CODE"],
                "product name": row["PRODUCT NAME"],
                "product details": row["PRODUCT DETAILS"],
                "gender": row["GENDER"],
                "list price in usd": round(list_price, 2),
                "customer price in usd": round(customer_price, 2),
                "discount": discount_str,
            }
        )

    return pd.DataFrame(word_results), pd.DataFrame(excel_results)


def generate_excel_file(df_excel):
    """Generates an Excel file in memory with yellow headers and two-decimal price formatting."""
    excel_io = io.BytesIO()

    with pd.ExcelWriter(excel_io, engine="openpyxl") as writer:
        df_excel.to_excel(writer, index=False, sheet_name="Price Offer")
        worksheet = writer.sheets["Price Offer"]

        yellow_fill = PatternFill(
            start_color="FFFF00", end_color="FFFF00", fill_type="solid"
        )
        header_font = Font(name="Calibri", bold=True)

        # Style Headers
        for cell in worksheet[1]:
            cell.fill = yellow_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Format price columns (columns 6 and 7) with 2 decimal places
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            row[5].number_format = "0.00"  # list price in usd
            row[6].number_format = "0.00"  # customer price in usd

        # Auto-adjust column widths
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(
                max_len + 3, 12
            )

    excel_io.seek(0)
    return excel_io


def update_word_document(
    template_path, customer_name, customer_address, date_str, df_word
):
    """Loads template.docx, updates metadata, populates the table, and forces Calibri font."""
    doc = docx.Document(template_path)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"

    for p in doc.paragraphs:
        if "CUSTOMER NAME:" in p.text:
            p.text = ""
            p.paragraph_format.tab_stops.clear_all()
            p.paragraph_format.tab_stops.add_tab_stop(
                Inches(6.5), WD_TAB_ALIGNMENT.RIGHT
            )

            r1 = p.add_run(f"CUSTOMER NAME: {customer_name.strip().upper()}")
            r1.font.name = "Calibri"

            r2 = p.add_run(f"\tDATE: {date_str}")
            r2.font.name = "Calibri"

        elif "CUSTOMER ADDRESS:" in p.text:
            p.text = ""
            addr_text = (
                customer_address.strip().upper()
                if customer_address.strip()
                else ""
            )
            r = p.add_run(f"CUSTOMER ADDRESS: {addr_text}")
            r.font.name = "Calibri"

        else:
            for run in p.runs:
                run.font.name = "Calibri"

    if doc.tables:
        table = doc.tables[0]
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        while len(table.rows) > 1:
            table._tbl.remove(table.rows[-1]._tr)

        hdr_cells = table.rows[0].cells
        for cell in hdr_cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(12)
                    run.font.bold = True

        for _, row in df_word.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row["SR NO."])
            row_cells[1].text = str(row["PRODUCT CODE"])
            row_cells[2].text = str(row["PRODUCT NAME"])
            row_cells[3].text = str(row["PRODUCT DETAILS"])
            row_cells[4].text = str(row["GENDER"])
            row_cells[5].text = str(row["PRICE/KG ($)"])

            for col_idx, cell in enumerate(row_cells):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.name = "Calibri"
                        run.font.size = Pt(10)
                        if col_idx == 5:
                            run.font.bold = True

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


def parse_discount_val(val_str):
    """Safely converts a string discount input to a float (blank = 0.0)."""
    if not val_str or not str(val_str).strip():
        return 0.0
    try:
        return float(str(val_str).replace("%", "").strip())
    except ValueError:
        return 0.0


# --- STREAMLIT UI ---
st.title("📋 Price Offer Generator")

# Initialize session states
if "base_fetched" not in st.session_state:
    st.session_state.base_fetched = False
if "files_generated" not in st.session_state:
    st.session_state.files_generated = False

# --- STEP 1: PRODUCT & CUSTOMER INPUTS ---
uploaded_file = st.file_uploader(
    "Upload Catalog Excel File (Kütüphane.xlsx)", type=["xlsx"]
)

customer_name = st.text_input(
    "Customer Name *", placeholder="e.g. AL KAHLAA PERFUMES"
)

customer_address = st.text_input(
    "Customer Address (Optional)", placeholder="e.g. SAJAA, SHARJAH"
)

codes_input = st.text_input(
    "Product Codes (comma separated) *",
    placeholder="e.g. 100416, 100419, 100421",
)

if st.button("1. Fetch Products & Preview Prices", type="primary"):
    if not uploaded_file:
        st.error("⚠️ Please upload the Excel file before proceeding.")
    elif not customer_name.strip():
        st.warning(
            "⚠️ Customer Name is required. Please enter a customer name."
        )
    elif not codes_input.strip():
        st.warning(
            "⚠️ Product Codes are required. Please enter at least one product code."
        )
    else:
        try:
            base_df = fetch_base_products(uploaded_file, codes_input)
            st.session_state.base_df = base_df
            st.session_state.saved_customer_name = customer_name.strip()
            st.session_state.saved_customer_address = customer_address.strip()
            st.session_state.base_fetched = True
            st.session_state.files_generated = (
                False  # Reset final files when new products are fetched
            )
        except Exception as e:
            st.error(f"An error occurred while reading the Excel file: {e}")

# --- STEP 2: PREVIEW & DISCOUNT SELECTION ---
if st.session_state.base_fetched:
    st.divider()
    st.subheader("🔍 Product Preview (Base List Prices)")

    preview_display = st.session_state.base_df.drop(columns=["FOUND"]).copy()
    preview_display["LIST PRICE USD"] = preview_display[
        "LIST PRICE USD"
    ].map(lambda x: f"{x:.2f}")
    st.dataframe(preview_display, use_container_width=True, hide_index=True)

    st.subheader("🏷️ Apply Discounts")
    discount_mode = st.radio(
        "Choose how you want to apply discounts:",
        options=[
            "Apply One Discount to All Products",
            "Apply Individual Discounts per Product",
        ],
        horizontal=True,
    )

    discounts_to_apply = []

    if discount_mode == "Apply One Discount to All Products":
        global_disc_input = st.text_input(
            "Discount % for All Products (Leave blank for 0%)",
            placeholder="e.g. 10",
            key="global_disc_box",
        )
        global_disc_val = parse_discount_val(global_disc_input)
        discounts_to_apply = [global_disc_val] * len(st.session_state.base_df)

    else:
        st.caption(
            "Enter a discount percentage for each product below (leave blank for 0%):"
        )
        for idx, row in st.session_state.base_df.iterrows():
            col_info, col_input = st.columns([3, 1])
            with col_info:
                st.markdown(
                    f"**{row['SR NO.']}. [{row['PRODUCT CODE']}] {row['PRODUCT NAME']}** "
                    f"— List Price: `${row['LIST PRICE USD']:.2f}`"
                )
            with col_input:
                indiv_input = st.text_input(
                    f"Discount % ({row['PRODUCT CODE']})",
                    placeholder="0%",
                    label_visibility="collapsed",
                    key=f"indiv_disc_{idx}_{row['PRODUCT CODE']}",
                )
                discounts_to_apply.append(parse_discount_val(indiv_input))

    if st.button("2. Generate Word & Excel Files", type="primary"):
        try:
            word_df, excel_df = apply_discounts_and_build_dfs(
                st.session_state.base_df, discounts_to_apply
            )

            now = datetime.now()
            date_str = f"{now.day}/{now.month}/{now.year}"

            clean_cust_name = st.session_state.saved_customer_name.upper()
            base_filename = (
                f"{clean_cust_name} PRICE OFFER {now.day}-{now.month}-{now.year}"
            )

            template_filename = "template.docx"
            word_file = update_word_document(
                template_filename,
                st.session_state.saved_customer_name,
                st.session_state.saved_customer_address,
                date_str,
                word_df,
            )
            excel_file = generate_excel_file(excel_df)

            # Store in session state so download buttons persist
            st.session_state.word_df = word_df
            st.session_state.excel_df = excel_df
            st.session_state.base_filename = base_filename
            st.session_state.word_file_data = word_file.getvalue()
            st.session_state.excel_file_data = excel_file.getvalue()
            st.session_state.files_generated = True

        except FileNotFoundError:
            st.error(
                "❌ Could not find 'template.docx'. Make sure your reference file is saved as 'template.docx' in the same folder as app.py."
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")

# --- STEP 3: FINAL PREVIEW & PERSISTENT DOWNLOAD BUTTONS ---
if st.session_state.files_generated:
    st.divider()
    st.subheader("📋 Final Output Preview")

    st.write("**Word Table Preview (Truncated Prices):**")
    st.dataframe(
        st.session_state.word_df, use_container_width=True, hide_index=True
    )

    st.write("**Excel Sheet Preview (Exact Two-Decimal Prices):**")
    excel_preview_display = st.session_state.excel_df.copy()
    excel_preview_display["list price in usd"] = excel_preview_display[
        "list price in usd"
    ].map(lambda x: f"{x:.2f}")
    excel_preview_display["customer price in usd"] = excel_preview_display[
        "customer price in usd"
    ].map(lambda x: f"{x:.2f}")
    st.dataframe(
        excel_preview_display, use_container_width=True, hide_index=True
    )

    st.subheader("📥 Downloads")
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="📄 Download Word File (.docx)",
            data=st.session_state.word_file_data,
            file_name=f"{st.session_state.base_filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="btn_download_word",
        )

    with col2:
        st.download_button(
            label="📊 Download Excel File (.xlsx)",
            data=st.session_state.excel_file_data,
            file_name=f"{st.session_state.base_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_download_excel",
        )

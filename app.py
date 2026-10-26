import streamlit as st
import openpyxl
import subprocess
import tempfile
import os
import io

st.set_page_config(page_title="GST Invoice Modifier", page_icon="📄", layout="centered")

st.title("📄 GST Invoice Modifier & Converter")
st.write("Upload your Excel template, update details, and download as Excel or PDF!")

# File uploader
uploaded_file = st.file_uploader("Upload Excel Invoice (.xlsx)", type=["xlsx"])

# Store uploaded file bytes and original filename safely in session state
if uploaded_file is not None:
    st.session_state["file_bytes"] = uploaded_file.getvalue()
    st.session_state["original_filename"] = os.path.splitext(uploaded_file.name)[0]

# Extract default values from cells if file bytes exist
inv_no_default = ""
inv_date_default = ""
rental_default = ""

if "file_bytes" in st.session_state and st.session_state["file_bytes"] is not None:
    try:
        wb_temp = openpyxl.load_workbook(io.BytesIO(st.session_state["file_bytes"]), data_only=True)
        sheet_temp = wb_temp.active
        inv_no_default = str(sheet_temp['G6'].value or "")
        inv_date_default = str(sheet_temp['G7'].value or "")
        rental_default = str(sheet_temp['C17'].value or "")
    except Exception:
        pass

invoice_no = st.text_input("Invoice Number (Cell G6)", value=inv_no_default)
invoice_date = st.text_input("Invoice Date (Cell G7)", value=inv_date_default)
rental_text = st.text_area("Rental Description & Month (Cell C17)", value=rental_default)

# Use only the original uploaded file name as the default output name
orig_name = st.session_state.get("original_filename", "Invoice")
output_filename = st.text_input("Output File Name", value=orig_name)

if st.button("🔄 Generate Files", type="primary"):
    if "file_bytes" not in st.session_state or st.session_state["file_bytes"] is None:
        st.warning("Please upload an Excel file first.")
    elif not invoice_no or not invoice_date or not rental_text:
        st.warning("Please fill in all required fields.")
    else:
        try:
            # Load workbook and update cells
            wb = openpyxl.load_workbook(io.BytesIO(st.session_state["file_bytes"]))
            sheet = wb.active
            sheet['G6'] = invoice_no
            sheet['G7'] = invoice_date
            sheet['C17'] = rental_text
            
            # Use temporary directory to handle both Excel and PDF safely
            with tempfile.TemporaryDirectory() as tmpdir:
                xlsx_path = os.path.join(tmpdir, "invoice.xlsx")
                wb.save(xlsx_path)

                # Read Excel bytes safely from the temp file
                with open(xlsx_path, "rb") as f:
                    st.session_state["xlsx_data"] = f.read()

                # Convert Excel to PDF using LibreOffice
                result = subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, xlsx_path],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=60
                )

                if result.returncode != 0:
                    st.error(f"LibreOffice Conversion Error:\n{result.stderr}")
                    st.session_state["pdf_data"] = None
                else:
                    pdf_path = os.path.join(tmpdir, "invoice.pdf")
                    if os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as f:
                            st.session_state["pdf_data"] = f.read()
                    else:
                        st.session_state["pdf_data"] = None

            st.session_state["filename_prefix"] = output_filename.strip() or "Invoice"
            st.success("Files generated successfully!")

        except Exception as e:
            st.error(f"Error processing file: {e}")

# Display download buttons side by side if data exists in session state
if "xlsx_data" in st.session_state and st.session_state["xlsx_data"] is not None:
    st.markdown("---")
    st.subheader("📥 Download Your Files")
    
    prefix = st.session_state.get("filename_prefix", "Invoice")
    xlsx_name = prefix if prefix.endswith('.xlsx') else prefix + '.xlsx'
    pdf_name = prefix if prefix.endswith('.pdf') else prefix + '.pdf'
    
    st.info(f"📁 **Current File Name:** `{prefix}` (.xlsx / .pdf)")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=st.session_state["xlsx_data"],
            file_name=xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    with col2:
        if "pdf_data" in st.session_state and st.session_state["pdf_data"] is not None:
            st.download_button(
                label="📥 Download PDF (.pdf)",
                data=st.session_state["pdf_data"],
                file_name=pdf_name,
                mime="application/pdf",
                use_container_width=True
            )

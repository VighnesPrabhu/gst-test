import streamlit as st
import openpyxl
import subprocess
import tempfile
import os
import io

st.set_page_config(page_title="GST Invoice Modifier & Converter", page_icon="📄", layout="centered")

st.title("📄 GST Invoice Modifier & Converter")
st.write("Upload your Excel template, update details, and download as Excel or PDF!")

uploaded_file = st.file_uploader("Upload Excel Invoice (.xlsx)", type=["xlsx"])

# Initialize session state for persistent downloads
if "xlsx_data" not in st.session_state:
    st.session_state.xlsx_data = None
if "pdf_data" not in st.session_state:
    st.session_state.pdf_data = None
if "output_filename" not in st.session_state:
    st.session_state.output_filename = "Invoice"

inv_no_default = ""
inv_date_default = ""
rental_default = ""

if uploaded_file is not None:
    try:
        wb_temp = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_temp = wb_temp.active
        inv_no_default = str(sheet_temp['G6'].value or "")
        inv_date_default = str(sheet_temp['G7'].value or "")
        rental_default = str(sheet_temp['C17'].value or "")
    except Exception:
        pass

invoice_no = st.text_input("Invoice Number (Cell G6)", value=inv_no_default)
invoice_date = st.text_input("Invoice Date (Cell G7)", value=inv_date_default)
rental_text = st.text_area("Rental Description & Month (Cell C17)", value=rental_default)
output_filename = st.text_input("Output File Name", value="Invoice")

if st.button("🔄 Generate Files", type="primary"):
    if uploaded_file is not None and invoice_no and invoice_date and rental_text:
        try:
            # 1. Handle Excel generation
            uploaded_file.seek(0)
            wb = openpyxl.load_workbook(uploaded_file)
            sheet = wb.active
            sheet['G6'] = invoice_no
            sheet['G7'] = invoice_date
            sheet['C17'] = rental_text
            
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            st.session_state.xlsx_data = excel_buffer.getvalue()

            # 2. Handle PDF generation using temp directory & LibreOffice
            with tempfile.TemporaryDirectory() as tmpdir:
                xlsx_path = os.path.join(tmpdir, "input.xlsx")
                wb.save(xlsx_path)

                result = subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, xlsx_path],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=60
                )

                if result.returncode != 0:
                    st.error(f"LibreOffice Error Output:\n{result.stderr}")
                    st.session_state.pdf_data = None
                else:
                    pdf_path = os.path.join(tmpdir, "input.pdf")
                    if os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as pdf_file:
                            st.session_state.pdf_data = pdf_file.read()
                    else:
                        st.session_state.pdf_data = None

            filename = output_filename.strip()
            if not filename:
                filename = "Invoice"
            st.session_state.output_filename = filename

            if st.session_state.xlsx_data is not None and st.session_state.pdf_data is not None:
                st.success("Files generated successfully!")
            elif st.session_state.xlsx_data is not None:
                st.warning("Excel generated, but PDF conversion encountered an issue.")
        except Exception as e:
            st.error(f"Error processing file: {e}")
            st.session_state.xlsx_data = None
            st.session_state.pdf_data = None
    else:
        st.warning("Please upload a file and fill in all required fields.")

# Display download buttons side by side if data exists in memory
if st.session_state.xlsx_data is not None or st.session_state.pdf_data is not None:
    st.markdown("---")
    st.subheader("📥 Download Your Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.session_state.xlsx_data is not None:
            xlsx_name = st.session_state.output_filename
            if not xlsx_name.endswith('.xlsx'):
                xlsx_name += '.xlsx'
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=st.session_state.xlsx_data,
                file_name=xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    with col2:
        if st.session_state.pdf_data is not None:
            pdf_name = st.session_state.output_filename
            if not pdf_name.endswith('.pdf'):
                pdf_name += '.pdf'
            st.download_button(
                label="📥 Download PDF (.pdf)",
                data=st.session_state.pdf_data,
                file_name=pdf_name,
                mime="application/pdf",
                use_container_width=True
            )

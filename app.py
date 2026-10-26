import streamlit as st
import openpyxl
import subprocess
import tempfile
import os

st.set_page_config(page_title="GST Invoice PDF Converter", page_icon="📄", layout="centered")

st.title("📄 GST Invoice to PDF Generator")
st.write("Upload your Excel template, update details, and download as a clean PDF!")

uploaded_file = st.file_uploader("Upload Excel Invoice (.xlsx)", type=["xlsx"])

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

if st.button("🔄 Generate & Download PDF", type="primary"):
    if uploaded_file is not None and invoice_no and invoice_date and rental_text:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                xlsx_path = os.path.join(tmpdir, "input.xlsx")
                
                # Save uploaded file to temp path
                uploaded_file.seek(0)
                with open(xlsx_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Modify workbook with openpyxl
                wb = openpyxl.load_workbook(xlsx_path)
                sheet = wb.active
                sheet['G6'] = invoice_no
                sheet['G7'] = invoice_date
                sheet['C17'] = rental_text
                wb.save(xlsx_path)

                # Convert Excel to PDF using LibreOffice headless
                subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, xlsx_path],
                    check=True,
                    timeout=60
                )

                pdf_path = os.path.join(tmpdir, "input.pdf")

                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()

                    filename = output_filename.strip()
                    if not filename.endswith('.pdf'):
                        filename += '.pdf'

                    st.success("PDF generated successfully!")
                    
                    st.download_button(
                        label="📥 Download Modified PDF",
                        data=pdf_bytes,
                        file_name=filename,
                        mime="application/pdf"
                    )
                else:
                    st.error("Failed to generate PDF conversion.")
        except Exception as e:
            st.error(f"Error processing file: {e}")
    else:
        st.warning("Please upload a file and fill in all required fields.")

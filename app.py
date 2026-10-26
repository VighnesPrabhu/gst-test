import streamlit as st
import openpyxl
import io

st.set_page_config(page_title="GST Invoice Modifier", page_icon="📄", layout="centered")

st.title("📄 GST Invoice Modifier")
st.write("Upload your Excel template, update details, and download instantly!")

uploaded_file = st.file_uploader("Upload Excel Invoice (.xlsx)", type=["xlsx"])

# Extract existing values from cells G6, G7, and C17 if a file is uploaded
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

# Input fields pre-filled with values from the Excel file
invoice_no = st.text_input("Invoice Number (Cell G6)", value=inv_no_default)
invoice_date = st.text_input("Invoice Date (Cell G7)", value=inv_date_default)
rental_text = st.text_area("Rental Description & Month (Cell C17)", value=rental_default)
output_filename = st.text_input("Output File Name", value="Invoice")

if st.button("🔄 Update & Download Excel", type="primary"):
    if uploaded_file is not None and invoice_no and invoice_date and rental_text:
        try:
            # Reload workbook for modification
            uploaded_file.seek(0)
            wb = openpyxl.load_workbook(uploaded_file)
            sheet = wb.active

            # Update specific target cells
            sheet['G6'] = invoice_no
            sheet['G7'] = invoice_date
            sheet['C17'] = rental_text

            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)

            filename = output_filename.strip()
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'

            st.success("Excel file updated successfully!")
            
            st.download_button(
                label="📥 Download Modified Excel",
                data=output_buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error processing file: {e}")
    else:
        st.warning("Please upload a file and fill in all required fields.")

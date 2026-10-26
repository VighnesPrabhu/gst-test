import streamlit as st
import openpyxl
import io
import os
import win32com.client
import pythoncom

# Page Configuration
st.set_page_config(
    page_title="GST Invoice Date and Name Converter",
    page_icon="📄",
    layout="centered"
)

# Styling of UI
st.markdown("""
    <style>
    .main {
        background-color: #f4f6f9;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .header-card {
        background: linear-gradient(135deg, #0066cc 0%, #004080 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header-card h1 {
        margin: 0;
        font-size: 2.2rem;
        font-weight: 700;
    }
    .header-card p {
        margin-top: 0.5rem;
        font-size: 1.1rem;
        opacity: 0.9;
    }
    .custom-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 1.5rem;
        border: 1px solid #e1e4e8;
    }
    .stButton>button {
        width: 100%;
        background-color: #0066cc;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.75rem;
        border: none;
        box-shadow: 0 2px 4px rgba(0,102,204,0.2);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #0052a3;
        box-shadow: 0 4px 8px rgba(0,102,204,0.3);
    }
    .success-badge {
        background-color: #e6f4ea;
        color: #137333;
        padding: 0.75rem 1rem;
        border-radius: 6px;
        font-weight: 600;
        border-left: 4px solid #34a853;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="header-card">
        <h1>📄 GST Invoice Date & Name Converter</h1>
    </div>
""", unsafe_allow_html=True)

# File Upload
st.markdown('<div class="custom-card">', unsafe_allow_html=True)
st.markdown("### 📁 Upload Invoice")
uploaded_file = st.file_uploader("Upload your Excel Invoice (.xlsx)", type=["xlsx"])
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is not None:
    # Save the uploaded file temporarily to disk so openpyxl can read images safely
    source_path = os.path.abspath("source_template.xlsx")
    with open(source_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Load workbook from the local disk path
    wb = openpyxl.load_workbook(source_path)
    sheet = wb.active

    # Read current values from specific cells
    current_no = sheet["G6"].value
    current_date = sheet["G7"].value
    current_rental_text = sheet["C17"].value

    st.markdown('<div class="success-badge">✅ File uploaded successfully! Modify the fields below:</div>', unsafe_allow_html=True)

    # Input form
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("### ✏️ Update the Invoice Details & File Name")

    with st.form("invoice_form"):
        col1, col2 = st.columns(2)

        with col1:
            new_no = st.text_input("Invoice Number (Cell G6)", value=str(current_no) if current_no else "")
        with col2:
            new_date = st.text_input("Invoice Date (Cell G7)", value=str(current_date) if current_date else "")
            
        new_rental_text = st.text_area("Rental Description and Month (Cell C17)", value=str(current_rental_text) if current_rental_text else "")
        
        # Custom File Name Input Field
        custom_filename = st.text_input("Download File Name (Without Extension)", value=str(current_no) if current_no else "Invoice")

        submitted = st.form_submit_button("🔄 Update & Generate Files")

    st.markdown('</div>', unsafe_allow_html=True)

    if submitted:
        # Update cells
        sheet["G6"] = new_no
        sheet["G7"] = new_date
        sheet["C17"] = new_rental_text

        # Save updated workbook to disk temporarily for processing
        temp_excel_path = os.path.abspath("temp_invoice.xlsx")
        temp_pdf_path = os.path.abspath("temp_invoice.pdf")

        wb.save(temp_excel_path)

        # Save Excel to memory buffer for download
        with open(temp_excel_path, "rb") as f:
            st.session_state["excel_data"] = f.read()
            
        # Clean up custom filename to remove invalid characters
        safe_name = "".join(c for c in custom_filename if c.isalnum() or c in (' ', '_', '-')).strip()
        st.session_state["file_name"] = safe_name if safe_name else "Invoice"

        # Convert to PDF
        pdf_success = False
        try:
            pythoncom.CoInitialize()
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False

            workbook = excel_app.Workbooks.Open(temp_excel_path)
            workbook.ExportAsFixedFormat(0, temp_pdf_path)
            workbook.Close(False)
            excel_app.Quit()

            if os.path.exists(temp_pdf_path):
                with open(temp_pdf_path, "rb") as pdf_file:
                    st.session_state["pdf_data"] = pdf_file.read()
                pdf_success = True
        except Exception as e:
            st.error(f"PDF Conversion Error: {e}")

        st.session_state["pdf_success"] = pdf_success

        # Cleanup temporary files safely
        for path in [source_path, temp_excel_path, temp_pdf_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

    # Download section
    if "excel_data" in st.session_state:
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
        st.markdown("### 📥 Download Updated Files")

        dl_col1, dl_col2 = st.columns(2)
        final_filename = st.session_state.get("file_name", "Invoice")

        with dl_col1:
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=st.session_state["excel_data"],
                file_name=f"{final_filename}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with dl_col2:
            if st.session_state.get("pdf_success", False):
                st.download_button(
                    label="📥 Download PDF (.pdf)",
                    data=st.session_state["pdf_data"],
                    file_name=f"{final_filename}.pdf",
                    mime="application/pdf"
                )
            else:
                st.warning("PDF generation failed or skipped.")

        st.markdown('</div>', unsafe_allow_html=True)
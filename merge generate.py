import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

def extract_specific_data(pdf_file):
    """
    Version 1: Extracts data based on fixed PDF rules.
    - Captures Kode Satker, Nama Satker (Title Case), Revisi Ke, and Digital Stamps.
    """
    extracted_records = []
    
    # Regex patterns
    satker_code_pattern = re.compile(r'\b(\d{6})\b')
    revisi_pattern = re.compile(r'REVISI\s+KE\s*:\s*(\d+)', re.IGNORECASE)
    ds_sebelum_pattern = re.compile(r'Digital\s+Stamp\s+Sebelum\s*:\s*(\d+)', re.IGNORECASE)
    ds_sesudah_pattern = re.compile(r'Digital\s+Stamp\s+Sesudah\s*:\s*(\d+)', re.IGNORECASE)
    
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
        
        # 1. Extract Kode Satker (6 digits)
        satker_match = satker_code_pattern.search(full_text)
        kode_satker = satker_match.group(1) if satker_match else ""
        
        # 2. Extract Nama Satker (Text after 6 digits, formatted to Title Case)
        nama_satker = ""
        if satker_match:
            for line in full_text.split('\n'):
                if kode_satker in line:
                    name_part = line.replace(kode_satker, "").strip()
                    nama_satker = name_part.title()
                    break
        
        # 3. Extract Revisi Ke
        revisi_match = revisi_pattern.search(full_text)
        revisi_ke = revisi_match.group(1) if revisi_match else ""

        # 4. Digital Stamp Logic
        ds_sebelum_match = ds_sebelum_pattern.search(full_text)
        ds_sesudah_match = ds_sesudah_pattern.search(full_text)
        
        ds_sebelum = ds_sebelum_match.group(1) if ds_sebelum_match else ""
        ds_sesudah = ds_sesudah_match.group(1) if ds_sesudah_match else ""
        
        ds_status = ""
        if ds_sebelum and ds_sesudah:
            if ds_sebelum == ds_sesudah:
                ds_status = "tidak berubah yaitu DS:"
            else:
                ds_status = "berubah yaitu DS:"

        # 5. Build record based on the template column order
        # We use dictionary keys that match your Excel headers
        record = {
            "-": "",
            "Kode Satker": kode_satker,
            "Revisi Ke": revisi_ke,
            "Nama Satker ": nama_satker, # Note the space after 'Satker' to match your template
            "DS berubah atau tidak": ds_status,
            "DS RAW": ds_sesudah,
            "-1": "", # Placeholder for column 7
            "-2": "", # Placeholder for column 8
            "-3": "", # Placeholder for column 9
            "-4": "", # Placeholder for column 10
            "Pejabat": "Kuasa Pengguna Anggaran",
            "-5": "", # Placeholder for column 12
            "-6": ""  # Placeholder for column 13
        }
        extracted_records.append(record)
    
    return extracted_records

def main():
    st.set_page_config(page_title="PDF to Excel Converter", layout="wide")
    st.title("DIPA Mail Merge Data Generator (Version 1)")
    st.write("Upload your PDFs to generate a new Excel file for your mail merge.")

    uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

    if uploaded_files:
        all_data = []
        for pdf_file in uploaded_files:
            try:
                data = extract_specific_data(pdf_file)
                all_data.extend(data)
            except Exception as e:
                st.error(f"Error processing {pdf_file.name}: {e}")

        if all_data:
            df = pd.DataFrame(all_data)
            
            # Rename the placeholder columns back to '-' for the final Excel file
            # This handles duplicate column names which DataFrames don't like internally
            final_columns = [
                "-", "Kode Satker", "Revisi Ke", "Nama Satker ", 
                "DS berubah atau tidak", "DS RAW", "-", "-", "-", "-", 
                "Pejabat", "-", "-"
            ]
            
            st.subheader("Data Preview")
            st.dataframe(df)

            # Convert to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # We write the dataframe but override the header to use '-' for all empty columns
                df.to_excel(writer, index=False, header=final_columns, sheet_name='Sheet1')
            
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Download Generated Excel",
                data=excel_data,
                file_name="Mail_Merge_Data_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
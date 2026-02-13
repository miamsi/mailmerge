import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import os

def format_kode_ds(ds_raw):
    """
    Implements formula: =LEFT(F2;4)&"-"&MID(F2;5;4)&"-"&MID(F2;9;4)&"-"&RIGHT(F2;4)
    """
    if not ds_raw:
        return ""
    ds = str(ds_raw).strip()
    if len(ds) >= 16:
        return f"{ds[0:4]}-{ds[4:8]}-{ds[8:12]}-{ds[12:16]}"
    return ds # Return as is if length doesn't match expected stamp

def extract_specific_data(pdf_file, reference_df, current_no):
    """
    Version 1: Extracts data from PDF and processes columns based on exact VLOOKUP/Formula logic.
    """
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
        
        # 1. Extract Kode Satker
        satker_match = satker_code_pattern.search(full_text)
        kode_satker = satker_match.group(1) if satker_match else ""
        
        # 2. Extract Revisi Ke
        revisi_match = revisi_pattern.search(full_text)
        revisi_ke = revisi_match.group(1) if revisi_match else ""

        # 3. Digital Stamp Logic
        ds_sebelum_match = ds_sebelum_pattern.search(full_text)
        ds_sesudah_match = ds_sesudah_pattern.search(full_text)
        ds_sebelum = ds_sebelum_match.group(1) if ds_sebelum_match else ""
        ds_sesudah = ds_sesudah_match.group(1) if ds_sesudah_match else ""
        
        ds_status = ""
        if ds_sebelum and ds_sesudah:
            ds_status = "tidak berubah yaitu DS:" if ds_sebelum == ds_sesudah else "berubah yaitu DS:"

        # 4. LOOKUP LOGIC (Mirroring Excel Formulas)
        nama_satker = "" # Output Col 4
        kppn = ""        # Output Col 8
        pejabat = ""     # Output Col 11
        tembusan_kl = "" # Output Col 12
        ref_val = ""     # Output Col 13
        
        if reference_df is not None and kode_satker:
            # Match Kode Satker (string based)
            reference_df['KODE SATKER'] = reference_df['KODE SATKER'].astype(str).str.split('.').str[0]
            ref_match = reference_df[reference_df['KODE SATKER'] == str(kode_satker)]
            
            if not ref_match.empty:
                # Column 4 (Nama Satker) uses 'Satker Fix' (Title Case)
                nama_satker = ref_match.iloc[0].iloc[5]
                # Column 8 (KPPN)
                kppn = ref_match.iloc[0].iloc[4]
                # Column 11 (Pejabat)
                pejabat = ref_match.iloc[0].iloc[6]
                # Column 12 (Tembusan KL) -> =VLOOKUP(B2;refsatker2!A:G;7;0)
                tembusan_kl = ref_match.iloc[0].iloc[6]
                # Column 13 (ref) -> =VLOOKUP(B2;refsatker2!A:B;2;0)
                ref_val = ref_match.iloc[0].iloc[1]

        # 5. ND PENGANTAR LOGIC (Updated)
        # Formula: =IF(E2="tidak berubah yaitu DS:";"tidak";"")
        ds_nd_pengantar = "tidak" if ds_status == "tidak berubah yaitu DS:" else ""

        # 6. Build Final Row Structure
        return {
            "No": current_no,
            "Kode Satker": kode_satker,
            "Revisi Ke": revisi_ke,
            "Nama Satker": nama_satker,
            "DS berubah atau tidak": ds_status,
            "DS RAW": ds_sesudah,
            "Kode DS": format_kode_ds(ds_sesudah),
            "KPPN": kppn,
            "No Surat": "",
            "Tgl Surat": "",
            "Pejabat": pejabat,
            "Tembusan KL": tembusan_kl,
            "ref": ref_val,
            "DS ND pengantar": ds_nd_pengantar
        }

def main():
    st.set_page_config(page_title="PDF to Excel V1", layout="wide")
    st.title("DIPA Mail Merge Generator (Version 1)")

    # Target path for refsatker.xlsx
    ref_path = r"C:\Users\michael.sidabutar\Documents\revisi\refsatker.xlsx"
    reference_df = None
    
    if os.path.exists(ref_path):
        try:
            # Target the specific sheet 'refsatker2'
            reference_df = pd.read_excel(ref_path, sheet_name='refsatker2')
            st.sidebar.success("✅ Master reference 'refsatker2' loaded successfully.")
        except Exception as e:
            st.sidebar.error(f"❌ Error reading refsatker.xlsx: {e}")
    else:
        st.sidebar.error(f"❌ File not found: {ref_path}")

    uploaded_files = st.file_uploader("Upload PDF Matrix Files", type="pdf", accept_multiple_files=True)

    if uploaded_files:
        all_data = []
        for i, pdf_file in enumerate(uploaded_files, start=1):
            try:
                record = extract_specific_data(pdf_file, reference_df, i)
                all_data.append(record)
            except Exception as e:
                st.error(f"Failed to process {pdf_file.name}: {e}")

        if all_data:
            df = pd.DataFrame(all_data)
            
            st.subheader("Final Data Preview")
            st.dataframe(df)

            # Export to Excel as a new download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Sheet name matches your template
                df.to_excel(writer, index=False, sheet_name='revisi')
            
            st.download_button(
                label="📥 Download Mail Merge Results",
                data=output.getvalue(),
                file_name="Mail_Merge_Results_V1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()

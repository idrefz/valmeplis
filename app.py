import streamlit as st
import pandas as pd
from pykml import parser
from io import BytesIO
import zipfile

st.title('Ekstrak Deskripsi dari KML ke Excel')
st.write('Unggah file KML dan ekstrak hanya konten deskripsi ke Excel')

def extract_descriptions(kml_file):
    """Mengekstrak deskripsi dari file KML"""
    try:
        doc = parser.parse(kml_file).getroot()
        descriptions = []
        
        # Cari semua Placemark dalam dokumen KML
        for pm in doc.Document.findall('.//{http://www.opengis.net/kml/2.2}Placemark'):
            desc = pm.description.text if hasattr(pm, 'description') and pm.description is not None else ''
            if desc:  # Hanya tambahkan jika deskripsi tidak kosong
                descriptions.append({'Deskripsi': desc})
        
        return pd.DataFrame(descriptions)
    except Exception as e:
        st.error(f"Error parsing KML: {str(e)}")
        return None

def save_to_excel(df):
    """Menyimpan DataFrame ke file Excel dalam memori"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Deskripsi KML')
    processed_data = output.getvalue()
    return processed_data

# Upload file
uploaded_file = st.file_uploader("Pilih file KML/KMZ", type=['kml', 'kmz'])

if uploaded_file is not None:
    try:
        # Untuk file KMZ (zip)
        if uploaded_file.name.endswith('.kmz'):
            with zipfile.ZipFile(uploaded_file) as kmz:
                kml_files = [f for f in kmz.namelist() if f.endswith('.kml')]
                if not kml_files:
                    st.error("Tidak menemukan file KML dalam KMZ")
                    st.stop()
                
                with kmz.open(kml_files[0]) as kml_file:
                    df = extract_descriptions(kml_file)
        else:
            df = extract_descriptions(uploaded_file)
        
        if df is not None:
            if not df.empty:
                st.success(f"Berhasil mengekstrak {len(df)} deskripsi!")
                st.dataframe(df)
                
                # Download Excel
                excel_data = save_to_excel(df)
                st.download_button(
                    label="Unduh Deskripsi sebagai Excel",
                    data=excel_data,
                    file_name='deskripsi_kml.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.warning("File KML tidak mengandung deskripsi atau deskripsi kosong.")
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")

st.markdown("""
**Petunjuk Penggunaan:**
1. Unggah file KML/KMZ
2. Aplikasi akan mengekstrak hanya konten deskripsi
3. Klik tombol "Unduh Deskripsi sebagai Excel" untuk mendapatkan file Excel
""")

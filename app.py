import streamlit as st
import pandas as pd
from pykml import parser
from io import BytesIO
import zipfile
import os

st.title('Konversi KML ke Excel')
st.write('Unggah file KML dan dapatkan data dalam format Excel')

def kml_to_dataframe(kml_file):
    """Mengubah file KML menjadi DataFrame"""
    try:
        doc = parser.parse(kml_file).getroot()
        data = []
        
        # Cari semua Placemark dalam dokumen KML
        for pm in doc.Document.findall('.//{http://www.opengis.net/kml/2.2}Placemark'):
            row = {
                'Nama': pm.name.text if hasattr(pm, 'name') and pm.name is not None else '',
                'Deskripsi': pm.description.text if hasattr(pm, 'description') and pm.description is not None else '',
                'Koordinat': pm.Point.coordinates.text if hasattr(pm, 'Point') and pm.Point is not None else '',
                'Tipe': 'Point' if hasattr(pm, 'Point') else 'Polygon' if hasattr(pm, 'Polygon') else 'Unknown'
            }
            data.append(row)
        
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Error parsing KML: {str(e)}")
        return None

def save_to_excel(df):
    """Menyimpan DataFrame ke file Excel dalam memori"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data KML')
    processed_data = output.getvalue()
    return processed_data

# Upload file
uploaded_file = st.file_uploader("Pilih file KML", type=['kml', 'kmz'])

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
                    df = kml_to_dataframe(kml_file)
        else:
            df = kml_to_dataframe(uploaded_file)
        
        if df is not None:
            st.success("File KML berhasil diproses!")
            st.dataframe(df.head())
            
            # Download Excel
            excel_data = save_to_excel(df)
            st.download_button(
                label="Unduh sebagai Excel",
                data=excel_data,
                file_name='kml_data.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")

st.markdown("""
**Petunjuk Penggunaan:**
1. Unggah file KML/KMZ
2. Aplikasi akan memproses dan menampilkan preview data
3. Klik tombol "Unduh sebagai Excel" untuk mendapatkan file Excel
""")

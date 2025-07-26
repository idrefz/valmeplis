import streamlit as st
import pandas as pd
from pykml import parser
from simplekml import Kml
from io import BytesIO
import zipfile
import os

st.set_page_config(page_title="KML ↔ Excel Converter", layout="wide")

# Fungsi-fungsi utilitas
def save_to_excel(df):
    """Menyimpan DataFrame ke file Excel dalam memori"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data KML')
    return output.getvalue()

def save_kml_to_bytes(kml):
    """Menyimpan KML ke bytes untuk diunduh"""
    kml_bytes = BytesIO()
    kml_bytes.write(kml.kml().encode('utf-8'))
    kml_bytes.seek(0)
    return kml_bytes

# Tab KML ke Excel
def kml_to_excel_tab():
    st.header("Konversi KML ke Excel")
    
    uploaded_file = st.file_uploader("Pilih file KML/KMZ", type=['kml', 'kmz'], key="kml_uploader")
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.kmz'):
                with zipfile.ZipFile(uploaded_file) as kmz:
                    kml_files = [f for f in kmz.namelist() if f.endswith('.kml')]
                    if not kml_files:
                        st.error("Tidak menemukan file KML dalam KMZ")
                        return
                    
                    with kmz.open(kml_files[0]) as kml_file:
                        doc = parser.parse(kml_file).getroot()
            else:
                doc = parser.parse(uploaded_file).getroot()
            
            data = []
            for pm in doc.Document.findall('.//{http://www.opengis.net/kml/2.2}Placemark'):
                row = {
                    'Name': pm.name.text if hasattr(pm, 'name') and pm.name is not None else '',
                    'Description': pm.description.text if hasattr(pm, 'description') and pm.description is not None else '',
                    'Coordinates': pm.Point.coordinates.text if hasattr(pm, 'Point') and pm.Point is not None else '',
                    'Type': 'Point' if hasattr(pm, 'Point') else 'Polygon' if hasattr(pm, 'Polygon') else 'Unknown'
                }
                data.append(row)
            
            df = pd.DataFrame(data)
            st.success("File KML berhasil diproses!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.dataframe(df.head())
            with col2:
                st.map(pd.DataFrame({
                    'lat': df['Coordinates'].str.extract(r'([\d.-]+),')[0].astype(float),
                    'lon': df['Coordinates'].str.extract(r',([\d.-]+)')[0].astype(float)
                }).dropna())
            
            excel_data = save_to_excel(df)
            st.download_button(
                label="Unduh sebagai Excel",
                data=excel_data,
                file_name='kml_data.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")

# Tab Excel ke KML
def excel_to_kml_tab():
    st.header("Konversi Excel ke KML")
    
    uploaded_file = st.file_uploader("Pilih file Excel/CSV", type=['xlsx', 'xls', 'csv'], key="excel_uploader")
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success("File Excel berhasil dibaca!")
            st.dataframe(df.head())
            
            # Pilih kolom
            cols = df.columns.tolist()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                name_col = st.selectbox("Pilih kolom untuk Nama", cols, index=cols.index('name') if 'name' in cols else 0)
            with col2:
                lon_col = st.selectbox("Pilih kolom untuk Longitude", cols, index=cols.index('longitude') if 'longitude' in cols else 
                                      cols.index('lon') if 'lon' in cols else 0)
            with col3:
                lat_col = st.selectbox("Pilih kolom untuk Latitude", cols, index=cols.index('latitude') if 'latitude' in cols else 
                                      cols.index('lat') if 'lat' in cols else 1 if len(cols) > 1 else 0)
            
            desc_cols = st.multiselect("Pilih kolom untuk Deskripsi (bisa pilih beberapa)", cols, 
                                     default=[c for c in ['description', 'desc', 'info'] if c in cols])
            
            # Buat KML
            kml = Kml()
            
            for _, row in df.iterrows():
                try:
                    lon = float(row[lon_col])
                    lat = float(row[lat_col])
                    
                    description = "<![CDATA[<table>"
                    for col in desc_cols:
                        description += f"<tr><td><b>{col}</b></td><td>{row[col]}</td></tr>"
                    description += "</table>]]>"
                    
                    kml.newpoint(
                        name=str(row[name_col]),
                        coords=[(lon, lat)],
                        description=description
                    )
                except (ValueError, KeyError):
                    continue
            
            # Download KML
            kml_bytes = save_kml_to_bytes(kml)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Unduh sebagai KML",
                    data=kml_bytes,
                    file_name='converted_data.kml',
                    mime='application/vnd.google-earth.kml+xml'
                )
            with col2:
                st.map(df.rename(columns={lat_col: 'lat', lon_col: 'lon'})[[lat_col, lon_col]].dropna())
            
            st.subheader("Preview KML")
            st.code(kml.kml(), language='xml')
        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")

# Antarmuka utama
st.title("Konverter KML ↔ Excel")

tab1, tab2 = st.tabs(["KML ke Excel", "Excel ke KML"])
with tab1:
    kml_to_excel_tab()
with tab2:
    excel_to_kml_tab()

st.markdown("""
### Petunjuk Penggunaan

**KML ke Excel:**
1. Unggah file KML/KMZ
2. Aplikasi akan mengekstrak data dan menampilkan peta
3. Unduh hasil dalam format Excel

**Excel ke KML:**
1. Unggah file Excel/CSV
2. Pilih kolom untuk: Nama, Longitude, Latitude
3. Pilih kolom tambahan untuk Deskripsi (opsional)
4. Unduh hasil dalam format KML
5. File KML bisa dibuka di Google Earth atau aplikasi GIS lainnya
""")

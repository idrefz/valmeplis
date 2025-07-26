import streamlit as st
import pandas as pd
from pykml import parser
from simplekml import Kml
from io import BytesIO
import zipfile
import os

st.set_page_config(page_title="KML ↔ Excel Converter Pro", layout="wide")

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

def detect_coordinate_columns(df):
    """Mendeteksi kolom yang mungkin berisi koordinat"""
    possible_lat = ['latitude', 'lat', 'y', 'ycoord', 'northing', 'lat1', 'lat2']
    possible_lon = ['longitude', 'lon', 'x', 'xcoord', 'easting', 'lon1', 'lon2']
    
    # Cari kolom yang cocok dengan pola
    lat_cols = [col for col in df.columns if any(p in col.lower() for p in possible_lat)]
    lon_cols = [col for col in df.columns if any(p in col.lower() for p in possible_lon)]
    
    # Default ke kolom pertama dan kedua jika tidak ditemukan
    default_lat = lat_cols[0] if lat_cols else df.columns[0]
    default_lon = lon_cols[0] if lon_cols else df.columns[1] if len(df.columns) > 1 else df.columns[0]
    
    return default_lat, default_lon

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
                try:
                    coord_df = pd.DataFrame({
                        'lat': df['Coordinates'].str.extract(r'([\d.-]+),')[0].astype(float),
                        'lon': df['Coordinates'].str.extract(r',([\d.-]+)')[0].astype(float)
                    }).dropna()
                    st.map(coord_df)
                except:
                    st.warning("Tidak dapat menampilkan peta")
            
            excel_data = save_to_excel(df)
            st.download_button(
                label="Unduh sebagai Excel",
                data=excel_data,
                file_name='kml_data.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")

# Tab Excel ke KML (Basic)
def excel_to_kml_basic_tab():
    st.header("Konversi Excel ke KML (Basic)")
    
    uploaded_file = st.file_uploader("Pilih file Excel/CSV", type=['xlsx', 'xls', 'csv'], key="excel_basic_uploader")
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"File berhasil dibaca! {len(df)} baris ditemukan.")
            st.dataframe(df.head())
            
            # Deteksi kolom otomatis
            default_lat, default_lon = detect_coordinate_columns(df)
            cols = df.columns.tolist()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                name_col = st.selectbox("Pilih kolom untuk Nama", cols, 
                                      index=cols.index('name') if 'name' in cols else 
                                      cols.index('ODP_NAME') if 'ODP_NAME' in cols else 0)
            with col2:
                lon_col = st.selectbox("Pilih kolom untuk Longitude", cols, 
                                      index=cols.index(default_lon) if default_lon in cols else 0)
            with col3:
                lat_col = st.selectbox("Pilih kolom untuk Latitude", cols, 
                                      index=cols.index(default_lat) if default_lat in cols else 1 if len(cols) > 1 else 0)
            
            # Validasi data
            df[lat_col] = pd.to_numeric(df[lat_col], errors='coerce')
            df[lon_col] = pd.to_numeric(df[lon_col], errors='coerce')
            df = df.dropna(subset=[lat_col, lon_col])
            
            desc_cols = st.multiselect("Pilih kolom untuk Deskripsi", cols)
            
            # Buat KML
            kml = Kml()
            points_added = 0
            
            for _, row in df.iterrows():
                try:
                    lon = float(row[lon_col])
                    lat = float(row[lat_col])
                    
                    description = "<![CDATA[<table border='1'>"
                    for col in desc_cols:
                        description += f"<tr><td><b>{col}</b></td><td>{row[col]}</td></tr>"
                    description += "</table>]]>"
                    
                    kml.newpoint(
                        name=str(row[name_col]),
                        coords=[(lon, lat)],
                        description=description
                    )
                    points_added += 1
                except (ValueError, TypeError):
                    continue
            
            st.success(f"Berhasil membuat {points_added} titik dalam KML")
            
            # Tampilkan peta
            try:
                st.map(df.rename(columns={lat_col: 'lat', lon_col: 'lon'})[['lat', 'lon']])
            except:
                st.warning("Tidak dapat menampilkan peta")
            
            # Download KML
            kml_bytes = save_kml_to_bytes(kml)
            st.download_button(
                label="Unduh sebagai KML",
                data=kml_bytes,
                file_name='converted_data.kml',
                mime='application/vnd.google-earth.kml+xml'
            )
            
            st.subheader("Preview KML")
            st.code(kml.kml(), language='xml')
            
        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")

# Tab Excel ke KML (Group by STO)
def excel_to_kml_sto_tab():
    st.header("Konversi Excel ke KML (Group by STO)")
    
    uploaded_file = st.file_uploader("Pilih file Excel/CSV", type=['xlsx', 'xls', 'csv'], key="excel_sto_uploader")
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"File berhasil dibaca! {len(df)} baris ditemukan.")
            
            # Pilih kolom
            cols = df.columns.tolist()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                name_col = st.selectbox("Kolom Nama Lokasi", cols, 
                                      index=cols.index('ODP_NAME') if 'ODP_NAME' in cols else 0)
            with col2:
                sto_col = st.selectbox("Kolom STO", cols, 
                                     index=cols.index('STO') if 'STO' in cols else 
                                     cols.index('STO_DESC') if 'STO_DESC' in cols else 0)
            with col3:
                lat_col = st.selectbox("Kolom Latitude", cols, 
                                     index=cols.index('LATITUDE') if 'LATITUDE' in cols else 0)
            with col4:
                lon_col = st.selectbox("Kolom Longitude", cols, 
                                     index=cols.index('LONGITUDE') if 'LONGITUDE' in cols else 1 if len(cols) > 1 else 0)
            
            desc_cols = st.multiselect("Kolom Deskripsi Tambahan", cols, 
                                     default=['ODP_INDEX', 'CLUSNAME', 'USED', 'RSV', 'KATEGORI'])
            
            # Validasi data
            df[lat_col] = pd.to_numeric(df[lat_col], errors='coerce')
            df[lon_col] = pd.to_numeric(df[lon_col], errors='coerce')
            df = df.dropna(subset=[lat_col, lon_col])
            
            if len(df) == 0:
                st.error("Tidak ada data dengan koordinat valid!")
                return
            
            # Buat KML dengan pengelompokan STO
            kml = Kml()
            sto_groups = df.groupby(sto_col)
            
            for sto_name, group in sto_groups:
                fol = kml.newfolder(name=f"STO {sto_name}")
                
                for _, row in group.iterrows():
                    try:
                        lon = float(row[lon_col])
                        lat = float(row[lat_col])
                        
                        description = "<![CDATA[<table border='1'>"
                        for col in desc_cols:
                            description += f"<tr><td><b>{col}</b></td><td>{row[col]}</td></tr>"
                        description += "</table>]]>"
                        
                        fol.newpoint(
                            name=str(row[name_col]),
                            coords=[(lon, lat)],
                            description=description
                        )
                    except (ValueError, TypeError):
                        continue
            
            st.success(f"Berhasil membuat {len(df)} titik dalam {len(sto_groups)} STO")
            
            # Tampilkan peta
            try:
                st.map(df.rename(columns={lat_col: 'lat', lon_col: 'lon'})[['lat', 'lon']])
            except:
                st.warning("Tidak dapat menampilkan peta")
            
            # Download KML
            kml_bytes = save_kml_to_bytes(kml)
            st.download_button(
                label="Unduh KML dengan Pengelompokan STO",
                data=kml_bytes,
                file_name='odp_by_sto.kml',
                mime='application/vnd.google-earth.kml+xml'
            )
            
            # Tampilkan statistik
            st.subheader("Statistik per STO")
            sto_stats = df[sto_col].value_counts().reset_index()
            sto_stats.columns = ['STO', 'Jumlah ODP']
            st.dataframe(sto_stats)
            
        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")

# Antarmuka utama
st.title("KML ↔ Excel Converter Pro")

tab1, tab2, tab3 = st.tabs(["KML ke Excel", "Excel ke KML (Basic)", "Excel ke KML (Group by STO)"])
with tab1:
    kml_to_excel_tab()
with tab2:
    excel_to_kml_basic_tab()
with tab3:
    excel_to_kml_sto_tab()

st.markdown("""
### Petunjuk Penggunaan

**KML ke Excel:**
1. Unggah file KML/KMZ
2. Aplikasi akan mengekstrak data dan menampilkan peta
3. Unduh hasil dalam format Excel

**Excel ke KML (Basic):**
1. Unggah file Excel/CSV
2. Pilih kolom untuk: Nama, Longitude, Latitude
3. Pilih kolom tambahan untuk Deskripsi (opsional)
4. Unduh hasil dalam format KML

**Excel ke KML (Group by STO):**
1. Unggah file Excel/CSV dengan data ODP
2. Pilih kolom untuk: Nama, STO, Latitude, Longitude
3. Pilih kolom tambahan untuk Deskripsi
4. Unduh KML yang sudah dikelompokkan per STO
""")

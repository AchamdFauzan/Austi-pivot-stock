import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auto Pivot Stock - SPR JABO Format", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (Format Persis SPR JABO)")
st.write("Aplikasi ini otomatis menyusun file mentahan menjadi Pivot Stock persis seperti format SPR JABO dan memisahkan Device & Accessories ke sheet berbeda.")

# 1. Upload File Mentahan
uploaded_file = st.file_uploader("Upload File MENTAHAN (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Membaca data mentahan
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Cleansing Data (Mengisi kolom Tsh yang kosong)
    if 'Tsh' in df.columns:
        df['Tsh'] = df['Tsh'].fillna('(kosong)')
        
    # Filter Kategori
    st.subheader("Filter Kategori Barang")
    kategori = st.radio("Pilih kategori untuk di-preview & di-download:", ["Semua Data", "Hanya Device", "Hanya Accessories"])
    
    # --- PENAMBAHAN KATA KUNCI TV, TABLET, DLL ---
    device_kw = ['oppo', 'samsung', 'vivo', 'iphone', 'macbook', 'acer', 'asus', 'lenovo', 'zyrex', 'infinix', 'realme', 'xiaomi', 'poco', 'ipad', 'tv', 'tablet', 'tab']
    acc_kw = ['watch', 'buds', 'enco', 'fan', 'case', 'charger', 'cable', 'strap', 'tws', 'adapter', 'powerbank', 'screen', 'tempered']
    
    # Memisahkan DataFrame sejak awal
    df_device = df[df['Article Description'].str.contains('|'.join(device_kw), case=False, na=False)]
    df_acc = df[df['Article Description'].str.contains('|'.join(acc_kw), case=False, na=False)]
        
    st.write("⏳ Memproses Pivot Data sesuai urutan SPR JABO...")
    
    # Define urutan TSH & Area sesuai file SPR JABO
    tsh_order = ['Mensi Alexander', 'Lukman Wibowo', 'Rendy Nur Setiawan', 'Febrian Tri Wibowo', 'Irman Permana', '(kosong)']
    area_order = [
        'Jakarta Pusat', 'Jakarta Barat', 'Bekasi', 'Bogor', 'Cilegon', 'Depok', 
        'Gudang', 'Jakarta Selatan', 'Jakarta Timur', 'Jakarta Utara', 
        'Kabupaten Tangerang', 'Pandeglang', 'Rangkasbitung', 'Serang', 
        'Tangerang', 'Tangerang Kabupaten', 'Tangerang Selatan'
    ]
    
    # --- FUNGSI UNTUK MEMBUAT PIVOT (Agar kode tidak berulang) ---
    def buat_pivot(data_df):
        if data_df.empty:
            return pd.DataFrame() # Kembalikan dataframe kosong jika tidak ada data
            
        data_df = data_df.copy()
        data_df['Tsh'] = pd.Categorical(data_df['Tsh'], categories=tsh_order, ordered=True)
        data_df['Area'] = pd.Categorical(data_df['Area'], categories=area_order, ordered=True)
        data_df = data_df.sort_values(['Tsh', 'Area', 'Name 1', 'Article Description'])
        
        pivot = pd.pivot_table(
            data_df, 
            index='Article Description', 
            columns=['Tsh', 'Area', 'Name 1'], 
            values='Quantity', 
            aggfunc='sum',
            fill_value=None
        )
        
        # Tambahkan Total Keseluruhan (Baris & Kolom)
        if not pivot.empty:
            pivot['Total Keseluruhan'] = pivot.sum(axis=1)
            total_row = pivot.sum(axis=0)
            total_row.name = 'Total Keseluruhan'
            pivot = pd.concat([pivot, pd.DataFrame(total_row).T])
            
        return pivot

    # Buat Pivot untuk masing-masing kategori
    pivot_device = buat_pivot(df_device)
    pivot_acc = buat_pivot(df_acc)

    # Style Header Tabel (Hijau)
    styles = [
        {
            'selector': 'th',
            'props': [
                ('background-color', '#4CAF50'), 
                ('color', 'white'),              
                ('font-weight', 'bold'),         
                ('text-align', 'center'),        
                ('border', '1px solid white')    
            ]
        }
    ]
    
    # --- MENAMPILKAN PREVIEW TABEL DI WEB ---
    if kategori in ["Semua Data", "Hanya Device"] and not pivot_device.empty:
        st.write("### 📱 Preview: Data Device (Handphone, Tablet, Laptop, TV)")
        st.dataframe(pivot_device.head(10).style.set_table_styles(styles), use_container_width=True)
        
    if kategori in ["Semua Data", "Hanya Accessories"] and not pivot_acc.empty:
        st.write("### 🎧 Preview: Data Accessories")
        st.dataframe(pivot_acc.head(10).style.set_table_styles(styles), use_container_width=True)
        
    # --- EXPORT KE EXCEL DENGAN SHEET BERBEDA ---
    if not pivot_device.empty or not pivot_acc.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            
            # Jika user memilih "Semua Data" atau "Hanya Device", buat sheet Device
            if kategori in ["Semua Data", "Hanya Device"] and not pivot_device.empty:
                pivot_device.to_excel(writer, sheet_name='Device')
                
            # Jika user memilih "Semua Data" atau "Hanya Accessories", buat sheet Accessories
            if kategori in ["Semua Data", "Hanya Accessories"] and not pivot_acc.empty:
                pivot_acc.to_excel(writer, sheet_name='Accessories')
                
        st.success("✅ File Excel siap diunduh! (Device & Accessories sudah dipisah per-sheet)")
        st.download_button(
            label="⬇️ Unduh Hasil Pivot (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"SPR_JABO_Stock_{kategori.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Data kosong setelah difilter. Coba periksa kembali file mentahannya.")

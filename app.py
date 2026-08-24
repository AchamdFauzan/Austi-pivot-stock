import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auto Pivot Stock", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (Device & Accessories)")
st.write("Aplikasi ini otomatis mengubah file mentahan menjadi pivot stock berformat SPR JABO.")

# 1. Upload File
uploaded_file = st.file_uploader("Upload File MENTAHAN (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Membaca data mentahan
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # 2. Cleansing Data (Mengisi kolom Tsh yang kosong)
    if 'Tsh' in df.columns:
        df['Tsh'] = df['Tsh'].fillna('(kosong)')
        
    # 3. Filter Kategori
    st.subheader("Pilih Kategori")
    kategori = st.radio("Filter data berdasarkan:", ["Semua Data", "Hanya Device", "Hanya Accessories"])
    
    # Keyword untuk mendeteksi Device vs Aksesoris (Bisa Anda tambah sendiri)
    device_kw = ['oppo', 'samsung', 'vivo', 'iphone', 'macbook', 'acer', 'asus', 'lenovo', 'zyrex', 'infinix', 'realme', 'xiaomi']
    acc_kw = ['watch', 'buds', 'enco', 'fan', 'case', 'charger', 'cable', 'strap', 'tws']
    
    if kategori == "Hanya Device":
        df = df[df['Article Description'].str.contains('|'.join(device_kw), case=False, na=False)]
    elif kategori == "Hanya Accessories":
        df = df[df['Article Description'].str.contains('|'.join(acc_kw), case=False, na=False)]
        
    # 4. Proses Pivot Data
    if not df.empty:
        st.write("Sedang memproses pivot data...")
        pivot_df = pd.pivot_table(
            df, 
            index='Article Description', 
            columns=['Tsh', 'Area', 'Name 1'], 
            values='Quantity', 
            aggfunc='sum'
        )
        
        st.success(f"Berhasil! Data siap diunduh.")
        st.dataframe(pivot_df.head(10)) # Menampilkan preview singkat
        
        # 5. Tombol Download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            pivot_df.to_excel(writer, sheet_name='Stock_Pivoted')
            
        st.download_button(
            label="⬇️ Download File Excel Hasil Pivot",
            data=buffer.getvalue(),
            file_name=f"Stock_Hasil_{kategori.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Data kosong setelah difilter. Periksa kembali keyword atau isi file Anda.")

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auto Pivot Stock - SPR JABO Format", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (Format Persis SPR JABO)")
st.write("Aplikasi ini otomatis menyusun file mentahan menjadi Pivot Stock persis seperti format SPR JABO (Lengkap dengan urutan TSH, Area, Toko, dan Subtotal).")

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
    kategori = st.radio("Pilih kategori:", ["Semua Data", "Hanya Device", "Hanya Accessories"])
    
    device_kw = ['oppo', 'samsung', 'vivo', 'iphone', 'macbook', 'acer', 'asus', 'lenovo', 'zyrex', 'infinix', 'realme', 'xiaomi', 'poco', 'ipad']
    acc_kw = ['watch', 'buds', 'enco', 'fan', 'case', 'charger', 'cable', 'strap', 'tws', 'adapter', 'powerbank', 'screen', 'tempered']
    
    if kategori == "Hanya Device":
        df = df[df['Article Description'].str.contains('|'.join(device_kw), case=False, na=False)]
    elif kategori == "Hanya Accessories":
        df = df[df['Article Description'].str.contains('|'.join(acc_kw), case=False, na=False)]
        
    if not df.empty:
        st.write("⏳ Memproses Pivot Data sesuai urutan SPR JABO...")
        
        # Define urutan TSH & Area sesuai file SPR JABO
        tsh_order = ['Mensi Alexander', 'Lukman Wibowo', 'Rendy Nur Setiawan', 'Febrian Tri Wibowo', 'Irman Permana', '(kosong)']
        area_order = [
            'Jakarta Pusat', 'Jakarta Barat', 'Bekasi', 'Bogor', 'Cilegon', 'Depok', 
            'Gudang', 'Jakarta Selatan', 'Jakarta Timur', 'Jakarta Utara', 
            'Kabupaten Tangerang', 'Pandeglang', 'Rangkasbitung', 'Serang', 
            'Tangerang', 'Tangerang Kabupaten', 'Tangerang Selatan'
        ]
        
        # Sort dataframe berdasarkan TSH, Area, dan Name 1
        df['Tsh'] = pd.Categorical(df['Tsh'], categories=tsh_order, ordered=True)
        df['Area'] = pd.Categorical(df['Area'], categories=area_order, ordered=True)
        df = df.sort_values(['Tsh', 'Area', 'Name 1', 'Article Description'])
        
        # Buat pivot dasar
        pivot_df = pd.pivot_table(
            df, 
            index='Article Description', 
            columns=['Tsh', 'Area', 'Name 1'], 
            values='Quantity', 
            aggfunc='sum',
            fill_value=None
        )
        
        # Tambahkan Total Keseluruhan (Baris & Kolom)
        pivot_df['Total Keseluruhan'] = pivot_df.sum(axis=1)
        
        # Tambahkan baris Total Keseluruhan di paling bawah
        total_row = pivot_df.sum(axis=0)
        total_row.name = 'Total Keseluruhan'
        pivot_df = pd.concat([pivot_df, pd.DataFrame(total_row).T])
        
        st.success("✅ Berhasil diproses!")
        st.dataframe(pivot_df.head(15))
        
        # Export ke Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            pivot_df.to_excel(writer, sheet_name='Stock_SPR_JABO')
            
        st.download_button(
            label="⬇️ Unduh Hasil Pivot Format SPR JABO (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"SPR_JABO_Stock_{kategori.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Data kosong setelah difilter.")

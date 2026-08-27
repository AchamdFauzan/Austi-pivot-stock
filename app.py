import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auto Pivot Stock - SPR JABO Format", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (Format Persis SPR JABO)")
st.write("Aplikasi ini otomatis menyusun file mentahan menjadi Pivot Stock persis seperti format SPR JABO.")

# 1. Upload File Mentahan
uploaded_file = st.file_uploader("Upload File MENTAHAN (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Membaca data mentahan
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Cleansing Data
    if 'Tsh' in df.columns:
        df['Tsh'] = df['Tsh'].fillna('(kosong)')
        
    # Filter Kategori
    st.subheader("Filter Kategori Barang")
    kategori = st.radio("Pilih kategori:", ["Semua Data", "Hanya Device", "Hanya Accessories"])
    
    device_kw = ['oppo', 'samsung', 'vivo', 'iphone', 'macbook', 'acer', 'asus', 'lenovo', 'zyrex', 'infinix', 'realme', 'xiaomi', 'poco', 'ipad', 'tv', 'tablet', 'tab']
    acc_kw = ['watch', 'buds', 'enco', 'fan', 'case', 'charger', 'cable', 'strap', 'tws', 'adapter', 'powerbank', 'screen', 'tempered']
    
    if kategori == "Hanya Device":
        df = df[df['Article Description'].str.contains('|'.join(device_kw), case=False, na=False)]
    elif kategori == "Hanya Accessories":
        df = df[df['Article Description'].str.contains('|'.join(acc_kw), case=False, na=False)]
        
    if not df.empty:
        st.write("⏳ Memproses Pivot Data sesuai urutan SPR JABO...")
        
        # Define urutan TSH & Area
        tsh_order = ['Mensi Alexander', 'Lukman Wibowo', 'Rendy Nur Setiawan', 'Febrian Tri Wibowo', 'Irman Permana', '(kosong)']
        area_order = [
            'Jakarta Pusat', 'Jakarta Barat', 'Bekasi', 'Bogor', 'Cilegon', 'Depok', 
            'Gudang', 'Jakarta Selatan', 'Jakarta Timur', 'Jakarta Utara', 
            'Kabupaten Tangerang', 'Pandeglang', 'Rangkasbitung', 'Serang', 
            'Tangerang', 'Tangerang Kabupaten', 'Tangerang Selatan'
        ]
        
        # Sort dataframe
        df['Tsh'] = pd.Categorical(df['Tsh'], categories=tsh_order, ordered=True)
        df['Area'] = pd.Categorical(df['Area'], categories=area_order, ordered=True)
        df = df.sort_values(['Tsh', 'Area', 'Name 1', 'Article Description'])
        
        # Buat pivot
        pivot_df = pd.pivot_table(
            df, 
            index='Article Description', 
            columns=['Tsh', 'Area', 'Name 1'], 
            values='Quantity', 
            aggfunc='sum',
            fill_value=None
        )
        
        # Tambahkan Total
        pivot_df['Total Keseluruhan'] = pivot_df.sum(axis=1)
        total_row = pivot_df.sum(axis=0)
        total_row.name = 'Total Keseluruhan'
        pivot_df = pd.concat([pivot_df, pd.DataFrame(total_row).T])
        
        st.success("✅ Berhasil diproses!")
        
        # ---------------- SOLUSI 1: TAMPILAN DI WEB ----------------
        # Membuat format warna header untuk Web
        styles = [
            {'selector': 'th', 'props': [('background-color', '#4CAF50'), ('color', 'white')]}
        ]
        styled_pivot_web = pivot_df.head(15).style.set_table_styles(styles)
        
        # Kita ubah dari st.dataframe() menjadi st.table() agar CSS warna terbaca oleh Streamlit
        st.table(styled_pivot_web)
        
        # ---------------- SOLUSI 2: EXPORT KE EXCEL BERWARNA ----------------
        buffer = io.BytesIO()
        # Kita menggunakan openpyxl agar warna bawaan pandas styler bisa diekspor ke Excel
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # Terapkan warna ke seluruh pivot data untuk Excel
            styled_pivot_excel = pivot_df.style.set_table_styles(styles)
            styled_pivot_excel.to_excel(writer, sheet_name='Stock_SPR_JABO')
            
        st.download_button(
            label="⬇️ Unduh Hasil Pivot Format SPR JABO (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"SPR_JABO_Stock_{kategori.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Data kosong setelah difilter.")

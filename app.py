import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auto Pivot Stock SPR JABO", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (Format SPR JABO)")
st.write("Sistem ini menyusun file mentahan menjadi format Pivot Excel dengan struktur kolom bertingkat (MultiIndex) persis seperti SPR.")

# Upload File Mentahan
uploaded_file = st.file_uploader("Upload File Mentahan (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Cleansing awal
    df.columns = df.columns.str.strip()
    
    required_cols = ['Tsh', 'Area', 'Name 1', 'Article Description', 'Quantity']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.error(f"❌ Gagal memproses! Kolom berikut tidak ditemukan: {', '.join(missing_cols)}")
    else:
        # Cleansing Isi Data
        for col in ['Tsh', 'Area', 'Name 1', 'Article Description']:
            df[col] = df[col].fillna('(kosong)').astype(str).str.strip()
            
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        
        # Filter Kategori
        kategori = st.radio("Pilih kategori:", ["Semua Data", "Hanya Device", "Hanya Accessories"], horizontal=True)
        
        device_kw = ['oppo', 'samsung', 'vivo', 'iphone', 'macbook', 'acer', 'asus', 'lenovo', 'zyrex', 'infinix', 'realme', 'xiaomi', 'poco', 'ipad', 'tv', 'tablet', 'tab']
        acc_kw = ['watch', 'buds', 'enco', 'fan', 'case', 'charger', 'cable', 'strap', 'tws', 'adapter', 'powerbank', 'screen', 'tempered']
        
        if kategori == "Hanya Device":
            df = df[df['Article Description'].str.contains('|'.join(device_kw), case=False, na=False)]
        elif kategori == "Hanya Accessories":
            df = df[df['Article Description'].str.contains('|'.join(acc_kw), case=False, na=False)]
            
        if not df.empty:
            st.success("✅ Memproses Pivot Data...")
            
            # Pembuatan Pivot Table (MultiIndex murni)
            pivot_df = pd.pivot_table(
                df, 
                index='Article Description', 
                columns=['Tsh', 'Area', 'Name 1'], 
                values='Quantity', 
                aggfunc='sum',
                fill_value=0
            )
            
            # Ubah nama index utama agar sesuai dengan gambar (Label Baris)
            pivot_df.index.name = 'Label Baris'
            
            # Tambahkan kolom Total Keseluruhan (Struktur 3 tingkat agar header tidak rusak)
            pivot_df[('Total Keseluruhan', '', '')] = pivot_df.sum(axis=1)
            
            # Tambahkan baris Total Keseluruhan di bawah
            pivot_df.loc['Total Keseluruhan'] = pivot_df.sum(axis=0)
            
            # Tampilkan wujud asli MultiIndex di Streamlit
            st.dataframe(pivot_df, use_container_width=True)
            
            # Export ke File Excel (.xlsx) 
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Pandas otomatis melakukan 'merge cell' pada header bertingkat saat diekspor
                pivot_df.to_excel(writer, sheet_name='Stock_SPR_JABO')
                
            buffer.seek(0)
            
            st.download_button(
                label="⬇️ Unduh Hasil Pivot SPR JABO (.xlsx)",
                data=buffer.getvalue(),
                file_name=f"SPR_JABO_{kategori.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

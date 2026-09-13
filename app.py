import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Auto Pivot Stock SPR JABO", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (Format SPR JABO)")
st.write("Sistem ini menyusun file mentahan menjadi Pivot Stock dengan format kolom bertingkat (Tsh, Area, Toko) persis SPR JABO.")

# 1. Fitur Upload File Mentahan
uploaded_file = st.file_uploader("Upload File Mentahan (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Cleansing awal: Hapus spasi tersembunyi di judul kolom
    df.columns = df.columns.str.strip()
    
    # Validasi keberadaan kolom wajib
    required_cols = ['Tsh', 'Area', 'Name 1', 'Article Description', 'Quantity']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.error(f"❌ Gagal memproses! Kolom berikut tidak ditemukan di Excel mentahan Anda: {', '.join(missing_cols)}")
        st.info("💡 Pastikan file mentahan Anda sudah memiliki kolom 'Tsh' dan 'Area'.")
    else:
        # Cleansing Isi Data
        for col in ['Tsh', 'Area', 'Name 1', 'Article Description']:
            df[col] = df[col].fillna('(kosong)').astype(str).str.strip()
            
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        
        # 2. Filter Kategori Barang
        st.subheader("Filter Kategori Barang")
        kategori = st.radio("Pilih kategori:", ["Semua Data", "Hanya Device", "Hanya Accessories"], horizontal=True)
        
        device_kw = ['oppo', 'samsung', 'vivo', 'iphone', 'macbook', 'acer', 'asus', 'lenovo', 'zyrex', 'infinix', 'realme', 'xiaomi', 'poco', 'ipad', 'tv', 'tablet', 'tab']
        acc_kw = ['watch', 'buds', 'enco', 'fan', 'case', 'charger', 'cable', 'strap', 'tws', 'adapter', 'powerbank', 'screen', 'tempered']
        
        if kategori == "Hanya Device":
            df = df[df['Article Description'].str.contains('|'.join(device_kw), case=False, na=False)]
        elif kategori == "Hanya Accessories":
            df = df[df['Article Description'].str.contains('|'.join(acc_kw), case=False, na=False)]
            
        if df.empty:
            st.warning("⚠️ Data kosong setelah difilter kategori.")
        else:
            st.success("⏳ Memproses Pivot Data...")
            
            # 3. Proses Pembuatan Pivot Table MultiIndex (Tsh, Area, Name 1)
            pivot_df = pd.pivot_table(
                df, 
                index='Article Description', 
                columns=['Tsh', 'Area', 'Name 1'], 
                values='Quantity', 
                aggfunc='sum',
                fill_value=0
            )
            
            # Tambahkan kolom Total Keseluruhan
            pivot_df[('Total Keseluruhan', '', '')] = pivot_df.sum(axis=1)
            
            # Tambahkan baris Total Keseluruhan paling bawah
            pivot_df.loc['Total Keseluruhan'] = pivot_df.sum(axis=0)
            
            st.info("💡 Pratinjau di web diringkas oleh sistem. Silakan klik tombol unduh untuk melihat struktur 3 tingkat (Tsh & Area) di Excel.")
            st.dataframe(pivot_df, use_container_width=True)
            
            # 4. Export ke File Excel (.xlsx) dengan Format Presisi SPR JABO
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pivot_df.to_excel(writer, sheet_name='Stock_SPR_JABO')
                
            buffer.seek(0)
            
            # --- PROSES MERAPIKAN FORMAT EXCEL DENGAN OPENPYXL ---
            wb = openpyxl.load_workbook(buffer)
            ws = wb['Stock_SPR_JABO']
            
            # Memastikan sel A1, A2, dan A3 terisi persis seperti format SPR JABO asli
            ws['A1'] = 'Tsh'
            ws['A2'] = 'Area'
            ws['A3'] = 'Label Baris'
            
            # Hapus baris ke-4 (baris jeda 'Article Description' bawaan Pandas)
            ws.delete_rows(4)
            
            # Tambahkan tombol Filter (Dropdown) di baris ke-3
            max_col = get_column_letter(ws.max_column)
            max_row = ws.max_row
            ws.auto_filter.ref = f"A3:{max_col}{max_row}"
            
            final_buffer = io.BytesIO()
            wb.save(final_buffer)
            final_buffer.seek(0)
            
            # Tombol Download Excel
            st.download_button(
                label="⬇️ Unduh Hasil Pivot SPR JABO (.xlsx)",
                data=final_buffer.getvalue(),
                file_name=f"SPR_JABO_{kategori.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

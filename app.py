import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Auto Pivot Stock SPR JABO", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (Format SPR JABO)")
st.write("Sistem ini otomatis menyusun file mentahan menjadi Pivot Stock dengan format kolom bertingkat persis SPR JABO.")

# 1. Fitur Upload
uploaded_file = st.file_uploader("Upload File Mentahan (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Cleansing awal: Hapus spasi tidak kasat mata di judul kolom
    df.columns = df.columns.str.strip()
    
    # Validasi keberadaan kolom wajib
    required_cols = ['Tsh', 'Area', 'Name 1', 'Article Description', 'Quantity']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.error(f"❌ Gagal memproses! Kolom berikut tidak ditemukan di Excel mentahan Anda: {', '.join(missing_cols)}")
        st.info("💡 Pastikan file mentahan Anda sudah memiliki kolom 'Tsh' dan 'Area' (Lakukan VLOOKUP terlebih dahulu jika diperlukan).")
    else:
        # Cleansing Isi Data: Menyeragamkan teks dan menghapus spasi
        for col in ['Tsh', 'Area', 'Name 1', 'Article Description']:
            df[col] = df[col].fillna('(kosong)').astype(str).str.strip()
            
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        
        # 2. Filter Kategori
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
            
            # 3. Proses Pembuatan Pivot Table (Format MultiIndex)
            pivot_df = pd.pivot_table(
                df, 
                index='Article Description', 
                columns=['Tsh', 'Area', 'Name 1'], 
                values='Quantity', 
                aggfunc='sum',
                fill_value=0
            )
            
           # Tambahkan kolom Total Keseluruhan (Struktur 3 Tingkat)
            pivot_df[('Total Keseluruhan', '', '')] = pivot_df.sum(axis=1)
            
            # Tambahkan baris Total Keseluruhan paling bawah
            pivot_df.loc['Total Keseluruhan'] = pivot_df.sum(axis=0)
            
            # --- TRIK UNTUK TAMPILAN WEB STREAMLIT AGAR RAPI ---
            # Kita buat salinan khusus untuk ditampilkan di layar saja
            preview_df = pivot_df.copy()
            
            # Ubah nama kolom yang tadinya bertingkat menjadi satu baris teks yang panjang dan jelas
            new_columns = []
            for col in preview_df.columns:
                if col[0] == 'Total Keseluruhan':
                    new_columns.append('Total Keseluruhan')
                else:
                    # Menampilkan urutan: Tsh -> Area -> Nama Toko
                    new_columns.append(f"Tsh: {col[0]} | Area: {col[1]} | Toko: {col[2]}")
            
            preview_df.columns = new_columns
            
            # Ubah nama index agar sesuai dengan permintaan Anda
            preview_df.index.name = "Label Baris"
            
            # Tampilkan data pratinjau di Streamlit (tampilan akan menjadi lebih panjang ke samping)
            st.dataframe(preview_df, use_container_width=True) 
            # 4. Export ke File Excel (.xlsx) dengan Format Custom (Persis Gambar)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pivot_df.to_excel(writer, sheet_name='Stock_SPR_JABO')
                
            buffer.seek(0)
            
            # --- PROSES MERAPIKAN FORMAT EXCEL MENGGUNAKAN OPENPYXL ---
            wb = openpyxl.load_workbook(buffer)
            ws = wb['Stock_SPR_JABO']
            
            # Ubah tulisan 'Name 1' di sel A3 menjadi 'Label Baris'
            ws['A3'] = 'Label Baris'
            
            # Hapus baris ke-4 (baris jeda/kosong bawaan Pandas)
            ws.delete_rows(4)
            
            # Tambahkan tombol Filter (Dropdown) di baris ke-3
            max_col = get_column_letter(ws.max_column)
            max_row = ws.max_row
            ws.auto_filter.ref = f"A3:{max_col}{max_row}"
            
            # Simpan hasil modifikasi ke buffer final
            final_buffer = io.BytesIO()
            wb.save(final_buffer)
            final_buffer.seek(0)
            
            st.download_button(
                label="⬇️ Unduh Hasil Pivot SPR JABO (.xlsx)",
                data=final_buffer.getvalue(),
                file_name=f"SPR_JABO_{kategori.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

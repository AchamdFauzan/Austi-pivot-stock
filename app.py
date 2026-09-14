import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils import get_column_letter
import os

st.set_page_config(page_title="Auto Pivot Stock SPR JABO", layout="wide")

st.title("📦 Aplikasi Auto-Pivot Stock (2 Sheet: Device & Acc)")
st.write("Sistem otomatis membagi data menjadi Sheet Device & Accessories. Master Kategori sudah terintegrasi otomatis di dalam sistem.")

# HANYA ADA 1 KOTAK UPLOAD SEKARANG
uploaded_raw = st.file_uploader("Upload File MENTAHAN Saja (.xlsx)", type=["xlsx"])

if uploaded_raw:
    # Nama file master yang sudah Anda tanam di GitHub
    master_file = "master_unique.xlsx"
    
    # Cek apakah file master benar-benar ada di GitHub Anda
    if not os.path.exists(master_file):
        st.error(f"❌ File '{master_file}' tidak ditemukan di sistem! Pastikan Anda sudah mengunggah file tersebut ke GitHub Anda.")
    else:
        # Membaca data
        df_raw = pd.read_excel(uploaded_raw, sheet_name=0)
        df_master = pd.read_excel(master_file, sheet_name=0)
        
        # Cleansing awal
        df_raw.columns = df_raw.columns.str.strip()
        df_master.columns = df_master.columns.str.strip()
        
        required_raw = ['Tsh', 'Area', 'Name 1', 'Article Description', 'Quantity', 'Material Grp Desc. 2']
        missing_raw = [col for col in required_raw if col not in df_raw.columns]
        
        if missing_raw:
            st.error(f"❌ File Mentahan kurang kolom: {', '.join(missing_raw)}.")
        else:
            st.success("✅ File Mentahan berhasil dibaca. Sedang memproses 2 Sheet (Device & Acc)...")
            
            # Cleansing Isi Data Raw
            for col in ['Tsh', 'Area', 'Name 1', 'Article Description', 'Material Grp Desc. 2']:
                df_raw[col] = df_raw[col].fillna('(kosong)').astype(str).str.strip()
            df_raw['Quantity'] = pd.to_numeric(df_raw['Quantity'], errors='coerce').fillna(0)
            
            # Cleansing Isi Data Master
            df_master['Material'] = df_master['Material'].astype(str).str.strip()
            df_master['Category'] = df_master['Category'].astype(str).str.strip()
            
            # Gabungkan Data (VLOOKUP otomatis)
            df_merged = pd.merge(df_raw, df_master, left_on='Material Grp Desc. 2', right_on='Material', how='left')
            
            # Pisahkan menjadi Device dan Acc
            df_device = df_merged[df_merged['Category'].str.contains('Device|Tablet', case=False, na=False)]
            df_acc = df_merged[df_merged['Category'].str.contains('Acc', case=False, na=False)]
            
            # Fungsi untuk membuat Pivot dan merapikan Sheet
            def generate_pivot_sheet(df_subset, sheet_name, excel_writer):
                if df_subset.empty:
                    return None
                    
                pivot_df = pd.pivot_table(
                    df_subset, 
                    index='Article Description', 
                    columns=['Tsh', 'Area', 'Name 1'], 
                    values='Quantity', 
                    aggfunc='sum'
                    # Perubahan: fill_value=0 dihapus agar sel yang kosong dibiarkan NaN (kosong) sementara
                )
                
                # Tambahkan Total Keseluruhan, hiraukan nilai kosong saat menjumlahkan
                pivot_df[('Total Keseluruhan', '', '')] = pivot_df.sum(axis=1, skipna=True)
                pivot_df.loc['Total Keseluruhan'] = pivot_df.sum(axis=0, skipna=True)
                
                # --- PERUBAHAN UTAMA: Ubah NaN/0 menjadi string kosong (blank cell) di Excel ---
                # Menggunakan trik fillna("") setelah semua penjumlahan selesai
                pivot_df = pivot_df.fillna("") 
                # Khusus untuk sel yang benar-benar bernilai angka 0 (jika ada dari mentahan), ubah juga ke kosong
                pivot_df = pivot_df.replace(0, "")
                
                pivot_df.to_excel(excel_writer, sheet_name=sheet_name)
                
                ws = excel_writer.sheets[sheet_name]
                ws['A1'] = 'Tsh'
                ws['A2'] = 'Area'
                ws['A3'] = 'Label Baris'
                ws.delete_rows(4)
                
                max_col = get_column_letter(ws.max_column)
                max_row = ws.max_row
                ws.auto_filter.ref = f"A3:{max_col}{max_row}"
                
                return pivot_df

            # Tulis ke dalam File Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pivot_dev = generate_pivot_sheet(df_device, 'Stock_Device', writer)
                pivot_acc = generate_pivot_sheet(df_acc, 'Stock_Accessories', writer)
                
            buffer.seek(0)
            
            # Tampilan Pratinjau (Opsional)
            if pivot_dev is not None:
                st.info("💡 Pratinjau Data (Sebagian data Stock_Device):")
                st.dataframe(pivot_dev.head(10), use_container_width=True)

            # Tombol Download (1 File Excel berisi 2 Sheet)
            st.download_button(
                label="⬇️ Unduh Hasil Pivot SPR JABO (2 Sheet).xlsx",
                data=buffer.getvalue(),
                file_name="SPR_JABO_Device_&_Acc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

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
    
    # Cleansing Nama Kolom (menghapus spasi tersembunyi di judul kolom Excel)
    df.columns = df.columns.str.strip()
    
    # Cek apakah kolom yang dibutuhkan ada
    required_cols = ['Tsh', 'Area', 'Name 1', 'Article Description', 'Quantity']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.error(f"Gagal memproses! Kolom berikut tidak ditemukan di Excel mentahan Anda: {', '.join(missing_cols)}")
    else:
        # Cleansing Data Isi (Hapus spasi depan/belakang, isi yang kosong dengan teks '(kosong)')
        df['Tsh'] = df['Tsh'].fillna('(kosong)').astype(str).str.strip()
        df['Area'] = df['Area'].fillna('(kosong)').astype(str).str.strip()
        df['Name 1'] = df['Name 1'].fillna('(kosong)').astype(str).str.strip()
        df['Article Description'] = df['Article Description'].fillna('(kosong)').astype(str).str.strip()
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
            
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
            
            # Define urutan bawaan
            tsh_order = ['Mensi Alexander', 'Lukman Wibowo', 'Rendy Nur Setiawan', 'Febrian Tri Wibowo', 'Irman Permana', '(kosong)']
            area_order = [
                'Jakarta Pusat', 'Jakarta Barat', 'Bekasi', 'Bogor', 'Cilegon', 'Depok', 
                'Gudang', 'Jakarta Selatan', 'Jakarta Timur', 'Jakarta Utara', 
                'Kabupaten Tangerang', 'Pandeglang', 'Rangkasbitung', 'Serang', 
                'Tangerang', 'Tangerang Kabupaten', 'Tangerang Selatan'
            ]
            
            # AMANKAN DATA: Masukkan TSH / Area yang ada di data tapi tidak ada di list bawaan
            for t in df['Tsh'].unique():
                if t not in tsh_order:
                    tsh_order.append(t)
            
            for a in df['Area'].unique():
                if a not in area_order:
                    area_order.append(a)
            
            # Sort dataframe menggunakan Categorical
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
                fill_value=0
            )
            
            # Validasi apakah pivot benar-benar memiliki data
            if pivot_df.empty or len(pivot_df.columns) == 0:
                st.warning("⚠️ Pivot berhasil dibuat tapi isinya kosong. Pastikan nilai 'Quantity' pada mentahan Anda tidak 0 semua.")
            else:
                # PERBAIKAN: Penambahan Kolom Total menggunakan format tuple 3 tingkat
                # Agar susunan Header 'Tsh', 'Area', 'Name 1' tidak rusak
                pivot_df[('Total Keseluruhan', '', '')] = pivot_df.sum(axis=1)
                
                # Penambahan Baris Total paling bawah
                pivot_df.loc['Total Keseluruhan'] = pivot_df.sum(axis=0)
                
                st.success("✅ Berhasil diproses!")
                
                # TAMPILAN DI WEB (Scrollable)
                st.dataframe(pivot_df, use_container_width=True)
                
                # EXPORT KE EXCEL
                buffer = io.BytesIO()
                
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    pivot_df.to_excel(writer, sheet_name='Stock_SPR_JABO')
                    
                buffer.seek(0)
                    
                st.download_button(
                    label="⬇️ Unduh Hasil Pivot Format SPR JABO (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"SPR_JABO_Stock_{kategori.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Data kosong setelah difilter kategori.")

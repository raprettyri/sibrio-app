import streamlit as st
import pandas as pd
import numpy as np
import docx
from docxtpl import DocxTemplate
import io
import os
import base64

# --- KONFIGURASI TAMPILAN WEB ---
st.set_page_config(page_title="SIBRIO - Generator BRS", page_icon="logobps.png", layout="wide")

# --- INJEKSI CUSTOM CSS ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;800;900&display=swap');

/* Mengatur font global dan background hitam */
.stApp {
    background-color: #0E1117;
    font-family: 'Poppins', sans-serif;
    color: white;
}

/* Menyembunyikan padding atas bawaan Streamlit */
.block-container {
    padding-top: 2rem !important;
}

/* Container untuk Logo & Judul Tengah */
.header-container {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 20px;
    margin-top: 10px;
}

/* Main Title SIBRIO */
.main-title {
    color: #FF5722;
    font-size: 65px;
    font-weight: 900;
    letter-spacing: 3px;
    margin: 0;
    line-height: 1.1;
}
.main-subtitle {
    text-align: center;
    color: #E0E0E0;
    font-size: 16px;
    font-weight: 400;
    margin-bottom: 40px;
    line-height: 1.4;
    margin-top: 10px;
}

/* Label Custom Rata Tengah */
.custom-label {
    text-align: center;
    font-weight: 600;
    font-size: 14px;
    color: #FAFAFA;
    margin-bottom: 5px;
}

/* Styling Tombol Generate */
div.stButton > button:first-child {
    background-color: #FF5722 !important;
    color: white !important;
    border-radius: 8px !important;
    border: none !important;
    height: 50px !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    width: 100%;
    margin-top: 10px;
}
div.stButton > button:first-child:hover {
    background-color: #E64A19 !important;
}

/* Styling Input Box */
div[data-baseweb="select"] > div, input[type="number"] {
    background-color: #262730 !important;
    color: white !important;
    border-radius: 6px !important;
    border: 1px solid #444 !important;
}
</style>
""", unsafe_allow_html=True)


# --- HELPER FUNCTIONS ---
def get_image_base64(path):
    if os.path.exists(path):
        with open(path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return ""

def clean_wisman(teks):
    if pd.isna(teks) or str(teks).strip() == "" or str(teks).strip() == "-": return 0
    try: return int(str(teks).strip().replace(".", "").replace(",", ""))
    except: return 0

def clean_hotel(teks):
    if pd.isna(teks) or str(teks).strip() == "" or str(teks).strip() == "-": return 0.0
    if isinstance(teks, (int, float, np.number)): return float(teks)
    teks_str = str(teks).strip()
    if "," in teks_str and "." not in teks_str: teks_str = teks_str.replace(",", ".")
    elif "," in teks_str and "." in teks_str: teks_str = teks_str.replace(".", "").replace(",", ".")
    try: return float(teks_str)
    except: return 0.0

def format_ribuan(angka): 
    return f"{int(round(float(angka))):,}".replace(",", ".")
def format_desimal(angka, digit=2): 
    return f"{float(angka):,.{digit}f}".replace(",", "X").replace(".", ",").replace("X", ".")
def get_nama_bulan(angka_bulan): 
    return ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][angka_bulan - 1]
def get_short_bulan(angka_bulan): 
    return ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"][angka_bulan - 1]
def get_arah_tren(angka): 
    return "peningkatan" if angka > 0 else "penurunan"
def get_arah_tren_singkat(angka): 
    return "naik" if angka > 0 else "turun"
def geser_tabel1(data_lama_negara, data_excel_baru):
    blok_lalu = data_lama_negara[1:9]          
    statis_jandes = [data_lama_negara[9]]      
    blok_ini = data_lama_negara[10:14]         
    bulan_ini = [data_excel_baru]              
    total_ini = [sum(blok_ini) + data_excel_baru] 
    return blok_lalu + statis_jandes + blok_ini + bulan_ini + total_ini


# --- KONTEN UTAMA (DI TENGAH LAYAR) ---
_, center_col, _ = st.columns([0.5, 6, 0.5])

with center_col:
    # --- HEADER: LOGO & JUDUL BERDAMPINGAN ---
    logo_base64 = get_image_base64("logosibrio.png")
    img_html = f'<img src="data:image/jpeg;base64,{logo_base64}" style="width: 85px; flex-shrink: 0; object-fit: contain;">' if logo_base64 else ''
    
    html_header = f"""
<div class="header-container">
    {img_html}
    <div class="main-title">SIBRIO</div>
</div>
<div class="main-subtitle">Platform Otomatis untuk Mendukung Penyusunan Berita Resmi Statistik yang<br>Cepat, Tepat, Efisien, dan Terstruktur</div>
"""
    st.markdown(html_header, unsafe_allow_html=True)

    # --- BARIS 1: BULAN & TAHUN ---
    col_b, col_t = st.columns(2)
    with col_b:
        st.markdown('<p class="custom-label">Bulan</p>', unsafe_allow_html=True)
        BULAN = st.selectbox("Bulan", [1,2,3,4,5,6,7,8,9,10,11,12], index=4, label_visibility="collapsed")
    with col_t:
        st.markdown('<p class="custom-label">Tahun</p>', unsafe_allow_html=True)
        TAHUN = st.number_input("Tahun", min_value=2020, max_value=2050, value=2026, label_visibility="collapsed")
        
    st.write("<br>", unsafe_allow_html=True)

    # --- BARIS 2: TPK YOY ---
    col_tpk1, col_tpk2 = st.columns(2)
    with col_tpk1:
        st.markdown('<p class="custom-label">TPK Bintang YOY (%)</p>', unsafe_allow_html=True)
        TPK_BINTANG_YOY = st.number_input("TPK Bintang", value=24.13, format="%.2f", label_visibility="collapsed")
    with col_tpk2:
        st.markdown('<p class="custom-label">TPK Non-Bintang YOY (%)</p>', unsafe_allow_html=True)
        TPK_NON_YOY = st.number_input("TPK Non", value=18.47, format="%.2f", label_visibility="collapsed")

    st.write("<br>", unsafe_allow_html=True)

    # --- BARIS 3: UPLOAD FILE EXCEL ---
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.markdown('<p class="custom-label">Data Wisman (Excel)</p>', unsafe_allow_html=True)
        file_wisman = st.file_uploader("Wisman", type=["xlsx", "xls"], label_visibility="collapsed")
    with col_up2:
        st.markdown('<p class="custom-label">Data VHTS (Excel)</p>', unsafe_allow_html=True)
        file_vhts = st.file_uploader("VHTS", type=["xlsx", "xls"], label_visibility="collapsed")

    st.write("<br>", unsafe_allow_html=True)

    # --- TOMBOL GENERATE ---
    if st.button("👇 Generate BRS Sekarang", use_container_width=True):
        # PROVINSI ACEH
        PROVINSI = 11 

        if file_wisman is not None and file_vhts is not None:
            # 1. Tentukan Nama File Statis (DIUPDATE SESUAI PERMINTAAN)
            file_template = "Template.docx"
            file_map_wisman = "wisman.xlsx"
            file_map_hotel = "hotel.xlsx"
            
            # 2. Hitung nama file BRS bulan lalu
            bln_lalu = 12 if BULAN == 1 else BULAN - 1
            thn_lalu = TAHUN - 1 if BULAN == 1 else TAHUN
            file_word_lama = f"BRS_Pariwisata_{bln_lalu}_{thn_lalu}.docx"
            
            # 3. Hitung nama file BRS yang akan dibuat
            nama_file_baru = f"BRS_Pariwisata_{BULAN}_{TAHUN}.docx"
            
            # 4. Validasi ketersediaan file
            if not os.path.exists(file_template):
                st.error(f"❌ File '{file_template}' tidak ditemukan di folder aplikasi!")
            elif not os.path.exists(file_map_wisman):
                st.error(f"❌ File '{file_map_wisman}' tidak ditemukan di folder aplikasi!")
            elif not os.path.exists(file_map_hotel):
                st.error(f"❌ File '{file_map_hotel}' tidak ditemukan di folder aplikasi!")
            elif not os.path.exists(file_word_lama):
                st.error(f"❌ File '{file_word_lama}' tidak ditemukan! Pastikan laporan bulan sebelumnya ada di folder yang sama.")
            else:
                with st.spinner(f'Memproses data dan menggunakan referensi {file_word_lama}...'):
                    try:
                        # EKSTRAKSI WORD LAMA
                        doc_lama = docx.Document(file_word_lama)
                        data_word = {"tabel1": {}, "tabel2": {}}
                        
                        tabel1 = doc_lama.tables[0]
                        target_negara = ["Malaysia", "Selandia Baru", "Perancis", "Jerman", "Australia", 
                                         "Belanda", "Amerika Serikat", "Singapura", "Inggris", "Thailand", "Lainnya", "JUMLAH"]
                        for row in tabel1.rows:
                            teks_baris = " ".join([str(c.text) for c in row.cells]).upper()
                            for target in target_negara:
                                if target.upper() in teks_baris and target not in data_word["tabel1"]:
                                    data_word["tabel1"][target] = [clean_wisman(sel.text) for sel in row.cells[-15:]]
                                    break
                        
                        tabel2 = doc_lama.tables[1]
                        data_rows = tabel2.rows[-12:] 
                        for i, kat in enumerate(["asing", "nusantara", "total", "tpk"]):
                            idx = i * 3
                            data_word["tabel2"][kat] = {
                                "baris_lama_1": [clean_hotel(data_rows[idx+1].cells[-2].text), clean_hotel(data_rows[idx+1].cells[-1].text)],
                                "baris_lama_2": [clean_hotel(data_rows[idx+2].cells[-2].text), clean_hotel(data_rows[idx+2].cells[-1].text)]
                            }

                        # EKSTRAKSI EXCEL WISMAN
                        wisman_map = pd.read_excel(file_map_wisman)
                        # Membuat nama sheet otomatis berdasarkan bulan yang dipilih (huruf kecil)
                        nama_bulan_sheet = get_short_bulan(BULAN).lower()
                        udara = pd.read_excel(file_wisman, sheet_name=f"udara {nama_bulan_sheet}", header=None)
                        non_udara = pd.read_excel(file_wisman, sheet_name=f"non udara {nama_bulan_sheet}", header=None)
                        
                        wisman_data = {}
                        for _, row in wisman_map.iterrows():
                            if row["tipe"] != "row": continue
                            nama, baris = row["variabel"], int(row["baris"]) - 1
                            wisman_data[nama] = clean_wisman(udara.iloc[baris, 11]) + clean_wisman(non_udara.iloc[baris, 11])
                        
                        top10_keys = ["malaysia", "selandia_baru", "perancis", "jerman", "australia", "belanda", "amerika_serikat", "singapura", "inggris", "thailand"]
                        wisman_data["lainnya"] = wisman_data["total_wisman"] - sum([wisman_data[k] for k in top10_keys])
                        wisman_sekarang = wisman_data["total_wisman"]

                        # EKSTRAKSI EXCEL HOTEL
                        hotel_df = pd.read_excel(file_vhts, sheet_name="Prov_Jenis")
                        hotel_df.columns = hotel_df.columns.str.strip().str.lower()
                        hotel_df["kd_prov"] = pd.to_numeric(hotel_df["kd_prov"], errors="coerce")
                        hotel_df["jenis_akomodasi"] = pd.to_numeric(hotel_df["jenis_akomodasi"], errors="coerce")
                        
                        aceh_bintang = hotel_df[(hotel_df["kd_prov"]==PROVINSI) & (hotel_df["jenis_akomodasi"]==1)]
                        aceh_non = hotel_df[(hotel_df["kd_prov"]==PROVINSI) & (hotel_df["jenis_akomodasi"]==2)]
                        
                        aceh_b = aceh_bintang.iloc[0] if not aceh_bintang.empty else pd.Series()
                        aceh_n = aceh_non.iloc[0] if not aceh_non.empty else pd.Series()
                        
                        hotel_data = {
                            "tpk_bintang": clean_hotel(aceh_b.get("tpk", 0)), "tpk_non": clean_hotel(aceh_n.get("tpk", 0)),
                            "rlm_asing": clean_hotel(aceh_b.get("rlmta", 0)), "rlm_asing_non": clean_hotel(aceh_n.get("rlmta", 0)),
                            "rlm_nus": clean_hotel(aceh_b.get("rlmtnus", 0)), "rlm_nus_non": clean_hotel(aceh_n.get("rlmtnus", 0)),
                            "rlm_total": clean_hotel(aceh_b.get("rlmtgab", 0)), "rlm_total_non": clean_hotel(aceh_n.get("rlmtgab", 0))
                        }

                        # KALKULASI NARASI
                        baris_baru_jumlah = geser_tabel1(data_word["tabel1"]["JUMLAH"], wisman_sekarang)
                        wisman_tahun_lalu, wisman_bulan_lalu, wisman_kumulatif_ini = baris_baru_jumlah[0], baris_baru_jumlah[12], baris_baru_jumlah[14]
                        wisman_kumulatif_lalu = baris_baru_jumlah[8] - sum(baris_baru_jumlah[1:8]) 
                        
                        mtm_w = ((wisman_sekarang - wisman_bulan_lalu) / wisman_bulan_lalu) * 100 if wisman_bulan_lalu else 0
                        yoy_w = ((wisman_sekarang - wisman_tahun_lalu) / wisman_tahun_lalu) * 100 if wisman_tahun_lalu else 0
                        ctc_w = ((wisman_kumulatif_ini - wisman_kumulatif_lalu) / wisman_kumulatif_lalu) * 100 if wisman_kumulatif_lalu else 0

                        tpk_bintang_lalu = data_word["tabel2"]["tpk"]["baris_lama_2"][0]
                        tpk_non_lalu = data_word["tabel2"]["tpk"]["baris_lama_2"][1]
                        
                        # SETUP CONTEXT
                        bln_lalu_angka = 12 if BULAN == 1 else BULAN - 1
                        thn_mtm = TAHUN - 1 if BULAN == 1 else TAHUN
                        bln_singkat = [
                            f"{get_short_bulan(bln_lalu_angka - 1 if bln_lalu_angka > 1 else 12)}'{str(thn_mtm - 1 if bln_lalu_angka == 1 else thn_mtm)[-2:]}",
                            f"{get_short_bulan(bln_lalu_angka)}'{str(thn_mtm)[-2:]}", f"{get_short_bulan(BULAN)}'{str(TAHUN)[-2:]}"
                        ]

                        context = {
                            "bulan": get_nama_bulan(BULAN), "tahun": str(TAHUN),
                            "bulan_lalu": get_nama_bulan(bln_lalu_angka), "tahun_lalu": str(TAHUN - 1), "tahun_mtm": str(thn_mtm),
                            "total_wisman": format_ribuan(wisman_sekarang), "total_wisman_lalu": format_ribuan(wisman_bulan_lalu),
                            "total_wisman_yoy": format_ribuan(wisman_tahun_lalu), "kumulatif_wisman": format_ribuan(wisman_kumulatif_ini),
                            "kumulatif_wisman_lalu": format_ribuan(wisman_kumulatif_lalu),
                            "mtm_wisman": format_desimal(abs(mtm_w)), "trend_mtm_wisman": get_arah_tren(mtm_w),
                            "yoy_wisman": format_desimal(abs(yoy_w)), "arah_yoy_wisman": get_arah_tren_singkat(yoy_w),
                            "trend_yoy_wisman": get_arah_tren(yoy_w), "ctc_wisman": format_desimal(abs(ctc_w)),
                            "trend_ctc_wisman": get_arah_tren(ctc_w), "negara_dominan": "Malaysia",
                            "persen_negara_dominan": format_desimal((wisman_data.get("malaysia", 0) / wisman_sekarang) * 100) if wisman_sekarang else "0",
                            "tpk_bintang": format_desimal(hotel_data.get("tpk_bintang", 0)), "tpk_bintang_lalu": format_desimal(tpk_bintang_lalu),
                            "selisih_tpk_bintang": format_desimal(abs(hotel_data.get("tpk_bintang", 0) - tpk_bintang_lalu)), "arah_tpk_bintang": get_arah_tren_singkat(hotel_data.get("tpk_bintang", 0) - tpk_bintang_lalu),
                            "trend_tpk_bintang_mtm": get_arah_tren(hotel_data.get("tpk_bintang", 0) - tpk_bintang_lalu),
                            "tpk_bintang_yoy": format_desimal(TPK_BINTANG_YOY), "selisih_tpk_bintang_yoy": format_desimal(abs(hotel_data.get("tpk_bintang", 0) - TPK_BINTANG_YOY)),
                            "trend_tpk_bintang_yoy": get_arah_tren(hotel_data.get("tpk_bintang", 0) - TPK_BINTANG_YOY),
                            "tpk_non": format_desimal(hotel_data.get("tpk_non", 0)), "selisih_tpk_non_mtm": format_desimal(abs(hotel_data.get("tpk_non", 0) - tpk_non_lalu)),
                            "trend_tpk_non_mtm": get_arah_tren(hotel_data.get("tpk_non", 0) - tpk_non_lalu), "selisih_tpk_non_yoy": format_desimal(abs(hotel_data.get("tpk_non", 0) - TPK_NON_YOY)),
                            "trend_tpk_non_yoy": get_arah_tren(hotel_data.get("tpk_non", 0) - TPK_NON_YOY),
                            "rlm_total": format_desimal(hotel_data.get("rlm_total", 0)), "rlm_total_lalu": format_desimal(data_word["tabel2"]["total"]["baris_lama_2"][0]),
                            "selisih_rlm_total_mtm": format_desimal(abs(hotel_data.get("rlm_total", 0) - data_word["tabel2"]["total"]["baris_lama_2"][0])), "trend_rlm_total_mtm": get_arah_tren(hotel_data.get("rlm_total", 0) - data_word["tabel2"]["total"]["baris_lama_2"][0]),
                            "rlm_asing": format_desimal(hotel_data.get("rlm_asing", 0)), "rlm_asing_lalu": format_desimal(data_word["tabel2"]["asing"]["baris_lama_2"][0]),
                            "selisih_rlm_asing_mtm": format_desimal(abs(hotel_data.get("rlm_asing", 0) - data_word["tabel2"]["asing"]["baris_lama_2"][0])), "trend_rlm_asing_mtm": get_arah_tren(hotel_data.get("rlm_asing", 0) - data_word["tabel2"]["asing"]["baris_lama_2"][0]),
                            "rlm_nus": format_desimal(hotel_data.get("rlm_nus", 0)), "selisih_rlm_nus_mtm": format_desimal(abs(hotel_data.get("rlm_nus", 0) - data_word["tabel2"]["nusantara"]["baris_lama_2"][0])),
                            "trend_rlm_nus_mtm": get_arah_tren(hotel_data.get("rlm_nus", 0) - data_word["tabel2"]["nusantara"]["baris_lama_2"][0]),
                            
                            "bln_1": bln_singkat[0], "bln_2": bln_singkat[1], "bln_3": bln_singkat[2],
                            "as_b1": format_desimal(data_word["tabel2"]["asing"]["baris_lama_1"][0]), "as_n1": format_desimal(data_word["tabel2"]["asing"]["baris_lama_1"][1]),
                            "as_b2": format_desimal(data_word["tabel2"]["asing"]["baris_lama_2"][0]), "as_n2": format_desimal(data_word["tabel2"]["asing"]["baris_lama_2"][1]),
                            "as_b3": format_desimal(hotel_data.get("rlm_asing", 0)), "as_n3": format_desimal(hotel_data.get("rlm_asing_non", 0)),
                            "nus_b1": format_desimal(data_word["tabel2"]["nusantara"]["baris_lama_1"][0]), "nus_n1": format_desimal(data_word["tabel2"]["nusantara"]["baris_lama_1"][1]),
                            "nus_b2": format_desimal(data_word["tabel2"]["nusantara"]["baris_lama_2"][0]), "nus_n2": format_desimal(data_word["tabel2"]["nusantara"]["baris_lama_2"][1]),
                            "nus_b3": format_desimal(hotel_data.get("rlm_nus", 0)), "nus_n3": format_desimal(hotel_data.get("rlm_nus_non", 0)),
                            "tot_b1": format_desimal(data_word["tabel2"]["total"]["baris_lama_1"][0]), "tot_n1": format_desimal(data_word["tabel2"]["total"]["baris_lama_1"][1]),
                            "tot_b2": format_desimal(data_word["tabel2"]["total"]["baris_lama_2"][0]), "tot_n2": format_desimal(data_word["tabel2"]["total"]["baris_lama_2"][1]),
                            "tot_b3": format_desimal(hotel_data.get("rlm_total", 0)), "tot_n3": format_desimal(hotel_data.get("rlm_total_non", 0)),
                            "tpk_b1": format_desimal(data_word["tabel2"]["tpk"]["baris_lama_1"][0]), "tpk_n1": format_desimal(data_word["tabel2"]["tpk"]["baris_lama_1"][1]),
                            "tpk_b2": format_desimal(data_word["tabel2"]["tpk"]["baris_lama_2"][0]), "tpk_n2": format_desimal(data_word["tabel2"]["tpk"]["baris_lama_2"][1]),
                            "tpk_b3": format_desimal(hotel_data.get("tpk_bintang", 0)), "tpk_n3": format_desimal(hotel_data.get("tpk_non", 0)),
                        }

                        # RENDER KE WORD
                        doc_tpl = DocxTemplate(file_template)
                        doc_tpl.render(context)
                        docx_obj = doc_tpl.docx 

                        tabel1_word = docx_obj.tables[0]
                        nama_bulan_s = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"]
                        headers_tabel1 = nama_bulan_s[BULAN - 1:] + ["Jan-Des"] + nama_bulan_s[:BULAN] + ["Total"]
                        row_header = tabel1_word.rows[1].cells[-15:] 
                        for col_idx in range(15): row_header[col_idx].text = headers_tabel1[col_idx]

                        baris_idx = 3
                        for key, negara in zip(top10_keys + ["lainnya"], target_negara[:-1]):
                            baris_baru = geser_tabel1(data_word["tabel1"][negara], wisman_data[key])
                            sel_tulis = tabel1_word.rows[baris_idx].cells[-15:] 
                            for col_idx in range(15): sel_tulis[col_idx].text = format_ribuan(baris_baru[col_idx])
                            baris_idx += 1

                        sel_tulis = tabel1_word.rows[baris_idx].cells[-15:]
                        for col_idx in range(15): sel_tulis[col_idx].text = format_ribuan(baris_baru_jumlah[col_idx])

                        # SIMPAN FILE LANGSUNG KE FOLDER
                        doc_tpl.save(nama_file_baru)

                        # SIMPAN KE MEMORY UNTUK FITUR DOWNLOAD BUTTON
                        output_stream = io.BytesIO()
                        doc_tpl.save(output_stream)
                        output_stream.seek(0)

                        st.success(f"🎉 Berhasil! Dokumen BRS telah dibuat dan tersimpan di folder ini sebagai: **{nama_file_baru}**")
                        st.download_button(
                            label="📥 Download Salinan Dokumen BRS",
                            data=output_stream,
                            file_name=nama_file_baru,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    
                    except Exception as e:
                        st.error(f"Terjadi kesalahan saat memproses data: {e}")
        else:
            st.warning("⚠️ Harap upload file Data Wisman dan Data VHTS terlebih dahulu!")
import streamlit as st
import pandas as pd
import numpy as np
import docx
from docxtpl import DocxTemplate
import io
import os
import base64
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials



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




# --- GOOGLE SHEETS HISTORIS ---
# Data historis disimpan di Google Sheets agar persisten di Streamlit Cloud.
# Spreadsheet harus dibagikan kepada email service account yang digunakan aplikasi.

HISTORIS_WISMAN_SHEET = "Historis Wisman"
HISTORIS_HOTEL_SHEET = "Historis Hotel"

WISMAN_HEADERS = [
    "tahun", "bulan", "periode", "negara", "nilai", "created_at", "updated_at"
]

HOTEL_HEADERS = [
    "tahun", "bulan", "periode", "kategori", "indikator",
    "hotel_bintang", "akomodasi_lainnya", "created_at", "updated_at"
]


@st.cache_resource(show_spinner=False)
def get_google_sheets_client():
    """Membuat koneksi gspread dari Streamlit Secrets."""
    if "gcp_service_account" not in st.secrets:
        raise RuntimeError(
            "Secrets 'gcp_service_account' belum dikonfigurasi di Streamlit Cloud."
        )

    spreadsheet_id = st.secrets.get("SIBRIO_SPREADSHEET_ID")
    if not spreadsheet_id:
        raise RuntimeError(
            "Secret 'SIBRIO_SPREADSHEET_ID' belum dikonfigurasi di Streamlit Cloud."
        )

    service_account_info = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes,
    )
    return gspread.authorize(credentials)


@st.cache_resource(show_spinner=False)
def get_historis_spreadsheet():
    """Membuka spreadsheet historis berdasarkan Spreadsheet ID."""
    client = get_google_sheets_client()
    spreadsheet_id = st.secrets["SIBRIO_SPREADSHEET_ID"]
    return client.open_by_key(spreadsheet_id)


def get_or_create_worksheet(title, headers):
    """Mengambil worksheet historis atau membuatnya jika belum ada."""
    spreadsheet = get_historis_spreadsheet()
    try:
        worksheet = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=title, rows=1000, cols=max(10, len(headers)))
        worksheet.append_row(headers, value_input_option="USER_ENTERED")
        return worksheet

    values = worksheet.get_all_values()
    if not values:
        worksheet.append_row(headers, value_input_option="USER_ENTERED")
    elif values[0] != headers:
        # Jangan menimpa data lama. Jika header berbeda, hentikan agar struktur data tidak rusak.
        raise RuntimeError(
            f"Header worksheet '{title}' tidak sesuai dengan struktur SIBRIO. "
            "Periksa worksheet tersebut sebelum melanjutkan."
        )
    return worksheet


def _find_matching_row(values, key_indexes, key_values):
    """Mencari nomor baris Google Sheets berdasarkan beberapa kolom kunci."""
    if not values:
        return None
    for row_number, row in enumerate(values[1:], start=2):
        # Pad row agar aman jika ada sel kosong di ujung.
        row = list(row) + [""] * (max(key_indexes) + 1 - len(row))
        if all(str(row[idx]).strip() == str(value).strip() for idx, value in zip(key_indexes, key_values)):
            return row_number
    return None


def _upsert_rows(worksheet, records, key_indexes):
    """UPSERT sederhana ke Google Sheets: update baris jika ada, append jika belum."""
    values = worksheet.get_all_values()
    updates = []
    appends = []

    for record in records:
        row_number = _find_matching_row(
            values,
            key_indexes,
            [record[idx] for idx in key_indexes],
        )
        if row_number is None:
            appends.append(record)
        else:
            updates.append((row_number, record))

    # Update record yang sudah ada.
    for row_number, record in updates:
        end_col = len(record)
        worksheet.update(
            f"A{row_number}:{gspread.utils.rowcol_to_a1(row_number, end_col).rstrip('0123456789')}{row_number}",
            [record],
            value_input_option="USER_ENTERED",
        )

    # Tambahkan record baru sekaligus.
    if appends:
        worksheet.append_rows(appends, value_input_option="USER_ENTERED")


def save_historis_wisman(tahun, bulan, wisman_data):
    """UPSERT data Wisman bulan berjalan ke Google Sheets tanpa menghapus bulan lain."""
    negara_map = [
        ("Malaysia", "malaysia"),
        ("Selandia Baru", "selandia_baru"),
        ("Perancis", "perancis"),
        ("Jerman", "jerman"),
        ("Australia", "australia"),
        ("Belanda", "belanda"),
        ("Amerika Serikat", "amerika_serikat"),
        ("Singapura", "singapura"),
        ("Inggris", "inggris"),
        ("Thailand", "thailand"),
        ("Lainnya", "lainnya"),
        ("JUMLAH", "total_wisman"),
    ]

    worksheet = get_or_create_worksheet(HISTORIS_WISMAN_SHEET, WISMAN_HEADERS)
    sekarang = datetime.now().isoformat(timespec="seconds")
    periode = f"{get_nama_bulan(int(bulan))} {int(tahun)}"

    records = []
    for negara, key in negara_map:
        nilai = int(round(float(wisman_data.get(key, 0) or 0)))
        records.append([
            int(tahun), int(bulan), periode, negara, nilai, sekarang, sekarang
        ])

    _upsert_rows(worksheet, records, key_indexes=[0, 1, 3])



def get_rolling_month_periods(brs_tahun, brs_bulan):
    """Menghasilkan 13 periode bulanan yang benar-benar tersimpan di Tabel 1 BRS.

    Struktur Tabel 1 SIBRIO:
    [bulan akhir tahun sebelumnya ... Des tahun sebelumnya,
     Jan ... bulan berjalan tahun berjalan]
    dengan satu kolom Jan-Des dan satu kolom Total yang tidak dihitung sebagai
    observasi bulanan.
    """
    periods = []
    # Bagian sebelum Jan-Des: bulan berjalan s.d. Desember tahun sebelumnya.
    for month in range(int(brs_bulan), 13):
        periods.append((int(brs_tahun) - 1, month))
    # Bagian setelah Jan-Des: Januari s.d. bulan BRS tahun berjalan.
    for month in range(1, int(brs_bulan) + 1):
        periods.append((int(brs_tahun), month))
    return periods


def backfill_historis_tabel1_dari_brs(doc_lama, brs_tahun, brs_bulan):
    """Mengimpor seluruh observasi bulanan Tabel 1 yang tersedia di BRS lama.

    Contoh BRS Juni 2026 akan mengisi data Tabel 1:
    Juni-Desember 2025 + Januari-Juni 2026.
    BRS Desember 2025 akan mengisi Desember 2024 + Januari-Desember 2025.
    Data yang sudah ada di Google Sheets tidak diduplikasi karena UPSERT.
    """
    tabel1 = doc_lama.tables[0]
    if len(tabel1.rows) < 2:
        return 0

    # Header dan data memakai 15 cell terakhir, sama seperti logika existing.
    header_cells = tabel1.rows[1].cells[-15:]
    headers = [str(cell.text).strip() for cell in header_cells]

    idx_jandes = 13 - int(brs_bulan)
    monthly_indices = [i for i in range(14) if i != idx_jandes]
    periods = get_rolling_month_periods(brs_tahun, brs_bulan)

    if len(monthly_indices) != len(periods):
        return 0

    target_negara = [
        "Malaysia", "Selandia Baru", "Perancis", "Jerman", "Australia",
        "Belanda", "Amerika Serikat", "Singapura", "Inggris", "Thailand",
        "Lainnya", "JUMLAH"
    ]

    # Ambil nilai per negara dari BRS lama.
    extracted = {}
    for row in tabel1.rows:
        teks_baris = " ".join(str(c.text) for c in row.cells).upper()
        for target in target_negara:
            if target.upper() in teks_baris and target not in extracted:
                cells = row.cells[-15:]
                values = [clean_wisman(cell.text) for cell in cells]
                extracted[target] = values
                break

    if not extracted:
        return 0

    # Konversi ke struktur wisman_data per periode agar fungsi save existing
    # tetap menjadi satu-satunya pintu UPSERT ke Google Sheets.
    total_saved = 0
    for period_index, (tahun, bulan) in enumerate(periods):
        cell_index = monthly_indices[period_index]
        wisman_data = {}
        key_map = {
            "Malaysia": "malaysia",
            "Selandia Baru": "selandia_baru",
            "Perancis": "perancis",
            "Jerman": "jerman",
            "Australia": "australia",
            "Belanda": "belanda",
            "Amerika Serikat": "amerika_serikat",
            "Singapura": "singapura",
            "Inggris": "inggris",
            "Thailand": "thailand",
            "Lainnya": "lainnya",
            "JUMLAH": "total_wisman",
        }
        for negara, values in extracted.items():
            if negara in key_map and cell_index < len(values):
                wisman_data[key_map[negara]] = values[cell_index]

        if wisman_data:
            save_historis_wisman(tahun, bulan, wisman_data)
            total_saved += 1

    return total_saved


def backfill_historis_tabel2_dari_brs(doc_lama, brs_tahun, brs_bulan, data_word):
    """Mengimpor periode Tabel 2 yang memang tersedia pada BRS lama.

    Struktur existing SIBRIO menggunakan tiga periode pada Tabel 2:
    dua bulan sebelumnya + bulan berjalan. Karena itu fungsi ini tidak
    mengarang data bulan yang tidak terdapat di dokumen.
    """
    # Dua periode lama + periode BRS saat ini.
    periode1_num = int(brs_tahun) * 12 + int(brs_bulan) - 2
    periode2_num = int(brs_tahun) * 12 + int(brs_bulan) - 1
    periode3_num = int(brs_tahun) * 12 + int(brs_bulan)

    def from_num(n):
        # Representasi bulan dengan basis 1.
        tahun = (n - 1) // 12
        bulan = ((n - 1) % 12) + 1
        return tahun, bulan

    p1 = from_num(periode1_num)
    p2 = from_num(periode2_num)
    p3 = from_num(periode3_num)

    # Nilai periode 3 diambil dari data Excel bulan berjalan BRS yang sedang
    # dibuat. Saat fungsi ini dipanggil sebelum save bulan berjalan, p3 belum
    # perlu disimpan di sini; fungsi utama akan melakukan save bulan berjalan.
    # Di sini kita hanya backfill dua periode lama yang ada di Word.
    periode_lama = [
        (p1, "baris_lama_1"),
        (p2, "baris_lama_2"),
    ]

    jumlah_disimpan = 0
    for (tahun, bulan), nama_baris in periode_lama:
        hotel_data_lama = {
            "rlm_asing": data_word["tabel2"]["asing"][nama_baris][0],
            "rlm_asing_non": data_word["tabel2"]["asing"][nama_baris][1],
            "rlm_nus": data_word["tabel2"]["nusantara"][nama_baris][0],
            "rlm_nus_non": data_word["tabel2"]["nusantara"][nama_baris][1],
            "rlm_total": data_word["tabel2"]["total"][nama_baris][0],
            "rlm_total_non": data_word["tabel2"]["total"][nama_baris][1],
            "tpk_bintang": data_word["tabel2"]["tpk"][nama_baris][0],
            "tpk_non": data_word["tabel2"]["tpk"][nama_baris][1],
        }
        save_historis_hotel(tahun, bulan, hotel_data_lama)
        jumlah_disimpan += 1

    return jumlah_disimpan


def backfill_historis_dari_brs(doc_lama, brs_tahun, brs_bulan, data_word):
    """Backfill histori dari BRS lama tanpa mengganggu data BRS berjalan."""
    hasil = {"tabel1": 0, "tabel2": 0}
    hasil["tabel1"] = backfill_historis_tabel1_dari_brs(
        doc_lama, brs_tahun, brs_bulan
    )
    hasil["tabel2"] = backfill_historis_tabel2_dari_brs(
        doc_lama, brs_tahun, brs_bulan, data_word
    )
    return hasil


def save_historis_hotel(tahun, bulan, hotel_data):
    """UPSERT data Hotel bulan berjalan ke Google Sheets tanpa menghapus bulan lain."""
    rows = [
        ("Rata-Rata Lama Menginap", "Asing",
         hotel_data.get("rlm_asing", 0), hotel_data.get("rlm_asing_non", 0)),
        ("Rata-Rata Lama Menginap", "Nusantara",
         hotel_data.get("rlm_nus", 0), hotel_data.get("rlm_nus_non", 0)),
        ("Rata-Rata Lama Menginap", "Total",
         hotel_data.get("rlm_total", 0), hotel_data.get("rlm_total_non", 0)),
        ("TPK", "TPK",
         hotel_data.get("tpk_bintang", 0), hotel_data.get("tpk_non", 0)),
    ]

    worksheet = get_or_create_worksheet(HISTORIS_HOTEL_SHEET, HOTEL_HEADERS)
    sekarang = datetime.now().isoformat(timespec="seconds")
    periode = f"{get_nama_bulan(int(bulan))} {int(tahun)}"

    records = []
    for kategori, indikator, hotel_bintang, akomodasi_lainnya in rows:
        records.append([
            int(tahun), int(bulan), periode, kategori, indikator,
            float(hotel_bintang or 0), float(akomodasi_lainnya or 0),
            sekarang, sekarang,
        ])

    _upsert_rows(worksheet, records, key_indexes=[0, 1, 3, 4])


def _worksheet_to_dataframe(title, headers):
    worksheet = get_or_create_worksheet(title, headers)
    values = worksheet.get_all_records()
    return pd.DataFrame(values, columns=headers) if values else pd.DataFrame(columns=headers)


def get_historis_wisman(tahun_awal, bulan_awal, tahun_akhir, bulan_akhir):
    """Mengambil historis Wisman dari Google Sheets sesuai rentang periode."""
    df = _worksheet_to_dataframe(HISTORIS_WISMAN_SHEET, WISMAN_HEADERS)
    if df.empty:
        return df

    df["tahun"] = pd.to_numeric(df["tahun"], errors="coerce")
    df["bulan"] = pd.to_numeric(df["bulan"], errors="coerce")
    df["nilai"] = pd.to_numeric(df["nilai"], errors="coerce").fillna(0).astype(int)
    df = df.dropna(subset=["tahun", "bulan"])

    periode_num = df["tahun"] * 12 + df["bulan"]
    awal = int(tahun_awal) * 12 + int(bulan_awal)
    akhir = int(tahun_akhir) * 12 + int(bulan_akhir)
    df = df[(periode_num >= awal) & (periode_num <= akhir)].copy()

    urutan_negara = [
        "Malaysia", "Selandia Baru", "Perancis", "Jerman", "Australia",
        "Belanda", "Amerika Serikat", "Singapura", "Inggris", "Thailand",
        "Lainnya", "JUMLAH"
    ]
    urutan = {nama: i for i, nama in enumerate(urutan_negara)}
    df["_urutan"] = df["negara"].map(urutan).fillna(999)
    return df.sort_values(["tahun", "bulan", "_urutan"]).drop(columns=["_urutan"])


def get_historis_hotel(tahun_awal, bulan_awal, tahun_akhir, bulan_akhir):
    """Mengambil historis Hotel dari Google Sheets sesuai rentang periode."""
    df = _worksheet_to_dataframe(HISTORIS_HOTEL_SHEET, HOTEL_HEADERS)
    if df.empty:
        return df

    for col in ["tahun", "bulan"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in ["hotel_bintang", "akomodasi_lainnya"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["tahun", "bulan"])

    periode_num = df["tahun"] * 12 + df["bulan"]
    awal = int(tahun_awal) * 12 + int(bulan_awal)
    akhir = int(tahun_akhir) * 12 + int(bulan_akhir)
    df = df[(periode_num >= awal) & (periode_num <= akhir)].copy()

    urutan_kategori = {"Rata-Rata Lama Menginap": 1, "TPK": 2}
    urutan_indikator = {"Asing": 1, "Nusantara": 2, "Total": 3, "TPK": 4}
    df["_kategori"] = df["kategori"].map(urutan_kategori).fillna(999)
    df["_indikator"] = df["indikator"].map(urutan_indikator).fillna(999)
    return df.sort_values(["tahun", "bulan", "_kategori", "_indikator"]).drop(
        columns=["_kategori", "_indikator"]
    )


def buat_excel_historis(df_wisman, df_hotel, pilihan):
    """Membuat file Excel historis sesuai pilihan user."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet_dibuat = False

        if pilihan in ("Tabel 1 - Wisman", "Tabel 1 + Tabel 2") and not df_wisman.empty:
            pivot_wisman = df_wisman.pivot(
                index="negara", columns="periode", values="nilai"
            )
            periode_order = (
                df_wisman[["tahun", "bulan", "periode"]]
                .drop_duplicates()
                .sort_values(["tahun", "bulan"])["periode"]
                .tolist()
            )
            urutan_negara = [
                "Malaysia", "Selandia Baru", "Perancis", "Jerman",
                "Australia", "Belanda", "Amerika Serikat", "Singapura",
                "Inggris", "Thailand", "Lainnya", "JUMLAH"
            ]
            pivot_wisman = pivot_wisman.reindex(
                index=[x for x in urutan_negara if x in pivot_wisman.index],
                columns=periode_order
            )
            pivot_wisman = pivot_wisman.fillna("")
            pivot_wisman.to_excel(writer, sheet_name="Historis Wisman")
            sheet_dibuat = True

        if pilihan in ("Tabel 2 - Hotel", "Tabel 1 + Tabel 2") and not df_hotel.empty:
            df_h = df_hotel.copy()
            periode_order = (
                df_h[["tahun", "bulan", "periode"]]
                .drop_duplicates()
                .sort_values(["tahun", "bulan"])["periode"]
                .tolist()
            )

            rows = []
            for _, r in df_h.iterrows():
                rows.append({
                    "Kategori": r["kategori"],
                    "Indikator": r["indikator"],
                    "Jenis": "Hotel Bintang",
                    "Nilai": r["hotel_bintang"],
                    "tahun": r["tahun"],
                    "bulan": r["bulan"],
                    "periode": r["periode"],
                })
                rows.append({
                    "Kategori": r["kategori"],
                    "Indikator": r["indikator"],
                    "Jenis": "Akomodasi Lainnya",
                    "Nilai": r["akomodasi_lainnya"],
                    "tahun": r["tahun"],
                    "bulan": r["bulan"],
                    "periode": r["periode"],
                })

            long_hotel = pd.DataFrame(rows)
            pivot_hotel = long_hotel.pivot_table(
                index=["Kategori", "Indikator", "Jenis"],
                columns="periode",
                values="Nilai",
                aggfunc="first"
            )
            pivot_hotel = pivot_hotel.reindex(columns=periode_order)
            pivot_hotel = pivot_hotel.reset_index()
            pivot_hotel.to_excel(writer, sheet_name="Historis Hotel", index=False)
            sheet_dibuat = True

        if not sheet_dibuat:
            pd.DataFrame({
                "Informasi": ["Tidak ada data historis untuk periode yang dipilih."]
            }).to_excel(writer, sheet_name="Informasi", index=False)

    output.seek(0)
    return output


def tampilkan_historis_data():
    """UI untuk melihat dan mengunduh data historis dari Google Sheets."""
    st.divider()
    st.markdown("## 📊 Historis Data")
    st.caption(
        "Data historis SIBRIO disimpan di Google Sheets sehingga tetap tersedia "
        "meskipun aplikasi Streamlit Cloud restart atau redeploy."
    )

    try:
        # Uji koneksi agar error konfigurasi tampil jelas di bagian historis.
        get_historis_spreadsheet()
    except Exception as e:
        st.warning(
            "⚠️ Fitur Historis Data belum terhubung ke Google Sheets. "
            f"Periksa Streamlit Secrets dan akses spreadsheet. Detail: {e}"
        )
        return

    pilihan = st.selectbox(
        "Pilih data yang ingin dilihat/download",
        ["Tabel 1 - Wisman", "Tabel 2 - Hotel", "Tabel 1 + Tabel 2"],
        key="historis_pilihan"
    )

    bulan_nama = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Periode Awal**")
        c1, c2 = st.columns(2)
        with c1:
            bulan_awal = st.selectbox(
                "Bulan awal", list(range(1, 13)),
                format_func=lambda x: bulan_nama[x - 1],
                index=0, key="hist_bulan_awal"
            )
        with c2:
            tahun_awal = st.number_input(
                "Tahun awal", min_value=2020, max_value=2050,
                value=2025, step=1, key="hist_tahun_awal"
            )

    with col2:
        st.markdown("**Periode Akhir**")
        c3, c4 = st.columns(2)
        with c3:
            bulan_akhir = st.selectbox(
                "Bulan akhir", list(range(1, 13)),
                format_func=lambda x: bulan_nama[x - 1],
                index=4, key="hist_bulan_akhir"
            )
        with c4:
            tahun_akhir = st.number_input(
                "Tahun akhir", min_value=2020, max_value=2050,
                value=2026, step=1, key="hist_tahun_akhir"
            )

    awal = int(tahun_awal) * 12 + int(bulan_awal)
    akhir = int(tahun_akhir) * 12 + int(bulan_akhir)

    if awal > akhir:
        st.error("❌ Periode awal tidak boleh lebih besar daripada periode akhir.")
        return

    try:
        df_wisman = pd.DataFrame()
        df_hotel = pd.DataFrame()

        if pilihan in ("Tabel 1 - Wisman", "Tabel 1 + Tabel 2"):
            df_wisman = get_historis_wisman(
                tahun_awal, bulan_awal, tahun_akhir, bulan_akhir
            )

        if pilihan in ("Tabel 2 - Hotel", "Tabel 1 + Tabel 2"):
            df_hotel = get_historis_hotel(
                tahun_awal, bulan_awal, tahun_akhir, bulan_akhir
            )
    except Exception as e:
        st.error(f"❌ Gagal membaca data historis dari Google Sheets: {e}")
        return

    ada_data = not df_wisman.empty or not df_hotel.empty

    if not ada_data:
        st.info(
            "ℹ️ Belum terdapat data historis untuk periode yang dipilih. "
            "Data akan otomatis tersimpan setiap kali Generate BRS berhasil dilakukan."
        )
        return

    if not df_wisman.empty:
        st.markdown("### Tabel 1 — Historis Wisman")
        pivot_wisman = df_wisman.pivot(
            index="negara", columns="periode", values="nilai"
        )
        periode_order = (
            df_wisman[["tahun", "bulan", "periode"]]
            .drop_duplicates()
            .sort_values(["tahun", "bulan"])["periode"]
            .tolist()
        )
        urutan_negara = [
            "Malaysia", "Selandia Baru", "Perancis", "Jerman",
            "Australia", "Belanda", "Amerika Serikat", "Singapura",
            "Inggris", "Thailand", "Lainnya", "JUMLAH"
        ]
        pivot_wisman = pivot_wisman.reindex(
            index=[x for x in urutan_negara if x in pivot_wisman.index],
            columns=periode_order
        )
        st.dataframe(pivot_wisman.fillna(""), use_container_width=True)

    if not df_hotel.empty:
        st.markdown("### Tabel 2 — Historis Hotel")
        df_h = df_hotel.copy()
        periode_order = (
            df_h[["tahun", "bulan", "periode"]]
            .drop_duplicates()
            .sort_values(["tahun", "bulan"])["periode"]
            .tolist()
        )
        preview_rows = []
        for _, r in df_h.iterrows():
            preview_rows.append({
                "Kategori": r["kategori"],
                "Indikator": r["indikator"],
                "Jenis": "Hotel Bintang",
                "periode": r["periode"],
                "nilai": r["hotel_bintang"],
            })
            preview_rows.append({
                "Kategori": r["kategori"],
                "Indikator": r["indikator"],
                "Jenis": "Akomodasi Lainnya",
                "periode": r["periode"],
                "nilai": r["akomodasi_lainnya"],
            })

        preview_df = pd.DataFrame(preview_rows)
        preview_pivot = preview_df.pivot_table(
            index=["Kategori", "Indikator", "Jenis"],
            columns="periode",
            values="nilai",
            aggfunc="first"
        ).reindex(columns=periode_order).reset_index()

        st.dataframe(preview_pivot.fillna(""), use_container_width=True)

    nama_awal = f"{get_short_bulan(int(bulan_awal))}{int(tahun_awal)}"
    nama_akhir = f"{get_short_bulan(int(bulan_akhir))}{int(tahun_akhir)}"
    nama_download = f"Historis_SIBRIO_{nama_awal}_{nama_akhir}.xlsx"

    excel_data = buat_excel_historis(df_wisman, df_hotel, pilihan)

    st.download_button(
        "📥 Download Historis Excel",
        data=excel_data,
        file_name=nama_download,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


# FUNGSI PERGESERAN TABEL YANG SUDAH DINAMIS

def geser_tabel1(data_lama_negara, data_excel_baru, bulan_sekarang):

    # Ambil 13 kolom sebelumnya (membuang bulan paling lama), lalu tambahkan data bulan ini

    array_baru = data_lama_negara[1:14] + [data_excel_baru]



    # Cari tahu di mana posisi kolom "Jan-Des" berada (Polanya selalu 13 - bulan)

    idx_jandes = 13 - bulan_sekarang



    # Total tahun ini adalah jumlah dari kolom setelah "Jan-Des" sampai akhir

    data_tahun_ini = array_baru[idx_jandes + 1 :]

    total_ini = sum(data_tahun_ini)



    return array_baru + [total_ini]





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

        st.markdown('<p class="custom-label">TPK Bintang y-o-y (%)</p>', unsafe_allow_html=True)

        TPK_BINTANG_YOY = st.number_input("TPK Bintang", value=24.13, format="%.2f", label_visibility="collapsed")

    with col_tpk2:

        st.markdown('<p class="custom-label">TPK Non-Bintang y-o-y (%)</p>', unsafe_allow_html=True)

        TPK_NON_YOY = st.number_input("TPK Non", value=18.47, format="%.2f", label_visibility="collapsed")



    st.write("<br>", unsafe_allow_html=True)



    # --- BARIS 3: UPLOAD FILE EXCEL & WORD ---

    col_up1, col_up2, col_up3 = st.columns(3)

    with col_up1:

        st.markdown('<p class="custom-label">Data Wisman (Excel)</p>', unsafe_allow_html=True)

        file_wisman = st.file_uploader("Wisman", type=["xlsx", "xls"], label_visibility="collapsed")

    with col_up2:

        st.markdown('<p class="custom-label">Data VHTS (Excel)</p>', unsafe_allow_html=True)

        file_vhts = st.file_uploader("VHTS", type=["xlsx", "xls"], label_visibility="collapsed")

    with col_up3:

        st.markdown('<p class="custom-label">BRS Bulan Lalu (Word)</p>', unsafe_allow_html=True)

        file_brs_lama = st.file_uploader("BRS Lama", type=["docx"], label_visibility="collapsed")



    st.write("<br>", unsafe_allow_html=True)



    # --- TOMBOL GENERATE ---

    if st.button("👇 Generate BRS Sekarang", use_container_width=True):

        PROVINSI = 11 



        if file_wisman is not None and file_vhts is not None and file_brs_lama is not None:

            file_template = "Template.docx"

            file_map_wisman = "wisman.xlsx"

            file_map_hotel = "hotel.xlsx"



            nama_file_baru = f"BRS_Pariwisata_{BULAN}_{TAHUN}.docx"



            if not os.path.exists(file_template):

                st.error(f"❌ File '{file_template}' tidak ditemukan di folder aplikasi!")

            elif not os.path.exists(file_map_wisman):

                st.error(f"❌ File '{file_map_wisman}' tidak ditemukan di folder aplikasi!")

            elif not os.path.exists(file_map_hotel):

                st.error(f"❌ File '{file_map_hotel}' tidak ditemukan di folder aplikasi!")

            else:

                with st.spinner('Memproses data dan merangkai dokumen BRS...'):

                    try:

                        # EKSTRAKSI WORD LAMA

                        doc_lama = docx.Document(file_brs_lama)

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

                        # BACKFILL HISTORIS DARI BRS LAMA + SIMPAN BULAN BERJALAN.
                        # BRS lama digunakan sebagai sumber histori yang belum ada.
                        # Setelah itu data Excel bulan berjalan di-UPSERT sebagai data terbaru.
                        # Logika rolling table dan perhitungan BRS tidak diubah.
                        try:
                            brs_bulan = 12 if BULAN == 1 else BULAN - 1
                            brs_tahun = TAHUN - 1 if BULAN == 1 else TAHUN

                            hasil_backfill = backfill_historis_dari_brs(
                                doc_lama, brs_tahun, brs_bulan, data_word
                            )

                            save_historis_wisman(TAHUN, BULAN, wisman_data)
                            save_historis_hotel(TAHUN, BULAN, hotel_data)

                            st.success(
                                "☁️ Historis berhasil diperbarui. "
                                f"Tabel 1: {hasil_backfill['tabel1']} periode dari BRS lama + {get_nama_bulan(BULAN)} {TAHUN}; "
                                f"Tabel 2: {hasil_backfill['tabel2']} periode lama + {get_nama_bulan(BULAN)} {TAHUN}."
                            )
                        except Exception as e_historis:
                            st.warning(
                                "⚠️ BRS berhasil diproses, tetapi proses penyimpanan/backfill historis ke Google Sheets gagal. "
                                f"Periksa konfigurasi Secrets/akses spreadsheet. Detail: {e_historis}"
                            )

                        # KALKULASI NARASI DAN INDEKS DINAMIS

                        baris_baru_jumlah = geser_tabel1(data_word["tabel1"]["JUMLAH"], wisman_sekarang, BULAN)



                        wisman_tahun_lalu = baris_baru_jumlah[0]

                        wisman_bulan_lalu = baris_baru_jumlah[11] if BULAN == 1 else baris_baru_jumlah[12]

                        wisman_kumulatif_ini = baris_baru_jumlah[14]



                        idx_jandes = 13 - BULAN

                        if BULAN == 1:

                            wisman_kumulatif_lalu = baris_baru_jumlah[0]

                        else:

                            wisman_kumulatif_lalu = baris_baru_jumlah[idx_jandes] - sum(baris_baru_jumlah[1:idx_jandes])



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

                            baris_baru = geser_tabel1(data_word["tabel1"][negara], wisman_data[key], BULAN)

                            sel_tulis = tabel1_word.rows[baris_idx].cells[-15:] 

                            for col_idx in range(15): sel_tulis[col_idx].text = format_ribuan(baris_baru[col_idx])

                            baris_idx += 1



                        sel_tulis = tabel1_word.rows[baris_idx].cells[-15:]

                        for col_idx in range(15): sel_tulis[col_idx].text = format_ribuan(baris_baru_jumlah[col_idx])



                        # SIMPAN KE MEMORY UNTUK FITUR DOWNLOAD BUTTON

                        output_stream = io.BytesIO()

                        doc_tpl.save(output_stream)

                        output_stream.seek(0)



                        st.success(f"🎉 Berhasil! Silakan download dokumen BRS Anda di bawah ini:")

                        st.download_button(

                            label=f"📥 Download {nama_file_baru}",

                            data=output_stream,

                            file_name=nama_file_baru,

                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",

                            use_container_width=True

                        )



                    except Exception as e:

                        st.error(f"Terjadi kesalahan saat memproses data: {e}")

        else:

            st.warning("⚠️ Harap lengkapi ketiga file (Wisman, VHTS, dan BRS Lama) terlebih dahulu!")

# --- FITUR HISTORIS DATA ---
# Diletakkan di luar blok Generate agar selalu dapat digunakan setelah database terisi.
tampilkan_historis_data()

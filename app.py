import streamlit as st
import pandas as pd
import numpy as np
import docx
from docxtpl import DocxTemplate
import io
import os
import base64

# --- KONFIGURASI TAMPILAN WEB ---
st.set_page_config(
    page_title="SIBRIO - Generator BRS",
    page_icon="logobps.png",
    layout="wide"
)

# --- KONFIGURASI FILE HISTORIS ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORIS_FILE = os.path.join(BASE_DIR, "historis_pariwisata.xlsx")

# --- INJEKSI CUSTOM CSS ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;800;900&display=swap');

.stApp {
    background-color: #0E1117;
    font-family: 'Poppins', sans-serif;
    color: white;
}

.block-container {
    padding-top: 2rem !important;
}

.header-container {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 20px;
    margin-top: 10px;
}

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

.custom-label {
    text-align: center;
    font-weight: 600;
    font-size: 14px;
    color: #FAFAFA;
    margin-bottom: 5px;
}

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

div[data-baseweb="select"] > div, input[type="number"] {
    background-color: #262730 !important;
    color: white !important;
    border-radius: 6px !important;
    border: 1px solid #444 !important;
}

.history-title {
    color: #FF5722;
    font-size: 26px;
    font-weight: 800;
    margin-top: 35px;
    margin-bottom: 15px;
}

.history-subtitle {
    color: #E0E0E0;
    font-size: 14px;
    margin-bottom: 15px;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def get_image_base64(path):
    if os.path.exists(path):
        with open(path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return ""


def clean_wisman(teks):
    if pd.isna(teks) or str(teks).strip() == "" or str(teks).strip() == "-":
        return 0
    try:
        return int(str(teks).strip().replace(".", "").replace(",", ""))
    except Exception:
        return 0


def clean_hotel(teks):
    if pd.isna(teks) or str(teks).strip() == "" or str(teks).strip() == "-":
        return 0.0

    if isinstance(teks, (int, float, np.number)):
        return float(teks)

    teks_str = str(teks).strip()

    if "," in teks_str and "." not in teks_str:
        teks_str = teks_str.replace(",", ".")
    elif "," in teks_str and "." in teks_str:
        teks_str = teks_str.replace(".", "").replace(",", ".")

    try:
        return float(teks_str)
    except Exception:
        return 0.0


def format_ribuan(angka):
    return f"{int(round(float(angka))):,}".replace(",", ".")


def format_desimal(angka, digit=2):
    return (
        f"{float(angka):,.{digit}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def get_nama_bulan(angka_bulan):
    return [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ][angka_bulan - 1]


def get_short_bulan(angka_bulan):
    return [
        "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
        "Jul", "Ags", "Sep", "Okt", "Nov", "Des"
    ][angka_bulan - 1]


def get_arah_tren(angka):
    return "peningkatan" if angka > 0 else "penurunan"


def get_arah_tren_singkat(angka):
    return "naik" if angka > 0 else "turun"


# ============================================================
# FUNGSI ROLLING TABEL 1 - TETAP DIPERTAHANKAN
# ============================================================

def geser_tabel1(data_lama_negara, data_excel_baru, bulan_sekarang):
    # Fungsi ini hanya untuk kebutuhan tampilan BRS.
    # Fungsi ini TIDAK digunakan untuk menghapus data historis.
    array_baru = data_lama_negara[1:14] + [data_excel_baru]

    idx_jandes = 13 - bulan_sekarang

    data_tahun_ini = array_baru[idx_jandes + 1:]
    total_ini = sum(data_tahun_ini)

    return array_baru + [total_ini]


# ============================================================
# FUNGSI DATABASE HISTORIS
# ============================================================

TARGET_NEGARA = [
    "Malaysia",
    "Selandia Baru",
    "Perancis",
    "Jerman",
    "Australia",
    "Belanda",
    "Amerika Serikat",
    "Singapura",
    "Inggris",
    "Thailand",
    "Lainnya",
    "JUMLAH"
]

TOP10_KEYS = [
    "malaysia",
    "selandia_baru",
    "perancis",
    "jerman",
    "australia",
    "belanda",
    "amerika_serikat",
    "singapura",
    "inggris",
    "thailand"
]

TABEL2_KATEGORI = [
    "Asing",
    "Nusantara",
    "Total",
    "TPK"
]


def empty_historis_tabel1():
    return pd.DataFrame(
        columns=["Tahun", "Bulan", "Periode", "Kebangsaan", "Jumlah Wisman"]
    )


def empty_historis_tabel2():
    return pd.DataFrame(
        columns=[
            "Tahun", "Bulan", "Periode", "Kategori",
            "Hotel Bintang", "Akomodasi Lainnya"
        ]
    )


def load_historis():
    """
    Membaca database historis lokal.
    Jika file belum tersedia, kembalikan DataFrame kosong.
    """
    if not os.path.exists(HISTORIS_FILE):
        return empty_historis_tabel1(), empty_historis_tabel2()

    try:
        xls = pd.ExcelFile(HISTORIS_FILE)

        if "Historis_Tabel1" in xls.sheet_names:
            df1 = pd.read_excel(HISTORIS_FILE, sheet_name="Historis_Tabel1")
        else:
            df1 = empty_historis_tabel1()

        if "Historis_Tabel2" in xls.sheet_names:
            df2 = pd.read_excel(HISTORIS_FILE, sheet_name="Historis_Tabel2")
        else:
            df2 = empty_historis_tabel2()

        # Pastikan struktur kolom tetap konsisten.
        for col in empty_historis_tabel1().columns:
            if col not in df1.columns:
                df1[col] = pd.Series(dtype="object")
        df1 = df1[empty_historis_tabel1().columns]

        for col in empty_historis_tabel2().columns:
            if col not in df2.columns:
                df2[col] = pd.Series(dtype="object")
        df2 = df2[empty_historis_tabel2().columns]

        # Konversi tipe data.
        if not df1.empty:
            df1["Tahun"] = pd.to_numeric(df1["Tahun"], errors="coerce").astype("Int64")
            df1["Bulan"] = pd.to_numeric(df1["Bulan"], errors="coerce").astype("Int64")
            df1["Jumlah Wisman"] = pd.to_numeric(
                df1["Jumlah Wisman"], errors="coerce"
            ).fillna(0)

        if not df2.empty:
            df2["Tahun"] = pd.to_numeric(df2["Tahun"], errors="coerce").astype("Int64")
            df2["Bulan"] = pd.to_numeric(df2["Bulan"], errors="coerce").astype("Int64")
            df2["Hotel Bintang"] = pd.to_numeric(
                df2["Hotel Bintang"], errors="coerce"
            ).fillna(0)
            df2["Akomodasi Lainnya"] = pd.to_numeric(
                df2["Akomodasi Lainnya"], errors="coerce"
            ).fillna(0)

        return df1, df2

    except Exception:
        # Jika file rusak/tidak dapat dibaca, jangan membuat aplikasi utama crash.
        return empty_historis_tabel1(), empty_historis_tabel2()


def save_historis(df1, df2):
    """
    Menyimpan seluruh histori ke Excel.
    Tidak ada proses rolling pada data ini.
    """
    with pd.ExcelWriter(HISTORIS_FILE, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="Historis_Tabel1", index=False)
        df2.to_excel(writer, sheet_name="Historis_Tabel2", index=False)


def update_historis(
    tahun,
    bulan,
    wisman_data,
    hotel_data
):
    """
    Menambah atau memperbarui satu periode pada database historis.

    Identitas Tabel 1:
        Tahun + Bulan + Kebangsaan

    Identitas Tabel 2:
        Tahun + Bulan + Kategori
    """
    df1, df2 = load_historis()

    tahun = int(tahun)
    bulan = int(bulan)
    periode = f"{get_nama_bulan(bulan)} {tahun}"

    # --------------------------------------------------------
    # HISTORIS TABEL 1
    # --------------------------------------------------------
    rows1 = []

    for key, negara in zip(TOP10_KEYS + ["lainnya"], TARGET_NEGARA[:-1]):
        rows1.append({
            "Tahun": tahun,
            "Bulan": bulan,
            "Periode": periode,
            "Kebangsaan": negara,
            "Jumlah Wisman": int(wisman_data.get(key, 0))
        })

    rows1.append({
        "Tahun": tahun,
        "Bulan": bulan,
        "Periode": periode,
        "Kebangsaan": "JUMLAH",
        "Jumlah Wisman": int(wisman_data.get("total_wisman", 0))
    })

    new_df1 = pd.DataFrame(rows1)

    # Hapus periode yang sama terlebih dahulu agar tidak duplikat.
    if not df1.empty:
        df1 = df1[
            ~(
                (df1["Tahun"].astype("Int64") == tahun)
                & (df1["Bulan"].astype("Int64") == bulan)
            )
        ]

    df1 = pd.concat([df1, new_df1], ignore_index=True)

    # --------------------------------------------------------
    # HISTORIS TABEL 2
    # --------------------------------------------------------
    rows2 = [
        {
            "Tahun": tahun,
            "Bulan": bulan,
            "Periode": periode,
            "Kategori": "Asing",
            "Hotel Bintang": float(hotel_data.get("rlm_asing", 0)),
            "Akomodasi Lainnya": float(hotel_data.get("rlm_asing_non", 0))
        },
        {
            "Tahun": tahun,
            "Bulan": bulan,
            "Periode": periode,
            "Kategori": "Nusantara",
            "Hotel Bintang": float(hotel_data.get("rlm_nus", 0)),
            "Akomodasi Lainnya": float(hotel_data.get("rlm_nus_non", 0))
        },
        {
            "Tahun": tahun,
            "Bulan": bulan,
            "Periode": periode,
            "Kategori": "Total",
            "Hotel Bintang": float(hotel_data.get("rlm_total", 0)),
            "Akomodasi Lainnya": float(hotel_data.get("rlm_total_non", 0))
        },
        {
            "Tahun": tahun,
            "Bulan": bulan,
            "Periode": periode,
            "Kategori": "TPK",
            "Hotel Bintang": float(hotel_data.get("tpk_bintang", 0)),
            "Akomodasi Lainnya": float(hotel_data.get("tpk_non", 0))
        }
    ]

    new_df2 = pd.DataFrame(rows2)

    if not df2.empty:
        df2 = df2[
            ~(
                (df2["Tahun"].astype("Int64") == tahun)
                & (df2["Bulan"].astype("Int64") == bulan)
            )
        ]

    df2 = pd.concat([df2, new_df2], ignore_index=True)

    # --------------------------------------------------------
    # SORTING
    # --------------------------------------------------------
    if not df1.empty:
        df1["Tahun"] = pd.to_numeric(df1["Tahun"], errors="coerce")
        df1["Bulan"] = pd.to_numeric(df1["Bulan"], errors="coerce")
        df1["Jumlah Wisman"] = pd.to_numeric(
            df1["Jumlah Wisman"], errors="coerce"
        ).fillna(0)

        df1 = df1.sort_values(
            ["Tahun", "Bulan", "Kebangsaan"],
            kind="stable"
        ).reset_index(drop=True)

    if not df2.empty:
        df2["Tahun"] = pd.to_numeric(df2["Tahun"], errors="coerce")
        df2["Bulan"] = pd.to_numeric(df2["Bulan"], errors="coerce")
        df2["Hotel Bintang"] = pd.to_numeric(
            df2["Hotel Bintang"], errors="coerce"
        ).fillna(0)
        df2["Akomodasi Lainnya"] = pd.to_numeric(
            df2["Akomodasi Lainnya"], errors="coerce"
        ).fillna(0)

        kategori_order = {
            "Asing": 1,
            "Nusantara": 2,
            "Total": 3,
            "TPK": 4
        }
        df2["_urutan"] = df2["Kategori"].map(kategori_order).fillna(99)

        df2 = (
            df2.sort_values(
                ["Tahun", "Bulan", "_urutan"],
                kind="stable"
            )
            .drop(columns="_urutan")
            .reset_index(drop=True)
        )

    save_historis(df1, df2)

    return df1, df2


def get_available_periods():
    df1, df2 = load_historis()
    periods = set()

    for df in [df1, df2]:
        if not df.empty:
            for _, row in df[["Tahun", "Bulan"]].drop_duplicates().iterrows():
                if pd.notna(row["Tahun"]) and pd.notna(row["Bulan"]):
                    periods.add(
                        pd.Period(
                            year=int(row["Tahun"]),
                            month=int(row["Bulan"]),
                            freq="M"
                        )
                    )

    return sorted(periods)


def format_period_label(period):
    return f"{get_short_bulan(period.month)} {period.year}"


def get_historis_by_period(start_period, end_period):
    df1, df2 = load_historis()

    if df1.empty and df2.empty:
        return df1, df2

    def filter_period(df):
        if df.empty:
            return df.copy()

        temp = df.copy()
        temp["_period"] = [
            pd.Period(
                year=int(y),
                month=int(m),
                freq="M"
            )
            for y, m in zip(temp["Tahun"], temp["Bulan"])
        ]

        temp = temp[
            (temp["_period"] >= start_period)
            & (temp["_period"] <= end_period)
        ].copy()

        temp.drop(columns="_period", inplace=True)
        return temp

    return filter_period(df1), filter_period(df2)


def create_historis_excel(df1, df2, start_period, end_period):
    """
    Membuat file Excel historis dalam format wide:
    - Tabel 1: Kebangsaan x Periode
    - Tabel 2: Rincian x Periode
    """
    buffer = io.BytesIO()

    periods = pd.period_range(
        start=start_period,
        end=end_period,
        freq="M"
    )

    period_columns = [
        f"{get_short_bulan(p.month)} {p.year}"
        for p in periods
    ]

    # --------------------------------------------------------
    # TABEL 1
    # --------------------------------------------------------
    if not df1.empty:
        pivot1 = df1.pivot_table(
            index="Kebangsaan",
            columns=["Tahun", "Bulan"],
            values="Jumlah Wisman",
            aggfunc="first"
        )

        desired_index1 = [
            "Malaysia",
            "Selandia Baru",
            "Perancis",
            "Jerman",
            "Australia",
            "Belanda",
            "Amerika Serikat",
            "Singapura",
            "Inggris",
            "Thailand",
            "Lainnya",
            "JUMLAH"
        ]

        pivot1 = pivot1.reindex(index=desired_index1)

        new_columns1 = []
        for p in periods:
            key = (p.year, p.month)
            if key in pivot1.columns:
                new_columns1.append(pivot1[key])
            else:
                new_columns1.append(pd.Series(np.nan, index=pivot1.index))

        export1 = pd.concat(new_columns1, axis=1)
        export1.columns = period_columns
        export1 = export1.reset_index()
        export1 = export1.rename(columns={"Kebangsaan": "Kebangsaan"})
    else:
        export1 = pd.DataFrame({"Kebangsaan": TARGET_NEGARA})

        for col in period_columns:
            export1[col] = np.nan

    # --------------------------------------------------------
    # TABEL 2
    # --------------------------------------------------------
    if not df2.empty:
        rows2 = []

        detail_map = [
            ("Asing", "Hotel Bintang", "Rata-Rata Lama Menginap - Asing - Hotel Bintang"),
            ("Asing", "Akomodasi Lainnya", "Rata-Rata Lama Menginap - Asing - Akomodasi Lainnya"),
            ("Nusantara", "Hotel Bintang", "Rata-Rata Lama Menginap - Nusantara - Hotel Bintang"),
            ("Nusantara", "Akomodasi Lainnya", "Rata-Rata Lama Menginap - Nusantara - Akomodasi Lainnya"),
            ("Total", "Hotel Bintang", "Rata-Rata Lama Menginap - Total - Hotel Bintang"),
            ("Total", "Akomodasi Lainnya", "Rata-Rata Lama Menginap - Total - Akomodasi Lainnya"),
            ("TPK", "Hotel Bintang", "TPK - Hotel Bintang"),
            ("TPK", "Akomodasi Lainnya", "TPK - Akomodasi Lainnya")
        ]

        for kategori, kolom_nilai, nama_rincian in detail_map:
            subset = df2[df2["Kategori"] == kategori].copy()

            row = {"Rincian": nama_rincian}

            for p, label in zip(periods, period_columns):
                nilai = subset[
                    (subset["Tahun"] == p.year)
                    & (subset["Bulan"] == p.month)
                ]

                if nilai.empty:
                    row[label] = np.nan
                else:
                    row[label] = float(nilai.iloc[0][kolom_nilai])

            rows2.append(row)

        export2 = pd.DataFrame(rows2)
    else:
        export2 = pd.DataFrame({
            "Rincian": [
                "Rata-Rata Lama Menginap - Asing - Hotel Bintang",
                "Rata-Rata Lama Menginap - Asing - Akomodasi Lainnya",
                "Rata-Rata Lama Menginap - Nusantara - Hotel Bintang",
                "Rata-Rata Lama Menginap - Nusantara - Akomodasi Lainnya",
                "Rata-Rata Lama Menginap - Total - Hotel Bintang",
                "Rata-Rata Lama Menginap - Total - Akomodasi Lainnya",
                "TPK - Hotel Bintang",
                "TPK - Akomodasi Lainnya"
            ]
        })

        for col in period_columns:
            export2[col] = np.nan

    # --------------------------------------------------------
    # TULIS EXCEL
    # --------------------------------------------------------
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export1.to_excel(
            writer,
            sheet_name="Tabel 1 - Wisman",
            index=False
        )

        export2.to_excel(
            writer,
            sheet_name="Tabel 2 - Hotel",
            index=False
        )

        # Styling sederhana menggunakan openpyxl.
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        workbook = writer.book

        for sheet_name in ["Tabel 1 - Wisman", "Tabel 2 - Hotel"]:
            ws = workbook[sheet_name]

            header_fill = PatternFill(
                fill_type="solid",
                fgColor="F9B77F"
            )

            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            ws.freeze_panes = "B2"

            for col_idx in range(1, ws.max_column + 1):
                max_length = 0
                for row_idx in range(1, ws.max_row + 1):
                    value = ws.cell(row=row_idx, column=col_idx).value
                    if value is not None:
                        max_length = max(max_length, len(str(value)))

                ws.column_dimensions[get_column_letter(col_idx)].width = min(
                    max(max_length + 2, 12),
                    45
                )

            # Format angka.
            for row in ws.iter_rows(min_row=2):
                for cell in row[1:]:
                    if isinstance(cell.value, (int, float, np.number)):
                        cell.number_format = '#,##0.00'

    buffer.seek(0)
    return buffer


# ============================================================
# KONTEN UTAMA
# ============================================================

_, center_col, _ = st.columns([0.5, 6, 0.5])

with center_col:

    # --- HEADER ---
    logo_base64 = get_image_base64("logosibrio.png")

    img_html = (
        f'<img src="data:image/jpeg;base64,{logo_base64}" '
        'style="width: 85px; flex-shrink: 0; object-fit: contain;">'
        if logo_base64 else ''
    )

    html_header = f"""
    <div class="header-container">
        {img_html}
        <div class="main-title">SIBRIO</div>
    </div>

    <div class="main-subtitle">
        Platform Otomatis untuk Mendukung Penyusunan Berita Resmi Statistik yang<br>
        Cepat, Tepat, Efisien, dan Terstruktur
    </div>
    """

    st.markdown(html_header, unsafe_allow_html=True)

    # --- BARIS 1: BULAN & TAHUN ---
    col_b, col_t = st.columns(2)

    with col_b:
        st.markdown(
            '<p class="custom-label">Bulan</p>',
            unsafe_allow_html=True
        )
        BULAN = st.selectbox(
            "Bulan",
            [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
            index=4,
            label_visibility="collapsed"
        )

    with col_t:
        st.markdown(
            '<p class="custom-label">Tahun</p>',
            unsafe_allow_html=True
        )
        TAHUN = st.number_input(
            "Tahun",
            min_value=2020,
            max_value=2050,
            value=2026,
            label_visibility="collapsed"
        )

    st.write("<br>", unsafe_allow_html=True)

    # --- BARIS 2: TPK YOY ---
    col_tpk1, col_tpk2 = st.columns(2)

    with col_tpk1:
        st.markdown(
            '<p class="custom-label">TPK Bintang y-o-y (%)</p>',
            unsafe_allow_html=True
        )
        TPK_BINTANG_YOY = st.number_input(
            "TPK Bintang",
            value=24.13,
            format="%.2f",
            label_visibility="collapsed"
        )

    with col_tpk2:
        st.markdown(
            '<p class="custom-label">TPK Non-Bintang y-o-y (%)</p>',
            unsafe_allow_html=True
        )
        TPK_NON_YOY = st.number_input(
            "TPK Non",
            value=18.47,
            format="%.2f",
            label_visibility="collapsed"
        )

    st.write("<br>", unsafe_allow_html=True)

    # --- BARIS 3: UPLOAD ---
    col_up1, col_up2, col_up3 = st.columns(3)

    with col_up1:
        st.markdown(
            '<p class="custom-label">Data Wisman (Excel)</p>',
            unsafe_allow_html=True
        )
        file_wisman = st.file_uploader(
            "Wisman",
            type=["xlsx", "xls"],
            label_visibility="collapsed"
        )

    with col_up2:
        st.markdown(
            '<p class="custom-label">Data VHTS (Excel)</p>',
            unsafe_allow_html=True
        )
        file_vhts = st.file_uploader(
            "VHTS",
            type=["xlsx", "xls"],
            label_visibility="collapsed"
        )

    with col_up3:
        st.markdown(
            '<p class="custom-label">BRS Bulan Lalu (Word)</p>',
            unsafe_allow_html=True
        )
        file_brs_lama = st.file_uploader(
            "BRS Lama",
            type=["docx"],
            label_visibility="collapsed"
        )

    st.write("<br>", unsafe_allow_html=True)

    # ========================================================
    # GENERATE BRS
    # ========================================================

    if st.button(
        "👇 Generate BRS Sekarang",
        use_container_width=True
    ):

        PROVINSI = 11

        if (
            file_wisman is not None
            and file_vhts is not None
            and file_brs_lama is not None
        ):

            file_template = "Template.docx"
            file_map_wisman = "wisman.xlsx"
            file_map_hotel = "hotel.xlsx"

            nama_file_baru = f"BRS_Pariwisata_{BULAN}_{TAHUN}.docx"

            if not os.path.exists(file_template):
                st.error(
                    f"❌ File '{file_template}' tidak ditemukan di folder aplikasi!"
                )

            elif not os.path.exists(file_map_wisman):
                st.error(
                    f"❌ File '{file_map_wisman}' tidak ditemukan di folder aplikasi!"
                )

            elif not os.path.exists(file_map_hotel):
                st.error(
                    f"❌ File '{file_map_hotel}' tidak ditemukan di folder aplikasi!"
                )

            else:

                with st.spinner(
                    "Memproses data dan merangkai dokumen BRS..."
                ):

                    try:

                        # ====================================================
                        # EKSTRAKSI WORD LAMA
                        # ====================================================

                        doc_lama = docx.Document(file_brs_lama)

                        data_word = {
                            "tabel1": {},
                            "tabel2": {}
                        }

                        tabel1 = doc_lama.tables[0]

                        for row in tabel1.rows:

                            teks_baris = " ".join(
                                [str(c.text) for c in row.cells]
                            ).upper()

                            for target in TARGET_NEGARA:

                                if (
                                    target.upper() in teks_baris
                                    and target not in data_word["tabel1"]
                                ):

                                    data_word["tabel1"][target] = [
                                        clean_wisman(sel.text)
                                        for sel in row.cells[-15:]
                                    ]

                                    break

                        tabel2 = doc_lama.tables[1]

                        data_rows = tabel2.rows[-12:]

                        for i, kat in enumerate(
                            ["asing", "nusantara", "total", "tpk"]
                        ):

                            idx = i * 3

                            data_word["tabel2"][kat] = {
                                "baris_lama_1": [
                                    clean_hotel(
                                        data_rows[idx + 1].cells[-2].text
                                    ),
                                    clean_hotel(
                                        data_rows[idx + 1].cells[-1].text
                                    )
                                ],
                                "baris_lama_2": [
                                    clean_hotel(
                                        data_rows[idx + 2].cells[-2].text
                                    ),
                                    clean_hotel(
                                        data_rows[idx + 2].cells[-1].text
                                    )
                                ]
                            }

                        # ====================================================
                        # EKSTRAKSI EXCEL WISMAN
                        # ====================================================

                        wisman_map = pd.read_excel(file_map_wisman)

                        nama_bulan_sheet = get_short_bulan(BULAN).lower()

                        udara = pd.read_excel(
                            file_wisman,
                            sheet_name=f"udara {nama_bulan_sheet}",
                            header=None
                        )

                        non_udara = pd.read_excel(
                            file_wisman,
                            sheet_name=f"non udara {nama_bulan_sheet}",
                            header=None
                        )

                        wisman_data = {}

                        for _, row in wisman_map.iterrows():

                            if row["tipe"] != "row":
                                continue

                            nama = row["variabel"]
                            baris = int(row["baris"]) - 1

                            wisman_data[nama] = (
                                clean_wisman(udara.iloc[baris, 11])
                                + clean_wisman(non_udara.iloc[baris, 11])
                            )

                        wisman_data["lainnya"] = (
                            wisman_data["total_wisman"]
                            - sum(
                                wisman_data[k]
                                for k in TOP10_KEYS
                            )
                        )

                        wisman_sekarang = wisman_data["total_wisman"]

                        # ====================================================
                        # EKSTRAKSI EXCEL HOTEL
                        # ====================================================

                        hotel_df = pd.read_excel(
                            file_vhts,
                            sheet_name="Prov_Jenis"
                        )

                        hotel_df.columns = (
                            hotel_df.columns
                            .str.strip()
                            .str.lower()
                        )

                        hotel_df["kd_prov"] = pd.to_numeric(
                            hotel_df["kd_prov"],
                            errors="coerce"
                        )

                        hotel_df["jenis_akomodasi"] = pd.to_numeric(
                            hotel_df["jenis_akomodasi"],
                            errors="coerce"
                        )

                        aceh_bintang = hotel_df[
                            (hotel_df["kd_prov"] == PROVINSI)
                            & (hotel_df["jenis_akomodasi"] == 1)
                        ]

                        aceh_non = hotel_df[
                            (hotel_df["kd_prov"] == PROVINSI)
                            & (hotel_df["jenis_akomodasi"] == 2)
                        ]

                        aceh_b = (
                            aceh_bintang.iloc[0]
                            if not aceh_bintang.empty
                            else pd.Series()
                        )

                        aceh_n = (
                            aceh_non.iloc[0]
                            if not aceh_non.empty
                            else pd.Series()
                        )

                        hotel_data = {
                            "tpk_bintang": clean_hotel(
                                aceh_b.get("tpk", 0)
                            ),
                            "tpk_non": clean_hotel(
                                aceh_n.get("tpk", 0)
                            ),
                            "rlm_asing": clean_hotel(
                                aceh_b.get("rlmta", 0)
                            ),
                            "rlm_asing_non": clean_hotel(
                                aceh_n.get("rlmta", 0)
                            ),
                            "rlm_nus": clean_hotel(
                                aceh_b.get("rlmtnus", 0)
                            ),
                            "rlm_nus_non": clean_hotel(
                                aceh_n.get("rlmtnus", 0)
                            ),
                            "rlm_total": clean_hotel(
                                aceh_b.get("rlmtgab", 0)
                            ),
                            "rlm_total_non": clean_hotel(
                                aceh_n.get("rlmtgab", 0)
                            )
                        }

                        # ====================================================
                        # KALKULASI
                        # ====================================================

                        baris_baru_jumlah = geser_tabel1(
                            data_word["tabel1"]["JUMLAH"],
                            wisman_sekarang,
                            BULAN
                        )

                        wisman_tahun_lalu = baris_baru_jumlah[0]

                        wisman_bulan_lalu = (
                            baris_baru_jumlah[11]
                            if BULAN == 1
                            else baris_baru_jumlah[12]
                        )

                        wisman_kumulatif_ini = baris_baru_jumlah[14]

                        idx_jandes = 13 - BULAN

                        if BULAN == 1:
                            wisman_kumulatif_lalu = (
                                baris_baru_jumlah[0]
                            )
                        else:
                            wisman_kumulatif_lalu = (
                                baris_baru_jumlah[idx_jandes]
                                - sum(
                                    baris_baru_jumlah[1:idx_jandes]
                                )
                            )

                        mtm_w = (
                            (
                                (wisman_sekarang - wisman_bulan_lalu)
                                / wisman_bulan_lalu
                            ) * 100
                            if wisman_bulan_lalu
                            else 0
                        )

                        yoy_w = (
                            (
                                (wisman_sekarang - wisman_tahun_lalu)
                                / wisman_tahun_lalu
                            ) * 100
                            if wisman_tahun_lalu
                            else 0
                        )

                        ctc_w = (
                            (
                                (
                                    wisman_kumulatif_ini
                                    - wisman_kumulatif_lalu
                                )
                                / wisman_kumulatif_lalu
                            ) * 100
                            if wisman_kumulatif_lalu
                            else 0
                        )

                        tpk_bintang_lalu = data_word["tabel2"]["tpk"][
                            "baris_lama_2"
                        ][0]

                        tpk_non_lalu = data_word["tabel2"]["tpk"][
                            "baris_lama_2"
                        ][1]

                        # ====================================================
                        # SETUP CONTEXT
                        # ====================================================

                        bln_lalu_angka = (
                            12 if BULAN == 1 else BULAN - 1
                        )

                        thn_mtm = (
                            TAHUN - 1
                            if BULAN == 1
                            else TAHUN
                        )

                        bln_singkat = [
                            f"{get_short_bulan(bln_lalu_angka - 1 if bln_lalu_angka > 1 else 12)}'"
                            f"{str(thn_mtm - 1 if bln_lalu_angka == 1 else thn_mtm)[-2:]}",

                            f"{get_short_bulan(bln_lalu_angka)}'"
                            f"{str(thn_mtm)[-2:]}",

                            f"{get_short_bulan(BULAN)}'"
                            f"{str(TAHUN)[-2:]}"
                        ]

                        context = {
                            "bulan": get_nama_bulan(BULAN),
                            "tahun": str(TAHUN),

                            "bulan_lalu": get_nama_bulan(
                                bln_lalu_angka
                            ),

                            "tahun_lalu": str(TAHUN - 1),
                            "tahun_mtm": str(thn_mtm),

                            "total_wisman": format_ribuan(
                                wisman_sekarang
                            ),

                            "total_wisman_lalu": format_ribuan(
                                wisman_bulan_lalu
                            ),

                            "total_wisman_yoy": format_ribuan(
                                wisman_tahun_lalu
                            ),

                            "kumulatif_wisman": format_ribuan(
                                wisman_kumulatif_ini
                            ),

                            "kumulatif_wisman_lalu": format_ribuan(
                                wisman_kumulatif_lalu
                            ),

                            "mtm_wisman": format_desimal(
                                abs(mtm_w)
                            ),

                            "trend_mtm_wisman": get_arah_tren(
                                mtm_w
                            ),

                            "yoy_wisman": format_desimal(
                                abs(yoy_w)
                            ),

                            "arah_yoy_wisman": get_arah_tren_singkat(
                                yoy_w
                            ),

                            "trend_yoy_wisman": get_arah_tren(
                                yoy_w
                            ),

                            "ctc_wisman": format_desimal(
                                abs(ctc_w)
                            ),

                            "trend_ctc_wisman": get_arah_tren(
                                ctc_w
                            ),

                            "negara_dominan": "Malaysia",

                            "persen_negara_dominan": format_desimal(
                                (
                                    wisman_data.get("malaysia", 0)
                                    / wisman_sekarang
                                ) * 100
                            ) if wisman_sekarang else "0",

                            "tpk_bintang": format_desimal(
                                hotel_data.get("tpk_bintang", 0)
                            ),

                            "tpk_bintang_lalu": format_desimal(
                                tpk_bintang_lalu
                            ),

                            "selisih_tpk_bintang": format_desimal(
                                abs(
                                    hotel_data.get("tpk_bintang", 0)
                                    - tpk_bintang_lalu
                                )
                            ),

                            "arah_tpk_bintang": get_arah_tren_singkat(
                                hotel_data.get("tpk_bintang", 0)
                                - tpk_bintang_lalu
                            ),

                            "trend_tpk_bintang_mtm": get_arah_tren(
                                hotel_data.get("tpk_bintang", 0)
                                - tpk_bintang_lalu
                            ),

                            "tpk_bintang_yoy": format_desimal(
                                TPK_BINTANG_YOY
                            ),

                            "selisih_tpk_bintang_yoy": format_desimal(
                                abs(
                                    hotel_data.get("tpk_bintang", 0)
                                    - TPK_BINTANG_YOY
                                )
                            ),

                            "trend_tpk_bintang_yoy": get_arah_tren(
                                hotel_data.get("tpk_bintang", 0)
                                - TPK_BINTANG_YOY
                            ),

                            "tpk_non": format_desimal(
                                hotel_data.get("tpk_non", 0)
                            ),

                            "selisih_tpk_non_mtm": format_desimal(
                                abs(
                                    hotel_data.get("tpk_non", 0)
                                    - tpk_non_lalu
                                )
                            ),

                            "trend_tpk_non_mtm": get_arah_tren(
                                hotel_data.get("tpk_non", 0)
                                - tpk_non_lalu
                            ),

                            "selisih_tpk_non_yoy": format_desimal(
                                abs(
                                    hotel_data.get("tpk_non", 0)
                                    - TPK_NON_YOY
                                )
                            ),

                            "trend_tpk_non_yoy": get_arah_tren(
                                hotel_data.get("tpk_non", 0)
                                - TPK_NON_YOY
                            ),

                            "rlm_total": format_desimal(
                                hotel_data.get("rlm_total", 0)
                            ),

                            "rlm_total_lalu": format_desimal(
                                data_word["tabel2"]["total"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "selisih_rlm_total_mtm": format_desimal(
                                abs(
                                    hotel_data.get("rlm_total", 0)
                                    - data_word["tabel2"]["total"][
                                        "baris_lama_2"
                                    ][0]
                                )
                            ),

                            "trend_rlm_total_mtm": get_arah_tren(
                                hotel_data.get("rlm_total", 0)
                                - data_word["tabel2"]["total"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "rlm_asing": format_desimal(
                                hotel_data.get("rlm_asing", 0)
                            ),

                            "rlm_asing_lalu": format_desimal(
                                data_word["tabel2"]["asing"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "selisih_rlm_asing_mtm": format_desimal(
                                abs(
                                    hotel_data.get("rlm_asing", 0)
                                    - data_word["tabel2"]["asing"][
                                        "baris_lama_2"
                                    ][0]
                                )
                            ),

                            "trend_rlm_asing_mtm": get_arah_tren(
                                hotel_data.get("rlm_asing", 0)
                                - data_word["tabel2"]["asing"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "rlm_nus": format_desimal(
                                hotel_data.get("rlm_nus", 0)
                            ),

                            "selisih_rlm_nus_mtm": format_desimal(
                                abs(
                                    hotel_data.get("rlm_nus", 0)
                                    - data_word["tabel2"]["nusantara"][
                                        "baris_lama_2"
                                    ][0]
                                )
                            ),

                            "trend_rlm_nus_mtm": get_arah_tren(
                                hotel_data.get("rlm_nus", 0)
                                - data_word["tabel2"]["nusantara"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "bln_1": bln_singkat[0],
                            "bln_2": bln_singkat[1],
                            "bln_3": bln_singkat[2],

                            "as_b1": format_desimal(
                                data_word["tabel2"]["asing"][
                                    "baris_lama_1"
                                ][0]
                            ),

                            "as_n1": format_desimal(
                                data_word["tabel2"]["asing"][
                                    "baris_lama_1"
                                ][1]
                            ),

                            "as_b2": format_desimal(
                                data_word["tabel2"]["asing"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "as_n2": format_desimal(
                                data_word["tabel2"]["asing"][
                                    "baris_lama_2"
                                ][1]
                            ),

                            "as_b3": format_desimal(
                                hotel_data.get("rlm_asing", 0)
                            ),

                            "as_n3": format_desimal(
                                hotel_data.get("rlm_asing_non", 0)
                            ),

                            "nus_b1": format_desimal(
                                data_word["tabel2"]["nusantara"][
                                    "baris_lama_1"
                                ][0]
                            ),

                            "nus_n1": format_desimal(
                                data_word["tabel2"]["nusantara"][
                                    "baris_lama_1"
                                ][1]
                            ),

                            "nus_b2": format_desimal(
                                data_word["tabel2"]["nusantara"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "nus_n2": format_desimal(
                                data_word["tabel2"]["nusantara"][
                                    "baris_lama_2"
                                ][1]
                            ),

                            "nus_b3": format_desimal(
                                hotel_data.get("rlm_nus", 0)
                            ),

                            "nus_n3": format_desimal(
                                hotel_data.get("rlm_nus_non", 0)
                            ),

                            "tot_b1": format_desimal(
                                data_word["tabel2"]["total"][
                                    "baris_lama_1"
                                ][0]
                            ),

                            "tot_n1": format_desimal(
                                data_word["tabel2"]["total"][
                                    "baris_lama_1"
                                ][1]
                            ),

                            "tot_b2": format_desimal(
                                data_word["tabel2"]["total"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "tot_n2": format_desimal(
                                data_word["tabel2"]["total"][
                                    "baris_lama_2"
                                ][1]
                            ),

                            "tot_b3": format_desimal(
                                hotel_data.get("rlm_total", 0)
                            ),

                            "tot_n3": format_desimal(
                                hotel_data.get("rlm_total_non", 0)
                            ),

                            "tpk_b1": format_desimal(
                                data_word["tabel2"]["tpk"][
                                    "baris_lama_1"
                                ][0]
                            ),

                            "tpk_n1": format_desimal(
                                data_word["tabel2"]["tpk"][
                                    "baris_lama_1"
                                ][1]
                            ),

                            "tpk_b2": format_desimal(
                                data_word["tabel2"]["tpk"][
                                    "baris_lama_2"
                                ][0]
                            ),

                            "tpk_n2": format_desimal(
                                data_word["tabel2"]["tpk"][
                                    "baris_lama_2"
                                ][1]
                            ),

                            "tpk_b3": format_desimal(
                                hotel_data.get("tpk_bintang", 0)
                            ),

                            "tpk_n3": format_desimal(
                                hotel_data.get("tpk_non", 0)
                            )
                        }

                        # ====================================================
                        # RENDER KE WORD
                        # ====================================================

                        doc_tpl = DocxTemplate(file_template)
                        doc_tpl.render(context)

                        docx_obj = doc_tpl.docx

                        tabel1_word = docx_obj.tables[0]

                        nama_bulan_s = [
                            "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
                            "Jul", "Ags", "Sep", "Okt", "Nov", "Des"
                        ]

                        headers_tabel1 = (
                            nama_bulan_s[BULAN - 1:]
                            + ["Jan-Des"]
                            + nama_bulan_s[:BULAN]
                            + ["Total"]
                        )

                        row_header = tabel1_word.rows[1].cells[-15:]

                        for col_idx in range(15):
                            row_header[col_idx].text = headers_tabel1[
                                col_idx
                            ]

                        baris_idx = 3

                        for key, negara in zip(
                            TOP10_KEYS + ["lainnya"],
                            TARGET_NEGARA[:-1]
                        ):

                            baris_baru = geser_tabel1(
                                data_word["tabel1"][negara],
                                wisman_data[key],
                                BULAN
                            )

                            sel_tulis = (
                                tabel1_word
                                .rows[baris_idx]
                                .cells[-15:]
                            )

                            for col_idx in range(15):
                                sel_tulis[col_idx].text = format_ribuan(
                                    baris_baru[col_idx]
                                )

                            baris_idx += 1

                        sel_tulis = (
                            tabel1_word
                            .rows[baris_idx]
                            .cells[-15:]
                        )

                        for col_idx in range(15):
                            sel_tulis[col_idx].text = format_ribuan(
                                baris_baru_jumlah[col_idx]
                            )

                        # ====================================================
                        # SIMPAN DOKUMEN KE MEMORY
                        # ====================================================

                        output_stream = io.BytesIO()

                        doc_tpl.save(output_stream)

                        output_stream.seek(0)

                        # ====================================================
                        # SIMPAN HISTORI
                        # HANYA SETELAH DOKUMEN BERHASIL DIBENTUK
                        # ====================================================

                        update_historis(
                            tahun=TAHUN,
                            bulan=BULAN,
                            wisman_data=wisman_data,
                            hotel_data=hotel_data
                        )

                        st.success(
                            "🎉 Berhasil! BRS berhasil dibuat dan data "
                            "periode ini telah disimpan ke histori."
                        )

                        st.download_button(
                            label=f"📥 Download {nama_file_baru}",
                            data=output_stream,
                            file_name=nama_file_baru,
                            mime=(
                                "application/"
                                "vnd.openxmlformats-officedocument."
                                "wordprocessingml.document"
                            ),
                            use_container_width=True
                        )

                    except Exception as e:

                        st.error(
                            f"Terjadi kesalahan saat memproses data: {e}"
                        )

        else:

            st.warning(
                "⚠️ Harap lengkapi ketiga file "
                "(Wisman, VHTS, dan BRS Lama) terlebih dahulu!"
            )

    # ========================================================
    # DATA HISTORIS
    # ========================================================

    st.markdown(
        '<div class="history-title">📊 Data Historis</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="history-subtitle">'
        'Pilih rentang periode untuk melihat dan mengunduh seluruh '
        'historis Tabel 1 dan Tabel 2 yang telah tersimpan.'
        '</div>',
        unsafe_allow_html=True
    )

    available_periods = get_available_periods()

    if not available_periods:

        st.info(
            "ℹ️ Belum terdapat data historis. "
            "Silakan generate BRS terlebih dahulu."
        )

    else:

        col_hist1, col_hist2 = st.columns(2)

        with col_hist1:

            start_period = st.selectbox(
                "Periode Awal",
                available_periods,
                index=0,
                format_func=format_period_label,
                key="hist_start_period"
            )

        with col_hist2:

            default_end_index = len(available_periods) - 1

            end_period = st.selectbox(
                "Periode Akhir",
                available_periods,
                index=default_end_index,
                format_func=format_period_label,
                key="hist_end_period"
            )

        tampilkan_histori = st.button(
            "🔎 Tampilkan Historis",
            use_container_width=True,
            key="btn_tampilkan_historis"
        )

        if tampilkan_histori:

            if start_period > end_period:

                st.error(
                    "⚠️ Periode awal tidak boleh lebih besar "
                    "dari periode akhir."
                )

            else:

                hist1, hist2 = get_historis_by_period(
                    start_period,
                    end_period
                )

                if hist1.empty and hist2.empty:

                    st.warning(
                        "⚠️ Tidak ditemukan data historis "
                        "pada rentang periode tersebut."
                    )

                else:

                    st.markdown(
                        f"### Tabel 1 - Wisman "
                        f"({format_period_label(start_period)} "
                        f"s/d {format_period_label(end_period)})"
                    )

                    periods_selected = pd.period_range(
                        start=start_period,
                        end=end_period,
                        freq="M"
                    )

                    labels_selected = [
                        f"{get_short_bulan(p.month)} {p.year}"
                        for p in periods_selected
                    ]

                    if not hist1.empty:

                        preview1 = hist1.pivot_table(
                            index="Kebangsaan",
                            columns=["Tahun", "Bulan"],
                            values="Jumlah Wisman",
                            aggfunc="first"
                        )

                        preview1 = preview1.reindex(
                            TARGET_NEGARA
                        )

                        preview1.columns = [
                            f"{get_short_bulan(int(bulan))} {int(tahun)}"
                            for tahun, bulan in preview1.columns
                        ]

                        preview1 = preview1.reindex(
                            columns=labels_selected
                        )

                        preview1 = preview1.reset_index()

                        st.dataframe(
                            preview1,
                            use_container_width=True,
                            hide_index=True
                        )

                    else:

                        st.info(
                            "Tidak ada data Tabel 1 pada rentang tersebut."
                        )

                    st.markdown("### Tabel 2 - Hotel")

                    if not hist2.empty:

                        detail_map = [
                            (
                                "Asing",
                                "Hotel Bintang",
                                "Rata-Rata Lama Menginap - Asing - Hotel Bintang"
                            ),
                            (
                                "Asing",
                                "Akomodasi Lainnya",
                                "Rata-Rata Lama Menginap - Asing - Akomodasi Lainnya"
                            ),
                            (
                                "Nusantara",
                                "Hotel Bintang",
                                "Rata-Rata Lama Menginap - Nusantara - Hotel Bintang"
                            ),
                            (
                                "Nusantara",
                                "Akomodasi Lainnya",
                                "Rata-Rata Lama Menginap - Nusantara - Akomodasi Lainnya"
                            ),
                            (
                                "Total",
                                "Hotel Bintang",
                                "Rata-Rata Lama Menginap - Total - Hotel Bintang"
                            ),
                            (
                                "Total",
                                "Akomodasi Lainnya",
                                "Rata-Rata Lama Menginap - Total - Akomodasi Lainnya"
                            ),
                            (
                                "TPK",
                                "Hotel Bintang",
                                "TPK - Hotel Bintang"
                            ),
                            (
                                "TPK",
                                "Akomodasi Lainnya",
                                "TPK - Akomodasi Lainnya"
                            )
                        ]

                        preview_rows = []

                        for kategori, kolom_nilai, nama_rincian in detail_map:

                            subset = hist2[
                                hist2["Kategori"] == kategori
                            ]

                            row = {
                                "Rincian": nama_rincian
                            }

                            for p, label in zip(
                                periods_selected,
                                labels_selected
                            ):

                                nilai = subset[
                                    (subset["Tahun"] == p.year)
                                    & (subset["Bulan"] == p.month)
                                ]

                                row[label] = (
                                    float(nilai.iloc[0][kolom_nilai])
                                    if not nilai.empty
                                    else np.nan
                                )

                            preview_rows.append(row)

                        preview2 = pd.DataFrame(preview_rows)

                        st.dataframe(
                            preview2,
                            use_container_width=True,
                            hide_index=True
                        )

                    else:

                        st.info(
                            "Tidak ada data Tabel 2 pada rentang tersebut."
                        )

                    # ====================================================
                    # DOWNLOAD HISTORIS
                    # ====================================================

                    excel_historis = create_historis_excel(
                        hist1,
                        hist2,
                        start_period,
                        end_period
                    )

                    nama_download = (
                        "Historis_Data_Pariwisata_"
                        f"{get_short_bulan(start_period.month)}"
                        f"{start_period.year}_"
                        f"{get_short_bulan(end_period.month)}"
                        f"{end_period.year}.xlsx"
                    )

                    st.download_button(
                        label="📥 Download Historis Excel",
                        data=excel_historis,
                        file_name=nama_download,
                        mime=(
                            "application/"
                            "vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        use_container_width=True,
                        key="download_historis"
                    )

"""
Guepedia AI Editor Assistant
----------------------------
Aplikasi internal berbasis web (Streamlit) untuk membantu tim editor:
  1. Mengunggah naskah (.docx/.pdf/.txt)
  2. Mengekstrak teks & menghitung jumlah kata
  3. Mengklasifikasi genre & menghasilkan 3 opsi blurb via Groq API
  4. Meninjau, memilih/mengedit, dan menyimpan blurb terpilih ke SQLite
  5. Menyesuaikan format dinamis & generate dokumen final (.docx)

Jalankan dengan:  streamlit run app.py

CATATAN TEMA (penting):
App ini didesain untuk SELALU tampil dark-green (.aig-*, gradient hijau
tua), TIDAK PEDULI apakah tema browser/OS user Light atau Dark.

Sebelumnya ini murni ditangani lewat CSS override manual, tapi beberapa
widget native Streamlit (terutama file uploader) merender sebagian
elemennya lewat variabel tema Streamlit sendiri -- yang tetap ikut
Light kalau tema global tidak dipaksa -- sehingga CSS manual kalah dan
elemen jadi tidak kebaca (putih di atas putih). Untuk itu tema Streamlit
di app INI sekarang dipaksa Dark lewat .streamlit/config.toml (base="dark"
+ palet warna hijau tua yang sama dengan CSS di bawah), dan menu
hamburger (Settings > Theme) disembunyikan (toolbarMode="minimal") biar
user tidak bisa switch balik ke Light dan merusak kontras. Ini hanya
berlaku untuk app ini -- app Streamlit lain milik user tetap bebas pakai
tema pilihan masing-masing.

TAMBAHAN (fix "input masih ikut tema browser"):
Selain widget Streamlit, elemen FORM NATIVE milik browser sendiri
(spinner number_input, checkbox/radio bawaan OS, scrollbar, autofill
popup) tidak diatur oleh tema Streamlit sama sekali -- itu dikontrol
browser lewat CSS property `color-scheme`. Kalau properti ini tidak
di-set eksplisit, browser tetap merender sebagian kontrol native itu
ikut preferensi OS/browser (Light), walau `.streamlit/config.toml`
sudah dipaksa Dark. Makanya sekarang ditambahkan `color-scheme: dark`
secara eksplisit di :root/html/body/.stApp, plus `accent-color` hijau
untuk checkbox/radio/range native, supaya BENAR-BENAR tidak ada elemen
input yang bisa balik ke Light lagi.

CSS override di bawah tetap dipertahankan sebagai lapisan kedua untuk
styling kartu/pill custom (.aig-*) dan sebagai jaga-jaga tambahan pada
dropdown popover selectbox, radio/checkbox indicator, dan slider --
karena elemen-elemen itu dirender BaseWeb dengan class ter-generate
(emotion-cache-xxxx) yang beda-beda tiap versi Streamlit, sehingga CSS
di bawah menarget banyak variasi selector (data-baseweb, data-testid,
role) sekaligus supaya tetap kena di berbagai versi.
"""
import io
import os
import tempfile
from datetime import datetime
from pathlib import Path
from cover_ai.prompt_builder import build_prompt, hitung_spine
from cover_ai.gemini_client import generate_cover

import streamlit as st
from dotenv import load_dotenv

from db import init_db, save_analysis, update_chosen_blurb
from docx_formatter import normalize_docx_layout
from groq_client import GENRE_LIST, GroqRequestError, analyze_naskah, generate_synopsis
from mailmerge import MailMergeError, generate_book_docx
from naskah_parser import parse_naskah
from utils import count_words, extract_text, guess_title

DEFAULT_TEMPLATE_PATH = Path(__file__).parent / "templates" / "template_buku.docx"

load_dotenv()

st.set_page_config(
    page_title="Guepedia AI Editor Assistant",
    page_icon="✨",
    layout="wide",
)

init_db()

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
/* ========================================================================
   PAKSA color-scheme BROWSER ke dark -- ini kunci supaya kontrol native
   (spinner number_input, checkbox/radio bawaan OS, scrollbar, autofill,
   dropdown-arrow native <select>) TIDAK PERNAH lagi ikut preferensi
   Light/Dark browser atau OS user. Tanpa baris ini, config.toml Streamlit
   saja tidak cukup karena color-scheme adalah properti CSS terpisah yang
   dibaca langsung oleh rendering engine browser.
   ======================================================================== */
:root, html, body, .stApp{
    color-scheme: dark !important;
}

/* accent-color menyamakan warna checkbox/radio/range NATIVE (bukan
   BaseWeb) ke hijau tema, kalau-kalau ada yang lolos jadi biru/abu
   default browser */
:root, html, body, .stApp{
    accent-color: #53d88d;
}

.stApp {
    background: radial-gradient(
        ellipse at 15% 0%,
        #1f5f4a 0%,
        #12392e 40%,
        #081814 100%
    );
    color: #eefcf6;
}

section[data-testid="stSidebar"] {
    display: none;
}

.aig-eyebrow{
    font-size:12px;
    letter-spacing:3px;
    color:#7ef0b4;
    font-weight:700;
    text-transform:uppercase;
    margin-bottom:2px;
}

.aig-title{
    font-size:28px;
    font-weight:800;
    color:white;
    margin:0 0 6px 0;
}

.aig-sub{
    color:#b9e7d1;
    font-size:14.5px;
    margin-bottom:24px;
}

.aig-card{
    background:rgba(255,255,255,0.05);
    border:1px solid rgba(126,240,180,.18);
    border-radius:16px;
    padding:22px;
    backdrop-filter:blur(8px);
}

.aig-panel{
    background:rgba(255,255,255,0.04);
    border:1px solid rgba(126,240,180,.14);
    border-radius:18px;
    padding:18px 20px 22px 20px;
    margin-bottom:18px;
}

.aig-step{
    background:rgba(255,255,255,0.04);
    border:1px solid rgba(126,240,180,.14);
    border-radius:14px;
    padding:14px;
    min-height:120px;
}

.aig-step-number{
    display:inline-block;
    width:28px;
    height:28px;
    line-height:28px;
    text-align:center;
    border-radius:999px;
    background:linear-gradient(135deg,#53d88d,#1ea96c);
    color:white;
    font-weight:700;
    margin-bottom:8px;
}

.aig-step-title{
    font-size:13px;
    font-weight:700;
    color:white;
    margin-bottom:4px;
}

.aig-step-desc{
    font-size:12px;
    color:#bfdccf;
    line-height:1.5;
}

.aig-pill{
    display:inline-block;
    padding:4px 8px;
    border-radius:999px;
    background:rgba(126,240,180,.16);
    color:#bff4d0;
    font-size:11px;
    font-weight:700;
    letter-spacing:.5px;
    text-transform:uppercase;
    margin-bottom:8px;
}

.aig-label{
    font-size:11px;
    color:#8fd7b5;
    text-transform:uppercase;
    letter-spacing:1px;
}

.aig-value{
    font-size:15px;
    margin-bottom:14px;
    color:white;
}

.aig-footnote{
    color:#8cb9a8;
    font-size:11.5px;
    text-align:center;
    margin-top:40px;
    line-height:1.6;
}

div[data-testid="stTextArea"] textarea{
    background:rgba(255,255,255,.05) !important;
    color:#eefcf6 !important;
    border:1px solid rgba(126,240,180,.25) !important;
    border-radius:12px !important;
}
/* Div pembungkus BaseWeb di belakang textarea (kadang masih putih
   dan nembus dari sisi/pinggir kalau tidak ikut ditarget) */
div[data-testid="stTextArea"] > div,
div[data-testid="stTextArea"] div[data-baseweb="textarea"]{
    background:rgba(255,255,255,.05) !important;
}

div[data-baseweb="select"]{
    background:rgba(255,255,255,.05);
    border-radius:10px;
}

.stButton > button{
    border-radius:10px;
    font-weight:600;
}

.stButton > button[kind="primary"]{
    background:linear-gradient(135deg,#53d88d,#1ea96c);
    color:white;
    border:none;
}

/* Tombol sekunder (bukan primary) -- sebelumnya tidak ditarget sama
   sekali sehingga bisa ikut warna default tema (biru/merah) tergantung
   Light/Dark. Sekarang disamakan ke gaya hijau-outline. */
.stButton > button:not([kind="primary"]){
    background:rgba(126,240,180,.10) !important;
    color:#eefcf6 !important;
    border:1px solid rgba(126,240,180,.35) !important;
}

/* Tombol submit form (st.form_submit_button) dan tombol unduh
   (st.download_button) -- keduanya punya wrapper testid TERPISAH dari
   .stButton biasa di versi Streamlit terbaru, jadi kalau tidak
   ditarget eksplisit, teks/border-nya bisa lolos ikut warna tema
   Light/Dark bawaan browser. */
div[data-testid="stFormSubmitButton"] button,
div[data-testid="stDownloadButton"] button{
    border-radius:10px !important;
    font-weight:600 !important;
}
div[data-testid="stFormSubmitButton"] button[kind="primary"]{
    background:linear-gradient(135deg,#53d88d,#1ea96c) !important;
    color:white !important;
    border:none !important;
}
div[data-testid="stDownloadButton"] button{
    background:linear-gradient(135deg,#53d88d,#1ea96c) !important;
    color:white !important;
    border:none !important;
}
div[data-testid="stDownloadButton"] button *,
div[data-testid="stFormSubmitButton"] button *{
    color:white !important;
}

div[data-testid="stAlert"]{
    border-radius:12px;
}

/* --------------------------------------------------------------------
   Kontras teks terlepas dari tema Streamlit user (Light/Dark/System).
   -------------------------------------------------------------------- */
div[data-testid="stTextInput"] label,
div[data-testid="stTextArea"] label,
div[data-testid="stNumberInput"] label,
div[data-testid="stSelectbox"] label,
div[data-testid="stRadio"] label,
div[data-testid="stFileUploader"] label,
div[data-testid="stFileUploaderDropzoneInstructions"],
div[data-testid="stMarkdownContainer"] p,
div[data-testid="stCaptionContainer"],
div[data-testid="stExpander"] summary,
div[data-testid="stExpander"] summary p,
div[data-testid="stMetricLabel"],
div[data-testid="stMetricValue"],
div[data-testid="stTabs"] button p,
div[data-testid="stRadio"] div[role="radiogroup"] label p,
div[data-testid="stForm"] label,
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
.stApp p, .stApp span, .stApp label{
    color:#eefcf6 !important;
}

div[data-testid="stCaptionContainer"]{
    color:#b9e7d1 !important;
}

div[data-testid="stFileUploaderDropzoneInstructions"] span,
div[data-testid="stFileUploaderDropzoneInstructions"] small{
    color:#cfeee0 !important;
}

/* Placeholder text pada input/textarea tetap kebaca */
div[data-testid="stTextInput"] input::placeholder,
div[data-testid="stTextArea"] textarea::placeholder{
    color:#9fd0b8 !important;
    opacity:1;
}

div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input{
    color:#eefcf6 !important;
    background:rgba(255,255,255,.05) !important;
    border:1px solid rgba(126,240,180,.25) !important;
    border-radius:10px !important;
}

/* Div pembungkus BaseWeb di belakang input (mis. "base-input"),
   kadang masih putih dan nembus dari sisi/pinggir kotak angka/teks
   kalau cuma elemen <input>-nya yang ditarget */
div[data-testid="stTextInput"] div[data-baseweb="input"],
div[data-testid="stTextInput"] div[data-baseweb="base-input"],
div[data-testid="stNumberInput"] div[data-baseweb="input"],
div[data-testid="stNumberInput"] div[data-baseweb="base-input"]{
    background:rgba(255,255,255,.05) !important;
}

/* Selector cadangan generik: div pembungkus langsung di bawah
   stNumberInput/stTextInput, apapun struktur internal BaseWeb-nya
   (beda versi Streamlit beda struktur), plus catch-all terakhir
   untuk elemen <input>/<select> di mana saja dalam app supaya tidak
   ada lagi kotak putih yang lolos. */
div[data-testid="stNumberInput"] > div,
div[data-testid="stTextInput"] > div{
    background:rgba(255,255,255,.05) !important;
    border-radius:10px !important;
}
.stApp input,
.stApp select{
    background-color:rgba(255,255,255,.05) !important;
    color:#eefcf6 !important;
}

/* --------------------------------------------------------------------
   Selectbox tertutup: teks value yang lagi kepilih & panah ikonnya.
   -------------------------------------------------------------------- */
div[data-baseweb="select"] *{
    color:#eefcf6 !important;
    fill:#eefcf6 !important;
}
div[data-baseweb="select"] > div{
    background:rgba(255,255,255,.05) !important;
    border-color:rgba(126,240,180,.25) !important;
}

/* --------------------------------------------------------------------
   Dropdown popover selectbox: menarget BANYAK varian selector sekaligus
   karena struktur DOM-nya beda-beda tergantung versi Streamlit.
   Elemen ini dirender lewat portal, jadi TIDAK diberi prefix .stApp --
   ditarget langsung dari root dokumen supaya tetap kena.
   -------------------------------------------------------------------- */
ul[data-testid="stSelectboxVirtualDropdown"],
div[data-baseweb="popover"] div[data-baseweb="menu"],
div[data-baseweb="popover"] ul,
div[data-baseweb="menu"],
ul[role="listbox"]{
    background-color:#0e2e25 !important;
    border:1px solid rgba(126,240,180,.25) !important;
}

ul[data-testid="stSelectboxVirtualDropdown"] li,
div[data-baseweb="popover"] li,
div[data-baseweb="menu"] li,
li[role="option"],
div[role="option"]{
    background-color:transparent !important;
    color:#eefcf6 !important;
}

ul[data-testid="stSelectboxVirtualDropdown"] li:hover,
div[data-baseweb="popover"] li:hover,
div[data-baseweb="menu"] li:hover,
li[role="option"]:hover,
div[role="option"]:hover{
    background-color:rgba(126,240,180,.18) !important;
}

ul[data-testid="stSelectboxVirtualDropdown"] *,
div[data-baseweb="popover"] *,
div[data-baseweb="menu"] *{
    color:#eefcf6 !important;
}

/* --------------------------------------------------------------------
   Radio & checkbox: lingkaran/kotak indikator dan teks labelnya.
   Ditambah state checked/hover eksplisit hijau supaya tidak jatuh ke
   warna primary default tema (yang bisa beda kalau Light/Dark bocor).
   -------------------------------------------------------------------- */
div[data-baseweb="radio"] label,
div[data-baseweb="checkbox"] label{
    color:#eefcf6 !important;
}
div[data-baseweb="radio"] svg,
div[data-baseweb="checkbox"] svg{
    fill:#eefcf6 !important;
}
div[data-baseweb="radio"] [aria-checked="true"] svg,
div[data-baseweb="checkbox"] [aria-checked="true"] svg,
div[data-baseweb="radio"] [data-checked="true"] svg,
div[data-baseweb="checkbox"] [data-checked="true"] svg{
    fill:#53d88d !important;
}
div[data-baseweb="radio"] div[role="radio"],
div[data-baseweb="checkbox"] span{
    border-color:rgba(126,240,180,.45) !important;
}

/* --------------------------------------------------------------------
   Slider (kalau dipakai): track, handle, dan angka label.
   -------------------------------------------------------------------- */
div[data-baseweb="slider"] *{
    color:#eefcf6 !important;
}
div[data-baseweb="slider"] div[role="slider"]{
    background-color:#53d88d !important;
}

/* --------------------------------------------------------------------
   File uploader: dropzone (kotak drag & drop + tombol Browse files).
   CATATAN: elemen ini dirender sebagai <section>, bukan <div> -- kalau
   cuma ditarget lewat div[...] selector-nya tidak pernah kena sama sekali.
   -------------------------------------------------------------------- */
section[data-testid="stFileUploaderDropzone"],
div[data-testid="stFileUploaderDropzone"]{
    background:rgba(255,255,255,.05) !important;
    border:1px dashed rgba(126,240,180,.35) !important;
    border-radius:12px !important;
}
section[data-testid="stFileUploaderDropzone"] *,
div[data-testid="stFileUploaderDropzone"] *{
    color:#eefcf6 !important;
    fill:#eefcf6 !important;
}
section[data-testid="stFileUploaderDropzone"] *:not(button),
div[data-testid="stFileUploaderDropzone"] *:not(button){
    background:transparent !important;
}
div[data-testid="stFileUploaderDropzone"] small,
div[data-testid="stFileUploaderDropzoneInstructions"] small{
    color:#bfe9d3 !important;
}
div[data-testid="stFileUploaderDropzone"] button,
div[data-testid="stFileUploader"] button{
    background:rgba(126,240,180,.16) !important;
    color:#eefcf6 !important;
    border:1px solid rgba(126,240,180,.35) !important;
    border-radius:8px !important;
}

/* File uploader: baris file yang sudah ter-upload (nama, ukuran, ikon hapus) */
div[data-testid="stFileUploaderFile"],
div[data-testid="stFileUploaderFileData"]{
    background:rgba(255,255,255,.05) !important;
    border:1px solid rgba(126,240,180,.25) !important;
    border-radius:12px !important;
}
div[data-testid="stFileUploaderFile"] *,
div[data-testid="stFileUploaderFileData"] *{
    color:#eefcf6 !important;
    fill:#eefcf6 !important;
}
small[data-testid="stFileUploaderFileErrorMessage"]{
    color:#ffb4b4 !important;
}

/* --------------------------------------------------------------------
   Label widget di versi Streamlit yang lebih baru (stWidgetLabel),
   tidak selalu ketangkep selector lama div[data-testid="st..."] label
   -------------------------------------------------------------------- */
label[data-testid="stWidgetLabel"] p,
label[data-testid="stWidgetLabel"] span,
div[data-testid="stWidgetLabel"] p,
div[data-testid="stWidgetLabel"] span{
    color:#eefcf6 !important;
}

/* Tombol +/- pada number_input */
div[data-testid="stNumberInput"] button{
    background:rgba(255,255,255,.05) !important;
    color:#eefcf6 !important;
    border:1px solid rgba(126,240,180,.25) !important;
}
div[data-testid="stNumberInput"] button svg{
    fill:#eefcf6 !important;
}

/* Reinforce warna teks isian, jaga-jaga tema Light user menang specificity */
div[data-testid="stTextArea"] textarea,
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input{
    color:#eefcf6 !important;
}

/* Expander (mis. "Sinopsis") beserta ikon panahnya */
div[data-testid="stExpander"] details{
    background:rgba(255,255,255,.04) !important;
    border:1px solid rgba(126,240,180,.14) !important;
    border-radius:12px !important;
}
div[data-testid="stExpander"] summary svg{
    fill:#eefcf6 !important;
}

/* Alert / info box (mis. "Naskah siap diproses...") */
div[data-testid="stAlert"]{
    background:rgba(255,255,255,.06) !important;
    border:1px solid rgba(126,240,180,.2) !important;
}
div[data-testid="stAlert"] *{
    color:#eefcf6 !important;
}

/* --------------------------------------------------------------------
   Tabs (dipakai di halaman hasil analisis, "Opsi 1/2/3") -- indikator
   garis bawah tab aktif kadang ikut warna primary tema bawaan.
   -------------------------------------------------------------------- */
div[data-testid="stTabs"] button[aria-selected="true"]{
    color:#7ef0b4 !important;
}
div[data-baseweb="tab-highlight"]{
    background-color:#53d88d !important;
}
div[data-baseweb="tab-border"]{
    background-color:rgba(126,240,180,.18) !important;
}

/* Blok kode (st.code untuk "Copy Teks" blurb) */
div[data-testid="stCodeBlock"] pre,
div[data-testid="stCodeBlock"] code{
    background-color:rgba(255,255,255,.05) !important;
    color:#eefcf6 !important;
    border:1px solid rgba(126,240,180,.18) !important;
}

/* ========================================================================
   FINAL OVERRIDE -- sengaja diletakkan PALING AKHIR di stylesheet supaya
   menang kalau ada "dasi" (tie) specificity dengan CSS bawaan Streamlit.
   Prefix "html body" dipakai supaya specificity-nya lebih tinggi dari
   selector Streamlit sendiri. Ini menutup semua kasus yang masih lolos:
   Cover Generator, Kategori Genre, dan kotak Margin Formatter.

   -webkit-text-fill-color ditambahkan karena Chrome/Safari kadang
   menerapkan warna teks lewat properti ini (bukan cuma "color"), khusus-
   nya untuk input yang nilainya di-set lewat value=... seperti pada
   field Cover Generator -- kalau cuma "color" yang ditarget, teksnya
   tetap invisible walau CSS sudah "match".
   ======================================================================== */
html body .stApp input,
html body .stApp textarea,
html body .stApp select{
    background-color:rgba(255,255,255,.06) !important;
    color:#eefcf6 !important;
    -webkit-text-fill-color:#eefcf6 !important;
    border-color:rgba(126,240,180,.25) !important;
}

html body .stApp input:disabled,
html body .stApp textarea:disabled,
html body .stApp input[readonly],
html body .stApp textarea[readonly]{
    background-color:rgba(255,255,255,.04) !important;
    color:#cfeee0 !important;
    -webkit-text-fill-color:#cfeee0 !important;
    opacity:1 !important;
}

/* Fokus (klik/tab ke dalam field) -- browser sering kasih outline biru
   bawaan sendiri (bukan dari Streamlit) terlepas dari tema apapun,
   ini yang bikin "kedip biru" kalau user klik input. Diganti hijau. */
html body .stApp input:focus,
html body .stApp textarea:focus,
html body .stApp select:focus,
html body .stApp div[data-baseweb="input"]:focus-within,
html body .stApp div[data-baseweb="base-input"]:focus-within,
html body .stApp div[data-baseweb="textarea"]:focus-within,
html body .stApp div[data-baseweb="select"]:focus-within{
    outline:none !important;
    border-color:#53d88d !important;
    box-shadow:0 0 0 1px rgba(83,216,141,.45) !important;
}

html body .stApp div[data-baseweb="input"],
html body .stApp div[data-baseweb="base-input"],
html body .stApp div[data-baseweb="textarea"]{
    background-color:rgba(255,255,255,.06) !important;
}

html body .stApp div[data-baseweb="select"] > div,
html body .stApp div[data-baseweb="select"] *{
    background-color:rgba(255,255,255,.06) !important;
    color:#eefcf6 !important;
}

/* Radio pill (Jenis Buku, Sumber Sinopsis dll) */
html body .stApp div[data-baseweb="radio"] *,
html body .stApp div[role="radiogroup"] *{
    color:#eefcf6 !important;
}

/* Kotak metric (mis. "Lebar Spine") */
html body .stApp div[data-testid="stMetric"]{
    background-color:rgba(255,255,255,.04) !important;
    border:1px solid rgba(126,240,180,.14) !important;
    border-radius:12px !important;
    padding:10px 14px !important;
}

/* Cabut semua efek hover custom (transform/scale/shadow) -- tombol
   statis saja, tidak ada animasi apapun saat di-hover */
html body .stApp button:hover,
html body .stApp .stButton > button:hover{
    transform:none !important;
    box-shadow:none !important;
    scale:1 !important;
}

/* Scrollbar (Chrome/Edge/Safari) -- tanpa ini scrollbar tetap abu-abu
   terang bawaan OS meskipun color-scheme sudah dark, kalau browser
   lama/tertentu tidak menghormati color-scheme untuk elemen ini. */
html body .stApp ::-webkit-scrollbar{
    width:10px;
    height:10px;
}
html body .stApp ::-webkit-scrollbar-track{
    background:#0e2e25;
}
html body .stApp ::-webkit-scrollbar-thumb{
    background-color:rgba(126,240,180,.35);
    border-radius:8px;
}
</style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# State Defaults
# ---------------------------------------------------------------------------
defaults = {
    "page": "dashboard",
    "cover_prompt": "",
    "raw_text": "",
    "file_name": "",
    "pending_genre_hint": "",
    "result": None,
    "active_blurb": 0,
    "error_msg": "",
    "formatted_docx_bytes": None,
    "editor_sections": [],
    "editor_source_text": "",
    "fmt_margin_top": 2.0,
    "fmt_margin_bottom": 2.0,
    "fmt_margin_left": 2.5,
    "fmt_margin_right": 2.5,
    "fmt_font_name": "Times New Roman",
    "fmt_font_size": 12,
    "fmt_line_spacing": 1.15,
    "fmt_alignment": "justify",
    "fmt_header_text": "NAMA DOKUMEN - DIFORMAT OTOMATIS",
    "fmt_header_font_name": "Arial",
    "fmt_header_font_size": 9,
    "mm_nama_penulis": "",
    "mm_isbn": "",
    "mm_qrcbn": "",
    "mm_judul_naskah": "",
    "mm_tahun_cetak": "Mei 2025",
    "mm_kata_pengantar": "",
    "mm_tentang_penulis": "",
    "fmt_heading_alignment": "left",
    "sinopsis_mode": "Tulis/Upload Manual",
    "sinopsis_manual_text": "",
    "sinopsis_ai_text": "",
}
for k, v in defaults.items():
    st.session_state.setdefault(k, v)


def convert_docx_bytes(uploaded_file, config: dict) -> bytes:
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in:
        tmp_in.write(uploaded_file.read())
        tmp_input_path = tmp_in.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
        tmp_output_path = tmp_out.name

    normalize_docx_layout(tmp_input_path, tmp_output_path, config)

    with open(tmp_output_path, "rb") as f:
        formatted_bytes = f.read()

    st.session_state.formatted_docx_bytes = formatted_bytes
    return formatted_bytes


def extract_text_from_docx_bytes(docx_bytes: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(docx_bytes))
    return "\n".join(p.text for p in doc.paragraphs)


def build_editor_sections(raw_text: str, fallback_title: str | None = None) -> list[dict[str, str]]:
    if not raw_text or not raw_text.strip():
        return []

    kata_pengantar, sections = parse_naskah(raw_text, fallback_title=fallback_title or "")

    if kata_pengantar and not st.session_state.get("mm_kata_pengantar"):
        st.session_state["mm_kata_pengantar"] = kata_pengantar

    return [{"title": title.strip(), "body": "\n".join(body).strip()} for title, body in sections]


def serialize_editor_sections(sections: list[dict[str, str]]) -> str:
    lines: list[str] = []
    for section in sections:
        title = (section.get("title") or "").strip()
        body = (section.get("body") or "").strip()
        if title:
            lines.append(title)
        if body:
            lines.extend(line.strip() for line in body.splitlines() if line.strip())
    return "\n".join(lines)


def generate_book_bytes(template_file, meta: dict, chapter_title: str, kata_pengantar_text: str, naskah_text: str, format_config: dict | None = None, source_docx_bytes: bytes | None = None, qrcbn: str = "", sinopsis_text: str = "", tentang_penulis_text: str = "") -> bytes:
    if template_file is not None:
        if hasattr(template_file, "seek"):
            template_file.seek(0)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_tpl:
            tmp_tpl.write(template_file.read())
            template_path = tmp_tpl.name
    else:
        template_path = str(DEFAULT_TEMPLATE_PATH)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
        output_path = tmp_out.name

    generate_book_docx(
        template_path=template_path,
        output_path=output_path,
        nama_penulis=meta["nama_penulis"],
        judul_naskah=meta["judul_naskah"],
        isbn=meta["isbn"],
        tahun_cetak=meta["tahun_cetak"],
        kata_pengantar_text=kata_pengantar_text,
        chapter_title=chapter_title,
        naskah_text=naskah_text,
        format_config=format_config,
        source_docx_bytes=source_docx_bytes,
        qrcbn=qrcbn,
        sinopsis_text=sinopsis_text,
        tentang_penulis_text=tentang_penulis_text,
    )

    with open(output_path, "rb") as f:
        return f.read()


def run_analysis():
    st.session_state.error_msg = ""
    text = st.session_state.raw_text
    if count_words(text) < 5:
        st.session_state.error_msg = "Naskah terlalu pendek untuk dianalisis. Unggah file atau tempel teks naskah terlebih dahulu."
        return
    with st.spinner("Menganalisis naskah & menyusun blurb via Groq…"):
        try:
            analysis = analyze_naskah(text, genre_hint=st.session_state.pending_genre_hint)
        except GroqRequestError as e:
            st.session_state.error_msg = str(e)
            return

    title = guess_title(text)
    word_count = count_words(text)
    record_id = save_analysis(
        st.session_state.file_name or "(tempel manual)",
        title,
        word_count,
        analysis["genre"],
        analysis["blurbs"],
    )
    st.session_state.result = {
        "record_id": record_id,
        "title": title,
        "word_count": word_count,
        "genre": analysis["genre"],
        "blurbs": analysis["blurbs"],
    }
    st.session_state.active_blurb = 0
    st.session_state.page = "result"


def back_to_dashboard():
    st.session_state.page = "dashboard"
    st.session_state.error_msg = ""


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
col_logo, col_title = st.columns([0.08, 0.92])
with col_logo:
    st.image("asset/logo.jpg", width=140)
with col_title:
    st.markdown("<div class='aig-eyebrow'>Guepedia</div>", unsafe_allow_html=True)
    st.markdown("<div class='aig-title'>AI Editor Assistant</div>", unsafe_allow_html=True)

if not os.getenv("GROQ_API_KEY"):
    st.warning(
        "GROQ_API_KEY belum ditemukan di environment. Salin `.env.example` menjadi `.env`, "
        "isi API key dari https://console.groq.com/keys, lalu jalankan ulang `streamlit run app.py`.",
        icon="⚠️",
    )

if not os.getenv("GEMINI_API_KEY"):
    st.warning(
        "GEMINI_API_KEY belum ditemukan di environment. Fitur AI Cover Generator butuh ini. "
        "Daftar gratis di https://aistudio.google.com/apikey, isi ke file `.env` "
        "(GEMINI_API_KEY=xxxxx), lalu jalankan ulang `streamlit run app.py`.",
        icon="⚠️",
    )

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
if st.session_state.page == "dashboard":
    st.markdown(
        "<div class='aig-panel'>"
        "<div class='aig-pill'>Workflow</div>"
        "<div class='aig-sub'>Pilih naskah, jalankan analisis, lalu review hasil blurb dalam satu alur yang lebih terarah.</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    col_step1, col_step2, col_step3 = st.columns(3)
    with col_step1:
        st.markdown(
            "<div class='aig-step'><div class='aig-step-number'>1</div><div class='aig-step-title'>Input Naskah</div><div class='aig-step-desc'>Unggah file atau tempel teks langsung ke area input.</div></div>",
            unsafe_allow_html=True,
        )
    with col_step2:
        st.markdown(
            "<div class='aig-step'><div class='aig-step-number'>2</div><div class='aig-step-title'>Analisis Otomatis</div><div class='aig-step-desc'>Sistem menghitung kata, menilai genre, dan menghasilkan opsi blurb.</div></div>",
            unsafe_allow_html=True,
        )
    with col_step3:
        st.markdown(
            "<div class='aig-step'><div class='aig-step-number'>3</div><div class='aig-step-title'>Review & Export</div><div class='aig-step-desc'>Pilih blurb terbaik, simpan hasil, lalu ekspor ke dokumen final.</div></div>",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    col_main, col_side = st.columns([1.2, 0.8], gap="large")

    with col_main:
        st.markdown("### 1. Input & Analisis")
        st.markdown(
            "<div class='aig-sub'>Unggah dokumen naskah atau tempel teks secara manual, lalu jalankan analisis.</div>",
            unsafe_allow_html=True,
        )

        uploaded = st.file_uploader(
            "Seret & lepas file naskah (.docx / .pdf / .txt)",
            type=["docx", "pdf", "txt"],
            key="main_naskah_uploader",
        )
        if uploaded is not None:
            try:
                st.session_state.raw_text = extract_text(uploaded)
                st.session_state.file_name = uploaded.name
                st.success(f"{uploaded.name} terbaca ({count_words(st.session_state.raw_text):,} kata)".replace(",", "."))
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")

        st.markdown("<div style='text-align:center;color:#5f7099;font-size:12px;margin:10px 0;'>— atau tempel teks naskah secara manual —</div>", unsafe_allow_html=True)

        st.session_state.raw_text = st.text_area(
            "Isi naskah",
            value=st.session_state.raw_text,
            height=180,
            label_visibility="collapsed",
            placeholder="Tempel isi naskah di sini…",
        )

        if st.session_state.raw_text.strip():
            st.info(f"Naskah siap diproses — sekitar {count_words(st.session_state.raw_text):,} kata".replace(",", "."))
        else:
            st.caption("Belum ada naskah yang dimasukkan. Silakan unggah file atau tempel teks.")

        if st.session_state.raw_text.strip():
            fallback_title = st.session_state.result["title"] if st.session_state.result else guess_title(st.session_state.raw_text)
            if st.session_state.get("editor_source_text") != st.session_state.raw_text:
                st.session_state.editor_sections = build_editor_sections(st.session_state.raw_text, fallback_title=fallback_title)
                st.session_state.editor_source_text = st.session_state.raw_text

            if st.session_state.editor_sections:
                st.markdown("### 3. Editor otomatis naskah")
                st.caption("Sistem memecah naskah menjadi bagian yang bisa diedit per bab atau sub-bab secara otomatis.")
                for idx, section in enumerate(st.session_state.editor_sections):
                    with st.expander(f"{section['title'] or f'Bagian {idx + 1}'}", expanded=idx == 0):
                        title_value = st.text_input(
                            "Judul bagian",
                            value=section.get("title", ""),
                            key=f"editor_title_{idx}",
                        )
                        body_value = st.text_area(
                            "Isi bagian",
                            value=section.get("body", ""),
                            height=140,
                            key=f"editor_body_{idx}",
                        )
                        st.session_state.editor_sections[idx]["title"] = title_value
                        st.session_state.editor_sections[idx]["body"] = body_value

        st.session_state.pending_genre_hint = st.selectbox(
            "Kategori genre (opsional)",
            options=[""] + GENRE_LIST,
            format_func=lambda g: "Pilih Genre (auto-deteksi)" if g == "" else g,
        )

        st.write("")
        st.button("✨ Mulai Analisis", type="primary", use_container_width=True, on_click=run_analysis)

        if st.session_state.error_msg:
            st.error(st.session_state.error_msg)
            if "Groq" in st.session_state.error_msg or "GROQ" in st.session_state.error_msg:
                st.button("🔁 Klik Retry", on_click=run_analysis)
        st.divider()

        st.subheader("🎨 AI Cover Generator")
        col1, col2 = st.columns(2)
        with col1:

            jenis_buku = st.radio(
                "Jenis Buku",
                ["nonfiksi", "fiksi"],
                horizontal=True,
            )

            judul = st.text_input(
                "Judul Buku",
                value=st.session_state.get("mm_judul_naskah", ""),
            )

            subjudul = st.text_input(
                "Sub Judul",
                placeholder="Kosongkan jika tidak ada"
            )

            penulis = st.text_input(
                "Nama Penulis",
                value=st.session_state.get("mm_nama_penulis", ""),
            )

            kategori = st.text_input(
                "Kategori Buku",
                value=st.session_state.get("pending_genre_hint", "")
            )

            genre = st.text_input(
                "Genre Buku",
                value=st.session_state.get("pending_genre_hint", "")
            )

            jumlah_halaman = st.number_input(
                "Jumlah Halaman",
                min_value=1,
                value=100,
            )
        with col2:
            spine = hitung_spine(jumlah_halaman)
            st.metric(
                "Lebar Spine",
                f"{spine:.2f} mm"
            )

        sinopsis_final = (
            st.session_state.get("sinopsis_manual_text", "")
            or st.session_state.get("sinopsis_ai_text", "")
        )

        if st.button("🎨 Generate Cover"):

            field_wajib = {
                "Judul Buku": judul,
                "Nama Penulis": penulis,
                "Sinopsis": sinopsis_final,
            }
            kosong = [nama for nama, isi in field_wajib.items() if not isi.strip()]

            if kosong:
                st.warning(
                    "Sesuai brief, field berikut wajib diisi dulu sebelum generate cover: "
                    + ", ".join(kosong)
                    + ". Sinopsis bisa diisi manual atau di-generate AI di panel "
                    "'📝 Sinopsis' pada sidebar kanan."
                )
            else:
                try:
                    prompt = build_prompt(
                        jenis_buku=jenis_buku,
                        judul=judul,
                        subjudul=subjudul,
                        penulis=penulis,
                        kategori=kategori,
                        genre=genre,
                        sinopsis=sinopsis_final,
                        jumlah_halaman=jumlah_halaman,
                    )
                except Exception as e:
                    st.error(f"Gagal membangun prompt: {e}")
                    prompt = None

                if prompt is not None:
                    # Simpan prompt ke session state
                    st.session_state.cover_prompt = prompt
                    st.success("✅ Prompt berhasil dibuat.")

                    try:
                        image = generate_cover(prompt)
                    except Exception as e:
                        st.error(f"Gagal generate gambar cover: {e}")
                        image = None

                    if image is not None:
                        st.image(
                            image,
                            caption="Preview Cover",
                            use_container_width=True
                        )

                    with st.expander("🔍 Lihat Prompt"):
                        st.text_area(
                            "Prompt Final",
                            value=st.session_state.cover_prompt,
                            height=450,
                            disabled=True,
                        )

    with col_side:
        st.markdown("### 2. Format & Ekspor")
        st.markdown(
            "<div class='aig-sub'>Fitur pendukung ini dipisahkan agar alur utama tetap fokus dan tidak terlalu padat.</div>",
            unsafe_allow_html=True,
        )

        with st.expander("📝 Sinopsis", expanded=False):
            st.caption("Pilih sumber sinopsis untuk halaman depan buku.")
            sinopsis_mode = st.radio(
                "Sumber Sinopsis",
                options=["Tulis/Upload Manual", "Generate dengan AI"],
                key="sinopsis_mode",
                horizontal=True,
            )

            if sinopsis_mode == "Tulis/Upload Manual":
                st.text_area(
                    "Isi Sinopsis",
                    height=150,
                    placeholder="Tempel atau tulis sinopsis di sini…",
                    key="sinopsis_manual_text",
                )
            else:
                st.caption("AI meringkas SEMUA bab (hasil Editor Otomatis Naskah di atas) jadi satu sinopsis 50–100 kalimat.")
                if st.button("🤖 Generate Sinopsis dengan AI", use_container_width=True):
                    if not st.session_state.editor_sections:
                        st.warning("Unggah/tempel naskah dulu supaya Editor Otomatis Naskah bisa memecahnya per bab.")
                    else:
                        with st.spinner("Meringkas seluruh bab menjadi sinopsis…"):
                            try:
                                sections_for_ai = [
                                    (s.get("title", ""), s.get("body", "").splitlines())
                                    for s in st.session_state.editor_sections
                                ]
                                judul_for_ai = st.session_state.get("mm_judul_naskah") or (
                                    st.session_state.result["title"] if st.session_state.result else ""
                                )
                                st.session_state.sinopsis_ai_text = generate_synopsis(sections_for_ai, judul_for_ai)
                                st.success("Sinopsis berhasil dibuat. Review/edit hasilnya di bawah kalau perlu.")
                            except GroqRequestError as e:
                                st.error(str(e))
                st.text_area(
                    "Hasil Sinopsis AI (bisa diedit)",
                    height=150,
                    key="sinopsis_ai_text",
                )

        with st.expander("📄 Formatter & Siapkan Dokumen", expanded=True):
            st.markdown(
                "<div class='aig-sub'>Satu alur: atur format dinamis, isi metadata halaman depan, lalu langsung siapkan dokumen final.</div>",
                unsafe_allow_html=True,
            )
            with st.form("formatter_form"):
                st.markdown("#### 📏 Layout & Formatter Dinamis")
                cfg_margin_top = st.number_input("Margin Atas (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_top")
                cfg_margin_bottom = st.number_input("Margin Bawah (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_bottom")
                cfg_margin_left = st.number_input("Margin Kiri (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_left")
                cfg_margin_right = st.number_input("Margin Kanan (cm)", min_value=0.5, max_value=5.0, step=0.1, key="fmt_margin_right")
                cfg_font_name = st.text_input("Nama Font Utama", key="fmt_font_name")
                cfg_font_size = st.number_input("Ukuran Font Utama (pt)", min_value=8, max_value=20, step=1, key="fmt_font_size")
                cfg_line_spacing = st.number_input("Line Spacing", min_value=1.0, max_value=2.0, step=0.05, key="fmt_line_spacing")
                cfg_alignment = st.selectbox("Alignment Isi Naskah", options=["justify", "left", "center", "right"], key="fmt_alignment")
                cfg_heading_alignment = st.selectbox(
                    "Alignment Header Bab",
                    options=["left", "center", "right"],
                    key="fmt_heading_alignment",
                    format_func=lambda a: {"left": "Kiri", "center": "Tengah", "right": "Kanan"}[a],
                )
                cfg_header_text = st.text_input("Teks Header", key="fmt_header_text")
                cfg_header_font_name = st.text_input("Font Header", key="fmt_header_font_name")
                cfg_header_font_size = st.number_input("Ukuran Font Header (pt)", min_value=6, max_value=16, step=1, key="fmt_header_font_size")

                st.markdown("#### ✍️ Halaman Depan & Redaksi")
                mm_template_file = st.file_uploader(
                    "Template buku kustom (opsional)",
                    type=["docx"],
                    key="mm_template_uploader",
                )

                col_mm1, col_mm2 = st.columns(2)
                with col_mm1:
                    mm_nama_penulis = st.text_input("Nama Penulis", key="mm_nama_penulis")
                    mm_isbn = st.text_input("ISBN", placeholder="978-623-xxx-xx-x", key="mm_isbn")
                    mm_qrcbn = st.text_input("QRCBN (opsional)", placeholder="62-xxxx-xxxx-xxx", key="mm_qrcbn")
                with col_mm2:
                    default_judul = st.session_state.result["title"] if st.session_state.result else ""
                    if default_judul and not st.session_state.get("mm_judul_naskah"):
                        st.session_state["mm_judul_naskah"] = default_judul
                    mm_judul_naskah = st.text_input("Judul Naskah", key="mm_judul_naskah")
                    mm_tahun_cetak = st.text_input("Tahun/Bulan Cetak", key="mm_tahun_cetak")

                mm_kata_pengantar = st.text_area(
                    "Isi Kata Pengantar (Opsional)",
                    height=120,
                    placeholder="Setiap manusia punya cara yang berbeda dalam mencintai, merelakan, dan bertahan...",
                    key="mm_kata_pengantar"
                )

                mm_tentang_penulis = st.text_area(
                    "Isi Tentang Penulis (Opsional)",
                    height=120,
                    placeholder="Ceritakan sedikit tentang penulis: latar belakang, karya sebelumnya, dll...",
                    key="mm_tentang_penulis"
                )

                submitted_format = st.form_submit_button("📄 Format & Siapkan Dokumen", type="primary", use_container_width=True)
                if submitted_format:
                    if count_words(st.session_state.raw_text) < 5:
                        st.warning("Unggah file naskah atau tempel teks naskah terlebih dahulu di atas.")
                    elif not mm_nama_penulis or not mm_judul_naskah:
                        st.warning("Nama Penulis dan Judul Naskah wajib diisi.")
                    else:
                        try:
                            chapter_title = (
                                st.session_state.editor_sections[0]["title"]
                                if st.session_state.editor_sections
                                else (st.session_state.result["title"] if st.session_state.result else guess_title(st.session_state.raw_text))
                            )

                            source_docx_bytes = None
                            if uploaded is not None and uploaded.name.lower().endswith(".docx"):
                                config = {
                                    "margin_top_cm": cfg_margin_top,
                                    "margin_bottom_cm": cfg_margin_bottom,
                                    "margin_left_cm": cfg_margin_left,
                                    "margin_right_cm": cfg_margin_right,
                                    "font_name": cfg_font_name,
                                    "font_size_pt": cfg_font_size,
                                    "line_spacing": cfg_line_spacing,
                                    "alignment": cfg_alignment,
                                    "header_text": cfg_header_text,
                                    "header_font_name": cfg_header_font_name,
                                    "header_font_size_pt": cfg_header_font_size,
                                }
                                source_docx_bytes = convert_docx_bytes(uploaded, config)
                                st.session_state["formatted_docx_bytes"] = source_docx_bytes
                                st.session_state["formatted_docx_name"] = f"{Path(uploaded.name).stem}_normalized.docx"
                                naskah_text = extract_text_from_docx_bytes(source_docx_bytes)
                            else:
                                naskah_text = serialize_editor_sections(st.session_state.editor_sections) if st.session_state.editor_sections else st.session_state.raw_text

                            format_config = {
                                "margin_top_cm": cfg_margin_top,
                                "margin_bottom_cm": cfg_margin_bottom,
                                "margin_left_cm": cfg_margin_left,
                                "margin_right_cm": cfg_margin_right,
                                "font_name": cfg_font_name,
                                "font_size_pt": cfg_font_size,
                                "line_spacing": cfg_line_spacing,
                                "alignment": cfg_alignment,
                                "heading_alignment": cfg_heading_alignment,
                                "header_text": cfg_header_text,
                                "header_font_name": cfg_header_font_name,
                                "header_font_size_pt": cfg_header_font_size,
                            }

                            final_sinopsis_text = (
                                st.session_state.get("sinopsis_manual_text", "")
                                if st.session_state.get("sinopsis_mode") == "Tulis/Upload Manual"
                                else st.session_state.get("sinopsis_ai_text", "")
                            )

                            book_bytes = generate_book_bytes(
                                mm_template_file,
                                meta={
                                    "nama_penulis": mm_nama_penulis,
                                    "judul_naskah": mm_judul_naskah,
                                    "isbn": mm_isbn,
                                    "tahun_cetak": mm_tahun_cetak,
                                },
                                chapter_title=chapter_title,
                                kata_pengantar_text=mm_kata_pengantar,
                                naskah_text=naskah_text,
                                format_config=format_config,
                                source_docx_bytes=source_docx_bytes or st.session_state.formatted_docx_bytes,
                                qrcbn=mm_qrcbn,
                                sinopsis_text=final_sinopsis_text,
                                tentang_penulis_text=mm_tentang_penulis,
                            )
                            st.session_state["generated_book_bytes"] = book_bytes
                            st.session_state["generated_book_name"] = f"{mm_judul_naskah.strip() or 'naskah'}_final.docx"
                            st.success("Dokumen final berhasil disiapkan. Silakan unduh hasilnya di bawah.")
                        except MailMergeError as e:
                            st.error(f"Template tidak sesuai: {e}")
                        except Exception as e:
                            st.error(f"Gagal menyiapkan dokumen: {e}")

            if "generated_book_bytes" in st.session_state and st.session_state.get("generated_book_bytes") is not None:
                st.download_button(
                    "⬇️ Unduh Dokumen Final",
                    data=st.session_state["generated_book_bytes"],
                    file_name=st.session_state.get("generated_book_name", "dokumen_final.docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
elif st.session_state.page == "result" and st.session_state.result:
    st.button("← Kembali ke Dashboard", on_click=back_to_dashboard)
    res = st.session_state.result

    st.markdown("### Hasil Analisis")
    st.markdown(
        "<div class='aig-sub'>Review tiap opsi blurb, pilih yang paling sesuai, lalu simpan hasilnya.</div>",
        unsafe_allow_html=True,
    )

    col_a, col_b = st.columns([0.9, 1.1], gap="large")

    with col_a:
        summary_html = (
            "<div class='aig-card'>"
            "<div class='aig-pill'>Summary</div>"
            f"<div class='aig-label'>Judul</div><div class='aig-value'>{res['title']}</div>"
            f"<div class='aig-label'>Jumlah Kata</div><div class='aig-value'>{res['word_count']:,}</div>"
            f"<div class='aig-label'>Klasifikasi Genre</div><div class='aig-value'>{res['genre']}</div>"
            "</div>"
        )
        summary_html = summary_html.replace(",", ".")
        st.markdown(summary_html, unsafe_allow_html=True)

    with col_b:
        st.markdown(
            "<div class='aig-card'>"
            "<div class='aig-pill'>Generated Blurb</div>"
            "</div>",
            unsafe_allow_html=True,
        )
        tabs = st.tabs([f"Opsi {i+1}" for i in range(len(res["blurbs"]))])
        for i, tab in enumerate(tabs):
            with tab:
                edited = st.text_area(
                    f"blurb_{i}",
                    value=res["blurbs"][i],
                    height=140,
                    label_visibility="collapsed",
                    key=f"blurb_edit_{res['record_id']}_{i}",
                )

                if st.session_state.active_blurb == i:
                    st.success("Blurb ini sudah tersimpan.")
                else:
                    st.caption("Belum disimpan ke database.")

                st.markdown("**Copy Teks**")
                st.code(edited, language=None)

                if st.button(
                    "💾 Simpan Blurb Ini",
                    key=f"save_{i}",
                    type="primary",
                    use_container_width=True,
                ):
                    res["blurbs"][i] = edited
                    update_chosen_blurb(res["record_id"], i)
                    st.session_state.active_blurb = i
                    st.success("Blurb tersimpan ke database.")

st.markdown(
    "<div class='aig-footnote'>Guepedia Editor Assistant 2026</div>",
    unsafe_allow_html=True,
)
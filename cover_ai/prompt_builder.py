from pathlib import Path

BASE_DIR = Path(__file__).parent

PROMPT_FILES = {
    "fiksi": BASE_DIR / "prompt_fiksi.txt",
    "nonfiksi": BASE_DIR / "prompt_nonfiksi.txt",
}


def hitung_spine(jumlah_halaman: int) -> float:
    """
    Menghitung ketebalan spine (mm)
    Rumus Guepedia:
    0.058 x jumlah halaman
    """
    return round(jumlah_halaman * 0.058, 2)


def load_prompt(jenis_buku: str) -> str:
    """
    Membaca template prompt txt
    """
    jenis = jenis_buku.lower()

    if jenis not in PROMPT_FILES:
        raise ValueError(f"Jenis buku '{jenis_buku}' tidak dikenali.")

    return PROMPT_FILES[jenis].read_text(
        encoding="utf-8"
    )


def build_prompt(
    jenis_buku: str,
    judul: str,
    subjudul: str,
    penulis: str,
    kategori: str,
    genre: str,
    sinopsis: str,
    jumlah_halaman: int,
) -> str:
    """
    Menyusun prompt final untuk AI image generator.

    PENTING: teks brief/prompt dari perusahaan (prompt_fiksi.txt /
    prompt_nonfiksi.txt) TIDAK PERNAH diubah atau di-replace sebagian.
    Brief tersebut ditempel apa adanya (verbatim). Data buku yang
    sebenarnya (judul, penulis, sinopsis, dll.) disisipkan sebagai
    blok data terpisah di atas brief, supaya AI tetap menerima
    informasi buku yang benar tanpa satu kata pun dari brief
    perusahaan berubah.
    """

    brief = load_prompt(jenis_buku)

    spine = hitung_spine(jumlah_halaman)

    data_block = (
        "=== DATA BUKU (WAJIB DIGUNAKAN UNTUK COVER INI) ===\n"
        f"Judul Buku: {judul or '(kosong)'}\n"
        f"Sub Judul: {subjudul or '(tidak ada)'}\n"
        f"Nama Penulis: {penulis or '(kosong)'}\n"
        f"Kategori Buku: {kategori or '(kosong)'}\n"
        f"Genre Buku: {genre or '(kosong)'}\n"
        f"Jumlah Halaman: {jumlah_halaman}\n"
        f"Ketebalan Punggung/Spine (hasil hitung, 0,058 x jumlah halaman): {spine} mm\n"
        f"Sinopsis:\n{sinopsis or '(kosong)'}\n"
    )

    catatan_penghubung = (
        "=== CATATAN PENTING UNTUK AI ===\n"
        "Instruksi desain di bawah ini berisi kata seperti '(Sesuaikan)' atau "
        "'(disesuaikan)' pada bagian Judul Buku, Sub Judul, Nama Penulis, Kategori, "
        "Genre, dan Sinopsis. Kata-kata itu BUKAN teks literal yang harus ditampilkan "
        "di cover. Itu adalah instruksi bahwa isian tersebut harus DIGANTI dengan data "
        "buku yang sudah diberikan pada DATA BUKU di atas. Jangan pernah menuliskan "
        "kata 'Sesuaikan' atau 'disesuaikan' di cover manapun.\n"
    )

    instruksi_desain = (
        "=== INSTRUKSI DESAIN COVER DARI PERUSAHAAN (JANGAN DIUBAH) ===\n"
        f"{brief}"
    )

    return f"{data_block}\n{catatan_penghubung}\n{instruksi_desain}"
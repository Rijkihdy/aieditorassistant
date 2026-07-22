"""
Parser bersama untuk naskah mentah: pisahkan front matter (Kata Pengantar &
Daftar Isi buatan penulis sendiri) dari isi bab asli, lalu pecah isi bab
tersebut jadi per-bab/sub-bab.

Dipakai bareng oleh app.py (editor per-bab di UI) dan mailmerge.py (generate
dokumen final) supaya logika pemisahan bab HANYA ada di satu tempat.

Kenapa perlu strip_front_matter():
    Naskah yang masuk ke penerbit hampir selalu sudah menyertakan Daftar Isi
    buatan sendiri, isinya listing "Bab 1 : Judul", "Bab 2 : Judul", dst.
    Kalau detektor bab cuma mencocokkan pola teks "Bab N" tanpa konteks,
    SETIAP baris di listing itu ikut dianggap bab baru -> muncul bab-bab
    "palsu" (isinya cuma judul, tanpa paragraf) di depan bab asli.

    Pembedanya: entri Daftar Isi selalu diikuti oleh entri "Bab N" berikutnya
    dalam jarak pendek (cuma diselingi 1-2 baris sub-poin). Bab asli diikuti
    oleh puluhan-ratusan baris paragraf sebelum "Bab N" berikutnya muncul.
    strip_front_matter() memakai jarak ini untuk menemukan awal bab asli yang
    pertama, dan membuang semua yang sebelum itu (judul buku, Kata Pengantar,
    listing Daftar Isi) karena dokumen final sudah punya slot sendiri untuk
    itu (placeholder Kata Pengantar + field TOC otomatis dari template).
"""
import re

CHAPTER_HEADING_RE = re.compile(r"^Bab\s+\d+", re.IGNORECASE)

# Kalau jarak (jumlah baris non-kosong) dari satu match "Bab N" ke match
# berikutnya < angka ini, dianggap entri Daftar Isi (bukan bab sungguhan).
MIN_REAL_CHAPTER_GAP = 12

# Baris pendek & mayoritas huruf besar tepat setelah "Bab N" polos (tanpa
# titik dua) dianggap sub-judul bab, digabung jadi satu heading.
_SUBTITLE_MAX_LEN = 140


def is_heading_line(line: str) -> bool:
    """True kalau baris ini terlihat seperti judul bab/sub-bab."""
    line = line.strip()
    return bool(CHAPTER_HEADING_RE.match(line) or (line.startswith("[") and line.endswith("]")))


def _looks_like_subtitle(line: str) -> bool:
    letters = [c for c in line if c.isalpha()]
    if not letters:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return len(line) <= _SUBTITLE_MAX_LEN and upper_ratio > 0.8


def strip_front_matter(raw_lines: list[str]) -> tuple[str, list[str]]:
    """
    Pisahkan naskah mentah -> (kata_pengantar_text, body_lines).

    body_lines dimulai dari baris "Bab N" pertama yang benar-benar diikuti
    isi (bukan cuma entri Daftar Isi). Semua yang sebelum itu dibuang.

    Kata Pengantar tetap terdeteksi baik naskah PUNYA Daftar Isi bawaan
    sendiri (diambil sampai baris "DAFTAR ISI") MAUPUN TIDAK PUNYA sama
    sekali (diambil sampai baris bab asli pertama ditemukan).
    """
    kp_idx = next((i for i, l in enumerate(raw_lines) if l.strip().upper() == "KATA PENGANTAR"), None)
    di_idx = next((i for i, l in enumerate(raw_lines) if l.strip().upper() == "DAFTAR ISI"), None)

    search_start = (di_idx + 1) if di_idx is not None else 0

    heading_idxs = [
        i for i in range(search_start, len(raw_lines)) if CHAPTER_HEADING_RE.match(raw_lines[i].strip())
    ]

    body_start = search_start
    for pos, idx in enumerate(heading_idxs):
        next_idx = heading_idxs[pos + 1] if pos + 1 < len(heading_idxs) else len(raw_lines)
        if next_idx - idx >= MIN_REAL_CHAPTER_GAP:
            body_start = idx
            break
    else:
        if heading_idxs:
            body_start = heading_idxs[0]

    kata_pengantar_text = ""
    if kp_idx is not None:
        # Kalau ada Daftar Isi bawaan, Kata Pengantar berhenti di situ.
        # Kalau tidak ada, berhenti tepat sebelum bab asli pertama.
        end_idx = di_idx if di_idx is not None else body_start
        if end_idx > kp_idx:
            kata_pengantar_text = "\n".join(raw_lines[kp_idx + 1 : end_idx]).strip()

    return kata_pengantar_text, raw_lines[body_start:]


def split_into_sections(body_lines: list[str], fallback_title: str = "") -> list[tuple[str, list[str]]]:
    """
    Pecah body_lines (SUDAH bebas dari front matter/Daftar Isi) jadi list
    (judul_bab, baris_isi). Baris "Bab N" polos yang langsung diikuti baris
    pendek huruf-kapital (mis. "BAB   1" lalu "MEMAHAMI KONSEP...") otomatis
    digabung jadi satu judul bab yang lebih rapi.
    """
    sections: list[tuple[str, list[str]]] = []
    current_title = ""
    current_body: list[str] = []
    saw_heading = False

    def flush() -> None:
        nonlocal current_title, current_body
        if current_title or (saw_heading and current_body):
            sections.append((current_title, current_body))
        current_title = ""
        current_body = []

    i = 0
    n = len(body_lines)
    while i < n:
        line = body_lines[i].strip()

        if not line:
            if current_title or current_body:
                current_body.append("")
            i += 1
            continue

        if is_heading_line(line):
            flush()
            saw_heading = True
            current_title = line
            i += 1

            has_colon = ":" in line
            if not has_colon and i < n and _looks_like_subtitle(body_lines[i].strip()):
                current_title = f"{line}: {body_lines[i].strip()}"
                i += 1

            current_body = []
            continue

        if not saw_heading:
            i += 1
            continue

        current_body.append(line)
        i += 1

    if current_title or current_body:
        sections.append((current_title, current_body))

    if not sections:
        return [(fallback_title.strip() or "Bab 1", [line for line in body_lines if line.strip()])]

    if len(sections) == 1 and not is_heading_line(sections[0][0]):
        return [(fallback_title.strip() or sections[0][0] or "Bab 1", sections[0][1])]

    return sections


def parse_naskah(raw_text: str, fallback_title: str = "") -> tuple[str, list[tuple[str, list[str]]]]:
    """Fungsi sekali-panggil: dari teks naskah mentah -> (kata_pengantar, sections)."""
    raw_lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    kata_pengantar, body_lines = strip_front_matter(raw_lines)
    sections = split_into_sections(body_lines, fallback_title=fallback_title)
    return kata_pengantar, sections
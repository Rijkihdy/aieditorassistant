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

# Pola "Bab N" diperluas: angka arab (1, 2, ...), angka romawi (I, II, ...),
# dan bilangan eja umum Bahasa Indonesia (Satu, Dua, ...) — sebelumnya HANYA
# angka arab yang dikenali, sehingga naskah yang menulis "Bab Satu" / "Bab I"
# tidak terdeteksi sebagai bab baru dan isinya tersambung ke bab sebelum/
# sesudahnya.
_ROMAN_NUM = r"[IVXLCDM]+"
_SPELLED_NUM = (
    r"(?:satu|dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|sepuluh|"
    r"sebelas|dua\s*belas|tiga\s*belas|empat\s*belas|lima\s*belas|"
    r"enam\s*belas|tujuh\s*belas|delapan\s*belas|sembilan\s*belas|"
    r"dua\s*puluh)"
)
CHAPTER_HEADING_RE = re.compile(
    rf"^bab\s+(?:\d+|{_ROMAN_NUM}|{_SPELLED_NUM})(?:\b|(?=[.:\-–]))",
    re.IGNORECASE,
)

# Kalau jarak (jumlah baris non-kosong) dari satu match "Bab N" ke match
# berikutnya < angka ini, dianggap entri Daftar Isi (bukan bab sungguhan).
MIN_REAL_CHAPTER_GAP = 12

# Baris pendek & mayoritas huruf besar tepat setelah "Bab N" polos (tanpa
# titik dua) dianggap sub-judul bab, digabung jadi satu heading.
_SUBTITLE_MAX_LEN = 140

# Marker internal (TIDAK PERNAH ditampilkan ke user / masuk ke dokumen akhir)
# yang dipasang oleh extract_docx_structured() pada baris yang terdeteksi
# sebagai judul bab lewat STYLE Word (Heading 1/2/Title) walau teksnya tidak
# cocok pola "Bab N" sama sekali (mis. judul bab cuma "Prolog", "Senja",
# dsb). Marker ini selalu dibuang lagi sebelum judul ditulis ke dokumen.
HEADING_MARKER = "\x01HEADING\x01"

# Marker internal untuk baris yang merepresentasikan sebuah gambar (dipasang
# oleh extract_docx_structured()); nilainya dipakai sebagai key ke dict
# `images` supaya gambar bisa disisipkan lagi di posisi yang sama.
IMAGE_MARKER_PREFIX = "\x02IMAGE:"


def strip_heading_marker(line: str) -> str:
    """Buang HEADING_MARKER dari sebuah baris (aman dipanggil walau tak ber-marker)."""
    return line[len(HEADING_MARKER):] if line.startswith(HEADING_MARKER) else line


def is_image_marker(line: str) -> bool:
    return line.startswith(IMAGE_MARKER_PREFIX)


def is_heading_line(line: str) -> bool:
    """True kalau baris ini terlihat seperti judul bab/sub-bab."""
    line = line.strip()
    if line.startswith(HEADING_MARKER):
        return True
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

    def _is_chapter_boundary(line: str) -> bool:
        line = line.strip()
        return bool(line.startswith(HEADING_MARKER) or CHAPTER_HEADING_RE.match(line))

    heading_idxs = [
        i for i in range(search_start, len(raw_lines)) if _is_chapter_boundary(raw_lines[i])
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
            current_title = strip_heading_marker(line)
            i += 1

            has_colon = ":" in current_title
            next_line = body_lines[i].strip() if i < n else ""
            if (
                not has_colon
                and i < n
                and not is_image_marker(next_line)
                and _looks_like_subtitle(next_line)
            ):
                current_title = f"{current_title}: {next_line}"
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


# Label front-matter yang punya penanganan khusus (pencocokan teks persis) di
# strip_front_matter(); kalau paragraf ber-style Heading kebetulan berisi
# label ini, JANGAN dipasangi HEADING_MARKER supaya pencocokan persis itu
# tetap berfungsi.
_FRONT_MATTER_LABELS = {"KATA PENGANTAR", "DAFTAR ISI"}


def _paragraph_image_blobs(paragraph) -> list[bytes]:
    """Ambil byte semua gambar inline (w:drawing > a:blip) di dalam sebuah paragraf."""
    from docx.oxml.ns import qn

    blobs: list[bytes] = []
    for blip in paragraph._p.findall(".//" + qn("a:blip")):
        rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
        if not rid:
            continue
        try:
            part = paragraph.part.related_parts[rid]
        except KeyError:
            continue
        blobs.append(part.blob)
    return blobs


def _paragraph_heading_text(paragraph) -> str | None:
    """
    Tentukan apakah sebuah paragraf python-docx adalah judul bab/sub-bab.

    Dua sinyal dipakai (salah satu cukup):
      1. STYLE Word paragraf tersebut Heading 1/Heading 2/.../Title — ini
         menangkap kasus penulis MEMANG memformat judul babnya sebagai
         heading di Word, apapun teksnya (tidak harus diawali kata "Bab").
      2. Pola teks klasik "Bab N" (lewat is_heading_line) — tetap dipakai
         untuk naskah yang judul babnya cuma teks biasa tanpa style heading.
    """
    text = paragraph.text.strip()
    if not text:
        return None

    if text.upper() in _FRONT_MATTER_LABELS:
        return text

    style_name = (paragraph.style.name if paragraph.style is not None else "") or ""
    is_style_heading = style_name.startswith("Heading") or style_name == "Title"

    if is_style_heading:
        return HEADING_MARKER + text
    if is_heading_line(text):
        return text
    return None


def extract_docx_structured(doc) -> tuple[str, list[tuple[str, list[str]]], dict[str, bytes]]:
    """
    Versi docx-aware dari parse_naskah(): membaca langsung objek python-docx
    Document (bukan string polos) supaya bisa:
      - Mengenali bab baru dari STYLE Word (Heading 1/2/Title), bukan cuma
        dari pola teks "Bab N" — judul bab dengan format apapun tetap
        terdeteksi sebagai bab baru, bukan tersambung ke bab lain.
      - Mempertahankan gambar yang ada di naskah asli: tiap gambar diberi
        marker unik pada urutan barisnya (dikembalikan lewat dict `images`)
        supaya bisa disisipkan lagi di posisi yang sama saat naskah ditulis
        ulang ke dokumen final, bukan hilang begitu saja.

    Return: (kata_pengantar_text, sections, images)
      images: dict {marker: bytes_gambar} — marker-nya adalah salah satu
              elemen dalam baris_isi (body_lines) tiap section.
    """
    raw_lines: list[str] = []
    images: dict[str, bytes] = {}
    img_counter = 0

    for paragraph in doc.paragraphs:
        for blob in _paragraph_image_blobs(paragraph):
            marker = f"{IMAGE_MARKER_PREFIX}{img_counter}"
            images[marker] = blob
            raw_lines.append(marker)
            img_counter += 1

        text = paragraph.text.strip()
        if not text:
            continue

        heading_text = _paragraph_heading_text(paragraph)
        raw_lines.append(heading_text if heading_text is not None else text)

    kata_pengantar, body_lines = strip_front_matter(raw_lines)
    sections = split_into_sections(body_lines)
    return kata_pengantar, sections, images
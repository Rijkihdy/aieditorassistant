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

# --- Deteksi heading dari listing Daftar Isi buatan penulis sendiri --------
#
# Banyak naskah TIDAK menandai judul babnya dengan kata "Bab" sama sekali
# (mis. "Mengenal Guru", "Peran Guru") DAN tidak memformatnya sebagai style
# Heading Word (cuma paragraf "Normal" biasa). Tanpa referensi tambahan,
# detektor bab (pola teks "Bab N" / style Heading) tidak akan pernah
# menganggap judul semacam itu sebagai batas bab baru -- seluruh naskah
# akan tersambung jadi satu "bab" raksasa.
#
# Naskah semacam ini hampir selalu tetap punya listing Daftar Isi buatan
# sendiri (mis. "3. Mengenal Guru……………………5"). extract_toc_titles()
# mengambil judul-judul bab dari listing itu, lalu is_heading_line() ikut
# mencocokkan baris body terhadap daftar judul ini (exact match setelah
# dinormalisasi) sebagai sinyal tambahan -- di LUAR pola "Bab N" dan style
# Heading yang sudah ada.
_TOC_LINE_RE = re.compile(r"^\d+\.\s*(.+)$")
# Buang ekor "dot leader" (....../……) + nomor halaman di ujung entri Daftar
# Isi, supaya yang tersisa cuma judul babnya sendiri.
_TOC_TRAILING_RE = re.compile(r"[.\u2026\s]{2,}\d*\s*$")

# Label listing Daftar Isi yang MEMANG bukan bab sungguhan (sudah punya
# penanganan sendiri, atau termasuk back-matter di luar cakupan pemisahan
# bab ini) -- jangan ikut dijadikan calon judul bab.
_TOC_NON_CHAPTER_LABELS = {
    "kata pengantar",
    "daftar isi",
    "daftar pustaka",
    "profil penulis",
    "tentang penulis",
    "biodata penulis",
    "sinopsis",
}

# Batas toleransi baris berturut-turut yang TIDAK cocok pola listing sebelum
# extract_toc_titles() berhenti scan (menandakan listing Daftar Isi sudah
# berakhir, bukan cuma diselingi baris ganjil).
_TOC_SCAN_MISS_TOLERANCE = 2


def _normalize_heading_text(text: str) -> str:
    """Normalisasi teks judul untuk pencocokan longgar (lowercase, spasi
    dirapikan, tanda baca di ujung dibuang) -- dipakai supaya judul di
    listing Daftar Isi ("Mengenal Guru") tetap cocok dengan judul yang
    sama persis di body walau beda kapitalisasi/spasi kecil."""
    text = re.sub(r"\s+", " ", text.strip().lower())
    return text.strip(" .:-–")


def _find_label_idx(raw_lines: list[str], label: str) -> int | None:
    """Cari index baris yang PERSIS sama (case-insensitive) dengan sebuah
    label front-matter (mis. "KATA PENGANTAR", "DAFTAR ISI")."""
    return next((i for i, l in enumerate(raw_lines) if l.strip().upper() == label), None)


def extract_toc_titles(raw_lines: list[str]) -> set[str]:
    """Ambil judul-judul bab dari listing Daftar Isi buatan penulis sendiri
    (kalau ada), dikembalikan sebagai set judul yang SUDAH dinormalisasi
    (lihat _normalize_heading_text) supaya langsung bisa dipakai untuk
    exact-match terhadap baris body ternormalisasi juga.

    Contoh baris yang dikenali: "3. Mengenal Guru……………………5" ->
    judul "Mengenal Guru" (dot leader + nomor halaman di ujung dibuang).

    Baris berlabel back-matter (Daftar Pustaka, Profil/Tentang Penulis, dst
    -- lihat _TOC_NON_CHAPTER_LABELS) sengaja TIDAK ikut, karena bukan bab
    dan sudah/akan punya penanganan sendiri di luar sini.
    """
    di_idx = _find_label_idx(raw_lines, "DAFTAR ISI")
    if di_idx is None:
        return set()

    titles: set[str] = set()
    misses = 0
    for line in raw_lines[di_idx + 1 :]:
        m = _TOC_LINE_RE.match(line.strip())
        if not m:
            misses += 1
            if misses > _TOC_SCAN_MISS_TOLERANCE:
                break
            continue
        misses = 0
        title = _TOC_TRAILING_RE.sub("", m.group(1)).strip(" .-–")
        normalized = _normalize_heading_text(title)
        if normalized and normalized not in _TOC_NON_CHAPTER_LABELS:
            titles.add(normalized)
    return titles

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


def is_heading_line(line: str, toc_titles: set[str] | None = None) -> bool:
    """True kalau baris ini terlihat seperti judul bab/sub-bab.

    toc_titles (opsional): set judul ternormalisasi hasil extract_toc_titles().
    Kalau diisi, baris yang PERSIS cocok (setelah dinormalisasi) dengan salah
    satu judul di listing Daftar Isi juga dianggap heading -- ini menangkap
    naskah yang judul babnya tidak diawali kata "Bab" dan tidak diformat
    sebagai style Heading Word sama sekali (lihat catatan di dekat
    extract_toc_titles())."""
    line = line.strip()
    if line.startswith(HEADING_MARKER):
        return True
    if CHAPTER_HEADING_RE.match(line) or (line.startswith("[") and line.endswith("]")):
        return True
    if toc_titles and _normalize_heading_text(line) in toc_titles:
        return True
    return False


def _looks_like_subtitle(line: str) -> bool:
    letters = [c for c in line if c.isalpha()]
    if not letters:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return len(line) <= _SUBTITLE_MAX_LEN and upper_ratio > 0.8


def _clean_front_matter_lines(lines: list[str]) -> list[str]:
    """Bersihkan baris-baris yang akan digabung jadi TEKS POLOS Kata Pengantar
    (bukan diproses lagi lewat split_into_sections/_insert_manuscript_sections,
    yang masing-masing sudah tahu cara menangani marker internal).

    Tanpa ini, dua marker internal bisa ikut lolos mentah-mentah ke dalam teks
    Kata Pengantar kalau kebetulan ada di rentang KATA PENGANTAR..DAFTAR ISI
    (atau KATA PENGANTAR..bab asli pertama, kalau naskah tidak punya Daftar
    Isi bawaan):
      - HEADING_MARKER ('\\x01HEADING\\x01'): dipasang extract_docx_structured()
        di paragraf ber-style Heading Word apapun teksnya -- kalau penulis
        kebetulan memformat salah satu baris di Kata Pengantar sebagai
        Heading, marker ini ikut kebawa.
      - IMAGE_MARKER_PREFIX ('\\x02IMAGE:...'): dipasang untuk gambar inline;
        Kata Pengantar tidak punya slot untuk menyisipkan gambar balik, jadi
        baris marker gambar di sini dibuang saja (bukan cuma di-strip).

    Karakter kontrol \\x01/\\x02 yang lolos ke paragraph.text akan bikin
    python-docx menolak dengan error "All strings must be XML compatible:
    ... no NULL bytes or control characters" saat dokumen ditulis.
    """
    cleaned: list[str] = []
    for line in lines:
        if is_image_marker(line):
            continue
        cleaned.append(strip_heading_marker(line))
    return cleaned


def strip_front_matter(
    raw_lines: list[str], toc_titles: set[str] | None = None
) -> tuple[str, list[str]]:
    """
    Pisahkan naskah mentah -> (kata_pengantar_text, body_lines).

    body_lines dimulai dari baris heading pertama yang benar-benar diikuti
    isi (bukan cuma entri Daftar Isi). Semua yang sebelum itu dibuang.
    "Heading pertama" di sini bisa berupa pola "Bab N" MAUPUN judul persis
    dari listing Daftar Isi (lihat extract_toc_titles) untuk naskah yang
    judul babnya tidak diawali kata "Bab" sama sekali.

    toc_titles (opsional): kalau tidak diberikan, dihitung otomatis dari
    raw_lines lewat extract_toc_titles() -- jadi caller lama yang belum
    tahu soal parameter ini tetap dapat manfaatnya tanpa perlu diubah.

    Kata Pengantar tetap terdeteksi baik naskah PUNYA Daftar Isi bawaan
    sendiri (diambil sampai baris "DAFTAR ISI") MAUPUN TIDAK PUNYA sama
    sekali (diambil sampai baris bab asli pertama ditemukan).
    """
    if toc_titles is None:
        toc_titles = extract_toc_titles(raw_lines)

    kp_idx = _find_label_idx(raw_lines, "KATA PENGANTAR")
    di_idx = _find_label_idx(raw_lines, "DAFTAR ISI")

    search_start = (di_idx + 1) if di_idx is not None else 0

    def _is_toc_title_match(line: str) -> bool:
        # strip_heading_marker() WAJIB dulu: baris yang cocok toc_titles di
        # _paragraph_heading_text() (jalur docx) sudah dipasangi
        # HEADING_MARKER juga (sama seperti style-heading biasa) supaya
        # is_heading_line() langsung mengenalinya -- tanpa dilepas dulu di
        # sini, perbandingan ke toc_titles selalu gagal (markernya ikut
        # kebanding) dan toc_match_idxs jadi kosong padahal ada match.
        line = strip_heading_marker(line.strip())
        return bool(toc_titles and _normalize_heading_text(line) in toc_titles)

    def _is_chapter_boundary(line: str) -> bool:
        line = line.strip()
        return bool(
            line.startswith(HEADING_MARKER)
            or CHAPTER_HEADING_RE.match(line)
            or _is_toc_title_match(line)
        )

    heading_idxs = [
        i for i in range(search_start, len(raw_lines)) if _is_chapter_boundary(raw_lines[i])
    ]

    # Match lewat toc_titles SUDAH DIKONFIRMASI cocok persis dengan listing
    # Daftar Isi (bukan cuma pola ambigu "Bab N" yang bisa juga match baris
    # listing itu sendiri) -- jadi TIDAK PERLU lewat heuristik jarak
    # MIN_REAL_CHAPTER_GAP di bawah. Tanpa pengecualian ini, dua bab ASLI
    # yang kebetulan pendek dan berurutan (mis. "Mengenal Guru" langsung
    # diikuti "Peran Guru" tanpa jeda panjang) akan salah dianggap entri
    # listing Daftar Isi dan malah DIBUANG dari body -- padahal keduanya
    # bab sungguhan.
    toc_match_idxs = [i for i in heading_idxs if _is_toc_title_match(raw_lines[i])]

    if toc_match_idxs:
        body_start = toc_match_idxs[0]
    else:
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
            kata_pengantar_text = "\n".join(
                _clean_front_matter_lines(raw_lines[kp_idx + 1 : end_idx])
            ).strip()

    return kata_pengantar_text, raw_lines[body_start:]


def split_into_sections(
    body_lines: list[str], fallback_title: str = "", toc_titles: set[str] | None = None
) -> list[tuple[str, list[str]]]:
    """
    Pecah body_lines (SUDAH bebas dari front matter/Daftar Isi) jadi list
    (judul_bab, baris_isi). Baris "Bab N" polos yang langsung diikuti baris
    pendek huruf-kapital (mis. "BAB   1" lalu "MEMAHAMI KONSEP...") otomatis
    digabung jadi satu judul bab yang lebih rapi.

    toc_titles (opsional): set judul ternormalisasi hasil extract_toc_titles(),
    dipakai sebagai sinyal tambahan di is_heading_line() untuk naskah yang
    judul babnya tidak diawali kata "Bab" dan tidak berstyle Heading Word.
    body_lines di sini sudah tidak punya baris "DAFTAR ISI" lagi (sudah
    dibuang strip_front_matter), jadi toc_titles TIDAK bisa dihitung ulang
    dari body_lines saja -- caller yang mau manfaat ini WAJIB hitung
    lewat extract_toc_titles(raw_lines) sebelum front matter dibuang, lalu
    oper ke sini (lihat parse_naskah / extract_docx_structured).
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

        if is_heading_line(line, toc_titles):
            flush()
            saw_heading = True
            current_title = strip_heading_marker(line)
            i += 1

            has_colon = ":" in current_title
            next_line_raw = body_lines[i].strip() if i < n else ""
            # PENTING: next_line_raw bisa jadi baris heading LAIN (mis.
            # sub-judul "BAGIAN III" yang juga di-style Heading terpisah di
            # Word) -- kalau begitu dia sudah bawa HEADING_MARKER sendiri.
            # Marker itu WAJIB dilepas SEBELUM digabung ke current_title,
            # kalau tidak, karakter kontrol \x01 di dalamnya ikut nempel ke
            # judul bab final (mis. jadi "BAGIAN III: \x01HEADING\x01Sub
            # Judul") dan bikin python-docx menolak saat document.save()
            # dengan error "All strings must be XML compatible ... no NULL
            # bytes or control characters".
            next_line = strip_heading_marker(next_line_raw)
            if (
                not has_colon
                and i < n
                and not is_image_marker(next_line_raw)
                and not is_heading_line(next_line, toc_titles)
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

    if len(sections) == 1 and not is_heading_line(sections[0][0], toc_titles):
        return [(fallback_title.strip() or sections[0][0] or "Bab 1", sections[0][1])]

    return sections


def parse_naskah(raw_text: str, fallback_title: str = "") -> tuple[str, list[tuple[str, list[str]]]]:
    """Fungsi sekali-panggil: dari teks naskah mentah -> (kata_pengantar, sections)."""
    raw_lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    toc_titles = extract_toc_titles(raw_lines)
    kata_pengantar, body_lines = strip_front_matter(raw_lines, toc_titles)
    sections = split_into_sections(body_lines, fallback_title=fallback_title, toc_titles=toc_titles)
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


# Judul bab/sub-bab asli selalu pendek (nama bab, bukan kalimat utuh).
# Kalau paragraf ber-style Heading ternyata sepanjang ini atau punya banyak
# tanda titik (ciri paragraf naratif biasa, bukan judul), style-nya
# kemungkinan besar salah ketik/salah pilih oleh penulis di Word -- JANGAN
# dipromosikan jadi bab baru walau style-nya "Heading". Tanpa pengaman ini,
# satu paragraf isi yang kebetulan ke-style Heading akan: (1) dianggap bab
# baru sendiri lengkap dengan page break, DAN (2) ikut tersedot ke field
# Daftar Isi otomatis Word (yang men-scan semua Heading 1-3), muncul sebagai
# entri TOC yang panjang/aneh dan bikin nomor halamannya dobel/wrap.
_STYLE_HEADING_MAX_LEN = 120
_STYLE_HEADING_MAX_SENTENCES = 1


def _looks_like_real_heading_text(text: str) -> bool:
    """True kalau teks ini pantas jadi judul bab/sub-bab (pendek, bukan
    paragraf naratif utuh dengan banyak kalimat)."""
    if len(text) > _STYLE_HEADING_MAX_LEN:
        return False
    # Hitung tanda akhir kalimat ('.', '!', '?') -- judul asli biasanya tidak
    # punya titik sama sekali, atau paling banyak satu di ujung.
    sentence_enders = text.count(".") + text.count("!") + text.count("?")
    if sentence_enders > _STYLE_HEADING_MAX_SENTENCES:
        return False
    return True


def _paragraph_heading_text(paragraph, toc_titles: set[str] | None = None) -> str | None:
    """
    Tentukan apakah sebuah paragraf python-docx adalah judul bab/sub-bab.

    Tiga sinyal dipakai (salah satu cukup):
      1. STYLE Word paragraf tersebut Heading 1/Heading 2/.../Title — ini
         menangkap kasus penulis MEMANG memformat judul babnya sebagai
         heading di Word, apapun teksnya (tidak harus diawali kata "Bab").
         TAPI hanya dipercaya kalau teksnya memang terlihat seperti judul
         (pendek, bukan paragraf naratif) -- lihat _looks_like_real_heading_text.
      2. Pola teks klasik "Bab N" (lewat is_heading_line) — tetap dipakai
         untuk naskah yang judul babnya cuma teks biasa tanpa style heading.
      3. Cocok persis (setelah dinormalisasi) dengan salah satu judul di
         listing Daftar Isi buatan penulis (toc_titles, lihat
         extract_toc_titles()) — menangkap naskah yang judul babnya BUKAN
         "Bab N" DAN cuma paragraf style "Normal" biasa (mis. "Mengenal
         Guru", "Peran Guru"), yang tanpa sinyal ini tidak akan pernah
         terdeteksi sebagai batas bab sama sekali.
    """
    text = paragraph.text.strip()
    if not text:
        return None

    if text.upper() in _FRONT_MATTER_LABELS:
        return text

    style_name = (paragraph.style.name if paragraph.style is not None else "") or ""
    is_style_heading = style_name.startswith("Heading") or style_name == "Title"

    if is_style_heading and _looks_like_real_heading_text(text):
        return HEADING_MARKER + text
    if is_heading_line(text):
        return text
    if toc_titles and _normalize_heading_text(text) in toc_titles:
        return HEADING_MARKER + text
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
    # Hitung dulu judul-judul dari listing Daftar Isi (kalau ada) dari teks
    # paragraf polos -- WAJIB sebelum loop utama di bawah, karena dipakai
    # sebagai sinyal tambahan supaya _paragraph_heading_text() ikut mengenali
    # judul bab yang cuma paragraf style "Normal" biasa (tidak diawali "Bab"
    # dan tidak ber-style Heading Word sama sekali). Lihat extract_toc_titles().
    toc_titles = extract_toc_titles([p.text.strip() for p in doc.paragraphs])

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

        heading_text = _paragraph_heading_text(paragraph, toc_titles)
        raw_lines.append(heading_text if heading_text is not None else text)

    kata_pengantar, body_lines = strip_front_matter(raw_lines, toc_titles)
    sections = split_into_sections(body_lines, toc_titles=toc_titles)
    return kata_pengantar, sections, images
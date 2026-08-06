"""
Parser naskah berbasis AI (Groq) -- PELENGKAP untuk naskah_parser.py.

KENAPA PERLU INI
-----------------
naskah_parser.py (versi regex/style) cuma bisa mengenali judul bab lewat dua
sinyal: pola teks "Bab N", ATAU style Word Heading 1/2/Title. Banyak naskah
TIDAK cocok dengan sinyal manapun -- contoh nyata yang ditemukan:

  - Naskah "Peran Guru Terhadap Perkembangan Jiwa Anak Pra Sekolah": SEMUA
    paragraf pakai style "Normal" (tidak ada satu pun Heading), dan judul
    babnya cuma frasa pendek biasa ("Mengenal Guru", "Peran Guru", dst) yang
    kebetulan sama persis dengan listing Daftar Isi-nya sendiri. Parser
    regex/style sama sekali tidak bisa mengenali ini -- hasilnya SELURUH
    naskah (termasuk listing Daftar Isi mentah dengan titik-titik & nomor
    halaman) numplek jadi SATU bab raksasa.

Untuk kasus seperti ini, satu-satunya cara yang robust adalah membaca
strukturnya seperti manusia baca buku -- itulah yang dilakukan modul ini.

DESAIN: AI CUMA MENUNJUK LOKASI, TIDAK MENULIS ULANG ISI
----------------------------------------------------------
Supaya HEMAT TOKEN, CEPAT, dan AMAN (tidak ada risiko AI mengubah/memotong
kata-kata asli penulis), AI TIDAK diminta menuliskan ulang isi naskah.
AI cuma dikirim daftar baris naskah yang sudah diberi NOMOR + potongan
pendek tiap baris (cukup untuk membedakan "ini judul pendek" vs "ini
paragraf isi yang panjang" tanpa perlu isi lengkap), lalu diminta menunjuk:
tiap bagian penting (Kata Pengantar, Sinopsis, tiap Bab, Tentang/Profil
Penulis, Daftar Pustaka, atau listing/sampah yang harus dibuang) MULAI di
baris nomor berapa. Isi teks aslinya lengkap tetap diambil langsung dari
baris asli lewat nomor itu -- BUKAN dari jawaban AI.

Bonus dari pendekatan ini: AI langsung mengelompokkan ke KATEGORI SEMANTIK
(mis. "tentang_penulis"), bukan cuma mencocokkan kata "TENTANG PENULIS"
secara harfiah -- jadi otomatis juga menangkap sebutan lain yang penulis
pakai (mis. "PROFIL PENULIS", "BIODATA PENULIS", dsb) tanpa perlu daftar
sinonim manual.

NASKAH PANJANG: DIPECAH JADI BEBERAPA PANGGILAN (CHUNKING)
-------------------------------------------------------------
Untuk naskah yang sangat panjang (banyak bab, ribuan baris), mengirim
SELURUH naskah dalam satu panggilan AI berisiko kena "lost in the middle"
-- LLM cenderung kehilangan fokus di bagian tengah/awal konteks yang
panjang, sehingga bab-bab awal ikut terlewat/tergabung jadi satu (kasus
nyata yang ditemukan: naskah 40 bab, yang berhasil terdeteksi cuma bab
21-40). Untuk menghindari ini, kalau jumlah baris naskah melebihi
`_CHUNK_SIZE`, `detect_structure_with_ai` otomatis memecah naskah jadi
beberapa potongan (chunk) dan memanggil AI SEKALI PER CHUNK secara
berurutan, bukan sekali borong semuanya. Supaya section yang terpotong di
batas chunk tetap nyambung, tiap chunk diberi tahu section terakhir yang
masih terbuka dari chunk sebelumnya (lihat `_build_chunk_user_content` &
`_CHUNK_SYSTEM_PROMPT`), dan tiap chunk juga dikasih beberapa baris
"konteks" (sudah diproses, hanya buat dibaca, TIDAK boleh dijadikan
landmark baru) dari sebelum batas chunk supaya AI tidak salah menandai
paragraf sambungan sebagai bab baru.

CARA PAKAI
----------
    from ai_naskah_parser import detect_structure_with_ai, build_sections_from_structure

    raw_lines = [...]  # baris-baris naskah (boleh berisi IMAGE_MARKER_PREFIX)
    structure = detect_structure_with_ai(raw_lines, api_key=GROQ_API_KEY)
    kata_pengantar, sinopsis, tentang_penulis, sections = build_sections_from_structure(
        raw_lines, structure
    )
    # `sections` punya format PERSIS SAMA seperti hasil naskah_parser.split_into_sections()
    # -- (judul, baris_isi) -- jadi tinggal disambung ke mailmerge.py apa adanya.
"""
from __future__ import annotations

import json
import logging
import os

from naskah_parser import IMAGE_MARKER_PREFIX, is_image_marker

_logger = logging.getLogger(__name__)

_DEFAULT_MODEL = "llama-3.3-70b-versatile"

# Potongan preview per baris yang dikirim ke AI -- cukup panjang untuk
# membedakan judul (pendek, biasanya utuh muncul) vs paragraf isi (panjang,
# akan kepotong "..." -- tapi itu TIDAK APA-APA karena AI cuma perlu tahu
# "ini kelihatannya paragraf naratif panjang", bukan meng-copy isinya).
_PREVIEW_CHARS = 160

# Di atas ambang ini (jumlah baris), naskah dipecah jadi beberapa chunk dan
# AI dipanggil beberapa kali (lihat docstring modul, bagian "NASKAH
# PANJANG"). Angka ini SENGAJA dibuat kecil (bukan cuma buat naskah yang
# "sangat" panjang) -- kegagalan "sebagian bab hilang" ternyata juga bisa
# muncul di naskah yang tidak terlalu panjang (mis. 40 bab tapi tiap bab
# pendek, total masih ratusan baris) begitu daftar landmark yang harus
# dikembalikan AI dalam SATU balasan jadi panjang (puluhan item) -- makin
# sedikit landmark yang diminta per panggilan, makin kecil kemungkinan AI
# "melewatkan" sebagian di antaranya.
_CHUNK_SIZE = 120

# Berapa baris SEBELUM batas chunk yang disertakan sebagai konteks
# baca-saja (bukan buat dilandmark) -- supaya AI tahu apakah awal chunk ini
# masih sambungan section sebelumnya atau betul-betul section baru.
_CONTEXT_OVERLAP_LINES = 15

_VALID_TYPES = {
    "kata_pengantar",
    "sinopsis",
    "tentang_penulis",
    "daftar_pustaka",
    "bab",
    "skip",
}

_SYSTEM_PROMPT = """\
Kamu menganalisis STRUKTUR sebuah naskah buku berbahasa Indonesia untuk \
diterbitkan. Naskah dikirim sebagai daftar baris bernomor (indeks dimulai \
dari 0), tiap baris adalah SATU paragraf (bisa berupa judul pendek ATAU \
paragraf isi yang panjang -- yang panjang dipotong dengan "..." di ujung, \
itu normal, ABAIKAN saja, kamu tidak perlu tahu isi lengkapnya).

TUGASMU: kembalikan JSON dengan TIGA hal:

1. "judul_buku" -- judul buku ini (biasanya ada di halaman pertama/cover,
   baris-baris paling awal naskah). Kosongkan ("") kalau benar-benar tidak
   ketemu, JANGAN mengarang.
2. "nama_penulis" -- nama penulis buku ini (biasanya di halaman
   pertama juga, sering ditandai "Oleh:", "Penulis:", atau cuma nama polos
   di bawah judul; JANGAN diambil dari isi "Tentang Penulis"/"Profil
   Penulis" kalau nama itu beda dari yang di halaman depan -- utamakan
   nama di halaman depan). Kosongkan ("") kalau tidak ketemu.
3. "landmarks" -- daftar titik-titik di mana sebuah bagian baru dimulai,
   dalam urutan kemunculan di naskah. Setiap landmark mewajibkan field:
  - "start": nomor baris (integer) tempat bagian ini MULAI (baris judul/\
label bagian itu sendiri, jika ada, atau baris konten pertama bagian itu \
jika tidak ada judul eksplisit)
  - "type": salah satu dari:
      "kata_pengantar"   -- kata pengantar/prakata penulis
      "sinopsis"         -- sinopsis/ringkasan buku
      "tentang_penulis"  -- biodata/profil penulis (apapun sebutannya: \
"Tentang Penulis", "Profil Penulis", "Biodata Penulis", dst)
      "daftar_pustaka"   -- daftar pustaka/referensi/bibliografi
      "bab"              -- satu bab/bagian isi buku (WAJIB isi field \
"title" dengan judul bab APA ADANYA dari naskah)
      "skip"             -- BUKAN bagian buku yang harus muncul di hasil \
akhir: contoh halaman judul/cover, listing Daftar Isi (baris-baris \
bernomor dengan titik-titik/nomor halaman), listing daftar isi yang \
mengulang label "Kata Pengantar"/"Sinopsis"/dst sebagai entri daftar \
(bukan section sungguhan), atau pengulangan bagian yang sudah pernah \
muncul sebelumnya
  - "title": HANYA untuk type "bab" -- judul bab persis seperti tertulis \
di naskah (jangan diringkas/diubah/diterjemahkan)

ATURAN PENTING UNTUK "landmarks":
- Urutkan landmark sesuai urutan kemunculan aslinya di naskah (start makin \
besar).
- SETIAP baris naskah otomatis jadi milik landmark TERAKHIR sebelum baris \
itu (sampai landmark berikutnya) -- jadi kamu TIDAK perlu menandai baris \
"end", cukup baris "start" tiap landmark baru.
- Landmark PERTAMA sebaiknya start=0 (atau index baris konten pertama) \
supaya tidak ada baris di awal yang tidak masuk kategori manapun. Kalau \
naskah diawali halaman judul/cover/nama penulis, tandai itu sebagai "skip" \
mulai dari baris 0 (tapi tetap baca isinya untuk mengisi "judul_buku" & \
"nama_penulis" di atas).
- Naskah sering menaruh listing Daftar Isi (mis. "3. Mengenal Guru...5") \
DI TENGAH naskah, bukan cuma di depan -- tandai SELURUH listing itu sebagai \
SATU landmark "skip", lalu landmark berikutnya (bab/section sungguhan yang \
pertama) dimulai lagi setelah listing itu selesai.
- Kalau ada bagian yang MUNCUL DUA KALI (mis. Sinopsis ditulis ulang di \
listing Daftar Isi, atau judul bab disebut ulang di halaman pembatas \
bagian/part), landmark KEDUA yang isinya cuma pengulangan singkat tanpa \
konten baru boleh ditandai "skip" -- gunakan konteks (baris pendek tanpa \
paragraf isi sesudahnya = kemungkinan cuma listing/pengulangan, baris \
diikuti banyak paragraf isi = section sungguhan).

Balas HANYA dengan JSON valid, tanpa teks lain, dengan format persis:
{"judul_buku": "...", "nama_penulis": "...", "landmarks": [{"start": 0, "type": "...", "title": "..."}, ...]}
"""

# Versi _SYSTEM_PROMPT khusus buat mode CHUNK (naskah panjang yang dipecah
# jadi beberapa potongan -- lihat docstring modul, bagian "NASKAH PANJANG").
# Bedanya dari _SYSTEM_PROMPT biasa:
#   - Dijelaskan bahwa yang dikirim cuma SEPOTONG naskah, bukan seluruhnya.
#   - Ada dua kelompok baris: "KONTEKS" (baca-saja, sudah diproses chunk
#     sebelumnya, JANGAN dilandmark) dan "CHUNK" (baris yang harus dianalisis
#     & boleh diberi landmark).
#   - Diberi tahu section apa yang masih terbuka dari chunk sebelumnya, jadi
#     kalau chunk ini diawali sambungan paragraf section itu, JANGAN dibuat
#     landmark baru untuk itu.
#   - Aturan "landmark pertama harus start=0" DIHAPUS (karena start=0 tiap
#     chunk sama sekali tidak berarti section baru).
_CHUNK_SYSTEM_PROMPT = """\
Kamu menganalisis STRUKTUR sebuah naskah buku berbahasa Indonesia untuk \
diterbitkan. Naskah ini PANJANG, jadi dikirim ke kamu SEPOTONG-SEPOTONG \
(beberapa kali panggilan terpisah) -- yang kamu terima sekarang HANYA satu \
potongan (chunk) dari keseluruhan naskah, BUKAN naskah lengkap.

Baris dikirim sebagai daftar bernomor (nomor baris GLOBAL, mengacu ke \
posisi aslinya di keseluruhan naskah, bukan direset per-chunk). Tiap baris \
adalah SATU paragraf (bisa judul pendek ATAU paragraf isi panjang -- yang \
panjang dipotong "..." di ujung, ABAIKAN saja, tidak perlu tahu isi \
lengkapnya).

Baris-baris itu dibagi jadi DUA kelompok, ditandai jelas di bawah:
  - "=== KONTEKS (sudah diproses, HANYA UNTUK DIBACA) ===" -- baris SEBELUM \
chunk ini, sudah ditangani oleh panggilan sebelumnya. JANGAN buat landmark \
apapun dengan "start" di rentang ini, ini cuma buat kamu paham kesinambungan.
  - "=== CHUNK (analisis & tandai landmark di sini) ===" -- baris yang \
harus kamu analisis. SEMUA landmark yang kamu hasilkan WAJIB punya "start" \
di dalam rentang CHUNK ini saja.

TUGASMU: kembalikan JSON dengan TIGA hal:

1. "judul_buku" -- judul buku, HANYA isi kalau kamu benar-benar melihatnya \
di baris chunk/konteks kali ini (biasanya cuma kelihatan di chunk paling \
awal naskah, halaman pertama/cover). Kosongkan ("") kalau tidak ketemu di \
potongan ini, JANGAN mengarang.
2. "nama_penulis" -- sama seperti di atas tapi buat nama penulis (biasanya \
ditandai "Oleh:"/"Penulis:" di halaman pertama). Kosongkan ("") kalau tidak \
ketemu di potongan ini.
3. "landmarks" -- daftar titik-titik di mana sebuah bagian BARU dimulai DI \
DALAM CHUNK ini saja, urut sesuai kemunculan. Setiap landmark:
  - "start": nomor baris GLOBAL (integer, WAJIB di dalam rentang CHUNK)
  - "type": salah satu dari "kata_pengantar", "sinopsis", "tentang_penulis" \
(apapun sebutannya: Tentang/Profil/Biodata Penulis), "daftar_pustaka", \
"bab" (WAJIB isi "title" persis seperti tertulis, jangan diringkas/diubah), \
atau "skip" (halaman judul/cover, listing Daftar Isi, atau pengulangan \
section yang sudah pernah muncul)
  - "title": HANYA untuk type "bab"

INFO PENTING soal kesinambungan dari chunk SEBELUMNYA:
{previous_section_info}

ATURAN PALING PENTING:
- Kalau baris PERTAMA di CHUNK ini cuma lanjutan paragraf isi dari section \
yang disebut di atas (bukan judul/heading baru), JANGAN buat landmark untuk \
itu sama sekali -- baris itu otomatis tetap dianggap milik section \
sebelumnya. Landmark baru HANYA dibuat kalau kamu betul-betul melihat \
judul/heading section BARU muncul di dalam rentang CHUNK.
- Kalau seluruh isi CHUNK ini ternyata masih sambungan section yang sama \
dari sebelumnya (tidak ada heading baru sama sekali di chunk ini), balas \
"landmarks": [] (array kosong) -- itu jawaban yang valid dan wajar.
- Naskah sering menaruh listing Daftar Isi DI TENGAH naskah -- kalau ada \
listing seperti itu MUNCUL DI DALAM chunk ini, tandai sebagai satu landmark \
"skip".
- Jangan pernah membuat landmark dengan "start" yang menunjuk ke baris di \
kelompok KONTEKS.

Balas HANYA dengan JSON valid, tanpa teks lain, dengan format persis:
{"judul_buku": "...", "nama_penulis": "...", "landmarks": [{"start": 0, "type": "...", "title": "..."}, ...]}
"""


def _build_numbered_preview(raw_lines: list[str], start: int = 0, end: int | None = None) -> str:
    """Bikin preview bernomor untuk raw_lines[start:end], tapi nomor barisnya
    tetap nomor GLOBAL (indeks asli di raw_lines) -- bukan direset dari 0 --
    supaya bisa dipakai juga buat kirim sepotong (chunk) naskah dan hasilnya
    (landmark "start") tetap nyambung/valid terhadap raw_lines aslinya.
    """
    if end is None:
        end = len(raw_lines)
    lines = []
    for i in range(start, end):
        line = raw_lines[i]
        if is_image_marker(line):
            lines.append(f"{i}: [GAMBAR]")
            continue
        preview = line if len(line) <= _PREVIEW_CHARS else line[:_PREVIEW_CHARS] + "…"
        lines.append(f"{i}: {preview}")
    return "\n".join(lines)


def _describe_previous_section(last_landmark: dict | None) -> str:
    """Deskripsi singkat (buat disisipkan ke prompt) tentang section terakhir
    yang masih terbuka dari chunk-chunk sebelumnya, supaya AI tahu apakah
    awal chunk sekarang itu sambungan atau betul-betul section baru.
    """
    if last_landmark is None:
        return (
            "Belum ada section manapun yang terdeteksi sebelum chunk ini -- "
            "ini kemungkinan bagian PALING AWAL naskah (mis. halaman "
            "judul/cover, atau section pertama yang belum pernah muncul)."
        )
    lm_type = last_landmark.get("type")
    if lm_type == "bab":
        judul = last_landmark.get("title", "")
        return (
            f'Section yang masih terbuka dari chunk sebelumnya: BAB berjudul '
            f'"{judul}". Kalau awal chunk ini masih paragraf isi bab '
            f'tersebut (belum ada judul bab baru), JANGAN buat landmark baru.'
        )
    return (
        f'Section yang masih terbuka dari chunk sebelumnya: type="{lm_type}". '
        f'Kalau awal chunk ini masih lanjutan isi section tersebut (belum ada '
        f'heading baru), JANGAN buat landmark baru.'
    )


def _build_chunk_user_content(
    raw_lines: list[str],
    context_start: int,
    chunk_start: int,
    chunk_end: int,
    last_landmark: dict | None,
) -> str:
    parts = []
    if context_start < chunk_start:
        parts.append("=== KONTEKS (sudah diproses, HANYA UNTUK DIBACA) ===")
        parts.append(_build_numbered_preview(raw_lines, context_start, chunk_start))
    parts.append("=== CHUNK (analisis & tandai landmark di sini) ===")
    parts.append(_build_numbered_preview(raw_lines, chunk_start, chunk_end))
    return "\n".join(parts)


def _call_groq_json(client, model: str, system_prompt: str, user_content: str) -> dict:
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content},
        ],
        response_format={"type": "json_object"},
        temperature=0,
        # Kasih ruang keluaran yang lega -- kalau ini kekecilan, balasan JSON
        # (daftar landmark) bisa TERPOTONG di tengah tanpa error yang jelas,
        # dan hasilnya persis kelihatan seperti "sebagian bab hilang" (JSON
        # jadi tidak valid lalu di-drop begitu saja oleh _call_groq_json).
        max_tokens=8000,
    )
    raw_json = response.choices[0].message.content
    try:
        parsed = json.loads(raw_json)
    except (json.JSONDecodeError, TypeError):
        _logger.warning(
            "ai_naskah_parser: balasan AI GAGAL di-parse sebagai JSON (kemungkinan "
            "terpotong/melebihi batas token keluaran). Cuplikan balasan: %r",
            (raw_json or "")[:300],
        )
        return {}
    return parsed if isinstance(parsed, dict) else {}


def detect_structure_with_ai(
    raw_lines: list[str],
    api_key: str | None = None,
    model: str = _DEFAULT_MODEL,
    chunk_size: int = _CHUNK_SIZE,
) -> dict:
    """Kirim naskah (sebagai preview bernomor) ke Groq, minta kembali daftar
    landmark struktur. Tidak menyentuh isi naskah asli sama sekali -- cuma
    baca balikan AI berupa INDEKS, validasi, lalu dikembalikan mentah-mentah
    (pemrosesan jadi `sections` dilakukan terpisah di
    `build_sections_from_structure`, supaya gampang di-unit-test tanpa
    perlu memanggil AI beneran).

    Kalau naskah lebih panjang dari `chunk_size` baris, otomatis dipecah
    jadi beberapa potongan dan AI dipanggil beberapa kali secara berurutan
    (lihat docstring modul, bagian "NASKAH PANJANG") supaya bab-bab di
    awal/tengah naskah panjang tidak ikut terlewat.
    """
    # Import di dalam fungsi supaya modul ini tetap bisa di-import & di-test
    # (mis. `build_sections_from_structure`) di lingkungan yang belum
    # menginstall `groq` / belum ada API key sama sekali.
    from groq import Groq

    client = Groq(api_key=api_key or os.environ["GROQ_API_KEY"])

    if len(raw_lines) <= chunk_size:
        _logger.info(
            "ai_naskah_parser: %d baris (<= chunk_size=%d) -> SATU panggilan AI, tidak di-chunk.",
            len(raw_lines), chunk_size,
        )
        numbered = _build_numbered_preview(raw_lines)
        structure = _call_groq_json(client, model, _SYSTEM_PROMPT, numbered)
        validated = _validate_structure(structure, total_lines=len(raw_lines))
        _logger.info(
            "ai_naskah_parser: hasil 1 panggilan -> %d landmark (%s)",
            len(validated["landmarks"]),
            [lm["type"] for lm in validated["landmarks"]],
        )
        return validated

    _logger.info(
        "ai_naskah_parser: %d baris (> chunk_size=%d) -> DI-CHUNK jadi %d panggilan AI.",
        len(raw_lines), chunk_size, -(-len(raw_lines) // chunk_size),
    )
    return _detect_structure_chunked(client, raw_lines, model=model, chunk_size=chunk_size)


def _detect_structure_chunked(client, raw_lines: list[str], model: str, chunk_size: int) -> dict:
    """Implementasi mode chunk: panggil AI berkali-kali (satu kali per
    potongan `chunk_size` baris), lalu gabungkan semua landmark jadi satu
    struktur utuh. Dipanggil berurutan (bukan paralel) karena tiap chunk
    perlu tahu section terakhir yang masih terbuka dari chunk sebelumnya
    (`last_landmark`) supaya tidak salah menandai sambungan paragraf
    sebagai bab/section baru.
    """
    n = len(raw_lines)
    all_landmarks: list[dict] = []
    judul_buku = ""
    nama_penulis = ""
    last_landmark: dict | None = None

    chunk_start = 0
    while chunk_start < n:
        chunk_end = min(chunk_start + chunk_size, n)
        context_start = max(0, chunk_start - _CONTEXT_OVERLAP_LINES)

        # .replace() dipakai (bukan str.format) karena prompt ini juga berisi
        # contoh literal JSON dengan kurung kurawal "{...}" yang akan
        # bentrok dengan placeholder ala str.format.
        system_prompt = _CHUNK_SYSTEM_PROMPT.replace(
            "{previous_section_info}", _describe_previous_section(last_landmark)
        )
        user_content = _build_chunk_user_content(
            raw_lines, context_start, chunk_start, chunk_end, last_landmark
        )
        structure = _call_groq_json(client, model, system_prompt, user_content)

        if not judul_buku:
            jb = structure.get("judul_buku")
            if isinstance(jb, str) and jb.strip():
                judul_buku = jb.strip()
        if not nama_penulis:
            np_ = structure.get("nama_penulis")
            if isinstance(np_, str) and np_.strip():
                nama_penulis = np_.strip()

        chunk_landmarks = structure.get("landmarks")
        chunk_landmark_count = 0
        if isinstance(chunk_landmarks, list):
            for lm in chunk_landmarks:
                if not isinstance(lm, dict):
                    continue
                start = lm.get("start")
                lm_type = lm.get("type")
                # Landmark yang menunjuk ke luar rentang CHUNK saat ini
                # (mis. AI keliru menunjuk baris konteks, atau chunk lain)
                # dibuang -- lebih aman kehilangan satu landmark daripada
                # merusak urutan/slicing global.
                if not isinstance(start, int) or not (chunk_start <= start < chunk_end):
                    continue
                if lm_type not in _VALID_TYPES:
                    continue
                entry = {"start": start, "type": lm_type}
                if lm_type == "bab":
                    entry["title"] = str(lm.get("title") or f"Bab (baris {start})").strip()
                all_landmarks.append(entry)
                last_landmark = entry
                chunk_landmark_count += 1

        _logger.info(
            "ai_naskah_parser: chunk baris %d-%d -> %d landmark ditemukan (%s)",
            chunk_start, chunk_end, chunk_landmark_count,
            [lm.get("type") for lm in (chunk_landmarks or []) if isinstance(lm, dict)],
        )

        chunk_start = chunk_end

    return _validate_structure(
        {"judul_buku": judul_buku, "nama_penulis": nama_penulis, "landmarks": all_landmarks},
        total_lines=n,
    )


def _validate_structure(structure: dict, total_lines: int) -> dict:
    """Buang/normalisasi landmark yang tidak masuk akal (indeks di luar
    jangkauan, type tidak dikenal, urutan tidak menaik, dst) supaya
    `build_sections_from_structure` tidak perlu ikut memvalidasi -- kalau
    hasil AI berantakan/kosong, mending fallback ke parser regex daripada
    dipaksakan dan menghasilkan buku yang lebih rusak.
    """
    if not isinstance(structure, dict):
        return {"landmarks": [], "judul_buku": "", "nama_penulis": ""}

    landmarks = structure.get("landmarks")
    if not isinstance(landmarks, list):
        landmarks = []

    cleaned = []
    last_start = -1
    for lm in landmarks:
        if not isinstance(lm, dict):
            continue
        start = lm.get("start")
        lm_type = lm.get("type")
        if not isinstance(start, int) or not (0 <= start < total_lines):
            continue
        if lm_type not in _VALID_TYPES:
            continue
        if start <= last_start:
            # Urutan tidak menaik / duplikat indeks -- skip, jangan biarkan
            # merusak logika slicing (yang mengasumsikan urutan naik).
            continue
        entry = {"start": start, "type": lm_type}
        if lm_type == "bab":
            entry["title"] = str(lm.get("title") or f"Bab (baris {start})").strip()
        cleaned.append(entry)
        last_start = start

    judul_buku = structure.get("judul_buku")
    nama_penulis = structure.get("nama_penulis")

    return {
        "landmarks": cleaned,
        "judul_buku": judul_buku.strip() if isinstance(judul_buku, str) else "",
        "nama_penulis": nama_penulis.strip() if isinstance(nama_penulis, str) else "",
    }


def build_sections_from_structure(
    raw_lines: list[str],
    structure: dict,
) -> tuple[str, str, str, list[tuple[str, list[str]]]]:
    """Ubah hasil deteksi AI (indeks landmark) jadi potongan teks/section
    asli, dengan MENGAMBIL LANGSUNG dari `raw_lines` (bukan dari jawaban
    AI) -- supaya isi naskah 100% utuh sama seperti aslinya.

    Return: (kata_pengantar_text, sinopsis_text, tentang_penulis_text, sections)
      `sections` formatnya PERSIS SAMA seperti naskah_parser.split_into_sections()
      -- list[(judul_bab, baris_isi)] -- supaya kompatibel langsung dengan
      alur mailmerge.py yang sudah ada.
    """
    landmarks = structure.get("landmarks", [])
    if not landmarks:
        # AI gagal / tidak mengembalikan apapun yang valid -- serahkan ke
        # pemanggil untuk fallback ke naskah_parser regex-based.
        return "", "", "", []

    n = len(raw_lines)
    kata_pengantar_lines: list[str] = []
    sinopsis_lines: list[str] = []
    tentang_penulis_lines: list[str] = []
    sections: list[tuple[str, list[str]]] = []

    for idx, lm in enumerate(landmarks):
        start = lm["start"]
        end = landmarks[idx + 1]["start"] if idx + 1 < len(landmarks) else n
        # Baris "start" itu sendiri biasanya adalah judul/label bagian ini
        # (mis. "SINOPSIS", "Mengenal Guru", "Daftar Pustaka") -- untuk
        # SEMUA tipe (bukan cuma "bab") baris label ini TIDAK ikut disertakan
        # lagi sebagai baris isi pertama, supaya tidak dobel (mis. "Daftar
        # Pustaka" muncul sebagai judul chapter SEKALIGUS baris pertama
        # isinya).
        content_start = start + 1
        chunk = [raw_lines[i] for i in range(content_start, end) if raw_lines[i] != ""]

        if lm["type"] == "kata_pengantar":
            kata_pengantar_lines.extend(chunk)
        elif lm["type"] == "sinopsis":
            sinopsis_lines.extend(chunk)
        elif lm["type"] == "tentang_penulis":
            tentang_penulis_lines.extend(chunk)
        elif lm["type"] == "daftar_pustaka":
            sections.append((raw_lines[start].strip() or "Daftar Pustaka", chunk))
        elif lm["type"] == "bab":
            sections.append((lm["title"], chunk))
        # type == "skip" -> baris-barisnya sengaja tidak dipakai sama sekali.

    kata_pengantar_text = "\n".join(kata_pengantar_lines).strip()
    sinopsis_text = "\n".join(sinopsis_lines).strip()
    tentang_penulis_text = "\n".join(tentang_penulis_lines).strip()
    return kata_pengantar_text, sinopsis_text, tentang_penulis_text, sections


def parse_naskah_with_ai(
    raw_lines: list[str],
    api_key: str | None = None,
    model: str = _DEFAULT_MODEL,
) -> dict:
    """Satu pintu masuk paling gampang dipakai: naskah masuk (`raw_lines`),
    hasil siap-pakai keluar -- dipakai baik oleh app.py (buat prefill form
    nama penulis/judul/kata pengantar/dst begitu naskah diunggah) MAUPUN
    mailmerge.py (buat generate dokumen final), supaya keduanya SELALU
    lihat struktur yang SAMA PERSIS (tidak ada risiko app.py & mailmerge.py
    parsing beda cara lalu hasilnya tidak sinkron).

    Return dict:
      {
        "ok": bool,               # False kalau AI gagal/hasil kosong -- pemanggil
                                   # WAJIB fallback ke parser regex kalau ini False
        "judul_buku": str,
        "nama_penulis": str,
        "kata_pengantar": str,
        "sinopsis": str,
        "tentang_penulis": str,
        "sections": list[(judul_bab, baris_isi)],
      }
    """
    empty_result = {
        "ok": False,
        "judul_buku": "",
        "nama_penulis": "",
        "kata_pengantar": "",
        "sinopsis": "",
        "tentang_penulis": "",
        "sections": [],
    }
    if not raw_lines:
        return empty_result

    try:
        structure = detect_structure_with_ai(raw_lines, api_key=api_key, model=model)
    except Exception:
        # API key belum di-set, quota habis, jaringan bermasalah, respons AI
        # bukan JSON valid, dst -- pemanggil HARUS fallback ke parser regex,
        # jangan sampai proses (generate dokumen ATAUPUN sekadar prefill
        # form) ikut gagal total gara-gara ini.
        return empty_result

    kp, sinopsis, tentang_penulis, sections = build_sections_from_structure(raw_lines, structure)
    if not sections:
        return empty_result

    return {
        "ok": True,
        "judul_buku": structure.get("judul_buku", ""),
        "nama_penulis": structure.get("nama_penulis", ""),
        "kata_pengantar": kp,
        "sinopsis": sinopsis,
        "tentang_penulis": tentang_penulis,
        "sections": sections,
    }
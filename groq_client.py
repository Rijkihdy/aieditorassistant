"""
Lapisan integrasi ke Groq API (External Service) sesuai Alur Proses (BPMN):
  1. Menyusun Prompt Engineering (gabung naskah + panduan genre + contoh blurb)
  2. Mengirim permintaan analisis (Request API ke Groq)
  3. Menerima hasil analisis genre & opsi blurb
"""
import json
import os

from groq import Groq

GENRE_LIST = [
    "Sejarah",
    "Roman / Percintaan",
    "Misteri / Thriller",
    "Fantasi",
    "Horor",
    "Pengembangan Diri",
    "Religi",
    "Bisnis",
    "Anak / Dongeng",
    "Puisi",
    "Umum / Fiksi",
]

# Berapa banyak karakter awal naskah yang dikirim ke Groq (jaga agar prompt tetap ringan/cepat)
MAX_CHARS_SENT = 6000


class GroqRequestError(Exception):
    """Dilempar saat request ke Groq API gagal, ditangkap di app.py untuk tombol Retry."""


def _get_client():
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        raise GroqRequestError(
            "GROQ_API_KEY belum diset. Isi file .env sesuai .env.example lalu restart aplikasi."
        )
    return Groq(api_key=api_key)


def _model_name():
    return os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")


def analyze_naskah(text: str, genre_hint: str = ""):
    """
    Satu kali panggilan ke Groq untuk sekaligus:
      - mengklasifikasi genre (jika genre_hint kosong)
      - menghasilkan 3 opsi teks blurb sesuai pola buku Guepedia

    Return: dict {"genre": str, "blurbs": [str, str, str]}
    """
    client = _get_client()
    excerpt = text[:MAX_CHARS_SENT]

    genre_instruction = (
        f'Genre sudah ditentukan pengguna sebagai "{genre_hint}", gunakan genre ini apa adanya.'
        if genre_hint
        else f"Klasifikasikan genre naskah ke SATU pilihan paling sesuai dari daftar berikut: {', '.join(GENRE_LIST)}."
    )

    system_prompt = (
        "Kamu adalah asisten editor digital untuk penerbit indie Guepedia. "
        "Tugasmu membaca cuplikan naskah lalu (1) menentukan genre, dan (2) menulis 3 alternatif "
        "teks blurb (sinopsis belakang buku) yang menjual, masing-masing 2-4 kalimat, gaya bahasa "
        "Indonesia yang menarik pembaca tapi tidak berlebihan (tidak clickbait murahan). "
        "Balas HANYA dalam format JSON valid tanpa markdown, tanpa penjelasan tambahan, dengan skema persis:\n"
        '{"genre": "<salah satu genre>", "blurbs": ["<opsi 1>", "<opsi 2>", "<opsi 3>"]}'
    )

    user_prompt = (
        f"{genre_instruction}\n\n"
        f"Cuplikan naskah:\n\"\"\"\n{excerpt}\n\"\"\""
    )

    try:
        completion = client.chat.completions.create(
            model=_model_name(),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.7,
            max_tokens=800,
            response_format={"type": "json_object"},
        )
        content = completion.choices[0].message.content
        data = json.loads(content)

        genre = genre_hint or data.get("genre", "Umum / Fiksi")
        blurbs = data.get("blurbs", [])
        # Jaga-jaga selalu ada tepat 3 opsi
        while len(blurbs) < 3:
            blurbs.append(blurbs[-1] if blurbs else "Blurb tidak tersedia.")
        return {"genre": genre, "blurbs": blurbs[:3]}

    except json.JSONDecodeError as e:
        raise GroqRequestError(f"Groq mengembalikan format tak terduga: {e}") from e
    except Exception as e:  # noqa: BLE001 - diteruskan sebagai pesan error yang ramah
        raise GroqRequestError(f"Request ke Groq API gagal: {e}") from e


# Berapa karakter cuplikan yang diambil dari tiap bab untuk disusun jadi sinopsis
CHARS_PER_CHAPTER_EXCERPT = 700
# Maksimum jumlah bab yang cuplikannya dikirim (jaga token tetap wajar untuk buku sangat panjang)
MAX_CHAPTERS_FOR_SYNOPSIS = 30


def generate_synopsis(
    chapter_sections: list[tuple[str, list[str]]],
    judul_naskah: str,
    min_kalimat: int = 50,
    max_kalimat: int = 100,
) -> str:
    """
    Satu kali panggilan ke Groq untuk menghasilkan SATU sinopsis utuh (bukan
    pilihan ganda) yang merangkum seluruh isi buku dari cuplikan tiap bab.

    chapter_sections: list (judul_bab, baris_isi) hasil naskah_parser.split_into_sections.
    Return: satu string sinopsis.
    """
    client = _get_client()

    excerpt_blocks = []
    for title, body_lines in chapter_sections[:MAX_CHAPTERS_FOR_SYNOPSIS]:
        excerpt = " ".join(body_lines)[:CHARS_PER_CHAPTER_EXCERPT]
        if excerpt.strip():
            excerpt_blocks.append(f"[{title or 'Bagian'}]\n{excerpt}")
    excerpt_text = "\n\n".join(excerpt_blocks)

    if not excerpt_text.strip():
        raise GroqRequestError("Naskah kosong, tidak ada cuplikan bab untuk dibuatkan sinopsis.")

    system_prompt = (
        "Kamu adalah editor senior penerbit buku non-fiksi. Dari cuplikan tiap bab yang diberikan, "
        f"tulis SATU sinopsis utuh (bukan beberapa opsi/pilihan) sepanjang {min_kalimat}-{max_kalimat} "
        "kalimat yang merangkum keseluruhan isi buku. Rangkai sebagai narasi yang mengalir (bukan "
        "poin-poin/bullet/penomoran bab), ambil benang merah & bagian penting dari SEMUA bab yang "
        "diberikan secara proporsional, gunakan gaya bahasa Indonesia yang membuat pembaca penasaran "
        "untuk membaca buku ini secara utuh, tapi tetap jujur terhadap isi (bukan clickbait). "
        "Balas HANYA dengan teks sinopsisnya saja — tanpa judul pembuka, tanpa markdown, tanpa "
        "penjelasan tambahan, tanpa tanda kutip."
    )

    user_prompt = f"Judul buku: {judul_naskah or '(tanpa judul)'}\n\nCuplikan tiap bab:\n{excerpt_text}"

    try:
        completion = client.chat.completions.create(
            model=_model_name(),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.7,
            max_tokens=4000,
        )
        synopsis = completion.choices[0].message.content
        if not synopsis or not synopsis.strip():
            raise GroqRequestError("Groq mengembalikan sinopsis kosong.")
        return synopsis.strip()

    except GroqRequestError:
        raise
    except Exception as e:  # noqa: BLE001
        raise GroqRequestError(f"Request sinopsis ke Groq API gagal: {e}") from e
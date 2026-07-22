import os
from io import BytesIO

from PIL import Image

try:
    from google import genai
    from google.genai import types
except ImportError as e:
    raise ImportError(
        "Package 'google-genai' belum terinstall. Jalankan: "
        "pip install google-genai"
    ) from e

# Model gratis (dengan limit kuota harian) dari Google AI Studio.
# Nama gaulnya "Nano Banana". Butuh GEMINI_API_KEY di environment
# (daftar gratis, tanpa kartu kredit, di https://aistudio.google.com/apikey).
MODEL_NAME = "gemini-2.5-flash-image"


def generate_cover(prompt: str) -> Image.Image:
    """
    Generate cover menggunakan Gemini 2.5 Flash Image (Nano Banana).

    Butuh environment variable GEMINI_API_KEY.
    """

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY belum di-set. Ambil API key gratis di "
            "https://aistudio.google.com/apikey lalu isi ke file .env "
            "(GEMINI_API_KEY=xxxxx)."
        )

    client = genai.Client(api_key=api_key)

    response = client.models.generate_content(
        model=MODEL_NAME,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_modalities=["IMAGE"],
        ),
    )

    candidates = response.candidates or []
    if not candidates:
        raise RuntimeError("Gemini tidak mengembalikan hasil apapun (kemungkinan diblokir safety filter).")

    for part in candidates[0].content.parts:
        if part.inline_data is not None:
            return Image.open(BytesIO(part.inline_data.data))

    # Kalau sampai sini berarti modelnya cuma balas teks, bukan gambar
    text_parts = [p.text for p in candidates[0].content.parts if p.text]
    raise RuntimeError(
        "Gemini tidak menghasilkan gambar. "
        + (f"Respon teks: {' '.join(text_parts)}" if text_parts else "")
    )


if __name__ == "__main__":

    prompt = """
    Premium professional non-fiction book cover.
    Artificial Intelligence.
    Modern.
    Blue.
    Clean.
    """

    img = generate_cover(prompt)

    img.save("cover.png")

    print("Berhasil")
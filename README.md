# Guepedia AI Editor Assistant

Aplikasi internal (Streamlit) untuk tim editor Guepedia: unggah naskah → dapat
jumlah kata, klasifikasi genre, dan 3 opsi blurb otomatis dari Groq API.

Sesuai Technical Stack di presentasi:
`Browser (Streamlit)` → `Web Server (Python Logic)` → `SQLite` & `Groq API`

## 1. Install dependencies

```bash
pip install -r requirements.txt
```

## 2. Set API key Groq (jangan taruh langsung di kode)

```bash
cp .env.example .env
```

Lalu buka `.env` dan isi:

```
GROQ_API_KEY=gsk_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
GROQ_MODEL=llama-3.3-70b-versatile
```

Dapatkan API key gratis di https://console.groq.com/keys

## 3. Jalankan aplikasi

```bash
streamlit run app.py
```

Buka http://localhost:8501 di browser.

## Struktur file

```
app.py            -> Halaman utama Streamlit (dashboard + hasil analisis)
groq_client.py     -> Pemanggilan Groq API (klasifikasi genre + generate blurb)
db.py             -> Penyimpanan riwayat analisis ke SQLite (guepedia_ai.db)
utils.py          -> Ekstraksi teks dari .docx/.pdf/.txt & hitung jumlah kata
requirements.txt   -> Daftar dependency Python
.env.example       -> Contoh isi file .env (API key TIDAK disertakan di sini)
```

## Alur pemakaian

1. Editor upload file naskah (.docx / .pdf / .txt) atau tempel teks langsung.
2. Klik **Mulai Analisis** — teks diekstrak, dikirim ke Groq API dengan prompt
   yang menggabungkan naskah + panduan genre + instruksi format blurb.
3. Jika request gagal (mis. API key salah/koneksi putus), muncul pesan error
   dan tombol **Klik Retry** untuk mencoba ulang — sesuai alur BPMN di presentasi.
4. Jika berhasil, halaman **Result** menampilkan Summary (judul, jumlah kata,
   genre) dan 3 opsi blurb yang bisa diedit lalu disimpan ke database.

## Catatan keamanan

- API key **tidak pernah** ditulis di kode — hanya dibaca dari environment
  variable `GROQ_API_KEY` lewat file `.env` (yang sebaiknya dimasukkan ke
  `.gitignore` bila proyek ini di-push ke Git).

# Skripsi Beres

> Beresin daftar pustaka dan format skripsi dalam satu menit. Gratis, tanpa registrasi, 100% browser-based.

[![Deploy on Vercel](https://img.shields.io/badge/Deploy-Vercel-black?style=flat&logo=vercel)](https://vercel.com)
[![Made for Mahasiswa](https://img.shields.io/badge/Made_for-Mahasiswa-EC4899?style=flat)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

**Demo:** https://skripsi-beres.vercel.app · **Audience:** mahasiswa tingkat akhir

---

## Apa Ini

Tool gratis buat mahasiswa Indonesia yang lagi nyusun skripsi/tugas akhir. Ada 2 fitur utama:

1. **Daftar Pustaka Generator** — input DOI / URL / judul jurnal, output sitasi dalam 5 format akademik (APA, MLA, IEEE, Harvard, Vancouver). Powered by [Crossref API](https://www.crossref.org/) (180+ juta jurnal).
2. **Format Skripsi Auto-Fix** — upload file `.docx`, pilih template kampus (UI, UGM, ITB, IPB, Telkom, Binus, dll), dapat dokumen yang udah rapi: margin, font, ukuran, spasi, semua sesuai pedoman.

## Fitur

### 📚 Daftar Pustaka Generator
- 5 format akademik: APA 7, MLA 9, IEEE, Harvard, Vancouver
- Input DOI lengkap atau judul bebas
- **Bulk mode** — paste banyak referensi sekaligus, generate semua dalam sekali klik
- Auto-deteksi DOI dari URL (e.g. `https://doi.org/10.1234/abcd`)
- Sumber data: Crossref (akurat & up-to-date)

### 📄 Format Skripsi Auto-Fix
- Upload `.docx` → download `.docx` yang udah rapi
- Auto-fix: margin (T/R/B/L), font, ukuran teks, line spacing
- Heading 1/2/3 di-style dengan ukuran progressive
- Body text justify alignment
- Template kampus tersedia:
  - Universitas Indonesia (UI)
  - Universitas Gadjah Mada (UGM)
  - Institut Teknologi Bandung (ITB)
  - Institut Pertanian Bogor (IPB)
  - Telkom University
  - Bina Nusantara (Binus)
  - Umum (Standar Indonesia)

## Tech Stack

- **Backend:** Flask 3.0 + Python 3.11
- **Document processing:** python-docx
- **Citation API:** Crossref REST API (no key required)
- **Frontend:** Vanilla JS + Crimson Pro / Atkinson Hyperlegible
- **Design system:** Swiss Modernism 2.0 (editorial black + accent pink)
- **Deploy:** Vercel (serverless) atau Railway (gunicorn)

## Setup Lokal

```bash
git clone https://github.com/moonzyr17/skripsi-beres.git
cd skripsi-beres
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

App jalan di `http://localhost:5099`.

## Deploy

### Vercel
```bash
npm i -g vercel
vercel
```

`vercel.json` + `api/index.py` udah disiapkan — tinggal `vercel --prod`.

### Railway
Repo udah ada `Procfile`. Connect repo di [railway.app](https://railway.app) → auto-deploy.

## Endpoint API

| Method | Path | Fungsi |
|---|---|---|
| GET  | `/` | Landing page |
| GET  | `/app` | Workflow app |
| POST | `/api/cite` | Generate single citation. Body: `{query, style}` |
| POST | `/api/cite-bulk` | Generate bulk citations. Body: `{text, style}` |
| POST | `/api/format-docx` | Format `.docx`. Form: `file`, `template` |
| GET  | `/api/templates` | List template kampus |
| GET  | `/health` | Health check |

## Lisensi

MIT — bebas dipake, dimodif, di-fork.

## Kontribusi

PR welcome. Mau nambah template kampus baru? Edit `KAMPUS_TEMPLATES` di `app.py` — formatnya self-explanatory.

---

Made with ☕ for mahasiswa yang lagi pusing sama format skripsi.

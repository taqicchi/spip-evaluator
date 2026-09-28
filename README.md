# Asisten Evaluasi SPIP

Aplikasi lokal (HTML + JavaScript, tanpa server, tanpa instalasi) untuk membaca
kertas kerja **KK PM/PK Evaluasi SPIP** (.xlsx) dan menampilkan hal yang perlu
dievaluasi, sehingga auditor tidak perlu menelusuri sel satu per satu.

## Cara pakai

1. Buka `index.html` dengan Chrome/Edge (klik dua kali cukup).
2. Tarik berkas .xlsx ke halaman, atau klik **Buka berkas…**.
3. Pilih sheet yang akan dibaca (semua, atau hanya KK 3.x / KKE / KK 4–8).
4. Telusuri per sheet, beri status & catatan, lalu **Ekspor hasil reviu (.xlsx)**.

Berkas **tidak diunggah ke mana pun**; semua diproses di peramban. Catatan reviu
disimpan di `localStorage` peramban per nama berkas. Untuk pindah komputer atau
berbagi dengan anggota tim, pakai **Sesi → Simpan sesi reviu (.json)**.

## Format kertas kerja yang dikenali

Deteksi berdasarkan isi (judul kolom), bukan nama sheet, jadi berkas K/L lain
dengan template yang sama ikut terbaca.

| Jenis | Contoh sheet | Cara dikenali | Tampilan |
|---|---|---|---|
| Struktur & Proses | KK3.1–KK3.4 | ada kolom "Uraian Parameter" dan "Uraian Hasil Pengujian" | kartu per parameter: kriteria A–E berdampingan dengan uraian, grade PM/PK/Evaluasi, AoI, penyebab |
| Rekap nilai | KKLEAD_SPIP | "PENYIMPULAN NILAI MATURITAS" + kolom "Skor" | skor per subunsur; tandai skor yang diketik manual |
| Tabel KK | KKE 1/2.x, KK 4, KK 5.x, KK 6–8 | baris penomoran kolom (1, 2, 3 …) di bawah header | tabel ringkas, beda PM→PK→Evaluasi disorot, detail vertikal per baris |
| Informasi | FAQ, petunjuk, referensi | selain di atas | teks |

KK 3.1 berisi blok per satker (Satker 1–10, kolom Grade PM/PK/Evaluasi dalam satu
blok, uraian dipakai bersama). KK 3.2–3.4 berisi tiga blok (PM, PK, Evaluasi)
masing-masing dengan uraiannya sendiri. Keduanya didukung.

## Pemeriksaan otomatis

| Kode | Arti |
|---|---|
| `BUKTI_KURANG` | Grade X berarti level X s.d. E terpenuhi, tetapi uraian di salah satu level itu kosong/masih template ("Bahwa ….. Telah …..") |
| `PERNYATAAN_NEGATIF` | Uraian pada level yang dianggap terpenuhi memuat "belum", "tidak ada", "kurang memadai", dst. |
| `URAIAN_DI_ATAS_GRADE` | Ada uraian di level lebih tinggi dari grade |
| `GRADE_KOSONG` / `GRADE_TIDAK_VALID` | Uraian ada tetapi grade kosong / bukan A–E |
| `AOI_KOSONG`, `PENYEBAB_KOSONG` | Grade PK/Evaluasi belum A tetapi AoI/penyebab tidak diisi (AoI adalah isian evaluator, tidak diperiksa pada tahap PM) |
| `GRADE_NAIK` / `GRADE_TURUN` | Grade berubah antar tahap (naik ditandai lebih serius) |
| `URAIAN_DUPLIKAT`, `URAIAN_DISALIN` | Teks uraian identik antar level/parameter, atau PK menyalin PM |
| `KODE_TANGGAL`, `KODE_SUBUNSUR_BEDA` | Kode subunsur (mis. 1.8) tersimpan sebagai tanggal, atau berbeda antar KK |
| `SKOR_MANUAL` | Skor di KKLEAD diketik manual, bukan rumus |
| `T_TANPA_AOI`, `YT_KOSONG`, `YT_TIDAK_VALID`, `BEDA_TAHAP`, `REALISASI_KOSONG` | Pemeriksaan tabel KKE/KK 5 |

Temuan adalah petunjuk untuk diperiksa, bukan kesimpulan.

## Hitung ulang

Berkas yang disimpan ulang oleh aplikasi selain Excel (mis. openpyxl) tidak
menyimpan hasil rumus, sehingga nilai di KKLEAD kosong saat dibaca. Aplikasi
menghitung ulang dengan logika yang sama dengan rumus KK:

- Kesimpulan parameter KK 3.1 = grade terbanyak antar satker; jika seri, grade terendah.
- Skor parameter A=5 … E=1; skor subunsur per KK = rata-rata parameter.
- Skor gabungan subunsur = rata-rata KK 3.1–3.4 (seperti KKLEAD II), **sebelum** veto KK 4.

## Berkas

- `index.html`, `style.css`: antarmuka
- `parser.js`: pembacaan format & pemeriksaan (bisa dipakai di Node: `require('./parser.js')`)
- `app.js`: tampilan, catatan reviu, ekspor
- `vendor/xlsx.full.min.js`: SheetJS 0.20.3 (lokal, agar bisa dipakai tanpa internet)

Jangan menyimpan kertas kerja asli di folder ini.

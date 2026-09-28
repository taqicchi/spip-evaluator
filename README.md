# Evaluator KK 3.1 SPIP (branch `kk31`)

Ruang kerja evaluator khusus **KK 3.1 – Penilaian Struktur dan Proses Efektivitas dan
Efisiensi Pencapaian Tujuan (T1), tingkat satker**. Aplikasi lokal (HTML + JavaScript),
tanpa server dan tanpa instalasi. Versi umum untuk semua sheet ada di branch `main`.

## Pembagian isian KK 3.1

| Kolom per satker | Diisi oleh |
|---|---|
| Uraian Hasil Pengujian (per level A–E), Grade PM | Satker (Penilaian Mandiri) |
| Grade PK, Grade Evaluasi | APIP / auditor / evaluator |
| Kluster AoI, Uraian AoI, Kluster Penyebab, Uraian Penyebab | APIP / auditor / evaluator |

## Cara pakai

1. Buka `index.html` dengan Chrome/Edge.
2. Tarik berkas KK (.xlsx) ke halaman. Bisa workbook lengkap atau berkas yang hanya berisi sheet KK3.1.
3. Menu **Bukti** → pilih folder atau ZIP bukti dari satker. Susunan yang dikenali:
   `…/1.1 - Penegakan Integritas dan Nilai Etika/C/berkas.pdf` (kode subunsur, lalu folder level A–E).
   ZIP dibaca langsung tanpa diekstrak; hanya berkas yang dibuka yang dibaca.
4. Pilih tahap (**PK** atau **Evaluasi**) dan satker.
5. Per parameter: baca kriteria, uraian satker, dan buka bukti tiap level (pratinjau PDF/gambar di panel kanan).
   Tandai **Terbukti / Sebagian / Tidak** per level dan beri catatan.
6. Aplikasi menyarankan grade = level tertinggi yang ia dan semua level di bawahnya terbukti.
   Tetapkan grade, Kluster & Uraian AoI (tombol *susun dari verifikasi*), Kluster & Uraian Penyebab.
7. **Unduh KK terisi**: isian ditulis ke salinan berkas asli pada kolom Grade PK/Evaluasi, AoI, dan
   penyebab satker yang bersangkutan. Aplikasi membaca ulang hasilnya untuk memastikan setiap sel
   tertulis benar sebelum mengunduh.

Pintasan: ← / → pindah parameter, A–E pilih grade (saat tidak sedang mengetik).

## Yang dijaga

- Berkas asli tidak diubah; hasil selalu salinan.
- Hanya sel yang diisi evaluator di aplikasi yang ditulis. Format, merge, validasi data (dropdown),
  rumus, dan sheet lain tetap byte-identik — bagian ZIP yang tidak berubah disalin apa adanya.
- Kluster AoI mengikuti daftar di sheet REF sesuai jenis parameter (SPIP/MRI/IEPK); penyebab dari
  daftar Kluster Penyebab.
- "Tidak dapat dinilai" mengosongkan grade (validasi KK hanya menerima A–E) dan menulis alasannya di Uraian AoI.
- KK 3.1 hanya punya satu set kolom AoI/penyebab per satker; isian tahap Evaluasi menimpa isian PK.

## Catatan otomatis atas isian satker

Ditampilkan sebagai bahan pengujian, bukan kesimpulan:

- grade PM mengklaim level yang uraiannya kosong/masih template, atau yang **folder buktinya kosong**;
- uraian pada level yang diklaim memuat "belum", "tidak ada", "kurang memadai", dst.;
- uraian identik antar level/parameter (salin-tempel);
- ada uraian di level lebih tinggi dari grade PM.

## Penyimpanan

Isian evaluator (verifikasi level, catatan, grade, AoI) tersimpan di peramban per nama berkas.
Untuk pindah komputer/berbagi: **Sesi → Simpan sesi (.json)**. **Sesi → Ekspor lembar kerja evaluator**
menghasilkan .xlsx berisi verifikasi per level, catatan, dan daftar bukti sebagai jejak pengujian.

## Berkas

- `index.html`, `style.css`, `kk31.js`: antarmuka
- `parser.js`: pembacaan format KK (dipakai bersama branch `main`)
- `zip.js`: baca ZIP secara lazy & tulis ulang dengan menyalin byte asli
- `xlsxpatch.js`: menulis sel ke XML sheet tanpa mengubah bagian lain
- `vendor/xlsx.full.min.js`: SheetJS 0.20.3

Jangan menyimpan kertas kerja atau bukti asli di folder ini.

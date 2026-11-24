# V5.2 · Elegant Hybrid UI Refresh

Desain baru: header premium navy-slate, sidebar dark, konten terang, aksen teal-biru, KPI cards, tab compact, tabel fokus keterbacaan. **Engine tidak berubah dari V5.1** dan kelipatan komisi tetap 1 poin persentase.

Tidak ada dependensi UI baru. Cukup ganti file `app.py` di GitHub dengan versi ini untuk update desain; file lain disertakan agar paket lengkap. Jangan upload file XLSX, SQLite, atau rahasia.

# INPERTEK AFFILIATE COMMISSION OPTIMIZER V5.2

## Multi-master + kelengkapan fitur V3

- Upload banyak master, identitas dan perubahan masing-masing master terpisah. Nama file harus sama saat memperbarui snapshot.
- Hanya sheet hijau, dengan validasi mapping. Sheet tidak cocok dilewati.
- Analisis lengkap, detail SKU dan rumus, penjelasan kolom, validasi, histori per master, ekspor Excel dan CSV.
- Mode agresif / seimbang / konservatif. Peluang khusus opsional tidak otomatis masuk ekspor. Batas kenaikan opsional.
- **Komisi USULAN SELALU kelipatan 1 poin persentase** (1%, 2%, 3%, dst). Persentase awal tetap dibaca apa adanya.
- Perlindungan margin normal: margin baru >= komisi baru; minimum margin; dan margin baru tidak negatif.
- Master asli read-only. No Supabase. SQLite sementara dan upload master di Streamlit Cloud akan diproses di server; selalu cadangkan ekspor.
- Tidak ada upload langsung marketplace; keluaran belum template resmi dan harus direview.

## Deploy

Replace app.py, engine.py, requirements.txt, README.md and .gitignore in the GitHub repository; jangan upload master, SQLite, atau data internal. Branch main, main file app.py.

## Test

`python -m pip install -r requirements.txt`
`python -m streamlit run app.py`
`python test_v51.py` (memerlukan lima master contoh pada /mnt/data di lingkungan pengembangan).

## Catatan

Migrasi SQLite V5 lama tidak otomatis terjadi apabila file database hilang pada cloud. Nama DB V5 tetap optimizer_v5.sqlite agar kompatibel selama database sebelumnya masih tersedia.

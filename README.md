# INPERTEK Affiliate INSERT Pilot V3

## Prinsip: master 100% read-only
File INSERT.xlsx hanya dibaca, tidak pernah ditulis. Snapshot SQLite lokal dan berkas kandidat Excel terpisah. Tidak ada perubahan otomatis ke marketplace.

## Menjalankan
1. Ekstrak ZIP -> File > Open Folder di VS Code -> pilih `INPERTEK_INSERT_PILOT`.
2. `python -m pip install -r requirements.txt`
3. `python -m streamlit run app.py`
4. Buka http://localhost:8501.
5. Bila punya data hasil sync lama, backup dan gunakan `insert_pilot.sqlite` milik aplikasi lama agar histori tidak tertimpa.

## Logika komisi
- Komisi otomatis dari sebagian ruang margin aman tiap produk (pengaturan default 50% headroom). Ini **parameter contoh**, bukan kebijakan resmi perusahaan.
- Normal: margin sesudah komisi baru >= 0, >= komisi baru, dan >= margin dilindungi (default 50% margin sebelum komisi).
- Bila ruang tidak mencukupi: TURUNKAN atau REVIEW NONAKTIF; file master tetap.
- Mode PELUANG KHUSUS hanya simulasi, margin tidak boleh negatif, tidak diikutkan ke kandidat ekspor.
- Angka Rupiah ditampilkan sebagai Rp1.234 dan di Excel sebagai angka dengan format Rupiah agar tetap dapat dihitung.
- Excel kandidat berisi perubahan komisi saja yang lolos pemeriksaan, mengeluarkan data validasi, tetap, nonaktif, peluang khusus, dan SKU sama di beberapa sheet pada marketplace sama.
- File kandidat **bukan** template upload resmi marketplace. Perlu pencocokan SKU aktif, aturan komisi, dan template Shopee/Tokopedia/TikTok.

## Batasan pilot
- Mapping warna hijau dan kolom pilot INSERT belum dinamis untuk 25 master lain.
- Harga jual dipakai sebagai estimasi basis biaya affiliate; validasi bila platform memakai basis berbeda.
- Margin diasumsikan sudah dikurangi biaya affiliate awal yang terdapat di master.
- Snapshot awal bawaan adalah data historis, bukan live connection Drive / marketplace.
- Kandidat dengan SKU ganda lintas-sheet ditahan dari ekspor.

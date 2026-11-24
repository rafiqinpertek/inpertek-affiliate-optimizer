# INPERTEK Affiliate Commission Optimizer V5 — Multi-Master

## Deploy (tanpa Supabase)

1. Backup file V3 dan download histori Excel apabila diperlukan.
2. Ganti `app.py`, `engine.py`, `requirements.txt`, `README.md`; tambahkan `.gitignore` pada GitHub repository **Private**.
3. Jangan upload file Excel master, HPP, kredensial, maupun `*.sqlite`.
4. Streamlit Cloud akan redeploy otomatis. Pastikan **Sharing** hanya untuk pengguna yang diizinkan.
5. Upload satu atau beberapa master XLSX lewat website. Utamakan tab hijau; aplikasi menahan sheet yang mapping-nya tidak lolos verifikasi.
6. Unduh Excel hasil review sebelum reboot server. SQLite di Streamlit Community Cloud sementara, tidak permanen.

## Kebijakan

- Master adalah **read-only**. Tidak ada penulisan atau perubahan terhadap master asli.
- Komisi diasumsikan dihitung dari **harga jual**. Tarif master dalam pecahan desimal (`0.03` = `3%`).
- Margin master diasumsikan **sudah dikurangi** biaya affiliate awal. Perhitungan mengembalikan biaya awal satu kali.
- Usulan normal mempertahankan margin baru ≥ biaya komisi baru, ≥ minimum margin yang dipilih, dan tidak negatif.
- Agresif menggunakan 85% ruang kenaikan aman. Batas nominal/komisi yang berlaku harus dicek ke marketplace.
- Data invalid, mapping meragukan, SKU duplikat lintas master tidak masuk kandidat perubahan.
- Ekspor KANDIDAT_REVIEW **bukan template resmi** untuk upload marketplace.

## Batasan

- Hanya worksheet yang bertab hijau RGB `92D050` yang dianalisis.
- Mapping standar saat ini: `D` SKU, `E` Nama, `N` HPP; Tokopedia `CC DA DB DH`; Shopee `EM FL FM FS`.
- Mapping diverifikasi melalui kecocokan **harga × rate = nominal fee** pada sampel per sheet. Sheet yang tidak cocok diskip, tidak diinterpretasi sembarangan.
- Identitas master berdasarkan **nama file**; jika upload ulang gunakan nama file sama untuk mendeteksi perubahan.
- Snapshot/riwayat disimpan sementara di `optimizer_v5.sqlite` pada server Streamlit; reboot/redeploy dapat menghilangkannya.
- V5 membangun baseline baru; SQLite V3 lama tidak otomatis dimigrasikan.
- Backend bukan database permanen / Google Drive sync.

## Lokal

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

# INPERTEK Affiliate Optimizer V4 — PostgreSQL (Supabase)

## Yang berubah dari V3
- **Master Excel tetap read-only**: file upload hanya dibaca sementara; tidak diubah maupun disimpan sebagai berkas di database.
- Snapshot produk, validasi, perubahan dan histori disimpan secara **persisten** di Supabase PostgreSQL, bukan SQLite lokal.
- Tabel disimpan dalam schema privat `inpertek_internal`, **bukan** `public`, agar tidak otomatis dipublikasikan oleh Supabase Data API.
- Parameter koneksi hanya melalui **Streamlit Cloud → Settings → Secrets**, tidak berada dalam kode GitHub.
- Rekomendasi komisi dan ekspor Excel tetap bersifat simulasi dan review; tidak mengubah komisi marketplace.

## 1. Siapkan koneksi database secara rahasia
Di Supabase → Project → **Connect** → **Session pooler**, salin connection string URI sesuai project Anda. Gunakan **Session pooler** bila akses IPv4 diperlukan. Pilih region dekat Indonesia (misalnya Singapore) bila tersedia. Password pada URI wajib di-*URL encode* bila memiliki karakter spesial seperti `@`, `#`, atau `/`.

Di Streamlit Cloud → aplikasi → **Settings → Secrets** masukkan:

```toml
DATABASE_URL = "postgresql://USER:PASSWORD_URL_ENCODED@POOLER_HOST:5432/postgres"
```

**Ganti semua nilai contoh dengan data Supabase yang asli pada halaman Secrets saja.** Jangan taruh URI di `app.py`, `README.md`, issue GitHub, screenshot, atau chat.

Atur aplikasi Streamlit sebagai **Private / Only specific people** dan uji kembali melalui Incognito. Repository Private **tidak** berarti aplikasi Private.

## 2. Update repository GitHub (file mana saja?)
Unggah/replace tepat file berikut ke **root repository**:
- `app.py`
- `engine.py`
- `cloud_store.py` **(baru)**
- `requirements.txt`
- `README.md`

**Jangan pernah mengunggah** `insert_pilot.sqlite`, `01. INSERT.xlsx`, file CSV data asli, `.streamlit/secrets.toml`, password, atau `DATABASE_URL` ke GitHub. ZIP ini sengaja tidak menyertakan database atau master asli.

Streamlit Cloud akan redeploy ketika ada commit baru. `Main file path` tetap `app.py` bila semua file berada di root repository.

## 3. Jalankan
Setelah Secrets tersimpan dan redeploy selesai, dashboard akan menampilkan **V4 · PostgreSQL**. Supabase membuat tabel otomatis di schema `inpertek_internal` pada koneksi pertama. Database awal masih kosong. Upload Excel asli **hanya jika** kontrol akses aplikasi telah dipastikan private dan kredensial database sudah aman. Upload pertama di V4 akan menjadi **baseline baru**.

Riwayat lama dari SQLite di server Streamlit tidak otomatis berpindah ke PostgreSQL dan mungkin sudah hilang akibat restart/redeploy. Jangan mengira histori lama termigrasi otomatis. Jika masih ada file `insert_pilot.sqlite` V3 yang valid di komputer lokal, cadangkan file tersebut untuk migrasi terpisah.

## 4. Perlindungan dan batasan
- Program hanya mengubah database histori terpisah, bukan master Excel.
- Kolom uang tetap diformat Rp pada dashboard dan ekspor.
- Sistem tidak menyimpan file Excel asli utuh di Supabase; yang disimpan adalah nilai produk termasuk HPP/margin, jadi Supabase tetap berisi data **rahasia**.
- Gunakan database khusus yang tidak dibagikan; aktifkan MFA akun GitHub/Streamlit/Supabase, atur project role, dan jangan beri akses Supabase ke pengguna yang tidak berwenang.
- Secara default masih belum ada proses persetujuan komisi per akun atau upload langsung ke marketplace.
- Rekomendasi mengasumsikan biaya affiliate awal dihitung dari harga jual: sesuaikan dengan ketentuan platform sebelum digunakan.
- Backup database: Supabase tier gratis memiliki batasan; siapkan rencana backup terpisah untuk produksi.

## 5. Troubleshooting
- Error `Database PostgreSQL belum diatur`: tambahkan Secrets `DATABASE_URL` dan reboot aplikasi.
- Error `Database belum bisa dihubungi`: cocokkan Connect → Session pooler, host, port, password, SSL, dan konektivitas project Supabase; jangan kirim URL/password kepada siapa pun.
- Dashboard kosong: snapshot V4 memang belum ada; upload master setelah akses private diverifikasi.
- File `cloud_store.py` tidak ditemukan: pastikan file baru berada di root repository, satu level dengan `app.py`.
- UI masih menunjukkan V3: tunggu redeploy commit terbaru dan periksa **Manage app → Logs**.

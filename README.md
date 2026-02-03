# Automasi-CSV-BCA-SUREL
Proyek ini adalah alat automasi berbasis Python yang dirancang untuk memproses file mutasi bank (format CSV dari BCA), mengonversinya menjadi format Excel yang rapi, mengirimkan laporan tersebut melalui email, dan memindahkan file ke direktori arsip yang ditentukan berdasarkan kode rekening unik.

# Fitur Utama
- Konversi CSV ke Excel: Mengubah file CSV mentah menjadi file .xlsx dengan format kolom otomatis (autofit).
- Parsing Data Cerdas: Secara otomatis mendeteksi nomor rekening (4 digit terakhir) dan periode tanggal dari isi file CSV untuk penamaan file yang dinamis.
- Pengiriman Email Tertarget: Mengirim file Excel sebagai lampiran ke penerima email yang berbeda-beda berdasarkan kode rekening yang terdeteksi.
- Routing File (Pengarsipan): Memindahkan file hasil ke folder jaringan atau lokal yang berbeda (Pusat/Depo) berdasarkan kode rekening.
- Pembersihan Otomatis: Secara otomatis membersihkan file sementara dan menghapus file sumber CSV setelah proses berhasil.

# Prasyarat
Sebelum menjalankan program, pastikan sistem Anda memiliki:
- Python 3.x terinstal.
- Pustaka Python yang diperlukan:
```bash
pip install pandas openpyxl
```

# Konfigurasi
Sebelum menjalankan program, Anda wajib menyunting file Dapur/config.conf. Sesuaikan dengan kredensial dan path lokal Anda.

Penjelasan config.conf
1. [DIRECTORY] Mengatur tujuan pemindahan file jika memilih Opsi 3.
- pusat: Path folder untuk rekening pusat (kode 3777). Ubah di kode python apabila 3777 ingin di sesuaikan dengan yang lain pada "3_MovingDoc.py".
- depo: Path folder untuk rekening depo (kode lainnya).

2. [EMAIL] Pengaturan akun Gmail pengirim.
- PENGIRIM: Alamat email Gmail Anda.
- PASSWORD: App Password (Bukan password login biasa). Pastikan 2FA aktif di Google Account dan buat App Password khusus.

3. [PENERIMA] Pemetaan kode rekening (4 digit terakhir) ke alamat email tujuan.
Format: KODE = email@tujuan.com
Contoh: 3777 = boss@kantor.com
Atau lihat contoh di dalam config.conf

4. [ISI EMAIL] Template subjek dan isi pesan email.
- SUBJECT: Judul email. Gunakan {nama_file} untuk menyisipkan nama file secara dinamis.
- CONTENT: Isi body email. Gunakan \n untuk baris baru.

# Cara Penggunaan
1. Siapkan File: Letakkan satu atau beberapa file .csv mutasi BCA ke dalam folder Input.

2. Jalankan Program: Buka terminal atau command prompt, lalu jalankan script utama:
```bash
python "Jalankan Proses CSV.py"
```
Atau dapat klik dua kali pada python tersebut.

3. Pilihan mode:
   1. Hanya Konversi
   2. Konversi dan Kirim Surel
   3. Konversi, Kirim Surel, & Pindahkan File

# WhatsApp-GSheets-Gateway

Aplikasi Google Apps Script untuk mengirim pesan WhatsApp melalui Google Sheets menggunakan API KirimWA.

## Fitur

- ✅ Kirim pesan teks WhatsApp dari Google Sheets
- ✅ Kirim media (gambar, video, audio, dokumen)
- ✅ Kirim stiker WhatsApp
- ✅ Kirim vCard (kontak)
- ✅ Cek keberadaan nomor WhatsApp
- ✅ Lihat informasi device
- ✅ Generate QR untuk login
- ✅ Logout device
- ✅ Kirim pesan massal dari data spreadsheet
- ✅ Pengaturan API Key dan nomor pengirim

## Instalasi

1. Buka Google Sheets dan buat spreadsheet baru
2. Pilih **Extensions > Apps Script**
3. Hapus semua kode default dan tempel kode dari file `whatsapp-gateway.gs`
4. Simpan project (beri nama misalnya "WhatsApp Gateway")
5. Kembali ke spreadsheet dan refresh halaman
6. Menu baru "WhatsApp Gateway" akan muncul di toolbar

## Penggunaan

1. Klik menu **WhatsApp Gateway > Pengaturan** untuk mengatur API Key dan nomor pengirim
2. Gunakan menu-menu lain sesuai kebutuhan:
   - **Kirim Pesan**: Mengirim pesan teks
   - **Kirim Media**: Mengirim gambar, video, audio, atau dokumen
   - **Kirim Stiker**: Mengirim stiker
   - **Kirim vCard**: Mengirim kontak
   - **Cek Nomor**: Memeriksa keberadaan nomor di WhatsApp
   - **Informasi Device**: Melihat informasi device pengirim
   - **Generate QR**: Membuat QR code untuk login WhatsApp
   - **Logout Device**: Logout dari device
   - **Kirim Pesan Massal dari Sheet**: Mengirim pesan massal dari data di spreadsheet

## Mengirim Pesan Massal

1. Siapkan data di spreadsheet dengan minimal 2 kolom (nomor tujuan dan pesan)
2. Klik **WhatsApp Gateway > Kirim Pesan Massal dari Sheet**
3. Pilih kolom yang berisi nomor dan pesan
4. Tentukan baris mulai dan baris akhir (opsional)
5. Klik "Kirim Pesan Massal"
6. Status pengiriman akan ditambahkan di kolom terakhir

## Persyaratan

- Akun Google
- API Key dari [MPedia](https://m-pedia.co.id/)
- Nomor WhatsApp yang terdaftar di layanan MPedia

## Catatan Penting

- Script ini menggunakan API dari MPedia v8
- Pastikan nomor pengirim sudah terdaftar di layanan MPedia
- Format nomor harus dimulai dengan kode negara (contoh: 628xxxxxxxxx untuk Indonesia)
- Untuk mengirim media, gunakan URL langsung ke file

## Lisensi

MIT

## Disclaimer

Aplikasi ini tidak terafiliasi dengan WhatsApp Inc

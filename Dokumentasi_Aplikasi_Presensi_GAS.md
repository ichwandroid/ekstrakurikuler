# Dokumentasi & Panduan Pengembangan Aplikasi Presensi (Google Apps Script)

Dokumen ini berisi panduan teknis, struktur kode, dan arsitektur sistem untuk membangun aplikasi presensi berbasis **Google Apps Script (GAS)**, **Google Sheets**, dengan antarmuka UI/UX modern, animasi **Anime.js**, serta sistem pembuat & pemindai **QR Code**.

---

## 🏗️ 1. Arsitektur System & Alur Kerja (Workflow)

1. **User Side (Halaman Presensi User):**
   - User membuka web app.
   - User memilih nama mereka dari daftar dropdown (diambil secara dinamis dari Google Sheet).
   - User mengklik tombol **"Generate QR Code"**.
   - Sistem membuat QR Code yang berisi identitas unik user (misal: ID Pegawai / Email / Nama + Token Waktu).
   - QR Code ditampilkan di layar HP/Perangkat user dengan efek animasi Glassmorphism.

2. **Admin Side (Halaman Pemindai Admin):**
   - Admin membuka web app mode Admin / Scanner.
   - Kamera perangkat admin terhubung menggunakan library pemindai QR (seperti `html5-qrcode`).
   - Admin memindai QR Code yang ditunjukkan oleh User.
   - Sistem mengirimkan payload hasil scan ke server (Google Apps Script).

3. **Backend & Database (Google Apps Script & Google Sheet):**
   - Server memverifikasi data QR Code.
   - Waktu scan (Timestamp), Nama, Status Presensi, dan Detail Perangkat/Lokasi dicatat secara otomatis ke **Google Sheets**.
   - Server mengembalikan respon sukses ke tampilan Admin.

---

## 📊 2. Struktur Google Sheets (Database)

Buat satu file Google Sheet dengan **2 Sheet (Tab)**:

### Tab 1: `Data_User`
| A (ID_User) | B (Nama_Lengkap) | C (Jabatan / Divisi) | D (Status_Aktif) |
| :--- | :--- | :--- | :--- |
| USR-001 | Ahmad Fauzi | Software Engineer | Aktif |
| USR-002 | Siti Rahma | Product Manager | Aktif |
| USR-003 | Budi Santoso | UI/UX Designer | Aktif |

### Tab 2: `Log_Presensi`
| A (Timestamp) | B (ID_User) | C (Nama_Lengkap) | D (Status) | E (Keterangan) |
| :--- | :--- | :--- | :--- | :--- |
| 2026-09-01 08:00:15 | USR-001 | Ahmad Fauzi | Hadir | Scan QR Admin |

---

## 🚀 3. Langkah-Langkah Deployment (Publishing)

1. **Buka Google Sheet** Anda.
2. Klik menu **Extensions** > **Apps Script**.
3. Buat 3 file dalam project Apps Script:
   - `Code.gs` (isikan kode Apps Script)
   - `index.html` (isikan kode halaman user)
   - `admin.html` (isikan kode halaman scanner admin)
4. Simpan semua file (`Ctrl + S`).
5. Klik tombol **Deploy** di pojok kanan atas > **New deployment**.
6. Pilih type: **Web app**.
7. Konfigurasi:
   - **Description**: `Aplikasi Presensi QR Code v1`
   - **Execute as**: `Me` (Akun Google Anda)
   - **Who has access**: `Anyone` (Siapa saja dapat mengakses)
8. Klik **Deploy**, lalu berikan izin akses (*Authorize Access*).
9. Salin URL Web App yang dihasilkan.

---

## 🔗 Cara Menggunakan
- **Akses User:** Buka URL Web App biasa (secara otomatis membuka halaman `index`).
- **Akses Admin Scanner:** Tambahkan parameter `?page=admin` di akhir URL Web App (contoh: `https://script.google.com/macros/s/XXXXX/exec?page=admin`).

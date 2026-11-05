# Product Requirements Document (PRD)
## Modul Omnichannel WhatsApp Inbox FU (Follow-Up)

| Atribut | Keterangan |
| :--- | :--- |
| **Nama Produk** | Omnichannel WhatsApp Inbox Follow-Up (FU) |
| **Versi Dokumen** | 2.1.0 |
| **Tanggal Pembaruan** | 5 November 2026 |
| **Status Dokumen** | Production Approved |
| **Target Pengguna** | Tim Sosmed, Admin Follow-Up, Tim Sales/PIC Cabang Fisik |
| **Repositori** | `REPORT` (`workspace-divisisosmed.vercel.app`) |

---

## 1. Latar Belakang & Analisis Masalah

### 1.1 Konteks Bisnis
Divisi Sosmed Gadget Mart mengelola komunikasi intensif dengan puluhan Person in Charge (PIC) di berbagai cabang fisik toko (Banjarmasin, Banjarbaru, Batulicin, Pontianak, Sampit, Pelaihari, dll.). Komunikasi ini bertujuan untuk menindaklanjuti (*follow-up*) data keep/prospek customer yang memesan unit secara online agar segera terealisasi menjadi transaksi faktur (AR) di toko offline.

### 1.2 Masalah Sistem Sebelumnya (Problem Statement)
1. **Pemisahan Antarmuka:** Admin sebelumnya harus berganti-ganti jendela antara WhatsApp Web resmi, spreadsheet Google Sheets keep data, dan laporan crosselling internal.
2. **Keterbatasan Gateway Pihak Ketiga (WatZap):**
   - Tidak mendukung pesan masuk balasan kutipan (*native quoted reply*).
   - Balasan yang diketik admin langsung dari ponsel fisik WhatsApp tidak tercatat di dashboard web (menyebabkan pesan menjadi *out-of-sync*).
   - URL media gambar sering kedaluwarsa (*broken image link*).
3. **WhatsApp Linked Identity (`@lid`):**
   - WhatsApp versi modern menggunakan enkripsi identitas terhubung (`@lid`). Tanpa resolusi, roomchat terpecah menjadi room duplikat/nomor anonim tanpa nomor telepon asli.
4. **Hilangnya Konteks Follow-Up:**
   - Admin kesulitan mencocokkan chat dari PIC tertentu dengan data format keep 2 hari terakhir yang belum diisi feedback dan statusnya.

---

## 2. Visi Produk & Tujuan (Goals & Objectives)

### 2.1 Visi
Menghadirkan antarmuka percakapan terpadu bergaya WhatsApp Web modern di dalam Dashboard Sosmed yang terintegrasi langsung dengan database transaksi Google Sheets dan PostgreSQL Serverless, sehingga proses follow-up hingga closing dapat diselesaikan dalam satu layar tanpa membuka aplikasi lain.

### 2.2 Metrik Keberhasilan (KPI)
- **Response Time Admin:** Kecepatan admin merespons update PIC cabang meningkat < 3 menit.
- **Closing Rate:** Peningkatan persentase konversi data keep menjadi faktur penjualan (WIN).
- **Data Integrity:** 100% obrolan, status keep, feedback PIC, dan dokumen Report AR tersimpan sinkron antara Dashboard web, WhatsApp, dan Google Sheets.

---

## 3. User Personas

| Persona | Peran | Kebutuhan Utama |
| :--- | :--- | :--- |
| **Admin Sosmed / FU** | Pengelola obrolan harian di kantor pusat | Memantau seluruh roomchat PIC, membalas chat secara instan, menyertakan foto, membaca mention/quote, serta menginput Status (Kolom V), Feedback (Kolom W), dan Report AR (Kolom N) ke Google Sheets. |
| **PIC Sales Cabang** | Sales toko di cabang fisik | Merespons follow-up admin melalui WhatsApp biasa di HP mereka (termasuk swipe-reply foto/pesan) tanpa perlu menginstal aplikasi tambahan. |
| **Supervisor / Leader** | Pemantau kinerja harian | Melihat daftar roomchat yang belum ditanggapi (*Unresponded*), obrolan selesai (*Resolved*), dan rekap closing transaksi. |

---

## 4. Ruang Lingkup Fitur (Feature Scope & Requirements)

### 4.1 Modul Obrolan (Inbox & Chat View)
- **FR-01 (Sinkronisasi Dua Arah Penuh):** Seluruh pesan masuk dari PIC dan pesan keluar dari Admin (baik dikirim via web maupun via aplikasi WhatsApp resmi di HP admin) tampil secara otomatis di roomchat yang bersangkutan.
- **FR-02 (Pembagi Tanggal Dinamis / Date Dividers):** Ruang obrolan menampilkan penanda tanggal yang rapi dan kontekstual (misal: `Today, 05/11/2026`, `Yesterday, 04/11/2026`, atau format tanggal penuh) untuk membedakan obrolan multi-hari.
- **FR-03 (Native Quoted Reply & Mention Jump):**
  - Balon chat yang membalas pesan tertentu menampilkan kartu kutipan (*quote card*) pengirim asli.
  - Mengklik kartu kutipan akan secara otomatis men-*scroll* layar ke balon chat target yang dikutip dan memberikan animasi sorotan kedip (*pulse highlight animation*).
- **FR-04 (Universal Media & Base64 Pipeline):** Foto dan gambar yang dikirim oleh PIC atau admin disimpan dalam format Base64 secara permanen di database, dapat di-zoom/preview, dan tidak akan kedaluwarsa.
- **FR-05 (Filter & Pencarian Kontak Cepat):**
  - Pencarian global berdasarkan nama kontak, nomor WhatsApp, cabang, atau teks obrolan terakhir.
  - Filter kategori obrolan: `Assigned chats`, `Unresponded`, dan `Resolved`.
  - Filter drop-down cabang (12 cabang toko).

### 4.2 Bilah Samping Format Keep PIC (Right Panel)
- **FR-06 (Integrasi Format Keep 2 Hari Terakhir):**
  - Bilah kanan menampilkan kartu format keep pelanggan yang terkait dengan PIC yang sedang dibuka obrolannya dalam 2 hari terakhir.
  - Sub-tab pemisahan: `Belum Diisi` (Pending) dan `Selesai` (Resolved).
- **FR-07 (Input Cepat Kolom N, V, W):**
  - Form interaktif untuk mengisi **Report AR (Kolom N)**, **Status (Kolom V: WIN, LOSE, ON PROGRESS HOT, dll.)**, dan **Feedback Tim FU (Kolom W)**.
  - Tombol **"Simpan Feedback & AR"** otomatis menyimpan perubahan ke Google Sheets secara real-time via Google Apps Script Web App dan menyiarkan update ke chat room.

### 4.3 Manajemen Status Obrolan (Chat Resolution)
- **FR-08 (Resolve / Reopen Chat):**
  - Admin dapat menandai percakapan sebagai `Resolved` setelah follow-up selesai.
  - Jika PIC mengirimkan chat baru di kemudian hari, status obrolan otomatis berubah kembali menjadi aktif/unresolved.

---

## 5. Non-Functional Requirements (NFR)

1. **Kecepatan & Performa (Performance):**
   - Pengiriman pesan keluar ke WhatsApp melalui WAHA VPS memiliki latensi $\le 500\text{ ms}$.
   - Incremental polling pesan baru memuat data dalam waktu $< 250\text{ ms}$.
2. **Keandalan & Ketahanan (Reliability & Failover):**
   - Arsitektur Dual-Gateway: Jika WAHA VPS mengalami gangguan jaringan (>6 detik), sistem otomatis mengalihkan pengiriman ke WatZap Official Gateway.
   - Database serverless PostgreSQL Neon memiliki ketersediaan tinggi (*high availability*) di Region Singapore.
3. **Keamanan (Security):**
   - Akses API WAHA dilindungi dengan header autentikasi `X-Api-Key`.
   - Endpoint webhook memvalidasi payload dan membersihkan format nomor telepon untuk mencegah injeksi data.
   - Seluruh koneksi database dan API menggunakan protokol HTTPS dan SSL terenkripsi.

---

## 6. Rencana Rilis & Fase Pengembangan

| Fase | Target Delivery | Status |
| :--- | :--- | :--- |
| **Fase 1** | Implementasi WAHA VPS & Sinkronisasi Dua Arah | Selesai (Live) |
| **Fase 2** | Media Base64, Quoted Reply, Resolusi Linked ID (`@lid`) | Selesai (Live) |
| **Fase 3** | Quoted Jump Animation, Date Dividers, Integrasi Format Keep Bilah Kanan | Selesai (Live) |
| **Fase 4** | Sinkronisasi Multi-Bulan Otomatis & Pemulihan Webhook Google Sheets | Selesai (Live) |

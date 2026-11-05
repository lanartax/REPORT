# Arsitektur & Spesifikasi Logika Teknis
## Modul Omnichannel WhatsApp Inbox FU (Follow-Up)

> **File Dokumentasi Teknis Sistem**  
> **Versi:** 2.1.0  
> **Terakhir Diperbarui:** 5 November 2026  
> **Repositori:** `REPORT` / Deployment: Vercel & Ubuntu 24.04 VPS

---

## 1. Ikhtisar Arsitektur Sistem (High-Level Architecture)

Modul ini menggunakan pola **Dual-Gateway Hybrid Microservices Architecture**. Frontend berbasis antarmuka reaktif (Vanilla JavaScript modern), backend beroperasi di serverless Node.js (Vercel) dan server lokal (Node.js runtime), didukung oleh **WAHA (WhatsApp HTTP API)** mandiri di VPS Ubuntu dan **Neon PostgreSQL Serverless (Region Singapore)** sebagai basis data persisten.

```mermaid
flowchart TB
    subgraph ClientTier["🖥️ Frontend Tier (Browser / PWA)"]
        ChatUI["Omnichannel Chat View\n(client.js / style.css)"]
        KeepPanel["Format Keep PIC Sidebar\n(Quick Form Kolom N, V, W)"]
        PollEngine["Smart Incremental Polling\n(?since=timestamp)"]
    end

    subgraph BackendTier["⚙️ Serverless / Backend Tier (Vercel / Local)"]
        SendAPI["api/watzap-send.js\n(Outgoing Router & Failover)"]
        WebhookAPI["api/watzap-webhook.js\n(Incoming Event Ingestion)"]
        KeepStatusAPI["api/update-keep-status.js\n(Status & Feedback Bridge)"]
        ReportARAPI["api/update-report-ar.js\n(AR Invoice Bridge)"]
        LidEngine["LID Cache & Resolving Engine"]
        MediaPipeline["Media Base64 Encoder Engine"]
    end

    subgraph GatewayTier["📡 Gateway & Infrastructure Tier"]
        subgraph VPS["VPS Ubuntu 24.04 (IP: 178.83.188.100)"]
            WAHA["WAHA Docker Container (Port 3088)\nEngine: Chromium WebJS\nSession: default"]
            MediaVol["Docker Volume: /opt/waha-media"]
            DataVol["Docker Volume: /opt/waha-data"]
        end
        WatZap["WatZap Official Cloud API\n(Fallback Secondary Gateway)"]
        GoogleSheetsGAS["Google Apps Script Web App\n(Spreadsheet Database Engine)"]
    end

    subgraph DatabaseTier["🗄️ Persistence Tier"]
        Postgres["Neon PostgreSQL Serverless (Singapore)\n- Tabel: watzap_inbox\n- Tabel: omni_chat_status\n- Tabel: watzap_webhook_logs"]
        LocalCSV["keep-data.csv\n(In-Memory & Local Fast Cache)"]
    end

    subgraph Devices["📱 Devices & Endpoints"]
        AdminPhone["Ponsel Fisik Admin WA\n(Nomor: 6285111426699)"]
        PICPhones["Ponsel PIC Sales Cabang\n(WhatsApp Mobile / Desktop)"]
    end

    %% Outgoing Message Flow
    ChatUI -->|1. Submit Pesan| SendAPI
    SendAPI -->|2. Rute Utama| WAHA
    SendAPI -.->|Failover jika WAHA Offline| WatZap
    WAHA -->|Kirim pesan WA| PICPhones
    SendAPI -->|Catat pesan ADMIN| Postgres

    %% Phone Outgoing Synchronization
    AdminPhone -->|Ketik chat langsung di HP| WAHA
    WAHA -->|Webhook message.any fromMe=true| WebhookAPI

    %% Incoming Message Flow
    PICPhones -->|Balas Pesan / Gambar / Quote| WAHA
    WAHA -->|3. Webhook message.any| WebhookAPI
    WebhookAPI --> LidEngine
    LidEngine <-->|Resolve @lid ke @c.us| WAHA
    WebhookAPI --> MediaPipeline
    MediaPipeline <-->|Download binary media| WAHA
    WebhookAPI -->|4. Insert / Upsert| Postgres

    %% Keep & Google Sheets Update Flow
    KeepPanel -->|Simpan Feedback, Status, AR| KeepStatusAPI
    KeepStatusAPI -->|Update Kolom N, V, W| GoogleSheetsGAS
    KeepStatusAPI -->|Update Cache| LocalCSV

    %% Polling Flow
    PollEngine <-->|Fetch chat baru ?since=ts| WebhookAPI
    WebhookAPI <-->|Query DB| Postgres
    PollEngine -->|Render balon chat & Quoted Jump| ChatUI
```

---

## 2. Rincian Alur Logika (Data Flow & Logic)

### 2.1 Alur Pesan Masuk (Incoming Webhook Flow)
1. PIC membalas pesan WhatsApp dari admin (bisa berupa teks polos, foto bukti fisik, atau balasan kutip/swipe-reply).
2. Kontainer WAHA di VPS menerima paket dari server WhatsApp dan memicu HTTP POST ke:
   `https://workspace-divisisosmed.vercel.app/api/watzap-webhook`
3. **Penyelarasan Arah (`fromMe`):**
   - Jika `fromMe === false`: Pesan berasal dari PIC. Lawan bicara diambil dari `p.from`.
   - Jika `fromMe === true`: Pesan diketik admin langsung dari ponsel fisik WhatsApp. Lawan bicara diambil dari `p.to` dan sender ditandai sebagai `ADMIN`.
4. **Resolusi Linked Identity (`@lid`):**
   - Jika JID pengirim berakhiran `@lid`, sistem memeriksa `wahaLidCache` di memori.
   - Jika belum ada di *cache*, backend melakukan panggilan internal ke:
     `GET http://178.83.188.100:3088/api/contacts?contactId=<LID>`
     untuk mendapatkan nomor telepon asli berakhiran `@c.us` (`628...`) dan nama profil pengirim (`pushname`).
5. **Pemrosesan Media Gambar:**
   - Jika pesan memiliki gambar (`hasMedia: true`), webhook mengunduh file biner dari WAHA menggunakan `X-Api-Key`, lalu mengonversinya menjadi string Base64 Data URI:
     $$\text{Data URI} = \text{"data:"} + \text{MimeType} + \text{";base64,"} + \text{Base64EncodedBuffer}$$
   - String ini disimpan langsung ke kolom `media_url` di PostgreSQL. Pendekatan ini menjamin gambar tidak akan pernah kedaluwarsa (*broken link*).
6. **Ekstraksi Quoted Reply:**
   - Membaca objek `p.replyTo` dari WAHA.
   - Mengekstrak teks yang dikutip, nama pengirim asli, dan JID pesan referensi.
7. **Penyimpanan:**
   - Menyimpan rekaman ke tabel `watzap_inbox` di Neon PostgreSQL menggunakan sintaks `ON CONFLICT (message_id) DO UPDATE`.

---

### 2.2 Alur Pesan Keluar (Outgoing Message Flow)
1. Admin mengetik pesan di bilah obrolan web dan menekan tombol kirim (atau tombol Enter).
2. Frontend memanggil `POST /api/watzap-send` dengan payload:
   ```json
   {
     "phone": "6281347986551",
     "message": "Halo kak, bagaimana update prospek kemarin?",
     "replyTo": {
       "messageId": "true_6281347986551@c.us_3EB0...",
       "text": "Kata nya hari ini kak"
     }
   }
   ```
3. **Rute Utama (WAHA VPS):**
   - Mengirim permintaan ke `POST http://178.83.188.100:3088/api/sendText`.
   - Latensi transmisi rata-rata: $\sim 200\text{ ms} - 450\text{ ms}$.
4. **Rute Failover (WatZap Gateway):**
   - Jika WAHA VPS mengalami gangguan jaringan atau tidak merespons dalam waktu 6 detik (`AbortSignal.timeout(6000)`), sistem otomatis mengalihkan permintaan ke endpoint WatZap Cloud API.
5. Catatan pesan admin langsung disimpan ke PostgreSQL dengan penanda `sender: 'ADMIN'`.

---

### 2.3 Alur Quoted Reply Jump & Focus Highlight
Fitur navigasi kutipan percakapan di frontend:
1. Setiap balon obrolan yang memiliki kutipan dibungkus dengan kartu kutipan interaktif dan atribut event:
   ```html
   <div class="omni-quoted-container clickable" onclick="handleOmniQuotedClick(${idx})">
   ```
2. Fungsi `handleOmniQuotedClick(idx)` mencari pesan referensi target di daftar riwayat aktif:
   - Pencocokan berbasis kesamaan teks kutipan (`cleanText`) dan pembuat pesan.
3. Setelah elemen baris obrolan target (`omniRow_${targetIndex}`) ditemukan:
   - Menjalankan `element.scrollIntoView({ behavior: 'smooth', block: 'center' })`.
   - Menambahkan kelas CSS `.omni-pulse-target` yang memicu animasi kedip latar hijau lembut (*pulse animation*) selama 2.4 detik agar mata admin langsung tertuju pada pesan yang dikutip.

---

### 2.4 Alur Update Format Keep & Google Sheets (Bilah Kanan)
1. Bilah kanan secara otomatis memfilter data keep PIC aktif dalam rentang 2 hari terakhir.
2. Ketika admin mengisi form dan menekan **"Simpan Feedback & AR"**:
   - Memanggil `POST /api/update-keep-status` dengan payload:
     ```json
     {
       "nomorWA": "6281347986551",
       "status": "ON PROGRESS HOT",
       "feedback": "user mau ngambil hari ini",
       "reportAR": "260278899"
     }
     ```
   - Server mengeksekusi Google Apps Script Web App yang aktif:
     `https://script.google.com/macros/s/AKfycbybKV-XtULQW6ct4GCNCCgk2f84AD4CVsSGCOOzweskJFaZtPk0jnEut8tJ80Aq1naoXA/exec`
   - Google Apps Script mencari nomor WA di sheet Keep dan menuliskan secara otomatis:
     - **Kolom N (14):** Nomor Dokumen Report AR
     - **Kolom O (15):** Tanggal Input AR (Otomatis tanggal hari ini)
     - **Kolom V (22):** Status Keep (WIN, LOSE, ON PROGRESS HOT, dll.)
     - **Kolom W (23):** Feedback Tim FU
   - Sistem memperbarui memori lokal dan menyiarkan pesan rekap update langsung ke roomchat PIC.

---

## 3. Skema Basis Data (Neon PostgreSQL)

### 3.1 Tabel `watzap_inbox`
Menyimpan seluruh lalu lintas pesan masuk dan pesan keluar.

```sql
CREATE TABLE IF NOT EXISTS watzap_inbox (
  id SERIAL PRIMARY KEY,
  message_id VARCHAR(100) UNIQUE,
  phone VARCHAR(30) NOT NULL,
  sender_name VARCHAR(100),
  message TEXT NOT NULL,
  sender VARCHAR(10) NOT NULL DEFAULT 'PIC', -- Nilai: 'PIC' atau 'ADMIN'
  is_group BOOLEAN DEFAULT FALSE,
  created_at BIGINT NOT NULL,
  media_url TEXT,                            -- Menyimpan Base64 Data URI (permanen)
  message_type VARCHAR(20) DEFAULT 'text',   -- Nilai: 'text' atau 'image'
  reply_to JSONB,                            -- Objek: {"senderName": "...", "text": "..."}
  raw_payload JSONB                          -- Payload mentah asli untuk audit trail
);

-- Indeks performa untuk kueri real-time
CREATE INDEX IF NOT EXISTS idx_watzap_phone ON watzap_inbox(phone);
CREATE INDEX IF NOT EXISTS idx_watzap_created ON watzap_inbox(created_at);
```

### 3.2 Tabel `omni_chat_status`
Menyimpan status penanganan obrolan oleh admin.

```sql
CREATE TABLE IF NOT EXISTS omni_chat_status (
  phone VARCHAR(30) PRIMARY KEY,
  is_resolved BOOLEAN NOT NULL DEFAULT FALSE,
  resolved_at BIGINT,
  resolved_by VARCHAR(100),
  updated_at BIGINT
);
```

---

## 4. Panduan Disaster Recovery & Rebuild Server dari Nol

Jika suatu saat server VPS berganti, rusak, atau ingin dibangun ulang di lingkungan baru, ikuti panduan langkah demi langkah berikut:

### Langkah 1: Persiapan VPS Baru (Sistem Operasi Ubuntu 22.04 / 24.04 LTS)
Buka terminal SSH VPS baru dan jalankan:
```bash
# 1. Perbarui sistem operasi
apt update && apt upgrade -y

# 2. Instal Docker dan perlengkapan utilitas
apt install -y docker.io curl git ufw

# 3. Aktifkan Docker service
systemctl enable --now docker

# 4. Konfigurasi Firewall Port (Buka port 3088 dan SSH)
ufw allow 22/tcp
ufw allow 3088/tcp
ufw --force enable

# 5. Siapkan direktori penyimpanan data persisten WAHA
mkdir -p /opt/waha-data
mkdir -p /opt/waha-media
chmod -R 777 /opt/waha-data /opt/waha-media
```

### Langkah 2: Jalankan Kontainer WAHA
Eksekusi perintah Docker run resmi berikut (sesuaikan `<IP_VPS_ANDA>` dengan alamat IP publik server):
```bash
docker run -d \
  --name waha \
  --restart unless-stopped \
  -p 3088:3000 \
  -v /opt/waha-data:/app/.sessions \
  -v /opt/waha-media:/tmp/whatsapp-files \
  -e WAHA_BASE_URL=http://<IP_VPS_ANDA>:3088 \
  -e WHATSAPP_FILES_LIFETIME=2592000 \
  -e WHATSAPP_HOOK_URL=https://workspace-divisisosmed.vercel.app/api/watzap-webhook \
  -e WHATSAPP_HOOK_EVENTS=message.any \
  -e WAHA_API_KEY=gadgetmart123 \
  -e WAHA_DASHBOARD_USERNAME=admin \
  -e WAHA_DASHBOARD_PASSWORD=gadgetmart123 \
  devlikeapro/waha
```

### Langkah 3: Sambungkan Sesi WhatsApp Resmi
1. Buka antarmuka dashboard WAHA di browser:  
   `http://<IP_VPS_ANDA>:3088/dashboard/`
2. Masukkan akun:
   - **Username:** `admin`
   - **Password:** `gadgetmart123`
3. Pilih sesi `default`, lalu klik tombol **Start / Scan QR**.
4. Scan kode QR menggunakan aplikasi WhatsApp nomor official kantor (`6285111426699`).
5. Pastikan status sesi telah berubah menjadi **WORKING** dan `engine: WEBJS`.

### Langkah 4: Sinkronisasi Konfigurasi Aplikasi (Environment Variables)
Pastikan berkas `.env` atau *Environment Variables* di platform deployment (Vercel) memuat konfigurasi:
```ini
DATABASE_URL="postgresql://neondb_owner:npg_QwTqPU9m6Yar@ep-round-night-aopjrfia-pooler.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require"
WAHA_BASE_URL="http://178.83.188.100:3088"
WAHA_API_KEY="gadgetmart123"
```

Modul omnichannel secara otomatis akan kembali beroperasi normal tanpa kehilangan riwayat pesan lama (karena riwayat tersimpan aman di Neon PostgreSQL).

# WhatsApp Omnichannel Module: PRD & Architecture Blueprint

> **Versi:** 2.0.0  
> **Tanggal:** 4 November 2026  
> **Status:** Production Ready  
> **Author:** Antigravity Engineering  
> **Target Pengguna:** Tim Sosmed / Admin Gadget Mart & Engineer Maintenance  

---

## 1. Executive Summary & Latar Belakang

Modul **Omnichannel WhatsApp** pada Dashboard Daily Achievement Divisi Sosmed dirancang sebagai pusat komunikasi terpusat (two-way chat interface) antara Admin Divisi Sosmed dengan para PIC Sales di toko/cabang fisik.

### Problem Statement Sebelumnya
1. **Keterbatasan Gateway WatZap Pihak Ketiga:**
   - WatZap API versi unofficial tidak meneruskan metadata quote reply (`replyTo` / `contextInfo`) secara konsisten pada webhook pesan masuk.
   - Pesan yang dikirim admin langsung dari aplikasi WhatsApp di HP tidak memicu webhook pada WatZap, menyebabkan riwayat chat di website menjadi tidak lengkap (*out-of-sync*).
   - Pengiriman gambar terbatasi oleh ukuran dan ketergantungan link hosting publik.
2. **Kendala WhatsApp Modern (Linked Identity / `@lid`):**
   - WhatsApp versi modern menggunakan enkripsi identitas berbasis Linked ID (`...41962333818961@lid`). Webhook yang membaca raw JID menyebabkan pembuatan room "hantu" alih-alih tergabung ke nomor telepon PIC asli (`628...`).

### Solusi Arsitektur Baru
Mengadopsi **Dual-Gateway Hybrid Architecture**:
- **Gateway Utama Omnichannel (Two-Way Interactive Chat):** Menggunakan **WAHA (WhatsApp HTTP API)** yang di-host mandiri pada VPS pribadi Ubuntu 24.04 (`178.83.188.100:3088`). Menyediakan native quoted reply, native image handling, resolusi kontak LID, dan notifikasi pesan instan baik yang dikirim dari web maupun langsung dari aplikasi WhatsApp di ponsel admin.
- **Gateway Sekunder (Fallback & Automated Bot):** Tetap menggunakan **WatZap API** untuk bot broadcast follow-up otomatis agar proses follow-up tidak terganggu, serta menjadi *failover route* jika VPS mengalami gangguan jaringan.

---

## 2. Product Requirements Document (PRD)

### 2.1 User Personas
1. **Admin Sosmed / CS Dashboard:**
   - Memantau obrolan masuk dari puluhan PIC cabang secara real-time.
   - Membalas pesan, mengirim gambar brosur/faktur, dan melihat konteks pesan mana yang sedang dibahas melalui kartu kutipan (*quote card*).
   - Menandai obrolan sebagai "Selesai" (*Resolved*) atau "Buka Kembali" (*Unresolved*).
2. **PIC Sales (Toko Cabang):**
   - Mengirim update keep barang, konfirmasi stok, dan foto bukti fisik ke nomor WhatsApp resmi kantor.
   - Menggunakan fitur bawaan WhatsApp (Swipe-to-Reply) untuk merespons pesan tertentu dari admin.
3. **Automated Follow-up Bot:**
   - Mengirim reminder otomatis berdasarkan data keep barang tanpa mencampuri session interaktif admin.

### 2.2 Functional Requirements (FR)

| ID | Fitur | Deskripsi Kebutuhan |
| :--- | :--- | :--- |
| **FR-01** | **Two-Way Synchronization** | Sinkronisasi penuh dua arah: Pesan masuk dari PIC dan pesan keluar dari admin (baik dikirim lewat web dashboard maupun lewat aplikasi WhatsApp resmi di ponsel) harus tercatat dan tampil di roomchat PIC yang tepat. |
| **FR-02** | **Native Quoted Reply** | Mendeteksi pesan yang di-swipe reply oleh lawan bicara, mengekstrak teks asli dan pengirimnya, lalu merendernya dalam bentuk kartu kutipan interaktif di atas balon chat. |
| **FR-03** | **Universal Media Rendering** | Gambar/foto yang dikirim PIC atau admin diunduh dan di-encode ke format Base64 Data URI (`data:image/jpeg;base64,...`) secara permanen ke database, mencegah *broken image* / link kedaluwarsa. |
| **FR-04** | **Automatic LID Resolution** | Kontak dengan format Linked Identity (`@lid`) wajib di-resolusikan secara otomatis ke nomor telepon resmi WhatsApp (`@c.us` / `628...`) dan nama profil pengirimnya secara real-time melalui WAHA Contacts API. |
| **FR-05** | **Permanent Chat Resolution** | Status `Resolved` / `Unresolved` chat disimpan di PostgreSQL Neon (`omni_chat_status`) agar status tidak hilang saat browser di-refresh atau server di-restart. |
| **FR-06** | **Search & Contact Filtering** | Admin dapat mencari roomchat berdasarkan nomor telepon, nama PIC, cabang, atau cuplikan isi pesan terakhir. |
| **FR-07** | **Dual-Gateway Failover** | Pengiriman pesan keluar dari web dashboard memprioritaskan WAHA VPS (~200ms latency). Jika WAHA tidak merespons dalam 6 detik, sistem otomatis melakukan fallback ke WatZap API. |
| **FR-08** | **Non-Intrusive Polling** | Frontend melakukan smart incremental polling (`?since=<timestamp>`) setiap 2.5–3 detik untuk menghemat bandwidth dan kuota serverless Neon DB. |

### 2.3 Non-Functional Requirements (NFR)
- **Latency:** Pengiriman pesan keluar via WAHA VPS harus terkirim ke WhatsApp user dalam waktu < 500 ms.
- **Reliability:** Sesi WAHA di VPS berjalan dalam kontainer Docker dengan konfigurasi `--restart unless-stopped` dan volume persisten.
- **Database Scalability:** Menyimpan riwayat chat di Neon PostgreSQL Serverless dengan indeks pada kolom `phone`, `created_at`, dan constraint unik pada `message_id`.

---

## 3. System Architecture & Diagram

### 3.1 High-Level Architecture Diagram

```mermaid
flowchart TB
    subgraph ClientLayer["🖥️ Frontend Tier (Browser)"]
        UI["Web Dashboard Omnichannel\n(Vanilla JS / CSS)"]
        PollEngine["Smart Incremental Polling\n(?since=timestamp)"]
    end

    subgraph BackendLayer["⚙️ Backend Tier (Vercel / Node.js)"]
        SendHandler["api/watzap-send.js\n(Outgoing Router)"]
        WebhookHandler["api/watzap-webhook.js\n(Incoming Ingestion)"]
        ResolveHandler["api/omni-resolve.js\n(Status Manager)"]
        LidResolver["LID Resolver & Cache Engine\n(Map + WAHA Contacts)"]
        MediaConverter["Media to Base64 Pipeline"]
    end

    subgraph GatewayLayer["📡 Message Gateway Tier"]
        subgraph VPS["VPS Ubuntu 24.04 (178.83.188.100)"]
            WAHA["WAHA Docker Container\n(Port 3088 / Chromium WebJS)"]
            MediaStorage["Persisted Volume\n/opt/waha-media"]
            SessionStorage["Persisted Volume\n/opt/waha-data"]
        end
        WatZapAPI["WatZap Official Cloud API\n(Fallback & Follow-up Bot)"]
    end

    subgraph DatabaseLayer["🗄️ Persistence Tier"]
        NeonDB["Neon PostgreSQL Serverless (Singapore)\n- watzap_inbox\n- omni_chat_status\n- watzap_webhook_logs"]
    end

    subgraph UserDevices["📱 User Devices"]
        AdminPhone["Admin WhatsApp Phone\n(6285111426699)"]
        PICPhone["PIC Sales Phone\n(Customer / Branch)"]
    end

    %% Outgoing Flow
    UI -->|1. POST send message| SendHandler
    SendHandler -->|2. Primary Route| WAHA
    SendHandler -.->|Failover Fallback| WatZapAPI
    WAHA -->|3. Send WhatsApp packet| PICPhone
    SendHandler -->|Save ADMIN record| NeonDB

    %% Phone Outgoing Flow
    AdminPhone -->|Chat directly from WA app| WAHA
    WAHA -->|Webhook event: message.any| WebhookHandler

    %% Incoming Flow
    PICPhone -->|4. Reply / Send Image / Quote| WAHA
    WAHA -->|5. HTTP Webhook Payload| WebhookHandler
    WebhookHandler --> LidResolver
    LidResolver <-->|Query Contact Info| WAHA
    WebhookHandler --> MediaConverter
    MediaConverter <-->|Fetch raw file binary| WAHA
    WebhookHandler -->|6. Store Ingested Message| NeonDB

    %% Polling Flow
    PollEngine <-->|7. GET ?since=ts| WebhookHandler
    WebhookHandler <-->|Query recent messages| NeonDB
    PollEngine -->|Render bubbles & quotes| UI
```

---

## 4. Sequence Diagrams

### 4.1 Pesan Masuk dari WhatsApp PIC (Termasuk Quote & Gambar)

```mermaid
sequenceDiagram
    autonumber
    actor PIC as PIC Sales (WhatsApp)
    participant WAHA as WAHA VPS (178.83.188.100)
    participant Webhook as api/watzap-webhook.js
    participant Neon as Neon PostgreSQL
    actor Admin as Admin Dashboard (Browser)

    PIC->>WAHA: Kirim chat dengan Swipe-Reply / Foto
    WAHA->>Webhook: POST Webhook (event: message.any)
    Note over Webhook: Deteksi fromMe = false (Lawan bicara = p.from)
    alt Kontak bertipe @lid
        Webhook->>WAHA: GET /api/contacts?contactId=...
        WAHA-->>Webhook: Return id=@c.us (628...), pushname
    end
    alt Pesan memiliki Media (Gambar)
        Webhook->>WAHA: GET /api/files/... (Headers: X-Api-Key)
        WAHA-->>Webhook: Return binary file
        Note over Webhook: Convert buffer ke Base64 Data URI
    end
    alt Pesan memiliki Quoted Reply (p.replyTo)
        Note over Webhook: Ekstraksi teks kutipan & pemetaan nama pengirim
    end
    Webhook->>Neon: INSERT INTO watzap_inbox (ON CONFLICT DO UPDATE)
    Admin->>Webhook: GET /api/watzap-webhook?since=...
    Webhook->>Neon: SELECT * FROM watzap_inbox WHERE created_at > since
    Neon-->>Webhook: Return new messages
    Webhook-->>Admin: Render balon chat + Quote Card + Gambar
```

### 4.2 Pesan Keluar dari Web Dashboard Admin

```mermaid
sequenceDiagram
    autonumber
    actor Admin as Admin Dashboard (Browser)
    participant SendAPI as api/watzap-send.js
    participant WAHA as WAHA VPS (178.83.188.100)
    participant WatZap as WatZap API (Fallback)
    participant Neon as Neon PostgreSQL
    actor PIC as PIC Sales (WhatsApp)

    Admin->>SendAPI: POST /api/watzap-send (phone, message, replyTo)
    Note over SendAPI: Normalisasi nomor ke format 628...
    SendAPI->>WAHA: POST /api/sendText (chatId: 628...@c.us)
    alt WAHA Berhasil (HTTP 201)
        WAHA-->>SendAPI: 201 Created (id: true_628...@c.us_3EB0...)
        WAHA->>PIC: Kirim pesan WhatsApp real-time
    else WAHA Timeout / Offline (>6s)
        SendAPI->>WatZap: POST /v1/send_message (Fallback)
        WatZap->>PIC: Kirim via WatZap Cloud
    end
    SendAPI->>Neon: INSERT INTO watzap_inbox (sender: ADMIN, message_id)
    SendAPI-->>Admin: 200 OK (success: true, gateway: WAHA_VPS)
```

---

## 5. Database Schema & Data Models

Tabel tersimpan pada database **Neon PostgreSQL Serverless (Region: Singapore)**.

### 5.1 Tabel `watzap_inbox` (Penyimpanan Chat Masuk & Keluar)
```sql
CREATE TABLE IF NOT EXISTS watzap_inbox (
  id SERIAL PRIMARY KEY,
  message_id VARCHAR(100) UNIQUE,
  phone VARCHAR(30) NOT NULL,
  sender_name VARCHAR(100),
  message TEXT NOT NULL,
  sender VARCHAR(10) NOT NULL DEFAULT 'PIC', -- 'PIC' atau 'ADMIN'
  is_group BOOLEAN DEFAULT FALSE,
  created_at BIGINT NOT NULL,
  media_url TEXT,                            -- Menyimpan Base64 Data URI atau HTTPS URL
  message_type VARCHAR(20) DEFAULT 'text',   -- 'text' atau 'image'
  reply_to JSONB,                            -- Format: {"senderName": "...", "text": "..."}
  raw_payload JSONB                          -- Full payload asli untuk audit & debugging
);

CREATE INDEX IF NOT EXISTS idx_watzap_phone ON watzap_inbox(phone);
CREATE INDEX IF NOT EXISTS idx_watzap_created ON watzap_inbox(created_at);
```

### 5.2 Tabel `omni_chat_status` (Status Resolusi Roomchat)
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

## 6. Detail Logika & Algoritma Utama

### 6.1 WhatsApp Linked Identity (`@lid`) Resolver
WhatsApp modern mengenkripsi nomor pengguna menjadi format `@lid` (contoh: `152419883311330@lid`).
- **Logika:**
  1. Jika JID diakhiri dengan `@lid`, cek `wahaLidCache` di memory.
  2. Jika *cache miss*, sistem melakukan panggilan internal ke:
     `GET http://178.83.188.100:3088/api/contacts?session=default&contactId=<str>`
  3. Mengambil properti `data.id` (yang memuat JID telepon resmi `628...@c.us`) dan properti nama (`data.pushname || data.name`).
  4. Menyimpan pemetaan ke *cache* agar panggilan berikutnya bernilai $O(1)$ tanpa latensi HTTP.

### 6.2 Normalisasi Arah Pesan (`fromMe` Handling)
- Jika `p.fromMe === false` (pesan masuk dari PIC):
  - Lawan bicara = `p.from` (atau `p._data.id.remote`).
  - Pengirim = `PIC`.
- Jika `p.fromMe === true` (pesan keluar yang diketik admin langsung dari aplikasi WhatsApp di HP):
  - Lawan bicara = `p.to` (atau `p._data.id.remote`).
  - Pengirim = `ADMIN`.
- **Hasil:** Admin dapat membalas chat langsung dari HP dan seluruh pesan balasan tetap rapi masuk ke roomchat PIC terkait di website.

### 6.3 Media-to-Base64 Conversion Pipeline
- URL lokal yang dihasilkan kontainer (`http://localhost:3000/api/files/...`) tidak dapat dibuka oleh browser klien dan dihapus WAHA setelah jangka waktu tertentu.
- **Logika Penanganan:**
  Saat webhook menerima `p.hasMedia = true`, webhook backend langsung menarik binary file dari WAHA dengan header `X-Api-Key: gadgetmart123`, mengubahnya menjadi string Base64 Data URI:
  $$\text{Data URI} = \text{"data:"} + \text{MimeType} + \text{";base64,"} + \text{Base64EncodedBuffer}$$
  String ini disimpan langsung ke kolom `media_url TEXT` di database. Gambar tersimpan permanen dan dapat di-render seketika di semua perangkat tanpa memerlukan akses ke VPS atau autentikasi.

---

## 7. Panduan Rebuild / Disaster Recovery (Cara Pasang Baru dari Nol)

Jika sewaktu-waktu VPS diganti atau Anda ingin membangun modul ini di server baru:

### Langkah 1: Persiapan VPS Baru (Ubuntu 24.04 LTS)
```bash
# Update paket & instalasi Docker
apt update && apt upgrade -y
apt install -y docker.io curl git
systemctl enable --now docker

# Buat folder penyimpanan persisten
mkdir -p /opt/waha-data
mkdir -p /opt/waha-media
chmod -R 777 /opt/waha-data /opt/waha-media
```

### Langkah 2: Jalankan Kontainer WAHA
Jalankan satu perintah Docker berikut:
```bash
docker run -d \
  --name waha \
  --restart unless-stopped \
  -p 3088:3000 \
  -v /opt/waha-data:/app/.sessions \
  -v /opt/waha-media:/tmp/whatsapp-files \
  -e WAHA_BASE_URL=http://<IP_VPS_ANDA>:3088 \
  -e WHATSAPP_FILES_LIFETIME=2592000 \
  -e WHATSAPP_HOOK_URL=https://<DOMAIN_WEB_ANDA>/api/watzap-webhook \
  -e WHATSAPP_HOOK_EVENTS=message.any \
  -e WAHA_API_KEY=gadgetmart123 \
  -e WAHA_DASHBOARD_USERNAME=admin \
  -e WAHA_DASHBOARD_PASSWORD=gadgetmart123 \
  devlikeapro/waha
```

### Langkah 3: Scan QR Code WhatsApp
1. Buka browser: `http://<IP_VPS_ANDA>:3088/dashboard/`
2. Login dengan kredensial:
   - **Username:** `admin`
   - **Password:** `gadgetmart123`
3. Masuk ke menu sesi `default`, lalu scan QR code menggunakan aplikasi WhatsApp nomor kantor (`6285111426699`).
4. Pastikan status sesi berubah menjadi **WORKING** dan `state: CONNECTED`.

### Langkah 4: Environment Variables Aplikasi
Pastikan file `.env` atau *Environment Variables* di Vercel memuat:
```ini
DATABASE_URL="postgresql://neondb_owner:npg_QwTqPU9m6Yar@ep-round-night-aopjrfia-pooler.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require"
WAHA_BASE_URL="http://178.83.188.100:3088"
WAHA_API_KEY="gadgetmart123"
```
Setelah itu, modul omnichannel akan langsung aktif dan berfungsi 100%.


-- ====================================================================
-- GADGETMART SOSMED WORKSPACE: NEON POSTGRESQL DATABASE SCHEMA
-- ====================================================================
-- Salin dan jalankan seluruh script ini di "SQL Editor" pada console Neon.tech

-- 1. Tabel Target Bulanan (Tim & Individual PIC Sales)
CREATE TABLE IF NOT EXISTS monthly_targets (
  key VARCHAR(50) PRIMARY KEY,
  month INT NOT NULL,
  year INT NOT NULL,
  team_target NUMERIC DEFAULT 8447356034,
  sales_targets JSONB DEFAULT '{}',
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Tabel Foto Profil PIC Sales
CREATE TABLE IF NOT EXISTS sales_photos (
  pic_name VARCHAR(100) PRIMARY KEY,
  photo_url TEXT NOT NULL,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 3. Tabel Arsip Pencapaian Bulanan
CREATE TABLE IF NOT EXISTS monthly_archives (
  id VARCHAR(100) PRIMARY KEY,
  month INT,
  year INT,
  title TEXT,
  data JSONB NOT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Data Awal Standar Target Oktober 2026
INSERT INTO monthly_targets (key, month, year, team_target, sales_targets, updated_at)
VALUES (
  '2026-10',
  10,
  2026,
  8447356034,
  '{
    "BJM.SOS.005": 1081261572,
    "BJM.SOS.015": 1081261572,
    "BJM.SOS.017": 1081261572,
    "BJM.SOS.004": 867261886,
    "BJM.SOS.007": 867261886,
    "BJM.SOS.016": 867261886,
    "BJM.SOS.018": 867261886,
    "BJM.SOS.019": 867261886,
    "BJM.SOS.022": 867261886
  }'::jsonb,
  NOW()
)
ON CONFLICT (key) DO NOTHING;

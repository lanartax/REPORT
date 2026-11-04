const { neon } = require('@neondatabase/serverless');

const DEFAULT_DATABASE_URL = 'postgresql://neondb_owner:npg_QwTqPU9m6Yar@ep-round-night-aopjrfia-pooler.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require';

let dbInitialized = false;

function getDb() {
  const databaseUrl = process.env.DATABASE_URL || DEFAULT_DATABASE_URL;
  if (!databaseUrl) return null;
  return neon(databaseUrl);
}

async function initDb(sql) {
  if (!sql || dbInitialized) return;
  try {
    await sql`
      CREATE TABLE IF NOT EXISTS monthly_targets (
        key VARCHAR(50) PRIMARY KEY,
        month INT NOT NULL,
        year INT NOT NULL,
        team_target NUMERIC,
        sales_targets JSONB,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `;
    await sql`
      CREATE TABLE IF NOT EXISTS sales_photos (
        pic_name VARCHAR(100) PRIMARY KEY,
        photo_url TEXT NOT NULL,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `;
    await sql`
      CREATE TABLE IF NOT EXISTS monthly_archives (
        id VARCHAR(100) PRIMARY KEY,
        month INT,
        year INT,
        title TEXT,
        data JSONB NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `;
    await sql`
      CREATE TABLE IF NOT EXISTS watzap_inbox (
        id SERIAL PRIMARY KEY,
        message_id VARCHAR(100) UNIQUE,
        phone VARCHAR(30) NOT NULL,
        sender_name VARCHAR(100),
        message TEXT NOT NULL,
        sender VARCHAR(10) NOT NULL DEFAULT 'PIC',
        is_group BOOLEAN DEFAULT FALSE,
        created_at BIGINT NOT NULL,
        media_url TEXT,
        message_type VARCHAR(20) DEFAULT 'text',
        reply_to JSONB,
        raw_payload JSONB
      );
    `;
    await sql`
      ALTER TABLE watzap_inbox ADD COLUMN IF NOT EXISTS media_url TEXT;
    `;
    await sql`
      ALTER TABLE watzap_inbox ADD COLUMN IF NOT EXISTS message_type VARCHAR(20) DEFAULT 'text';
    `;
    await sql`
      ALTER TABLE watzap_inbox ADD COLUMN IF NOT EXISTS reply_to JSONB;
    `;
    await sql`
      ALTER TABLE watzap_inbox ADD COLUMN IF NOT EXISTS raw_payload JSONB;
    `;
    await sql`
      CREATE TABLE IF NOT EXISTS watzap_webhook_logs (
        id SERIAL PRIMARY KEY,
        payload JSONB,
        created_at BIGINT
      );
    `;
    await sql`
      CREATE TABLE IF NOT EXISTS omni_chat_status (
        phone VARCHAR(32) PRIMARY KEY,
        is_resolved BOOLEAN DEFAULT FALSE,
        resolved_at BIGINT,
        resolved_by VARCHAR(128),
        updated_at BIGINT
      );
    `;
    await sql`
      CREATE INDEX IF NOT EXISTS idx_omni_chat_resolved ON omni_chat_status(is_resolved);
    `;
    await sql`
      CREATE INDEX IF NOT EXISTS idx_watzap_phone_ts ON watzap_inbox(phone, created_at);
    `;
    await sql`
      CREATE INDEX IF NOT EXISTS idx_watzap_created_at ON watzap_inbox(created_at);
    `;
    dbInitialized = true;
  } catch (err) {
    console.error("Neon DB Init Error:", err);
  }
}

module.exports = { getDb, initDb };

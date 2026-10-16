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
    dbInitialized = true;
  } catch (err) {
    console.error("Neon DB Init Error:", err);
  }
}

module.exports = { getDb, initDb };

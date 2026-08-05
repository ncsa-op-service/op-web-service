import { Pool } from "pg";

if (!process.env.DATABASE_URL) {
  throw new Error("ไม่พบ DATABASE_URL");
}

export const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
});

export async function initializeDatabase(): Promise<void> {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS uploads (
      id SERIAL PRIMARY KEY,
      original_file_name VARCHAR(255) NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS line_links (
      id SERIAL PRIMARY KEY,
      upload_id INTEGER REFERENCES uploads(id) ON DELETE CASCADE,
      row_number INTEGER NOT NULL,
      name VARCHAR(255),
      original_url TEXT NOT NULL,
      final_url TEXT,
      status VARCHAR(30) NOT NULL DEFAULT 'WAITING',
      message TEXT,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      checked_at TIMESTAMPTZ
    );
  `);
}
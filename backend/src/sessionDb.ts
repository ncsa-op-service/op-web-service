import { pool } from "./db.js";

export async function initializeSessionDatabase(): Promise<void> {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS user_sessions (
      id UUID PRIMARY KEY,
      user_id INTEGER NOT NULL
        REFERENCES users(id)
        ON DELETE CASCADE,
      token_hash CHAR(64) UNIQUE NOT NULL,
      login_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      last_seen_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      logout_at TIMESTAMPTZ,
      ip_address VARCHAR(100),
      user_agent TEXT,
      legacy_log_id INTEGER
    );
  `);

  await pool.query(`
    ALTER TABLE user_sessions
      ADD COLUMN IF NOT EXISTS legacy_log_id INTEGER;
  `);

  await pool.query(`
    ALTER TABLE admin_logs
      ADD COLUMN IF NOT EXISTS session_id UUID
        REFERENCES user_sessions(id)
        ON DELETE SET NULL;
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS idx_user_sessions_user_login
    ON user_sessions (user_id, login_at DESC);
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS idx_user_sessions_presence
    ON user_sessions (logout_at, last_seen_at DESC);
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS idx_admin_logs_session
    ON admin_logs (session_id, created_at DESC);
  `);

  await pool.query(`
    CREATE UNIQUE INDEX IF NOT EXISTS idx_user_sessions_legacy_log
    ON user_sessions (legacy_log_id)
    WHERE legacy_log_id IS NOT NULL;
  `);

  /*
    นำ Login Log เก่าเข้ามาเป็น Session แบบ Offline
    เพื่อให้หน้า Super Admin นับประวัติเก่าได้ทันที
  */
  await pool.query(`
    INSERT INTO user_sessions (
      id,
      user_id,
      token_hash,
      login_at,
      last_seen_at,
      logout_at,
      ip_address,
      user_agent,
      legacy_log_id
    )
    SELECT
      (
        SUBSTRING(MD5('legacy-session-' || l.id::TEXT), 1, 8) || '-' ||
        SUBSTRING(MD5('legacy-session-' || l.id::TEXT), 9, 4) || '-' ||
        SUBSTRING(MD5('legacy-session-' || l.id::TEXT), 13, 4) || '-' ||
        SUBSTRING(MD5('legacy-session-' || l.id::TEXT), 17, 4) || '-' ||
        SUBSTRING(MD5('legacy-session-' || l.id::TEXT), 21, 12)
      )::UUID,
      l.admin_id,
      MD5('legacy-token-a-' || l.id::TEXT) ||
        MD5('legacy-token-b-' || l.id::TEXT),
      l.created_at,
      l.created_at,
      l.created_at,
      l.ip_address,
      'legacy-import',
      l.id
    FROM admin_logs l
    WHERE l.action_type = 'login'
      AND l.admin_id IS NOT NULL
    ON CONFLICT (legacy_log_id)
      WHERE legacy_log_id IS NOT NULL
    DO NOTHING;
  `);

  await pool.query(`
    UPDATE admin_logs l
    SET session_id = s.id
    FROM user_sessions s
    WHERE s.legacy_log_id = l.id
      AND l.session_id IS NULL;
  `);
}

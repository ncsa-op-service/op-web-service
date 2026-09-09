import { createHash } from "node:crypto";
import { Router, type Request } from "express";
import { pool } from "../db.js";

const router = Router();

type AuthenticatedSession = {
  sessionId: string;
  userId: number;
  name: string;
  email: string;
  role: "super_admin" | "editor" | "viewer";
};

function hashToken(token: string): string {
  return createHash("sha256").update(token).digest("hex");
}

function bearerToken(req: Request): string | null {
  const authorization = req.get("authorization");
  if (!authorization?.startsWith("Bearer ")) return null;
  return authorization.slice(7).trim() || null;
}

async function authenticate(req: Request): Promise<AuthenticatedSession | null> {
  const token = bearerToken(req);
  if (!token) return null;

  const result = await pool.query(
    `SELECT
       s.id AS session_id,
       u.id AS user_id,
       u.name,
       u.email,
       u.role
     FROM user_sessions s
     JOIN users u ON u.id = s.user_id
     WHERE s.token_hash = $1
       AND s.logout_at IS NULL
     LIMIT 1`,
    [hashToken(token)]
  );

  if (result.rows.length === 0) return null;

  const row = result.rows[0];
  return {
    sessionId: row.session_id,
    userId: row.user_id,
    name: row.name,
    email: row.email,
    role: row.role,
  };
}

router.post("/heartbeat", async (req, res) => {
  try {
    const session = await authenticate(req);
    if (!session) return res.status(401).json({ message: "Session ไม่ถูกต้องหรือหมดอายุ" });

    await pool.query(
      `UPDATE user_sessions
       SET last_seen_at = NOW()
       WHERE id = $1 AND logout_at IS NULL`,
      [session.sessionId]
    );

    return res.json({ message: "Heartbeat updated", serverTime: new Date().toISOString() });
  } catch (error) {
    console.error("Heartbeat error:", error);
    return res.status(500).json({ message: "ไม่สามารถอัปเดตสถานะได้" });
  }
});

router.post("/logout", async (req, res) => {
  const client = await pool.connect();

  try {
    const session = await authenticate(req);
    if (!session) return res.status(204).end();

    await client.query("BEGIN");
    await client.query(
      `UPDATE user_sessions
       SET logout_at = NOW(), last_seen_at = NOW()
       WHERE id = $1 AND logout_at IS NULL`,
      [session.sessionId]
    );
    await client.query(
      `INSERT INTO admin_logs (
         admin_id, session_id, action_type, message, ip_address
       )
       SELECT $1, $2, 'logout', $3, ip_address
       FROM user_sessions
       WHERE id = $2`,
      [session.userId, session.sessionId, `${session.name} ออกจากระบบ`]
    );
    await client.query("COMMIT");
    return res.json({ message: "Logout สำเร็จ" });
  } catch (error) {
    await client.query("ROLLBACK").catch(() => undefined);
    console.error("Logout error:", error);
    return res.status(500).json({ message: "ไม่สามารถ Logout ได้" });
  } finally {
    client.release();
  }
});

router.get("/overview", async (req, res) => {
  try {
    const currentSession = await authenticate(req);
    if (!currentSession) return res.status(401).json({ message: "กรุณาเข้าสู่ระบบ" });
    if (currentSession.role !== "super_admin") {
      return res.status(403).json({ message: "ไม่มีสิทธิ์ดูสถานะผู้ใช้งาน" });
    }

    /*
      Login ที่เกิดจากผู้ใช้คนเดียวกันห่างกันไม่เกิน 15 นาที
      จะถูกนับเป็นการเข้าใช้งานรอบเดียวกัน
    */
    const result = await pool.query(`
      WITH ordered_sessions AS (
        SELECT
          s.*,
          LAG(s.login_at) OVER (
            PARTITION BY s.user_id
            ORDER BY s.login_at
          ) AS previous_login_at
        FROM user_sessions s
      ), marked_sessions AS (
        SELECT
          *,
          CASE
            WHEN previous_login_at IS NULL
              OR login_at - previous_login_at > INTERVAL '15 minutes'
            THEN 1 ELSE 0
          END AS starts_new_visit
        FROM ordered_sessions
      ), numbered_sessions AS (
        SELECT
          *,
          SUM(starts_new_visit) OVER (
            PARTITION BY user_id
            ORDER BY login_at
            ROWS UNBOUNDED PRECEDING
          ) AS visit_number
        FROM marked_sessions
      ), visits AS (
        SELECT
          user_id,
          visit_number,
          MIN(login_at) AS started_at,
          MAX(last_seen_at) AS last_seen_at,
          CASE
            WHEN COUNT(*) FILTER (WHERE logout_at IS NULL) > 0 THEN NULL
            ELSE MAX(logout_at)
          END AS logout_at,
          (ARRAY_AGG(ip_address ORDER BY login_at DESC))[1] AS ip_address,
          COUNT(*)::INTEGER AS merged_logins,
          BOOL_OR(
            logout_at IS NULL
            AND last_seen_at > NOW() - INTERVAL '90 seconds'
          ) AS is_online,
          BOOL_OR(
            logout_at IS NULL
            AND last_seen_at <= NOW() - INTERVAL '90 seconds'
            AND last_seen_at > NOW() - INTERVAL '10 minutes'
          ) AS is_idle
        FROM numbered_sessions
        GROUP BY user_id, visit_number
      ), account_summary AS (
        SELECT
          user_id,
          COUNT(*)::INTEGER AS visit_count,
          MAX(last_seen_at) AS last_seen_at,
          BOOL_OR(is_online) AS is_online,
          BOOL_OR(is_idle) AS is_idle
        FROM visits
        GROUP BY user_id
      )
      SELECT
        u.id,
        u.name,
        u.email,
        u.role,
        COALESCE(a.visit_count, 0) AS visit_count,
        a.last_seen_at,
        COALESCE(a.is_online, FALSE) AS is_online,
        COALESCE(a.is_idle, FALSE) AS is_idle,
        COALESCE(
          (
            SELECT JSON_AGG(
              JSON_BUILD_OBJECT(
                'visitNumber', v.visit_number,
                'startedAt', v.started_at,
                'lastSeenAt', v.last_seen_at,
                'logoutAt', v.logout_at,
                'ipAddress', v.ip_address,
                'mergedLogins', v.merged_logins,
                'isOnline', v.is_online,
                'isIdle', v.is_idle
              ) ORDER BY v.started_at DESC
            )
            FROM visits v
            WHERE v.user_id = u.id
          ),
          '[]'::JSON
        ) AS visits
      FROM users u
      LEFT JOIN account_summary a ON a.user_id = u.id
      ORDER BY
        COALESCE(a.visit_count, 0) DESC,
        COALESCE(a.is_online, FALSE) DESC,
        COALESCE(a.is_idle, FALSE) DESC,
        a.last_seen_at DESC NULLS LAST,
        u.email ASC;
    `);

    return res.json({ accounts: result.rows, generatedAt: new Date().toISOString() });
  } catch (error) {
    console.error("Session overview error:", error);
    return res.status(500).json({ message: "ไม่สามารถโหลดสถานะผู้ใช้งานได้" });
  }
});

export default router;

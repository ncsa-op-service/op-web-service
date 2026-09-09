import { createHash, randomBytes, randomUUID } from "node:crypto";
import { Router } from "express";
import bcrypt from "bcryptjs";
import { pool } from "../db.js";

const router = Router();

function hashToken(token: string): string {
  return createHash("sha256").update(token).digest("hex");
}

router.post("/login", async (req, res) => {
  const client = await pool.connect();

  try {
    const email = String(req.body?.email ?? "").trim().toLowerCase();
    const password = String(req.body?.password ?? "");

    if (!email || !password) {
      return res.status(400).json({ message: "กรุณากรอก Email และ Password" });
    }

    const result = await client.query(
      `SELECT id, name, email, password_hash, role
       FROM users
       WHERE LOWER(email) = $1
       LIMIT 1`,
      [email]
    );

    if (result.rows.length === 0) {
      return res.status(401).json({ message: "Email หรือ Password ไม่ถูกต้อง" });
    }

    const user = result.rows[0];
    const passwordMatch = await bcrypt.compare(password, user.password_hash);

    if (!passwordMatch) {
      return res.status(401).json({ message: "Email หรือ Password ไม่ถูกต้อง" });
    }

    const token = randomBytes(32).toString("hex");
    const sessionId = randomUUID();
    const ipAddress = req.ip ?? null;
    const userAgent = req.get("user-agent") ?? null;

    await client.query("BEGIN");

    await client.query(
      `INSERT INTO user_sessions (
         id, user_id, token_hash, ip_address, user_agent
       ) VALUES ($1, $2, $3, $4, $5)`,
      [sessionId, user.id, hashToken(token), ipAddress, userAgent]
    );

    await client.query(
      `INSERT INTO admin_logs (
         admin_id, session_id, action_type, message, ip_address
       ) VALUES ($1, $2, $3, $4, $5)`,
      [user.id, sessionId, "login", `${user.name} เข้าสู่ระบบ`, ipAddress]
    );

    await client.query("COMMIT");

    return res.status(200).json({
      message: "Login สำเร็จ",
      token,
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        role: user.role,
      },
    });
  } catch (error) {
    await client.query("ROLLBACK").catch(() => undefined);
    console.error("Login error:", error);
    return res.status(500).json({ message: "เกิดข้อผิดพลาดภายใน Server" });
  } finally {
    client.release();
  }
});

export default router;

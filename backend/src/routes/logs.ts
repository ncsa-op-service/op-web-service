import { Router } from "express";
import { pool } from "../db.js";

const router = Router();

router.get("/", async (_req, res) => {
  try {
    const result = await pool.query(`
      SELECT
        l.id,
        l.admin_id,
        l.session_id,
        u.name AS admin_name,
        u.email AS admin_email,
        l.action_type,
        l.message,
        l.ip_address,
        l.created_at
      FROM admin_logs l
      LEFT JOIN users u ON u.id = l.admin_id
      ORDER BY l.created_at DESC, l.id DESC
    `);

    return res.json({ logs: result.rows });
  } catch (error) {
    console.error("Get logs error:", error);
    return res.status(500).json({ message: "ไม่สามารถโหลด Log ได้" });
  }
});

export default router;

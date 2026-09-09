import { Router } from "express";
import bcrypt from "bcryptjs";
import { pool } from "../db.js";

const router = Router();

/* =========================
   GET /api/users
   ดูผู้ใช้ทั้งหมด
========================= */
router.get("/", async (_req, res) => {
  try {
    const result = await pool.query(`
      SELECT
        id,
        name,
        email,
        role,
        created_at,
        updated_at
      FROM users
      ORDER BY id ASC
    `);

    return res.json({
      users: result.rows,
    });
  } catch (error) {
    console.error("Get users error:", error);

    return res.status(500).json({
      message: "ไม่สามารถโหลดข้อมูลผู้ใช้ได้",
    });
  }
});

/* =========================
   POST /api/users
   เพิ่มผู้ใช้ใหม่
========================= */
router.post("/", async (req, res) => {
  try {
    const { name, email, password, role } = req.body;

    if (!name || !email || !password || !role) {
      return res.status(400).json({
        message: "กรุณากรอกข้อมูลให้ครบ",
      });
    }

    const allowedRoles = ["super_admin", "editor", "viewer"];

    if (!allowedRoles.includes(role)) {
      return res.status(400).json({
        message: "Role ไม่ถูกต้อง",
      });
    }

    const existingUser = await pool.query(
      "SELECT id FROM users WHERE email = $1",
      [email]
    );

    if (existingUser.rows.length > 0) {
      return res.status(409).json({
        message: "Email นี้มีผู้ใช้งานแล้ว",
      });
    }

    const passwordHash = await bcrypt.hash(password, 12);

    const result = await pool.query(
      `
      INSERT INTO users (
        name,
        email,
        password_hash,
        role
      )
      VALUES ($1, $2, $3, $4)
      RETURNING
        id,
        name,
        email,
        role,
        created_at,
        updated_at
      `,
      [name, email, passwordHash, role]
    );

    return res.status(201).json({
      message: "เพิ่มผู้ใช้สำเร็จ",
      user: result.rows[0],
    });
  } catch (error) {
    console.error("Create user error:", error);

    return res.status(500).json({
      message: "ไม่สามารถเพิ่มผู้ใช้ได้",
    });
  }
});

/* =========================
   PUT /api/users/:id
   แก้ไขผู้ใช้
========================= */
router.put("/:id", async (req, res) => {
  try {
    const id = Number(req.params.id);
    const { name, email, password, role } = req.body;

    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({
        message: "User ID ไม่ถูกต้อง",
      });
    }

    if (!name || !email || !role) {
      return res.status(400).json({
        message: "กรุณากรอก Name, Email และ Role",
      });
    }

    const allowedRoles = ["super_admin", "editor", "viewer"];

    if (!allowedRoles.includes(role)) {
      return res.status(400).json({
        message: "Role ไม่ถูกต้อง",
      });
    }

    const existingEmail = await pool.query(
      `
      SELECT id
      FROM users
      WHERE email = $1
        AND id <> $2
      `,
      [email, id]
    );

    if (existingEmail.rows.length > 0) {
      return res.status(409).json({
        message: "Email นี้มีผู้ใช้งานแล้ว",
      });
    }

    let result;

    if (password) {
      const passwordHash = await bcrypt.hash(password, 12);

      result = await pool.query(
        `
        UPDATE users
        SET
          name = $1,
          email = $2,
          password_hash = $3,
          role = $4,
          updated_at = NOW()
        WHERE id = $5
        RETURNING
          id,
          name,
          email,
          role,
          created_at,
          updated_at
        `,
        [name, email, passwordHash, role, id]
      );
    } else {
      result = await pool.query(
        `
        UPDATE users
        SET
          name = $1,
          email = $2,
          role = $3,
          updated_at = NOW()
        WHERE id = $4
        RETURNING
          id,
          name,
          email,
          role,
          created_at,
          updated_at
        `,
        [name, email, role, id]
      );
    }

    if (result.rows.length === 0) {
      return res.status(404).json({
        message: "ไม่พบผู้ใช้งาน",
      });
    }

    return res.json({
      message: "แก้ไขผู้ใช้สำเร็จ",
      user: result.rows[0],
    });
  } catch (error) {
    console.error("Update user error:", error);

    return res.status(500).json({
      message: "ไม่สามารถแก้ไขผู้ใช้ได้",
    });
  }
});

/* =========================
   DELETE /api/users/:id
   ลบผู้ใช้
========================= */
router.delete("/:id", async (req, res) => {
  try {
    const id = Number(req.params.id);

    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({
        message: "User ID ไม่ถูกต้อง",
      });
    }

    const result = await pool.query(
      `
      DELETE FROM users
      WHERE id = $1
      RETURNING id, name, email, role
      `,
      [id]
    );

    if (result.rows.length === 0) {
      return res.status(404).json({
        message: "ไม่พบผู้ใช้งาน",
      });
    }

    return res.json({
      message: "ลบผู้ใช้สำเร็จ",
      user: result.rows[0],
    });
  } catch (error) {
    console.error("Delete user error:", error);

    return res.status(500).json({
      message: "ไม่สามารถลบผู้ใช้ได้",
    });
  }
});

export default router;
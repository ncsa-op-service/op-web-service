import "dotenv/config";
import bcrypt from "bcryptjs";
import { pool } from "./db.js";

async function createSuperAdmin() {
  try {
    const name = process.env.SUPER_ADMIN_NAME;
    const email = process.env.SUPER_ADMIN_EMAIL;
    const password = process.env.SUPER_ADMIN_PASSWORD;

    if (!name || !email || !password) {
      throw new Error(
        "กรุณากำหนด SUPER_ADMIN_NAME, SUPER_ADMIN_EMAIL และ SUPER_ADMIN_PASSWORD ใน .env"
      );
    }

    const existingUser = await pool.query(
      "SELECT id FROM users WHERE email = $1",
      [email]
    );

    if (existingUser.rows.length > 0) {
      console.log("Super Admin นี้มีอยู่แล้ว");
      return;
    }

    const passwordHash = await bcrypt.hash(password, 12);

    await pool.query(
      `
      INSERT INTO users (name, email, password_hash, role)
      VALUES ($1, $2, $3, $4)
      `,
      [name, email, passwordHash, "super_admin"]
    );

    console.log("สร้าง Super Admin สำเร็จ");
    console.log(`Email: ${email}`);
  } catch (error) {
    console.error("สร้าง Super Admin ไม่สำเร็จ:", error);
  } finally {
    await pool.end();
  }
}

createSuperAdmin();
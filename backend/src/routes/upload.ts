import { Router } from "express";
import multer from "multer";
import * as XLSX from "xlsx";

import { pool } from "../db.js";
import type { ExcelRow } from "../types.js";

const router = Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024,
  },
  fileFilter: (_req, file, callback) => {
    const allowedExtensions = [".xlsx", ".xls"];
    const lowerName = file.originalname.toLowerCase();

    const isAllowed = allowedExtensions.some((extension) =>
      lowerName.endsWith(extension)
    );

    if (!isAllowed) {
      callback(new Error("รองรับเฉพาะไฟล์ .xlsx และ .xls"));
      return;
    }

    callback(null, true);
  },
});

router.post("/", upload.single("file"), async (req, res) => {
  const client = await pool.connect();
  let transactionStarted = false;

  try {
    if (!req.file) {
      res.status(400).json({
        message: "กรุณาเลือกไฟล์ Excel",
      });
      return;
    }

    const workbook = XLSX.read(req.file.buffer, {
      type: "buffer",
    });

    const firstSheetName = workbook.SheetNames[0];

    if (!firstSheetName) {
      res.status(400).json({
        message: "ไม่พบ Sheet ในไฟล์ Excel",
      });
      return;
    }

    const worksheet = workbook.Sheets[firstSheetName];

    if (!worksheet) {
      res.status(400).json({
        message: "ไม่พบข้อมูลใน Sheet",
      });
      return;
    }

    const rows = XLSX.utils.sheet_to_json<ExcelRow>(worksheet, {
      defval: "",
    });

    if (rows.length === 0) {
      res.status(400).json({
        message: "ไม่พบข้อมูลในไฟล์ Excel",
      });
      return;
    }

    await client.query("BEGIN");
    transactionStarted = true;

    const uploadResult = await client.query<{
      id: number;
      original_file_name: string;
      created_at: Date;
    }>(
      `
        INSERT INTO uploads (original_file_name)
        VALUES ($1)
        RETURNING id, original_file_name, created_at
      `,
      [req.file.originalname]
    );

    const uploadedFile = uploadResult.rows[0];

    if (!uploadedFile) {
      throw new Error("ไม่สามารถบันทึกข้อมูลไฟล์อัปโหลดได้");
    }

    let insertedCount = 0;

    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];

      if (!row) {
        continue;
      }

      const rawRow = row as ExcelRow & Record<string, unknown>;

      const name = String(
        rawRow.name ??
          rawRow["ชื่อ"] ??
          rawRow["ชื่อร้าน"] ??
          rawRow["name"] ??
          ""
      ).trim();

      const originalUrl = String(
        rawRow.line_url ??
          rawRow.url ??
          rawRow["ลิงก์"] ??
          rawRow["ลิงก์ LINE"] ??
          rawRow["LINE URL"] ??
          rawRow["line_url"] ??
          rawRow["url"] ??
          ""
      ).trim();

      if (!originalUrl) {
        continue;
      }

      await client.query(
        `
          INSERT INTO line_links (
            upload_id,
            row_number,
            name,
            original_url,
            status
          )
          VALUES ($1, $2, $3, $4, 'WAITING')
        `,
        [
          uploadedFile.id,
          index + 2,
          name || null,
          originalUrl,
        ]
      );

      insertedCount += 1;
    }

    if (insertedCount === 0) {
      throw new Error(
        "ไม่พบลิงก์ในไฟล์ กรุณาตรวจสอบชื่อคอลัมน์ เช่น line_url, url, ลิงก์ หรือ LINE URL"
      );
    }

    await client.query("COMMIT");
    transactionStarted = false;

    res.status(201).json({
      upload: uploadedFile,
      insertedCount,
      message: "อัปโหลดและอ่านไฟล์สำเร็จ",
    });
  } catch (error) {
    if (transactionStarted) {
      await client.query("ROLLBACK");
    }

    console.error("Upload error:", error);

    res.status(500).json({
      message:
        error instanceof Error
          ? error.message
          : "เกิดข้อผิดพลาดในการอัปโหลด",
    });
  } finally {
    client.release();
  }
});

export default router;
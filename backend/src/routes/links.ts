import { Router } from "express";
import { pool } from "../db.js";
import type { LinkStatus } from "../types.js";

const router = Router();

const LINE_DOMAINS = [
  "line.me",
  "lin.ee",
  "page.line.me",
  "access.line.me",
];

function isLineDomain(urlText: string): boolean {
  try {
    const url = new URL(urlText);
    const hostname = url.hostname.toLowerCase();

    return LINE_DOMAINS.some(
      (domain) =>
        hostname === domain || hostname.endsWith(`.${domain}`)
    );
  } catch {
    return false;
  }
}

async function checkUrl(originalUrl: string): Promise<{
  finalUrl: string | null;
  status: LinkStatus;
  message: string;
}> {
  let parsedUrl: URL;

  try {
    parsedUrl = new URL(originalUrl);
  } catch {
    return {
      finalUrl: null,
      status: "INVALID_URL",
      message: "รูปแบบ URL ไม่ถูกต้อง",
    };
  }

  if (!["http:", "https:"].includes(parsedUrl.protocol)) {
    return {
      finalUrl: null,
      status: "INVALID_URL",
      message: "รองรับเฉพาะ HTTP และ HTTPS",
    };
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10_000);

  try {
    const response = await fetch(originalUrl, {
      method: "GET",
      redirect: "follow",
      signal: controller.signal,
      headers: {
        "User-Agent":
          "Mozilla/5.0 (compatible; LineQrChecker/1.0)",
      },
    });

    const finalUrl = response.url || originalUrl;

    if (response.status === 404) {
      return {
        finalUrl,
        status: "NOT_FOUND",
        message: "ไม่พบหน้าเว็บ",
      };
    }

    const foundLine =
      isLineDomain(originalUrl) || isLineDomain(finalUrl);

    if (foundLine) {
      return {
        finalUrl,
        status: "FOUND",
        message: "พบลิงก์ที่นำไปยัง LINE",
      };
    }

    return {
      finalUrl,
      status: "NOT_LINE",
      message: "ปลายทางไม่ใช่เว็บไซต์ LINE",
    };
  } catch (error) {
    const timedOut =
      error instanceof Error && error.name === "AbortError";

    return {
      finalUrl: null,
      status: timedOut ? "TIMEOUT" : "ERROR",
      message: timedOut
        ? "ลิงก์ไม่ตอบสนองภายในเวลาที่กำหนด"
        : "ไม่สามารถเชื่อมต่อลิงก์ได้",
    };
  } finally {
    clearTimeout(timeout);
  }
}

router.get("/", async (_req, res) => {
  try {
    const result = await pool.query(`
      SELECT
        id,
        upload_id,
        row_number,
        name,
        original_url,
        final_url,
        status,
        message,
        created_at,
        checked_at
      FROM line_links
      ORDER BY id DESC
    `);

    res.json(result.rows);
  } catch (error) {
    console.error(error);

    res.status(500).json({
      message: "ไม่สามารถอ่านรายการได้",
    });
  }
});

router.post("/:id/check", async (req, res) => {
  try {
    const id = Number(req.params.id);

    if (!Number.isInteger(id)) {
      res.status(400).json({
        message: "รหัสรายการไม่ถูกต้อง",
      });
      return;
    }

    const findResult = await pool.query(
      `
        SELECT id, original_url
        FROM line_links
        WHERE id = $1
      `,
      [id]
    );

    if (findResult.rowCount === 0) {
      res.status(404).json({
        message: "ไม่พบรายการ",
      });
      return;
    }

    const originalUrl = findResult.rows[0].original_url;
    const checked = await checkUrl(originalUrl);

    const updateResult = await pool.query(
      `
        UPDATE line_links
        SET
          final_url = $1,
          status = $2,
          message = $3,
          checked_at = NOW()
        WHERE id = $4
        RETURNING *
      `,
      [
        checked.finalUrl,
        checked.status,
        checked.message,
        id,
      ]
    );

    res.json(updateResult.rows[0]);
  } catch (error) {
    console.error(error);

    res.status(500).json({
      message: "เกิดข้อผิดพลาดในการตรวจสอบลิงก์",
    });
  }
});

router.post("/check-all", async (_req, res) => {
  try {
    const result = await pool.query(`
      SELECT id, original_url
      FROM line_links
      ORDER BY id ASC
    `);

    const checkedItems = [];

    for (const row of result.rows) {
      const checked = await checkUrl(row.original_url);

      const updateResult = await pool.query(
        `
          UPDATE line_links
          SET
            final_url = $1,
            status = $2,
            message = $3,
            checked_at = NOW()
          WHERE id = $4
          RETURNING *
        `,
        [
          checked.finalUrl,
          checked.status,
          checked.message,
          row.id,
        ]
      );

      checkedItems.push(updateResult.rows[0]);
    }

    res.json({
      checkedCount: checkedItems.length,
      items: checkedItems,
    });
  } catch (error) {
    console.error(error);

    res.status(500).json({
      message: "ตรวจสอบรายการทั้งหมดไม่สำเร็จ",
    });
  }
});

export default router;
import {
  Router,
} from "express";

import {
  pool,
} from "../db.js";

const router =
  Router();

type UserRole =
  | "super_admin"
  | "editor"
  | "viewer";

type SavedMerge = {
  s: {
    r: number;
    c: number;
  };
  e: {
    r: number;
    c: number;
  };
};

type SaveSheet = {
  originalSheetName: string;
  detectedType: string;
  detectedLabel: string;
  confidence: number;
  headers: string[];
  rows: unknown[][];
  sheetRef?: string | null;
  merges?: SavedMerge[];
};

async function getUserRole(
  userId: number
): Promise<UserRole | null> {
  const result =
    await pool.query<{
      role: UserRole;
    }>(
      `
        SELECT role
        FROM users
        WHERE id = $1
        LIMIT 1
      `,
      [userId]
    );

  return (
    result.rows[0]?.role ??
    null
  );
}

function readUserId(
  value:
    | string
    | string[]
    | undefined
) {
  const raw =
    Array.isArray(value)
      ? value[0]
      : value;

  const userId =
    Number(raw);

  return Number.isInteger(
    userId
  ) && userId > 0
    ? userId
    : null;
}

function normalizeMerges(
  value: unknown
): SavedMerge[] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value
    .map((item) => {
      const source =
        item as Partial<SavedMerge>;

      const sr =
        Number(source.s?.r);
      const sc =
        Number(source.s?.c);
      const er =
        Number(source.e?.r);
      const ec =
        Number(source.e?.c);

      if (
        !Number.isInteger(sr) ||
        !Number.isInteger(sc) ||
        !Number.isInteger(er) ||
        !Number.isInteger(ec) ||
        sr < 0 ||
        sc < 0 ||
        er < sr ||
        ec < sc
      ) {
        return null;
      }

      return {
        s: {
          r: sr,
          c: sc,
        },
        e: {
          r: er,
          c: ec,
        },
      };
    })
    .filter(
      (
        item
      ): item is SavedMerge =>
        item !== null
    );
}

async function assertCanManage(
  userId: number
) {
  const role =
    await getUserRole(
      userId
    );

  if (
    role !== "super_admin" &&
    role !== "editor"
  ) {
    return false;
  }

  return true;
}

/* ==============================
   GET LATEST SNAPSHOT
================================ */

router.get(
  "/latest",
  async (_req, res) => {
    try {
      const snapshotResult =
        await pool.query<{
          id: number;
          original_file_name:
            | string
            | null;
          saved_by: number | null;
          saved_by_name:
            | string
            | null;
          created_at: string;
        }>(`
          SELECT
            s.id,
            s.original_file_name,
            s.saved_by,
            u.name AS saved_by_name,
            s.created_at
          FROM clab_snapshots s
          LEFT JOIN users u
            ON u.id = s.saved_by
          ORDER BY
            s.created_at DESC,
            s.id DESC
          LIMIT 1
        `);

      const snapshot =
        snapshotResult.rows[0] ??
        null;

      if (!snapshot) {
        return res.json({
          snapshot: null,
          sheets: [],
        });
      }

      const sheetsResult =
        await pool.query(
          `
            SELECT
              id,
              snapshot_id,
              original_sheet_name,
              detected_type,
              detected_label,
              confidence,
              headers,
              rows,
              sheet_ref,
              merges,
              created_at,
              updated_at
            FROM clab_snapshot_sheets
            WHERE snapshot_id = $1
            ORDER BY id ASC
          `,
          [snapshot.id]
        );

      return res.json({
        snapshot,
        sheets:
          sheetsResult.rows,
      });
    } catch (error) {
      console.error(
        "GET /api/clab/latest error:",
        error
      );

      return res.status(500).json({
        error:
          "โหลดข้อมูล CLAB ล่าสุดไม่สำเร็จ",
      });
    }
  }
);

/* ==============================
   GET LATEST SHEET PER TYPE
================================ */

router.get(
  "/latest-by-type",
  async (_req, res) => {
    try {
      const result =
        await pool.query(`
          SELECT DISTINCT ON (
            cs.detected_type
          )
            cs.id,
            cs.snapshot_id,
            cs.original_sheet_name,
            cs.detected_type,
            cs.detected_label,
            cs.confidence,
            cs.headers,
            cs.rows,
            cs.created_at,
            cs.updated_at,
            s.original_file_name,
            s.saved_by,
            u.name AS saved_by_name
          FROM clab_snapshot_sheets cs
          INNER JOIN clab_snapshots s
            ON s.id = cs.snapshot_id
          LEFT JOIN users u
            ON u.id = s.saved_by
          ORDER BY
            cs.detected_type,
            COALESCE(
              cs.updated_at,
              cs.created_at
            ) DESC,
            cs.id DESC
        `);

      return res.json({
        sheets:
          result.rows,
      });
    } catch (error) {
      console.error(
        "GET /api/clab/latest-by-type error:",
        error
      );

      return res.status(500).json({
        error:
          "โหลดข้อมูล CLAB ล่าสุดแยกตามหัวข้อไม่สำเร็จ",
      });
    }
  }
);

/* ==============================
   SAVE ONE SHEET
   Super Admin / Editor only
================================ */

router.post(
  "/sheets",
  async (req, res) => {
    const client =
      await pool.connect();

    let transactionStarted =
      false;

    try {
      const userId =
        readUserId(
          req.headers[
            "x-user-id"
          ]
        );

      if (!userId) {
        return res
          .status(401)
          .json({
            error:
              "ไม่พบข้อมูลผู้ใช้งาน",
          });
      }

      const canManage =
        await assertCanManage(
          userId
        );

      if (!canManage) {
        return res
          .status(403)
          .json({
            error:
              "ไม่มีสิทธิ์บันทึกข้อมูล CLAB",
          });
      }

      const {
        fileName,
        sheet,
      } = req.body as {
        fileName?: string;
        sheet?: SaveSheet;
      };

      if (
        !sheet ||
        typeof sheet.originalSheetName !==
          "string" ||
        typeof sheet.detectedType !==
          "string" ||
        !Array.isArray(
          sheet.rows
        )
      ) {
        return res
          .status(400)
          .json({
            error:
              "ข้อมูล Sheet ไม่ถูกต้อง",
          });
      }

      await client.query(
        "BEGIN"
      );

      transactionStarted =
        true;

      const snapshotResult =
        await client.query<{
          id: number;
          created_at: string;
        }>(
          `
            INSERT INTO clab_snapshots (
              original_file_name,
              saved_by
            )
            VALUES ($1, $2)
            RETURNING
              id,
              created_at
          `,
          [
            fileName?.trim() ||
              null,
            userId,
          ]
        );

      const snapshot =
        snapshotResult.rows[0];

      if (!snapshot) {
        throw new Error(
          "ไม่สามารถสร้าง CLAB Snapshot ได้"
        );
      }

      const sheetResult =
        await client.query<{
          id: number;
          created_at: string;
        }>(
          `
            INSERT INTO clab_snapshot_sheets (
              snapshot_id,
              original_sheet_name,
              detected_type,
              detected_label,
              confidence,
              headers,
              rows,
              sheet_ref,
              merges,
              updated_at
            )
            VALUES (
              $1,
              $2,
              $3,
              $4,
              $5,
              $6::jsonb,
              $7::jsonb,
              $8,
              $9::jsonb,
              NOW()
            )
            RETURNING
              id,
              created_at
          `,
          [
            snapshot.id,
            sheet.originalSheetName,
            sheet.detectedType,
            sheet.detectedLabel ||
              sheet.detectedType,
            Number.isFinite(
              Number(
                sheet.confidence
              )
            )
              ? Math.max(
                  0,
                  Math.min(
                    100,
                    Number(
                      sheet.confidence
                    )
                  )
                )
              : 0,
            JSON.stringify(
              sheet.headers ??
                []
            ),
            JSON.stringify(
              sheet.rows
            ),
            typeof sheet.sheetRef ===
              "string" &&
            sheet.sheetRef.trim()
              ? sheet.sheetRef.trim()
              : null,
            JSON.stringify(
              normalizeMerges(
                sheet.merges
              )
            ),
          ]
        );

      const savedSheet =
        sheetResult.rows[0];

      if (!savedSheet) {
        throw new Error(
          "บันทึก Sheet ไม่สำเร็จ"
        );
      }

      await client.query(
        `
          INSERT INTO admin_logs (
            admin_id,
            action_type,
            message,
            ip_address
          )
          VALUES (
            $1,
            'clab_sheet_save',
            $2,
            $3
          )
        `,
        [
          userId,
          `บันทึก CLAB ${sheet.detectedLabel || sheet.detectedType} จากไฟล์ ${fileName ?? "-"}`,
          req.ip ?? null,
        ]
      );

      await client.query(
        "COMMIT"
      );

      transactionStarted =
        false;

      return res.status(201).json({
        message:
          "บันทึก CLAB สำเร็จ",
        snapshotId:
          snapshot.id,
        sheetId:
          savedSheet.id,
        createdAt:
          savedSheet.created_at,
      });
    } catch (error) {
      if (
        transactionStarted
      ) {
        await client.query(
          "ROLLBACK"
        );
      }

      console.error(
        "POST /api/clab/sheets error:",
        error
      );

      return res.status(500).json({
        error:
          "บันทึกข้อมูล CLAB ไม่สำเร็จ",
      });
    } finally {
      client.release();
    }
  }
);

/* ==============================
   UPDATE SAVED SHEET
   Super Admin / Editor only
================================ */

router.put(
  "/sheets/:id",
  async (req, res) => {
    try {
      const userId =
        readUserId(
          req.headers[
            "x-user-id"
          ]
        );

      if (!userId) {
        return res
          .status(401)
          .json({
            error:
              "ไม่พบข้อมูลผู้ใช้งาน",
          });
      }

      const canManage =
        await assertCanManage(
          userId
        );

      if (!canManage) {
        return res
          .status(403)
          .json({
            error:
              "ไม่มีสิทธิ์แก้ไขข้อมูล CLAB",
          });
      }

      const sheetId =
        Number(
          req.params.id
        );

      if (
        !Number.isInteger(
          sheetId
        ) ||
        sheetId <= 0
      ) {
        return res
          .status(400)
          .json({
            error:
              "Sheet ID ไม่ถูกต้อง",
          });
      }

      const {
        headers,
        rows,
        originalSheetName,
        detectedLabel,
        confidence,
        sheetRef,
        merges,
      } = req.body as {
        headers?: string[];
        rows?: unknown[][];
        originalSheetName?: string;
        detectedLabel?: string;
        confidence?: number;
        sheetRef?: string | null;
        merges?: SavedMerge[];
      };

      const result =
        await pool.query(
          `
            UPDATE clab_snapshot_sheets
            SET
              headers =
                COALESCE(
                  $1::jsonb,
                  headers
                ),
              rows =
                COALESCE(
                  $2::jsonb,
                  rows
                ),
              original_sheet_name =
                COALESCE(
                  $3,
                  original_sheet_name
                ),
              detected_label =
                COALESCE(
                  $4,
                  detected_label
                ),
              confidence =
                COALESCE(
                  $5,
                  confidence
                ),
              sheet_ref =
                COALESCE(
                  $6,
                  sheet_ref
                ),
              merges =
                COALESCE(
                  $7::jsonb,
                  merges
                ),
              updated_at =
                NOW()
            WHERE id = $8
            RETURNING *
          `,
          [
            headers
              ? JSON.stringify(
                  headers
                )
              : null,
            rows
              ? JSON.stringify(
                  rows
                )
              : null,
            originalSheetName ??
              null,
            detectedLabel ??
              null,
            Number.isFinite(
              Number(
                confidence
              )
            )
              ? Number(
                  confidence
                )
              : null,
            typeof sheetRef ===
              "string" &&
            sheetRef.trim()
              ? sheetRef.trim()
              : null,
            Array.isArray(
              merges
            )
              ? JSON.stringify(
                  normalizeMerges(
                    merges
                  )
                )
              : null,
            sheetId,
          ]
        );

      const updated =
        result.rows[0];

      if (!updated) {
        return res
          .status(404)
          .json({
            error:
              "ไม่พบข้อมูล CLAB ที่ต้องการแก้ไข",
          });
      }

      await pool.query(
        `
          INSERT INTO admin_logs (
            admin_id,
            action_type,
            message,
            ip_address
          )
          VALUES (
            $1,
            'clab_sheet_update',
            $2,
            $3
          )
        `,
        [
          userId,
          `แก้ไข CLAB Sheet #${sheetId}`,
          req.ip ?? null,
        ]
      );

      return res.json({
        message:
          "แก้ไขข้อมูล CLAB สำเร็จ",
        sheet:
          updated,
      });
    } catch (error) {
      console.error(
        "PUT /api/clab/sheets/:id error:",
        error
      );

      return res.status(500).json({
        error:
          "แก้ไขข้อมูล CLAB ไม่สำเร็จ",
      });
    }
  }
);

/* ==============================
   HISTORY
================================ */

router.get(
  "/snapshots",
  async (req, res) => {
    try {
      const userId =
        readUserId(
          req.headers[
            "x-user-id"
          ]
        );

      if (!userId) {
        return res
          .status(401)
          .json({
            error:
              "ไม่พบข้อมูลผู้ใช้งาน",
          });
      }

      const canManage =
        await assertCanManage(
          userId
        );

      if (!canManage) {
        return res
          .status(403)
          .json({
            error:
              "ไม่มีสิทธิ์ดูประวัติ CLAB",
          });
      }

      const result =
        await pool.query(`
          SELECT
            s.id,
            s.original_file_name,
            s.saved_by,
            u.name AS saved_by_name,
            s.created_at,
            COUNT(cs.id)::int
              AS sheet_count
          FROM clab_snapshots s
          LEFT JOIN users u
            ON u.id = s.saved_by
          LEFT JOIN clab_snapshot_sheets cs
            ON cs.snapshot_id = s.id
          GROUP BY
            s.id,
            u.name
          ORDER BY
            s.created_at DESC,
            s.id DESC
          LIMIT 50
        `);

      return res.json({
        snapshots:
          result.rows,
      });
    } catch (error) {
      console.error(
        "GET /api/clab/snapshots error:",
        error
      );

      return res.status(500).json({
        error:
          "โหลดประวัติ CLAB ไม่สำเร็จ",
      });
    }
  }
);

export default router;

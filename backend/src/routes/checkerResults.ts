import { Router } from "express";
import { pool } from "../db.js";

const router = Router();

type CaseType =
  | "fake_domain"
  | "fake_web"
  | "c2_oss"
  | "fake_line";

type CaseStatus = "active" | "inactive";

type ProviderResultStatus =
  | "success"
  | "nxdomain"
  | "failed"
  | null;

type SaveItem = {
  caseId: string;
  urlSms: string;
  checked: {
    ais?: string;
    trueDtac?: string;
    nt?: string;
    cloudflare?: string;
  };
  checkState?: {
    ais?: ProviderResultStatus;
    trueDtac?: ProviderResultStatus;
    nt?: ProviderResultStatus;
    cloudflare?: ProviderResultStatus;
  };
};

type SaveRequestBody = {
  fileName?: string;
  caseType?: CaseType;
  caseStatus?: CaseStatus;
  detectedDate?: string;
  selectedProvider?: string;
  checkedBy?: number | null;
  networkInfo?: {
    ip?: string | null;
    org?: string | null;
    provider?: string | null;
  } | null;
  items?: SaveItem[];
};

type ReferenceBody = {
  url?: string;
  note?: string | null;
  createdBy?: number | null;
};

type CompareRowStatus = "same" | "changed" | "new" | "missing";

const allowedCaseTypes: CaseType[] = [
  "fake_domain",
  "fake_web",
  "c2_oss",
  "fake_line",
];

const allowedCaseStatuses: CaseStatus[] = ["active", "inactive"];

const allowedResultStatuses = ["success", "nxdomain", "failed"] as const;

function getToday(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function isDateKey(value: unknown): value is string {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function cleanResult(value: unknown): string | null {
  const text = String(value ?? "").trim();
  return text || null;
}

function cleanResultStatus(value: unknown): ProviderResultStatus {
  if (
    typeof value === "string" &&
    allowedResultStatuses.includes(
      value as "success" | "nxdomain" | "failed"
    )
  ) {
    return value as "success" | "nxdomain" | "failed";
  }

  return null;
}

function normalizeLineReferenceUrl(rawValue: unknown): {
  url: string;
  token: string;
} | null {
  try {
    const parsed = new URL(String(rawValue ?? "").trim());
    const host = parsed.hostname.toLowerCase();
    const match = parsed.pathname.match(/^\/ti\/p\/([^/?#]+)\/?$/i);

    if (
      parsed.protocol !== "https:" ||
      !(host === "line.me" || host === "www.line.me") ||
      !match
    ) {
      return null;
    }

    const token = match[1] ?? "";
    if (!token || !/^[A-Za-z0-9_-]+$/.test(token)) {
      return null;
    }

    return {
      url: `https://line.me/ti/p/${token}`,
      token,
    };
  } catch {
    return null;
  }
}

function extractLineUrl(rawValue: unknown): string | null {
  const normalized = normalizeLineReferenceUrl(rawValue);
  return normalized?.url ?? null;
}

function collectLineUrls(row: {
  url_sms?: string | null;
  ais_result?: string | null;
  true_dtac_result?: string | null;
  nt_result?: string | null;
  cloudflare_result?: string | null;
}): string[] {
  const values = [
    row.url_sms,
    row.ais_result,
    row.true_dtac_result,
    row.nt_result,
    row.cloudflare_result,
  ];

  return Array.from(
    new Set(
      values
        .map(extractLineUrl)
        .filter((value): value is string => Boolean(value))
    )
  ).sort();
}

async function getLatestBatchForDate(
  date: string,
  caseType: CaseType
): Promise<number | null> {
  const result = await pool.query<{ id: number }>(
    `
      SELECT id
      FROM url_fake_batches
      WHERE detected_date = $1::date
        AND case_type = $2
      ORDER BY created_at DESC, id DESC
      LIMIT 1
    `,
    [date, caseType]
  );

  return result.rows[0]?.id ?? null;
}

async function getBatchLineMap(batchId: number | null) {
  const map = new Map<string, Set<string>>();

  if (batchId === null) {
    return map;
  }

  const result = await pool.query<{
    id: number;
    case_id: string;
    url_sms: string;
    ais_result: string | null;
    true_dtac_result: string | null;
    nt_result: string | null;
    cloudflare_result: string | null;
  }>(
    `
      SELECT
        id,
        case_id,
        url_sms,
        ais_result,
        true_dtac_result,
        nt_result,
        cloudflare_result
      FROM url_fake_results
      WHERE batch_id = $1
      ORDER BY id ASC
    `,
    [batchId]
  );

  for (const row of result.rows) {
    const lineUrls = collectLineUrls(row);
    if (lineUrls.length === 0) continue;

    const caseId = String(row.case_id ?? "").trim() || `RESULT-${row.id}`;
    const existing = map.get(caseId) ?? new Set<string>();

    for (const url of lineUrls) {
      existing.add(url);
    }

    map.set(caseId, existing);
  }

  return map;
}

/* =========================================================
   SAVE CHECKER BATCH
========================================================= */

router.post("/", async (req, res) => {
  const client = await pool.connect();
  let transactionStarted = false;

  try {
    const {
      fileName,
      caseType,
      caseStatus,
      detectedDate,
      selectedProvider,
      checkedBy,
      networkInfo,
      items,
    } = req.body as SaveRequestBody;

    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ message: "ไม่มีข้อมูลสำหรับบันทึก" });
    }

    const safeCaseType: CaseType =
      caseType && allowedCaseTypes.includes(caseType)
        ? caseType
        : "fake_domain";

    const safeCaseStatus: CaseStatus =
      caseStatus && allowedCaseStatuses.includes(caseStatus)
        ? caseStatus
        : "active";

    const safeDetectedDate = isDateKey(detectedDate)
      ? detectedDate
      : getToday();

    const safeCheckedBy =
      typeof checkedBy === "number" &&
      Number.isInteger(checkedBy) &&
      checkedBy > 0
        ? checkedBy
        : null;

    await client.query("BEGIN");
    transactionStarted = true;

    const batchResult = await client.query<{
      id: number;
      created_at: Date;
    }>(
      `
        INSERT INTO url_fake_batches (
          original_file_name,
          selected_provider,
          public_ip,
          isp,
          detected_provider,
          case_type,
          detected_date,
          case_status,
          checked_by
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7::date, $8, $9)
        RETURNING id, created_at
      `,
      [
        fileName || null,
        selectedProvider || null,
        networkInfo?.ip || null,
        networkInfo?.org || null,
        networkInfo?.provider || null,
        safeCaseType,
        safeDetectedDate,
        safeCaseStatus,
        safeCheckedBy,
      ]
    );

    const batch = batchResult.rows[0];
    if (!batch) {
      throw new Error("ไม่สามารถสร้าง Batch ได้");
    }

    let insertedCount = 0;

    for (const item of items) {
      const caseId = String(item.caseId ?? "").trim();
      const urlSms = String(item.urlSms ?? "").trim();

      if (!urlSms) continue;

      const aisResult = cleanResult(item.checked?.ais);
      const trueDtacResult = cleanResult(item.checked?.trueDtac);
      const ntResult = cleanResult(item.checked?.nt);
      const cloudflareResult = cleanResult(item.checked?.cloudflare);

      const aisStatus = cleanResultStatus(item.checkState?.ais);
      const trueDtacStatus = cleanResultStatus(item.checkState?.trueDtac);
      const ntStatus = cleanResultStatus(item.checkState?.nt);
      const cloudflareStatus = cleanResultStatus(item.checkState?.cloudflare);

      await client.query(
        `
          INSERT INTO url_fake_results (
            batch_id,
            case_id,
            url_sms,
            ais_result,
            true_dtac_result,
            nt_result,
            cloudflare_result,
            ais_status,
            true_dtac_status,
            nt_status,
            cloudflare_status
          )
          VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
        `,
        [
          batch.id,
          caseId,
          urlSms,
          aisResult,
          trueDtacResult,
          ntResult,
          cloudflareResult,
          aisStatus,
          trueDtacStatus,
          ntStatus,
          cloudflareStatus,
        ]
      );

      insertedCount += 1;
    }

    if (insertedCount === 0) {
      throw new Error("ไม่พบ URL สำหรับบันทึก");
    }

    await client.query("COMMIT");
    transactionStarted = false;

    return res.status(201).json({
      message: "บันทึกข้อมูลสำเร็จ",
      batchId: batch.id,
      insertedCount,
      createdAt: batch.created_at,
      caseType: safeCaseType,
      caseStatus: safeCaseStatus,
      detectedDate: safeDetectedDate,
    });
  } catch (error) {
    if (transactionStarted) {
      await client.query("ROLLBACK");
    }

    console.error("Save checker results error:", error);

    return res.status(500).json({
      message:
        error instanceof Error
          ? error.message
          : "บันทึกข้อมูลไม่สำเร็จ",
    });
  } finally {
    client.release();
  }
});

/* =========================================================
   HISTORY
   ทุกครั้งที่กดบันทึกจะเป็นคนละ Batch จึงไม่เขียนทับข้อมูลเดิม
========================================================= */

router.get("/history", async (_req, res) => {
  try {
    const result = await pool.query(
      `
        SELECT
          b.*,
          COUNT(r.id)::int AS result_count,
          NULL::text AS save_type,
          NULL::text AS saved_provider,
          NULL::boolean AS publish_to_summary,
          b.created_at AS saved_at
        FROM url_fake_batches b
        LEFT JOIN url_fake_results r
          ON r.batch_id = b.id
        GROUP BY b.id
        ORDER BY b.detected_date DESC, b.created_at DESC, b.id DESC
        LIMIT 1000
      `
    );

    return res.json({ batches: result.rows });
  } catch (error) {
    console.error("Load checker history error:", error);
    return res.status(500).json({ message: "โหลดประวัติไม่สำเร็จ" });
  }
});

router.get("/line-references", async (_req, res) => {
  try {
    const result = await pool.query(
      `
        SELECT
          ref.id,
          ref.url,
          ref.token,
          ref.verified_scan,
          ref.is_active,
          ref.note,
          ref.created_by,
          creator.name AS created_by_name,
          ref.created_at,
          ref.updated_at,
          COALESCE(stats.match_count, 0)::int AS match_count,
          stats.latest_detected_at
        FROM line_qr_references ref
        LEFT JOIN users creator
          ON creator.id = ref.created_by
        LEFT JOIN LATERAL (
          SELECT
            COUNT(DISTINCT r.id)::int AS match_count,
            MAX(COALESCE(r.created_at, b.created_at)) AS latest_detected_at
          FROM url_fake_results r
          INNER JOIN url_fake_batches b
            ON b.id = r.batch_id
          CROSS JOIN LATERAL (
            VALUES
              (r.url_sms),
              (r.ais_result),
              (r.true_dtac_result),
              (r.nt_result),
              (r.cloudflare_result)
          ) AS candidate(line_url)
          WHERE BTRIM(COALESCE(candidate.line_url, '')) = ref.url
        ) stats ON TRUE
        WHERE ref.is_active = TRUE
        ORDER BY ref.id ASC
      `
    );

    return res.json({ references: result.rows });
  } catch (error) {
    console.error("Load LINE references error:", error);
    return res.status(500).json({ message: "โหลด LINE Reference ไม่สำเร็จ" });
  }
});

// ประวัติว่า Reference นี้ถูกตรวจพบจาก Batch ไหน / วันเวลาใด / ใครตรวจ / พบจากช่องไหน
router.get("/line-references/:id/matches", async (req, res) => {
  try {
    const id = Number(req.params.id);

    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({ message: "Reference ID ไม่ถูกต้อง" });
    }

    const rawLimit = Number(req.query.limit ?? 200);
    const limit =
      Number.isInteger(rawLimit) && rawLimit > 0
        ? Math.min(rawLimit, 1000)
        : 200;

    const referenceResult = await pool.query(
      `
        SELECT
          ref.*,
          creator.name AS created_by_name
        FROM line_qr_references ref
        LEFT JOIN users creator
          ON creator.id = ref.created_by
        WHERE ref.id = $1
        LIMIT 1
      `,
      [id]
    );

    const reference = referenceResult.rows[0] ?? null;

    if (!reference) {
      return res.status(404).json({ message: "ไม่พบ LINE Reference" });
    }

    const matchesResult = await pool.query(
      `
        SELECT
          r.id AS result_id,
          r.batch_id,
          r.case_id,
          candidate.source_provider,
          candidate.line_url,
          b.original_file_name,
          b.detected_date,
          COALESCE(r.created_at, b.created_at) AS detected_at,
          b.created_at AS batch_created_at,
          b.selected_provider,
          b.checked_by,
          checker.name AS checked_by_name,
          checker.email AS checked_by_email
        FROM url_fake_results r
        INNER JOIN url_fake_batches b
          ON b.id = r.batch_id
        LEFT JOIN users checker
          ON checker.id = b.checked_by
        CROSS JOIN LATERAL (
          VALUES
            ('URL SMS'::text, r.url_sms),
            ('AIS'::text, r.ais_result),
            ('TRUE / DTAC'::text, r.true_dtac_result),
            ('NT'::text, r.nt_result),
            ('Cloudflare'::text, r.cloudflare_result)
        ) AS candidate(source_provider, line_url)
        WHERE BTRIM(COALESCE(candidate.line_url, '')) = $1
        ORDER BY
          COALESCE(r.created_at, b.created_at) DESC,
          r.id DESC,
          candidate.source_provider ASC
        LIMIT $2
      `,
      [reference.url, limit]
    );

    return res.json({
      reference,
      total: matchesResult.rowCount ?? matchesResult.rows.length,
      matches: matchesResult.rows,
    });
  } catch (error) {
    console.error("Load LINE reference matches error:", error);
    return res.status(500).json({
      message: "โหลดประวัติการตรวจพบ LINE Reference ไม่สำเร็จ",
    });
  }
});

router.post("/line-references", async (req, res) => {
  try {
    const body = req.body as ReferenceBody;
    const normalized = normalizeLineReferenceUrl(body.url);

    if (!normalized) {
      return res.status(400).json({
        message: "URL ต้องเป็นรูปแบบ https://line.me/ti/p/{token}",
      });
    }

    const createdBy =
      typeof body.createdBy === "number" && Number.isInteger(body.createdBy)
        ? body.createdBy
        : null;

    const result = await pool.query(
      `
        INSERT INTO line_qr_references (
          url,
          token,
          verified_scan,
          is_active,
          note,
          created_by
        )
        VALUES ($1, $2, TRUE, TRUE, $3, $4)
        ON CONFLICT (token)
        DO UPDATE SET
          url = EXCLUDED.url,
          verified_scan = TRUE,
          is_active = TRUE,
          note = EXCLUDED.note,
          updated_at = NOW()
        RETURNING *
      `,
      [
        normalized.url,
        normalized.token,
        cleanResult(body.note),
        createdBy,
      ]
    );

    return res.status(201).json({
      message: "บันทึก LINE Reference สำเร็จ",
      reference: result.rows[0],
    });
  } catch (error) {
    console.error("Create LINE reference error:", error);
    return res.status(500).json({ message: "บันทึก LINE Reference ไม่สำเร็จ" });
  }
});

router.put("/line-references/:id", async (req, res) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({ message: "Reference ID ไม่ถูกต้อง" });
    }

    const body = req.body as ReferenceBody;
    const normalized = normalizeLineReferenceUrl(body.url);

    if (!normalized) {
      return res.status(400).json({
        message: "URL ต้องเป็นรูปแบบ https://line.me/ti/p/{token}",
      });
    }

    const duplicate = await pool.query<{ id: number }>(
      `
        SELECT id
        FROM line_qr_references
        WHERE token = $1
          AND id <> $2
        LIMIT 1
      `,
      [normalized.token, id]
    );

    if (duplicate.rows.length > 0) {
      return res.status(409).json({ message: "LINE URL นี้มีอยู่ใน Reference แล้ว" });
    }

    const result = await pool.query(
      `
        UPDATE line_qr_references
        SET
          url = $1,
          token = $2,
          note = $3,
          verified_scan = TRUE,
          is_active = TRUE,
          updated_at = NOW()
        WHERE id = $4
        RETURNING *
      `,
      [normalized.url, normalized.token, cleanResult(body.note), id]
    );

    if (result.rows.length === 0) {
      return res.status(404).json({ message: "ไม่พบ LINE Reference" });
    }

    return res.json({
      message: "แก้ไข LINE Reference สำเร็จ",
      reference: result.rows[0],
    });
  } catch (error) {
    console.error("Update LINE reference error:", error);
    return res.status(500).json({ message: "แก้ไข LINE Reference ไม่สำเร็จ" });
  }
});

router.delete("/line-references/:id", async (req, res) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({ message: "Reference ID ไม่ถูกต้อง" });
    }

    const result = await pool.query(
      `
        UPDATE line_qr_references
        SET
          is_active = FALSE,
          updated_at = NOW()
        WHERE id = $1
          AND is_active = TRUE
        RETURNING id
      `,
      [id]
    );

    if (result.rows.length === 0) {
      return res.status(404).json({ message: "ไม่พบ LINE Reference" });
    }

    return res.json({ message: "ลบ LINE Reference สำเร็จ" });
  } catch (error) {
    console.error("Delete LINE reference error:", error);
    return res.status(500).json({ message: "ลบ LINE Reference ไม่สำเร็จ" });
  }
});

/* =========================================================
   DAILY LINE COMPARISON
   ใช้ Batch ล่าสุดของแต่ละวันที่เลือกมาเทียบกันตาม CASE ID
========================================================= */

router.get("/line-compare", async (req, res) => {
  try {
    const currentDate = isDateKey(req.query.currentDate)
      ? req.query.currentDate
      : getToday();

    const previousDate = isDateKey(req.query.previousDate)
      ? req.query.previousDate
      : (() => {
          const date = new Date(`${currentDate}T12:00:00`);
          date.setDate(date.getDate() - 1);
          return [
            date.getFullYear(),
            String(date.getMonth() + 1).padStart(2, "0"),
            String(date.getDate()).padStart(2, "0"),
          ].join("-");
        })();

    const rawCaseType = String(req.query.caseType ?? "fake_domain");
    const caseType: CaseType = allowedCaseTypes.includes(rawCaseType as CaseType)
      ? (rawCaseType as CaseType)
      : "fake_domain";

    const [currentBatchId, previousBatchId] = await Promise.all([
      getLatestBatchForDate(currentDate, caseType),
      getLatestBatchForDate(previousDate, caseType),
    ]);

    const [currentMap, previousMap] = await Promise.all([
      getBatchLineMap(currentBatchId),
      getBatchLineMap(previousBatchId),
    ]);

    const allCaseIds = Array.from(
      new Set([...currentMap.keys(), ...previousMap.keys()])
    ).sort((a, b) => a.localeCompare(b, "th"));

    const rows = allCaseIds.map((caseId) => {
      const previousUrls = Array.from(previousMap.get(caseId) ?? []).sort();
      const currentUrls = Array.from(currentMap.get(caseId) ?? []).sort();

      let status: CompareRowStatus;

      if (previousUrls.length === 0 && currentUrls.length > 0) {
        status = "new";
      } else if (previousUrls.length > 0 && currentUrls.length === 0) {
        status = "missing";
      } else if (JSON.stringify(previousUrls) === JSON.stringify(currentUrls)) {
        status = "same";
      } else {
        status = "changed";
      }

      return {
        caseId,
        previousUrls,
        currentUrls,
        status,
      };
    });

    const summary = {
      same: rows.filter((row) => row.status === "same").length,
      changed: rows.filter((row) => row.status === "changed").length,
      new: rows.filter((row) => row.status === "new").length,
      missing: rows.filter((row) => row.status === "missing").length,
      total: rows.length,
    };

    return res.json({
      currentDate,
      previousDate,
      currentBatchId,
      previousBatchId,
      summary,
      rows,
    });
  } catch (error) {
    console.error("Compare LINE daily data error:", error);
    return res.status(500).json({ message: "เปรียบเทียบ LINE URL รายวันไม่สำเร็จ" });
  }
});

/* =========================================================
   OPEN SAVED BATCH
   ต้องอยู่หลัง route /history, /line-references, /line-compare
========================================================= */

router.get("/:id", async (req, res) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({ message: "Batch ID ไม่ถูกต้อง" });
    }

    const batchResult = await pool.query(
      `
        SELECT *
        FROM url_fake_batches
        WHERE id = $1
        LIMIT 1
      `,
      [id]
    );

    const batch = batchResult.rows[0] ?? null;
    if (!batch) {
      return res.status(404).json({ message: "ไม่พบ Batch" });
    }

    const resultRows = await pool.query(
      `
        SELECT *
        FROM url_fake_results
        WHERE batch_id = $1
        ORDER BY id ASC
      `,
      [id]
    );

    return res.json({
      batch,
      results: resultRows.rows,
    });
  } catch (error) {
    console.error("Open checker batch error:", error);
    return res.status(500).json({ message: "โหลด Batch ไม่สำเร็จ" });
  }
});

export default router;

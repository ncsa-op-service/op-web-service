import { Router } from "express";
import { pool } from "../db.js";

const router = Router();

type CaseType =
  | "fake_domain"
  | "fake_web"
  | "c2_oss"
  | "fake_line";

type CaseStatus =
  | "active"
  | "inactive";

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

  networkInfo?: {
    ip?: string | null;
    org?: string | null;
    provider?: string | null;
  } | null;

  items?: SaveItem[];
};

const allowedCaseTypes: CaseType[] = [
  "fake_domain",
  "fake_web",
  "c2_oss",
  "fake_line",
];

const allowedCaseStatuses: CaseStatus[] = [
  "active",
  "inactive",
];

const allowedResultStatuses = [
  "success",
  "nxdomain",
  "failed",
] as const;

function getToday(): string {
  const now = new Date();

  const year = now.getFullYear();

  const month = String(
    now.getMonth() + 1
  ).padStart(2, "0");

  const day = String(
    now.getDate()
  ).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function cleanResult(
  value: unknown
): string | null {
  const text = String(value ?? "").trim();
  return text || null;
}

function cleanResultStatus(
  value: unknown
): ProviderResultStatus {
  if (
    typeof value === "string" &&
    allowedResultStatuses.includes(
      value as
        | "success"
        | "nxdomain"
        | "failed"
    )
  ) {
    return value as
      | "success"
      | "nxdomain"
      | "failed";
  }

  return null;
}

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
      networkInfo,
      items,
    } = req.body as SaveRequestBody;

    /* =========================
       VALIDATE ITEMS
    ========================= */

    if (
      !Array.isArray(items) ||
      items.length === 0
    ) {
      return res.status(400).json({
        message: "ไม่มีข้อมูลสำหรับบันทึก",
      });
    }

    /* =========================
       DEFAULT VALUES
    ========================= */

    const safeCaseType: CaseType =
      caseType &&
      allowedCaseTypes.includes(caseType)
        ? caseType
        : "fake_domain";

    const safeCaseStatus: CaseStatus =
      caseStatus &&
      allowedCaseStatuses.includes(caseStatus)
        ? caseStatus
        : "active";

    const safeDetectedDate =
      detectedDate &&
      /^\d{4}-\d{2}-\d{2}$/.test(
        detectedDate
      )
        ? detectedDate
        : getToday();

    await client.query("BEGIN");
    transactionStarted = true;

    /* =========================
       CREATE BATCH
    ========================= */

    const batchResult =
      await client.query<{
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
            case_status
          )
          VALUES (
            $1,
            $2,
            $3,
            $4,
            $5,
            $6,
            $7::date,
            $8
          )
          RETURNING
            id,
            created_at
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
        ]
      );

    const batch = batchResult.rows[0];

    if (!batch) {
      throw new Error(
        "ไม่สามารถสร้าง Batch ได้"
      );
    }

    let insertedCount = 0;

    /* =========================
       SAVE ITEMS
    ========================= */

    for (const item of items) {
      const caseId = String(
        item.caseId ?? ""
      ).trim();

      const urlSms = String(
        item.urlSms ?? ""
      ).trim();

      if (!urlSms) {
        continue;
      }

      const aisResult =
        cleanResult(
          item.checked?.ais
        );

      const trueDtacResult =
        cleanResult(
          item.checked?.trueDtac
        );

      const ntResult =
        cleanResult(
          item.checked?.nt
        );

      const cloudflareResult =
        cleanResult(
          item.checked?.cloudflare
        );

      const aisStatus =
        cleanResultStatus(
          item.checkState?.ais
        );

      const trueDtacStatus =
        cleanResultStatus(
          item.checkState?.trueDtac
        );

      const ntStatus =
        cleanResultStatus(
          item.checkState?.nt
        );

      const cloudflareStatus =
        cleanResultStatus(
          item.checkState?.cloudflare
        );

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
          VALUES (
            $1,
            $2,
            $3,

            $4,
            $5,
            $6,
            $7,

            $8,
            $9,
            $10,
            $11
          )
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
      throw new Error(
        "ไม่พบ URL สำหรับบันทึก"
      );
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
      await client.query(
        "ROLLBACK"
      );
    }

    console.error(
      "Save checker results error:",
      error
    );

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

export default router;

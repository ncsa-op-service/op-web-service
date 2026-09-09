import {
  Router,
  type Request,
  type Response,
} from "express";
import { pool } from "../db.js";

const router = Router();

type CaseType =
  | "fake_domain"
  | "fake_web"
  | "c2_oss"
  | "fake_line";

type ProviderKey =
  | "ais"
  | "trueDtac"
  | "nt"
  | "cloudflare";

type ProviderResultStatus =
  | "success"
  | "nxdomain"
  | "failed"
  | null;

type ResultRow = {
  batch_id: number;
  case_type: CaseType;
  created_at: string;
  ais_result: string | null;
  true_dtac_result: string | null;
  nt_result: string | null;
  cloudflare_result: string | null;
  ais_status: ProviderResultStatus;
  true_dtac_status: ProviderResultStatus;
  nt_status: ProviderResultStatus;
  cloudflare_status: ProviderResultStatus;
};


type UrlFakeWebSummaryRow = {
  id: number;
  batch_id: number;
  case_id: string;
  url_sms: string;
  ais_result: string | null;
  true_dtac_result: string | null;
  nt_result: string | null;
  cloudflare_result: string | null;
  ais_status: ProviderResultStatus;
  true_dtac_status: ProviderResultStatus;
  nt_status: ProviderResultStatus;
  cloudflare_status: ProviderResultStatus;
  created_at: string;
};

type SimpleGraphCount = {
  found: number;
  notFound: number;
  pending: number;
};

type SimpleGraphMap = Record<
  ProviderKey,
  SimpleGraphCount
>;

type ProviderSummary = {
  reachable: number;
  blocked: number;
  pending: number;
};

type ProviderSummaryMap =
  Record<
    ProviderKey,
    ProviderSummary
  >;

const caseTypes: CaseType[] = [
  "fake_domain",
  "fake_web",
  "c2_oss",
  "fake_line",
];

const providers: ProviderKey[] = [
  "ais",
  "trueDtac",
  "nt",
  "cloudflare",
];

function clean(
  value:
    | string
    | null
    | undefined
) {
  return String(
    value ?? ""
  ).trim();
}

function getSimpleResultKind(
  value: string | null,
  status: ProviderResultStatus
) {
  const text = clean(value);

  if (!text) {
    return "pending" as const;
  }

  if (
    status === "nxdomain" ||
    text
      .toUpperCase()
      .includes(
        "DNS_PROBE_FINISHED_NXDOMAIN"
      )
  ) {
    return "notFound" as const;
  }

  if (status === "failed") {
    return "notFound" as const;
  }

  if (status === "success") {
    return "found" as const;
  }

  // fallback สำหรับข้อมูลเก่าที่ยังไม่มี *_status
  if (
    text.startsWith("http://") ||
    text.startsWith("https://")
  ) {
    return "found" as const;
  }

  return "notFound" as const;
}

function createSimpleGraph():
  SimpleGraphMap {
  return {
    ais: {
      found: 0,
      notFound: 0,
      pending: 0,
    },
    trueDtac: {
      found: 0,
      notFound: 0,
      pending: 0,
    },
    nt: {
      found: 0,
      notFound: 0,
      pending: 0,
    },
    cloudflare: {
      found: 0,
      notFound: 0,
      pending: 0,
    },
  };
}

function isReachable(
  value: string | null,
  caseType: CaseType
) {
  const text =
    clean(value).toLowerCase();

  if (
    text.startsWith("เปิดได้")
  ) {
    return true;
  }

  if (
    caseType === "fake_line" &&
    (
      text.startsWith("http://") ||
      text.startsWith("https://")
    )
  ) {
    return true;
  }

  return false;
}

function getProviderValue(
  row: ResultRow,
  provider: ProviderKey
) {
  if (provider === "ais") {
    return row.ais_result;
  }

  if (provider === "trueDtac") {
    return row.true_dtac_result;
  }

  if (provider === "nt") {
    return row.nt_result;
  }

  return row.cloudflare_result;
}

function getProviderStatus(
  row: ResultRow,
  provider: ProviderKey
): ProviderResultStatus {
  if (provider === "ais") {
    return row.ais_status ?? null;
  }

  if (provider === "trueDtac") {
    return row.true_dtac_status ?? null;
  }

  if (provider === "nt") {
    return row.nt_status ?? null;
  }

  return row.cloudflare_status ?? null;
}

function createProviderSummary():
  ProviderSummary {
  return {
    reachable: 0,
    blocked: 0,
    pending: 0,
  };
}

function createProviderMap():
  ProviderSummaryMap {
  return {
    ais: createProviderSummary(),
    trueDtac:
      createProviderSummary(),
    nt: createProviderSummary(),
    cloudflare:
      createProviderSummary(),
  };
}

async function requireSummaryEditor(
  req: Request,
  res: Response
): Promise<{
  id: number;
  role: string;
} | null> {
  const userId =
    Number(
      req.header("x-user-id")
    );

  if (
    !Number.isInteger(userId) ||
    userId <= 0
  ) {
    res.status(401).json({
      message:
        "ไม่พบข้อมูลผู้ใช้งาน",
    });

    return null;
  }

  const result =
    await pool.query<{
      id: number;
      role: string;
    }>(
      `
        SELECT id, role
        FROM users
        WHERE id = $1
        LIMIT 1
      `,
      [userId]
    );

  if (
    result.rows.length === 0
  ) {
    res.status(401).json({
      message:
        "ไม่พบผู้ใช้งานในระบบ",
    });

    return null;
  }

  const user =
    result.rows[0]!;

  if (
    user.role !==
      "super_admin" &&
    user.role !==
      "editor"
  ) {
    res.status(403).json({
      message:
        "เฉพาะ Super Admin และ Editor เท่านั้นที่แก้ไข Summary ได้",
    });

    return null;
  }

  return user;
}

router.get(
  "/",
  async (_req, res) => {
    try {
      const reportDateResult =
        await pool.query<{
          report_date:
            | string
            | null;
        }>(`
          SELECT
            MAX(detected_date)::text
              AS report_date
          FROM url_fake_batches
        `);

      const reportDate =
        reportDateResult.rows[0]
          ?.report_date ?? null;

      const latestRowsResult =
        await pool.query<ResultRow>(`
          WITH latest_batches AS (
            SELECT DISTINCT ON (
              case_type
            )
              id,
              case_type,
              created_at
            FROM url_fake_batches
            WHERE
              case_status =
                'active'
            ORDER BY
              case_type,
              detected_date DESC,
              id DESC
          )

          SELECT
            latest_batches.id
              AS batch_id,
            latest_batches.case_type,
            latest_batches.created_at,
            results.ais_result,
            results.true_dtac_result,
            results.nt_result,
            results.cloudflare_result,
            results.ais_status,
            results.true_dtac_status,
            results.nt_status,
            results.cloudflare_status

          FROM latest_batches

          JOIN url_fake_results
            AS results
            ON results.batch_id =
              latest_batches.id

          ORDER BY
            latest_batches.case_type,
            results.id
        `);

      const additionalByType:
        Record<
          CaseType,
          number
        > = {
          fake_domain: 0,
          fake_web: 0,
          c2_oss: 0,
          fake_line: 0,
        };

      if (reportDate) {
        const additionalResult =
          await pool.query<{
            case_type: CaseType;
            total: number;
          }>(
            `
              SELECT
                batch.case_type,
                COUNT(
                  DISTINCT
                  result.url_sms
                )::int AS total

              FROM url_fake_batches
                AS batch

              JOIN url_fake_results
                AS result
                ON result.batch_id =
                  batch.id

              WHERE
                batch.detected_date =
                  $1::date

                AND NOT EXISTS (
                  SELECT 1

                  FROM url_fake_batches
                    AS old_batch

                  JOIN url_fake_results
                    AS old_result
                    ON old_result.batch_id =
                      old_batch.id

                  WHERE
                    old_batch.case_type =
                      batch.case_type
                    AND
                    old_batch.detected_date <
                      batch.detected_date
                    AND LOWER(
                      TRIM(
                        old_result.url_sms
                      )
                    ) =
                    LOWER(
                      TRIM(
                        result.url_sms
                      )
                    )
                )

              GROUP BY
                batch.case_type
            `,
            [reportDate]
          );

        for (
          const row
          of additionalResult.rows
        ) {
          additionalByType[
            row.case_type
          ] = row.total;
        }
      }

      const categories = {
        fakeDomain: {
          active: 0,
          providers:
            createProviderMap(),
        },

        fakeWeb: {
          active: 0,
          providers:
            createProviderMap(),
        },

        c2Oss: {
          active: 0,
          providers:
            createProviderMap(),
        },

        fakeLine: {
          active: 0,
          reachable: 0,
          blocked: 0,
          pending: 0,
        },
      };

      const categoryKey = {
        fake_domain:
          "fakeDomain",
        fake_web:
          "fakeWeb",
        c2_oss:
          "c2Oss",
      } as const;

      for (
        const caseType
        of caseTypes
      ) {
        const rows =
          latestRowsResult.rows.filter(
            (row) =>
              row.case_type ===
              caseType
          );

        if (
          caseType ===
          "fake_line"
        ) {
          categories.fakeLine.active =
            rows.length;

          for (
            const row
            of rows
          ) {
            const values =
              providers.map(
                (provider) =>
                  getProviderValue(
                    row,
                    provider
                  )
              );

            const reachable =
              values.some(
                (value) =>
                  isReachable(
                    value,
                    caseType
                  )
              );

            const hasResult =
              values.some(
                (value) =>
                  clean(value) !==
                  ""
              );

            if (reachable) {
              categories.fakeLine
                .reachable += 1;
            } else if (
              hasResult
            ) {
              categories.fakeLine
                .blocked += 1;
            } else {
              categories.fakeLine
                .pending += 1;
            }
          }

          continue;
        }

        const key =
          categoryKey[
            caseType
          ];

        const category =
          categories[key];

        category.active =
          rows.length;

        for (
          const provider
          of providers
        ) {
          for (
            const row
            of rows
          ) {
            const value =
              getProviderValue(
                row,
                provider
              );

            if (
              isReachable(
                value,
                caseType
              )
            ) {
              category
                .providers[
                  provider
                ]
                .reachable += 1;
            } else if (
              clean(value) !== ""
            ) {
              category
                .providers[
                  provider
                ]
                .blocked += 1;
            } else {
              category
                .providers[
                  provider
                ]
                .pending += 1;
            }
          }
        }
      }

      const additionalTotal =
        Object.values(
          additionalByType
        ).reduce(
          (
            total,
            value
          ) =>
            total + value,
          0
        );

      const totalCases =
        categories.fakeDomain
          .active +
        categories.fakeWeb
          .active +
        categories.c2Oss
          .active +
        categories.fakeLine
          .active;

      const updateTimes =
        latestRowsResult.rows
          .map(
            (row) =>
              row.created_at
          )
          .sort();

      const updatedAt =
        updateTimes.length > 0
          ? updateTimes[
              updateTimes.length -
                1
            ]
          : null;

      /*
        ข้อมูล URL Fake Web สำหรับ Main Summary
        ใช้ batch ล่าสุดที่บันทึกจากหน้า Checker
        โดยรองรับทั้ง fake_domain และ fake_web
      */
      const urlFakeWebBatchResult =
        await pool.query<{
          id: number;
          original_file_name:
            | string
            | null;
          selected_provider:
            | string
            | null;
          public_ip:
            | string
            | null;
          isp:
            | string
            | null;
          detected_provider:
            | string
            | null;
          case_type: CaseType;
          detected_date:
            | string
            | null;
          case_status:
            | string
            | null;
          created_at: string;
        }>(`
          SELECT
            id,
            original_file_name,
            selected_provider,
            public_ip,
            isp,
            detected_provider,
            case_type,
            detected_date::text
              AS detected_date,
            case_status,
            created_at
          FROM url_fake_batches
          WHERE
            case_status = 'active'
            AND case_type IN (
              'fake_domain',
              'fake_web'
            )
          ORDER BY
            detected_date DESC,
            id DESC
          LIMIT 1
        `);

      const urlFakeWebBatch =
        urlFakeWebBatchResult.rows[0] ??
        null;

      let urlFakeWebResults:
        UrlFakeWebSummaryRow[] = [];

      const urlFakeWebGraph =
        createSimpleGraph();

      if (urlFakeWebBatch) {
        const result =
          await pool.query<
            UrlFakeWebSummaryRow
          >(
            `
              SELECT
                id,
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
                cloudflare_status,
                created_at
              FROM url_fake_results
              WHERE batch_id = $1
              ORDER BY id ASC
            `,
            [urlFakeWebBatch.id]
          );

        urlFakeWebResults =
          result.rows;

        for (
          const row
          of urlFakeWebResults
        ) {
          const values: Array<{
            provider: ProviderKey;
            value: string | null;
            status: ProviderResultStatus;
          }> = [
            {
              provider: "ais",
              value: row.ais_result,
              status: row.ais_status,
            },
            {
              provider: "trueDtac",
              value: row.true_dtac_result,
              status: row.true_dtac_status,
            },
            {
              provider: "nt",
              value: row.nt_result,
              status: row.nt_status,
            },
            {
              provider: "cloudflare",
              value: row.cloudflare_result,
              status: row.cloudflare_status,
            },
          ];

          for (
            const item
            of values
          ) {
            const kind =
              getSimpleResultKind(
                item.value,
                item.status
              );

            urlFakeWebGraph[
              item.provider
            ][kind] += 1;
          }
        }
      }

      return res.json({
        reportDate,
        updatedAt,

        additional: {
          total:
            additionalTotal,

          fakeDomain:
            additionalByType
              .fake_domain,

          fakeWeb:
            additionalByType
              .fake_web,

          c2Oss:
            additionalByType
              .c2_oss,

          fakeLine:
            additionalByType
              .fake_line,
        },

        urlFakeWeb: {
          batch:
            urlFakeWebBatch,

          results:
            urlFakeWebResults,

          graph:
            urlFakeWebGraph,
        },

        totalCases,
        categories,
      });
    } catch (error) {
      console.error(
        "Summary API error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "ไม่สามารถโหลด Summary ได้",
        });
    }
  }
);

router.get(
  "/edit/:caseType",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const caseType =
        String(
          req.params.caseType
        ) as CaseType;

      if (
        !caseTypes.includes(
          caseType
        )
      ) {
        return res
          .status(400)
          .json({
            message:
              "ประเภทข้อมูลไม่ถูกต้อง",
          });
      }

      const batchResult =
        await pool.query(
          `
            SELECT
              id,
              original_file_name,
              selected_provider,
              public_ip,
              isp,
              detected_provider,
              case_type,
              detected_date,
              case_status,
              created_at
            FROM url_fake_batches
            WHERE
              case_type = $1
              AND case_status =
                'active'
            ORDER BY
              detected_date DESC,
              id DESC
            LIMIT 1
          `,
          [caseType]
        );

      if (
        batchResult.rows.length ===
        0
      ) {
        return res.json({
          batch: null,
          results: [],
        });
      }

      const batch =
        batchResult.rows[0];

      const resultsResult =
        await pool.query(
          `
            SELECT
              id,
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
              cloudflare_status,
              created_at
            FROM url_fake_results
            WHERE batch_id = $1
            ORDER BY id ASC
          `,
          [batch.id]
        );

      return res.json({
        batch,
        results:
          resultsResult.rows,
      });
    } catch (error) {
      console.error(
        "Summary edit load error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "โหลดข้อมูลสำหรับแก้ไขไม่สำเร็จ",
        });
    }
  }
);

router.patch(
  "/edit/results/:resultId",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const editor =
        await requireSummaryEditor(
          req,
          res
        );

      if (!editor) {
        return;
      }

      const resultId =
        Number(
          req.params.resultId
        );

      if (
        !Number.isInteger(
          resultId
        ) ||
        resultId <= 0
      ) {
        return res
          .status(400)
          .json({
            message:
              "resultId ไม่ถูกต้อง",
          });
      }

      const caseId =
        clean(
          req.body?.caseId
        );

      const urlSms =
        clean(
          req.body?.urlSms
        );

      if (
        !caseId ||
        !urlSms
      ) {
        return res
          .status(400)
          .json({
            message:
              "Case ID และ URL ต้องไม่ว่าง",
          });
      }

      const toNullable = (
        value: unknown
      ) => {
        const text =
          clean(
            value as
              | string
              | null
              | undefined
          );

        return text || null;
      };

      const toResultStatus = (
        value: unknown
      ): ProviderResultStatus => {
        const status = clean(
          value as
            | string
            | null
            | undefined
        ).toLowerCase();

        if (
          status === "success" ||
          status === "nxdomain" ||
          status === "failed"
        ) {
          return status;
        }

        return null;
      };

      const updated =
        await pool.query(
          `
            UPDATE url_fake_results
            SET
              case_id = $2,
              url_sms = $3,
              ais_result = $4,
              true_dtac_result = $5,
              nt_result = $6,
              cloudflare_result = $7,
              ais_status =
                CASE
                  WHEN $8::text IS NULL
                    THEN ais_status
                  ELSE $8
                END,
              true_dtac_status =
                CASE
                  WHEN $9::text IS NULL
                    THEN true_dtac_status
                  ELSE $9
                END,
              nt_status =
                CASE
                  WHEN $10::text IS NULL
                    THEN nt_status
                  ELSE $10
                END,
              cloudflare_status =
                CASE
                  WHEN $11::text IS NULL
                    THEN cloudflare_status
                  ELSE $11
                END
            WHERE id = $1
            RETURNING *
          `,
          [
            resultId,
            caseId,
            urlSms,
            toNullable(
              req.body?.aisResult
            ),
            toNullable(
              req.body
                ?.trueDtacResult
            ),
            toNullable(
              req.body?.ntResult
            ),
            toNullable(
              req.body
                ?.cloudflareResult
            ),
            req.body?.aisStatus === undefined
              ? null
              : toResultStatus(
                  req.body?.aisStatus
                ),
            req.body?.trueDtacStatus === undefined
              ? null
              : toResultStatus(
                  req.body?.trueDtacStatus
                ),
            req.body?.ntStatus === undefined
              ? null
              : toResultStatus(
                  req.body?.ntStatus
                ),
            req.body?.cloudflareStatus === undefined
              ? null
              : toResultStatus(
                  req.body?.cloudflareStatus
                ),
          ]
        );

      if (
        updated.rows.length ===
        0
      ) {
        return res
          .status(404)
          .json({
            message:
              "ไม่พบข้อมูลที่ต้องการแก้ไข",
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
            $2,
            $3,
            $4
          )
        `,
        [
          editor.id,
          "summary_result_update",
          `แก้ไขข้อมูล Summary result #${resultId}`,
          req.ip ?? null,
        ]
      );

      return res.json({
        message:
          "บันทึกการแก้ไขสำเร็จ",
        result:
          updated.rows[0],
      });
    } catch (error) {
      console.error(
        "Summary edit result error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "บันทึกการแก้ไขไม่สำเร็จ",
        });
    }
  }
);

router.patch(
  "/edit/batches/:batchId",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const editor =
        await requireSummaryEditor(
          req,
          res
        );

      if (!editor) {
        return;
      }

      const batchId =
        Number(
          req.params.batchId
        );

      if (
        !Number.isInteger(
          batchId
        ) ||
        batchId <= 0
      ) {
        return res
          .status(400)
          .json({
            message:
              "batchId ไม่ถูกต้อง",
          });
      }

      const detectedDate =
        clean(
          req.body
            ?.detectedDate
        );

      const caseStatus =
        clean(
          req.body
            ?.caseStatus
        );

      if (
        !detectedDate
      ) {
        return res
          .status(400)
          .json({
            message:
              "กรุณาระบุวันที่ตรวจพบ",
          });
      }

      if (
        caseStatus !==
          "active" &&
        caseStatus !==
          "inactive"
      ) {
        return res
          .status(400)
          .json({
            message:
              "caseStatus ไม่ถูกต้อง",
          });
      }

      const updated =
        await pool.query(
          `
            UPDATE url_fake_batches
            SET
              detected_date =
                $2::date,
              case_status = $3
            WHERE id = $1
            RETURNING *
          `,
          [
            batchId,
            detectedDate,
            caseStatus,
          ]
        );

      if (
        updated.rows.length ===
        0
      ) {
        return res
          .status(404)
          .json({
            message:
              "ไม่พบ Batch ที่ต้องการแก้ไข",
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
            $2,
            $3,
            $4
          )
        `,
        [
          editor.id,
          "summary_batch_update",
          `แก้ไขข้อมูล Summary batch #${batchId}`,
          req.ip ?? null,
        ]
      );

      return res.json({
        message:
          "บันทึกข้อมูล Batch สำเร็จ",
        batch:
          updated.rows[0],
      });
    } catch (error) {
      console.error(
        "Summary edit batch error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "บันทึกข้อมูล Batch ไม่สำเร็จ",
        });
    }
  }
);

export default router;

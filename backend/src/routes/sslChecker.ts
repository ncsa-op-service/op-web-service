import {
  Router,
  type Request,
  type Response,
} from "express";
import tls from "tls";
import { pool } from "../db.js";

const router = Router();

type SSLStatus =
  | "VALID"
  | "EXPIRING_SOON"
  | "EXPIRED"
  | "ERROR";

type SSLResult = {
  url: string;
  hostname: string;
  port: number;
  validFrom: string | null;
  expirationDate: string | null;
  daysLeft: number | null;
  status: SSLStatus;
  issuer: string;
  subject: string;
  checkedAt: string;
  error: string | null;
};

type JobStatus =
  | "READY"
  | "RUNNING"
  | "PAUSED"
  | "COMPLETED"
  | "ERROR";

type RoundStatus =
  | "WAITING"
  | "RUNNING"
  | "PAUSED"
  | "COMPLETED"
  | "ERROR";

const ROUND_SIZE = 500;
const REQUEST_SIZE = 100;
const MAX_CREATE_URLS = 500_000;

const CONNECT_TIMEOUT_MS =
  10_000;

const CONCURRENCY =
  12;

/*
  Job ที่กำลังทำงานอยู่ใน process นี้

  ข้อมูลผลจริงยังบันทึกใน PostgreSQL
  Map นี้ใช้แค่ป้องกัน start ซ้ำใน process เดียว
*/
const activeJobs =
  new Map<
    number,
    Promise<void>
  >();

const pauseRequested =
  new Set<number>();

function cleanText(
  value: unknown
) {
  return String(
    value ?? ""
  )
    .normalize("NFKC")
    .replace(
      /[\u200B-\u200D\uFEFF]/g,
      ""
    )
    .replace(/\u00A0/g, " ")
    .trim();
}

function normalizeUrl(
  value: unknown
) {
  const raw =
    cleanText(value);

  if (!raw) {
    throw new Error(
      "URL ว่าง"
    );
  }

  const candidate =
    /^https?:\/\//i.test(
      raw
    )
      ? raw
      : `https://${raw}`;

  const parsed =
    new URL(candidate);

  if (!parsed.hostname) {
    throw new Error(
      "URL ไม่ถูกต้อง"
    );
  }

  const port =
    parsed.port
      ? Number(
          parsed.port
        )
      : 443;

  if (
    !Number.isInteger(
      port
    ) ||
    port <= 0 ||
    port > 65535
  ) {
    throw new Error(
      "Port ไม่ถูกต้อง"
    );
  }

  return {
    originalUrl:
      raw,
    hostname:
      parsed.hostname,
    port,
  };
}

function formatDistinguishedName(
  value:
    | Record<
        string,
        string
      >
    | undefined
) {
  if (!value) {
    return "";
  }

  const preferred =
    [
      "O",
      "OU",
      "CN",
      "C",
    ];

  const parts:
    string[] = [];

  for (
    const key of preferred
  ) {
    const item =
      value[key];

    if (item) {
      parts.push(
        `${key}=${item}`
      );
    }
  }

  for (
    const [
      key,
      item,
    ] of
      Object.entries(
        value
      )
  ) {
    if (
      preferred.includes(
        key
      ) ||
      !item
    ) {
      continue;
    }

    parts.push(
      `${key}=${item}`
    );
  }

  return parts.join(
    ", "
  );
}

function calculateStatus(
  expirationDate: Date
): {
  status: SSLStatus;
  daysLeft: number;
} {
  const now =
    Date.now();

  const diff =
    expirationDate.getTime() -
    now;

  const daysLeft =
    Math.ceil(
      diff /
        86_400_000
    );

  if (daysLeft < 0) {
    return {
      status:
        "EXPIRED",
      daysLeft,
    };
  }

  if (daysLeft <= 30) {
    return {
      status:
        "EXPIRING_SOON",
      daysLeft,
    };
  }

  return {
    status:
      "VALID",
    daysLeft,
  };
}

async function checkCertificate(
  input: unknown
): Promise<SSLResult> {
  const checkedAt =
    new Date()
      .toISOString();

  let originalUrl =
    cleanText(input);

  let hostname = "";
  let port = 443;

  try {
    const normalized =
      normalizeUrl(input);

    originalUrl =
      normalized.originalUrl;

    hostname =
      normalized.hostname;

    port =
      normalized.port;
  } catch (error) {
    return {
      url:
        originalUrl,
      hostname,
      port,
      validFrom:
        null,
      expirationDate:
        null,
      daysLeft:
        null,
      status:
        "ERROR",
      issuer:
        "",
      subject:
        "",
      checkedAt,
      error:
        error instanceof
          Error
          ? error.message
          : "URL ไม่ถูกต้อง",
    };
  }

  return new Promise(
    (resolve) => {
      let settled =
        false;

      const finish = (
        result: SSLResult
      ) => {
        if (settled) {
          return;
        }

        settled = true;

        resolve(result);
      };

      const socket =
        tls.connect({
          host:
            hostname,
          port,
          servername:
            hostname,

          /*
            false เพื่อให้ยังอ่าน cert ที่
            expired / self-signed ได้
          */
          rejectUnauthorized:
            false,
        });

      socket.setTimeout(
        CONNECT_TIMEOUT_MS
      );

      socket.once(
        "secureConnect",
        () => {
          try {
            const certificate =
              socket.getPeerCertificate(
                true
              );

            if (
              !certificate ||
              Object.keys(
                certificate
              ).length === 0
            ) {
              finish({
                url:
                  originalUrl,
                hostname,
                port,
                validFrom:
                  null,
                expirationDate:
                  null,
                daysLeft:
                  null,
                status:
                  "ERROR",
                issuer:
                  "",
                subject:
                  "",
                checkedAt,
                error:
                  "ไม่พบ SSL Certificate",
              });

              socket.destroy();
              return;
            }

            const validFromDate =
              new Date(
                certificate.valid_from
              );

            const validToDate =
              new Date(
                certificate.valid_to
              );

            if (
              Number.isNaN(
                validFromDate.getTime()
              ) ||
              Number.isNaN(
                validToDate.getTime()
              )
            ) {
              throw new Error(
                "Certificate ไม่มีวันที่ valid_from / valid_to ที่อ่านได้"
              );
            }

            const {
              status,
              daysLeft,
            } =
              calculateStatus(
                validToDate
              );

            finish({
              url:
                originalUrl,
              hostname,
              port,
              validFrom:
                validFromDate.toISOString(),
              expirationDate:
                validToDate.toISOString(),
              daysLeft,
              status,
              issuer:
                formatDistinguishedName(
                  certificate.issuer as
                    | Record<
                        string,
                        string
                      >
                    | undefined
                ),
              subject:
                formatDistinguishedName(
                  certificate.subject as
                    | Record<
                        string,
                        string
                      >
                    | undefined
                ),
              checkedAt,
              error:
                null,
            });
          } catch (error) {
            finish({
              url:
                originalUrl,
              hostname,
              port,
              validFrom:
                null,
              expirationDate:
                null,
              daysLeft:
                null,
              status:
                "ERROR",
              issuer:
                "",
              subject:
                "",
              checkedAt,
              error:
                error instanceof
                  Error
                  ? error.message
                  : "อ่าน Certificate ไม่สำเร็จ",
            });
          } finally {
            socket.destroy();
          }
        }
      );

      socket.once(
        "timeout",
        () => {
          finish({
            url:
              originalUrl,
            hostname,
            port,
            validFrom:
              null,
            expirationDate:
              null,
            daysLeft:
              null,
            status:
              "ERROR",
            issuer:
              "",
            subject:
              "",
            checkedAt,
            error:
              `Connection timeout (${CONNECT_TIMEOUT_MS / 1000}s)`,
          });

          socket.destroy();
        }
      );

      socket.once(
        "error",
        (
          error
        ) => {
          finish({
            url:
              originalUrl,
            hostname,
            port,
            validFrom:
              null,
            expirationDate:
              null,
            daysLeft:
              null,
            status:
              "ERROR",
            issuer:
              "",
            subject:
              "",
            checkedAt,
            error:
              error.message,
          });

          socket.destroy();
        }
      );
    }
  );
}

async function mapWithConcurrency<
  T,
  R
>(
  items: T[],
  concurrency: number,
  worker: (
    item: T,
    index: number
  ) => Promise<R>
): Promise<R[]> {
  const results:
    R[] =
    new Array<R>(
      items.length
    );

  let nextIndex = 0;

  async function runWorker() {
    while (true) {
      const currentIndex =
        nextIndex;

      nextIndex += 1;

      if (
        currentIndex >=
        items.length
      ) {
        return;
      }

      const item =
        items[currentIndex]!;

      results[currentIndex] =
        await worker(
          item,
          currentIndex
        );
    }
  }

  const workerCount =
    Math.min(
      Math.max(
        1,
        concurrency
      ),
      items.length
    );

  const workers =
    Array.from(
      {
        length:
          workerCount,
      },
      () =>
        runWorker()
    );

  await Promise.all(
    workers
  );

  return results;
}


async function requireSuperAdmin(
  req: Request,
  res: Response
): Promise<number | null> {
  const userId =
    Number(
      req.header(
        "x-user-id"
      )
    );

  if (
    !Number.isInteger(
      userId
    ) ||
    userId <= 0
  ) {
    res
      .status(401)
      .json({
        message:
          "ไม่พบข้อมูลผู้ใช้งานสำหรับการแก้ไข",
      });

    return null;
  }

  const userResult =
    await pool.query(
      `
        SELECT
          id,
          role
        FROM users
        WHERE id = $1
        LIMIT 1
      `,
      [userId]
    );

  if (
    userResult.rows.length ===
    0
  ) {
    res
      .status(401)
      .json({
        message:
          "ไม่พบผู้ใช้งานในระบบ",
      });

    return null;
  }

  if (
    userResult.rows[0]
      .role !==
    "super_admin"
  ) {
    res
      .status(403)
      .json({
        message:
          "เฉพาะ Super Admin เท่านั้นที่แก้ไขผล SSL/TLS ได้",
      });

    return null;
  }

  return userId;
}

function parsePositiveInt(
  value: unknown,
  fallback: number
) {
  const parsed =
    Number(value);

  if (
    !Number.isInteger(
      parsed
    ) ||
    parsed <= 0
  ) {
    return fallback;
  }

  return parsed;
}

async function refreshJobProgress(
  jobId: number
) {
  const progress =
    await pool.query(
      `
        SELECT
          COUNT(*)::int
            AS checked_count,

          COUNT(*) FILTER (
            WHERE r.status =
              'COMPLETED'
          )::int
            AS completed_rounds

        FROM ssl_check_rounds r

        LEFT JOIN
          ssl_check_results sr
        ON
          sr.round_id = r.id

        WHERE
          r.job_id = $1
      `,
      [jobId]
    );

  /*
    Query ด้านบน join แล้วทำให้
    completed_rounds ซ้ำตามจำนวน result
    จึงคำนวณแยกให้แม่นยำ
  */
  const counts =
    await pool.query(
      `
        SELECT
          (
            SELECT COUNT(*)::int
            FROM ssl_check_results
            WHERE job_id = $1
          ) AS checked_count,

          (
            SELECT COUNT(*)::int
            FROM ssl_check_rounds
            WHERE
              job_id = $1
              AND status =
                'COMPLETED'
          ) AS completed_rounds
      `,
      [jobId]
    );

  const checkedCount =
    Number(
      counts.rows[0]
        ?.checked_count ?? 0
    );

  const completedRounds =
    Number(
      counts.rows[0]
        ?.completed_rounds ?? 0
    );

  await pool.query(
    `
      UPDATE ssl_check_jobs
      SET
        checked_count = $2,
        completed_rounds = $3,
        updated_at = NOW()
      WHERE id = $1
    `,
    [
      jobId,
      checkedCount,
      completedRounds,
    ]
  );

  return {
    checkedCount,
    completedRounds,
  };
}

async function saveChunkResults(
  jobId: number,
  roundId: number,
  items: Array<{
    file_order: number;
    url: string;
  }>,
  results: SSLResult[]
) {
  const client =
    await pool.connect();

  try {
    await client.query(
      "BEGIN"
    );

    for (
      let index = 0;
      index < results.length;
      index += 1
    ) {
      const item =
        items[index];

      const result =
        results[index];

      if (
        !item ||
        !result
      ) {
        continue;
      }

      await client.query(
        `
          INSERT INTO
            ssl_check_results (
              job_id,
              round_id,
              file_order,
              url,
              hostname,
              port,
              valid_from,
              expiration_date,
              days_left,
              status,
              issuer,
              subject,
              checked_at,
              error,
              updated_at
            )
          VALUES (
            $1, $2, $3, $4, $5,
            $6, $7, $8, $9, $10,
            $11, $12, $13, $14,
            NOW()
          )
          ON CONFLICT (
            job_id,
            file_order
          )
          DO UPDATE SET
            round_id =
              EXCLUDED.round_id,
            url =
              EXCLUDED.url,
            hostname =
              EXCLUDED.hostname,
            port =
              EXCLUDED.port,
            valid_from =
              EXCLUDED.valid_from,
            expiration_date =
              EXCLUDED.expiration_date,
            days_left =
              EXCLUDED.days_left,
            status =
              EXCLUDED.status,
            issuer =
              EXCLUDED.issuer,
            subject =
              EXCLUDED.subject,
            checked_at =
              EXCLUDED.checked_at,
            error =
              EXCLUDED.error,
            updated_at =
              NOW()
        `,
        [
          jobId,
          roundId,
          item.file_order,
          result.url,
          result.hostname,
          result.port,
          result.validFrom,
          result.expirationDate,
          result.daysLeft,
          result.status,
          result.issuer,
          result.subject,
          result.checkedAt,
          result.error,
        ]
      );
    }

    const roundCountResult =
      await client.query(
        `
          SELECT
            COUNT(*)::int
              AS checked_count
          FROM ssl_check_results
          WHERE round_id = $1
        `,
        [roundId]
      );

    const roundChecked =
      Number(
        roundCountResult
          .rows[0]
          ?.checked_count ?? 0
      );

    await client.query(
      `
        UPDATE ssl_check_rounds
        SET
          checked_count = $2,
          updated_at = NOW()
        WHERE id = $1
      `,
      [
        roundId,
        roundChecked,
      ]
    );

    await client.query(
      "COMMIT"
    );

    return roundChecked;
  } catch (error) {
    await client.query(
      "ROLLBACK"
    );

    throw error;
  } finally {
    client.release();
  }
}

async function runJob(
  jobId: number,
  roundLimit:
    number | null
) {
  try {
    pauseRequested.delete(
      jobId
    );

    const jobResult =
      await pool.query(
        `
          SELECT *
          FROM ssl_check_jobs
          WHERE id = $1
          LIMIT 1
        `,
        [jobId]
      );

    if (
      jobResult.rows.length ===
      0
    ) {
      return;
    }

    const job =
      jobResult.rows[0];

    await pool.query(
      `
        UPDATE ssl_check_jobs
        SET
          status = 'RUNNING',
          started_at =
            COALESCE(
              started_at,
              NOW()
            ),
          paused_at = NULL,
          last_error = NULL,
          updated_at = NOW()
        WHERE id = $1
      `,
      [jobId]
    );

    const roundsResult =
      await pool.query(
        `
          SELECT *
          FROM ssl_check_rounds
          WHERE
            job_id = $1
            AND status <>
              'COMPLETED'
          ORDER BY
            round_number ASC
        `,
        [jobId]
      );

    const rounds =
      roundLimit === null
        ? roundsResult.rows
        : roundsResult.rows.slice(
            0,
            roundLimit
          );

    for (
      const round of rounds
    ) {
      if (
        pauseRequested.has(
          jobId
        )
      ) {
        await pool.query(
          `
            UPDATE ssl_check_jobs
            SET
              status = 'PAUSED',
              paused_at = NOW(),
              updated_at = NOW()
            WHERE id = $1
          `,
          [jobId]
        );

        return;
      }

      await pool.query(
        `
          UPDATE ssl_check_jobs
          SET
            current_round = $2,
            status = 'RUNNING',
            updated_at = NOW()
          WHERE id = $1
        `,
        [
          jobId,
          round.round_number,
        ]
      );

      await pool.query(
        `
          UPDATE ssl_check_rounds
          SET
            status = 'RUNNING',
            started_at =
              COALESCE(
                started_at,
                NOW()
              ),
            paused_at = NULL,
            updated_at = NOW()
          WHERE id = $1
        `,
        [round.id]
      );

      while (true) {
        if (
          pauseRequested.has(
            jobId
          )
        ) {
          await pool.query(
            `
              UPDATE ssl_check_rounds
              SET
                status = 'PAUSED',
                paused_at = NOW(),
                updated_at = NOW()
              WHERE id = $1
            `,
            [round.id]
          );

          await pool.query(
            `
              UPDATE ssl_check_jobs
              SET
                status = 'PAUSED',
                paused_at = NOW(),
                updated_at = NOW()
              WHERE id = $1
            `,
            [jobId]
          );

          return;
        }

        /*
          ดึงเฉพาะ URL ในรอบนี้
          ที่ยังไม่มีผลตรวจ
          ครั้งละ 100 URL
        */
        const pendingResult =
          await pool.query(
            `
              SELECT
                i.file_order,
                i.url
              FROM ssl_check_items i

              LEFT JOIN
                ssl_check_results r
              ON
                r.job_id =
                  i.job_id
                AND r.file_order =
                  i.file_order

              WHERE
                i.job_id = $1
                AND i.round_number =
                  $2
                AND r.id IS NULL

              ORDER BY
                i.file_order ASC

              LIMIT $3
            `,
            [
              jobId,
              round.round_number,
              REQUEST_SIZE,
            ]
          );

        const items =
          pendingResult.rows as Array<{
            file_order: number;
            url: string;
          }>;

        if (
          items.length === 0
        ) {
          break;
        }

        const checkedResults =
          await mapWithConcurrency(
            items,
            CONCURRENCY,
            (
              item
            ) =>
              checkCertificate(
                item.url
              )
          );

        const roundChecked =
          await saveChunkResults(
            jobId,
            round.id,
            items,
            checkedResults
          );

        await refreshJobProgress(
          jobId
        );

        if (
          roundChecked >=
          Number(
            round.total_urls
          )
        ) {
          break;
        }
      }

      /*
        รอบนี้ครบจริงแล้ว
        จึงค่อย mark COMPLETED
      */
      await pool.query(
        `
          UPDATE ssl_check_rounds
          SET
            status = 'COMPLETED',
            checked_count =
              total_urls,
            completed_at = NOW(),
            updated_at = NOW()
          WHERE id = $1
        `,
        [round.id]
      );

      await refreshJobProgress(
        jobId
      );
    }

    const remaining =
      await pool.query(
        `
          SELECT
            COUNT(*)::int
              AS remaining
          FROM ssl_check_rounds
          WHERE
            job_id = $1
            AND status <>
              'COMPLETED'
        `,
        [jobId]
      );

    const remainingCount =
      Number(
        remaining.rows[0]
          ?.remaining ?? 0
      );

    if (
      remainingCount === 0
    ) {
      await pool.query(
        `
          UPDATE ssl_check_jobs
          SET
            status = 'COMPLETED',
            checked_count =
              total_urls,
            completed_rounds =
              total_rounds,
            completed_at = NOW(),
            paused_at = NULL,
            updated_at = NOW()
          WHERE id = $1
        `,
        [jobId]
      );
    } else {
      /*
        ตรวจครบตามจำนวนรอบที่ผู้ใช้เลือกแล้ว
        ให้หยุดพักโดยอัตโนมัติ
      */
      await pool.query(
        `
          UPDATE ssl_check_jobs
          SET
            status = 'PAUSED',
            paused_at = NOW(),
            updated_at = NOW()
          WHERE id = $1
        `,
        [jobId]
      );
    }
  } catch (error) {
    console.error(
      `SSL Job ${jobId} error:`,
      error
    );

    await pool.query(
      `
        UPDATE ssl_check_jobs
        SET
          status = 'ERROR',
          last_error = $2,
          updated_at = NOW()
        WHERE id = $1
      `,
      [
        jobId,
        error instanceof
          Error
          ? error.message
          : "เกิดข้อผิดพลาดระหว่างตรวจสอบ SSL/TLS",
      ]
    );
  } finally {
    activeJobs.delete(
      jobId
    );

    pauseRequested.delete(
      jobId
    );
  }
}

function startBackgroundJob(
  jobId: number,
  roundLimit:
    number | null
) {
  if (
    activeJobs.has(
      jobId
    )
  ) {
    return false;
  }

  const promise =
    runJob(
      jobId,
      roundLimit
    );

  activeJobs.set(
    jobId,
    promise
  );

  void promise;

  return true;
}

// =========================================================
// HEALTH
// =========================================================

router.get(
  "/health",
  (
    _req: Request,
    res: Response
  ) => {
    return res.json({
      ok: true,
      service:
        "ssl-checker",
      roundSize:
        ROUND_SIZE,
      requestSize:
        REQUEST_SIZE,
    });
  }
);

// =========================================================
// LEGACY CHECK
// ใช้กับ Frontend เก่าได้ต่อ
// =========================================================

router.post(
  "/check",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const urls =
        Array.isArray(
          req.body?.urls
        )
          ? req.body.urls
          : [];

      if (
        urls.length === 0
      ) {
        return res
          .status(400)
          .json({
            message:
              "กรุณาส่ง urls อย่างน้อย 1 รายการ",
          });
      }

      if (
        urls.length >
        REQUEST_SIZE
      ) {
        return res
          .status(400)
          .json({
            message:
              `ตรวจสอบได้สูงสุด ${REQUEST_SIZE} URL ต่อครั้ง`,
          });
      }

      const results =
        await mapWithConcurrency(
          urls,
          CONCURRENCY,
          (
            url
          ) =>
            checkCertificate(
              url
            )
        );

      return res.json({
        count:
          results.length,
        results,
      });
    } catch (error) {
      console.error(
        "SSL checker route error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "เกิดข้อผิดพลาดระหว่างตรวจสอบ SSL/TLS",
        });
    }
  }
);

// =========================================================
// CREATE JOB
// =========================================================

router.post(
  "/jobs",
  async (
    req: Request,
    res: Response
  ) => {
    const client =
      await pool.connect();

    try {
      const fileName =
        cleanText(
          req.body?.fileName
        );

      const urls =
        Array.isArray(
          req.body?.urls
        )
          ? req.body.urls
              .map(
                (
                  item:
                    unknown
                ) =>
                  cleanText(
                    item
                  )
              )
              .filter(
                Boolean
              )
          : [];

      const createdByRaw =
        Number(
          req.body?.createdBy
        );

      const createdBy =
        Number.isInteger(
          createdByRaw
        ) &&
        createdByRaw > 0
          ? createdByRaw
          : null;

      if (!fileName) {
        return res
          .status(400)
          .json({
            message:
              "ไม่พบชื่อไฟล์",
          });
      }

      if (
        urls.length === 0
      ) {
        return res
          .status(400)
          .json({
            message:
              "ไม่พบ URL สำหรับสร้างงานตรวจ",
          });
      }

      if (
        urls.length >
        MAX_CREATE_URLS
      ) {
        return res
          .status(400)
          .json({
            message:
              `สร้างงานได้สูงสุด ${MAX_CREATE_URLS.toLocaleString()} URL`,
          });
      }

      const totalRounds =
        Math.ceil(
          urls.length /
            ROUND_SIZE
        );

      await client.query(
        "BEGIN"
      );

      const jobResult =
        await client.query(
          `
            INSERT INTO
              ssl_check_jobs (
                original_file_name,
                total_urls,
                round_size,
                request_size,
                total_rounds,
                created_by,
                status
              )
            VALUES (
              $1, $2, $3, $4,
              $5, $6, 'READY'
            )
            RETURNING *
          `,
          [
            fileName,
            urls.length,
            ROUND_SIZE,
            REQUEST_SIZE,
            totalRounds,
            createdBy,
          ]
        );

      const job =
        jobResult.rows[0];

      /*
        Insert URL ต้นฉบับตามลำดับ
        ใช้ batch SQL ทีละ 1000
        ลดจำนวน query สำหรับไฟล์ใหญ่
      */
      const insertChunkSize =
        1000;

      for (
        let offset = 0;
        offset < urls.length;
        offset +=
          insertChunkSize
      ) {
        const chunk =
          urls.slice(
            offset,
            offset +
              insertChunkSize
          );

        const params:
          unknown[] = [];

        const values =
          chunk.map(
            (
              url:
                string,
              index:
                number
            ) => {
              const fileOrder =
                offset +
                index +
                1;

              const roundNumber =
                Math.ceil(
                  fileOrder /
                    ROUND_SIZE
                );

              const base =
                index * 4;

              params.push(
                job.id,
                fileOrder,
                roundNumber,
                url
              );

              return `(
                $${base + 1},
                $${base + 2},
                $${base + 3},
                $${base + 4}
              )`;
            }
          );

        await client.query(
          `
            INSERT INTO
              ssl_check_items (
                job_id,
                file_order,
                round_number,
                url
              )
            VALUES
              ${values.join(",")}
          `,
          params
        );
      }

      for (
        let roundNumber = 1;
        roundNumber <=
        totalRounds;
        roundNumber += 1
      ) {
        const startOrder =
          (roundNumber - 1) *
            ROUND_SIZE +
          1;

        const endOrder =
          Math.min(
            roundNumber *
              ROUND_SIZE,
            urls.length
          );

        const totalUrls =
          endOrder -
          startOrder +
          1;

        await client.query(
          `
            INSERT INTO
              ssl_check_rounds (
                job_id,
                round_number,
                start_order,
                end_order,
                total_urls,
                status
              )
            VALUES (
              $1, $2, $3, $4,
              $5, 'WAITING'
            )
          `,
          [
            job.id,
            roundNumber,
            startOrder,
            endOrder,
            totalUrls,
          ]
        );
      }

      await client.query(
        "COMMIT"
      );

      return res
        .status(201)
        .json({
          message:
            "สร้างงานตรวจ SSL/TLS สำเร็จ",
          job,
        });
    } catch (error) {
      await client.query(
        "ROLLBACK"
      );

      console.error(
        "Create SSL job error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "สร้างงานตรวจ SSL/TLS ไม่สำเร็จ",
        });
    } finally {
      client.release();
    }
  }
);

// =========================================================
// LIST JOBS
// =========================================================

router.get(
  "/jobs",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const limit =
        Math.min(
          100,
          parsePositiveInt(
            req.query.limit,
            20
          )
        );

      const result =
        await pool.query(
          `
            SELECT
              j.*,
              u.name AS created_by_name,
              u.email AS created_by_email
            FROM ssl_check_jobs j

            LEFT JOIN users u
            ON u.id =
              j.created_by

            ORDER BY
              j.created_at DESC,
              j.id DESC

            LIMIT $1
          `,
          [limit]
        );

      return res.json({
        jobs:
          result.rows,
      });
    } catch (error) {
      console.error(
        "List SSL jobs error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "โหลดรายการงาน SSL/TLS ไม่สำเร็จ",
        });
    }
  }
);

// =========================================================
// GET JOB DETAIL
// =========================================================

router.get(
  "/jobs/:jobId",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const jobId =
        Number(
          req.params.jobId
        );

      if (
        !Number.isInteger(
          jobId
        ) ||
        jobId <= 0
      ) {
        return res
          .status(400)
          .json({
            message:
              "jobId ไม่ถูกต้อง",
          });
      }

      const result =
        await pool.query(
          `
            SELECT
              j.*,

              COALESCE(
                s.valid_count,
                0
              )::int
                AS valid_count,

              COALESCE(
                s.expiring_count,
                0
              )::int
                AS expiring_count,

              COALESCE(
                s.expired_count,
                0
              )::int
                AS expired_count,

              COALESCE(
                s.error_count,
                0
              )::int
                AS error_count

            FROM ssl_check_jobs j

            LEFT JOIN (
              SELECT
                job_id,

                COUNT(*) FILTER (
                  WHERE status =
                    'VALID'
                ) AS valid_count,

                COUNT(*) FILTER (
                  WHERE status =
                    'EXPIRING_SOON'
                ) AS expiring_count,

                COUNT(*) FILTER (
                  WHERE status =
                    'EXPIRED'
                ) AS expired_count,

                COUNT(*) FILTER (
                  WHERE status =
                    'ERROR'
                ) AS error_count

              FROM ssl_check_results
              GROUP BY job_id
            ) s
            ON s.job_id = j.id

            WHERE j.id = $1
            LIMIT 1
          `,
          [jobId]
        );

      if (
        result.rows.length ===
        0
      ) {
        return res
          .status(404)
          .json({
            message:
              "ไม่พบงานตรวจ SSL/TLS",
          });
      }

      return res.json({
        job:
          result.rows[0],
        activeInServer:
          activeJobs.has(
            jobId
          ),
        pauseRequested:
          pauseRequested.has(
            jobId
          ),
      });
    } catch (error) {
      console.error(
        "Get SSL job error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "โหลดข้อมูลงาน SSL/TLS ไม่สำเร็จ",
        });
    }
  }
);

// =========================================================
// GET ROUNDS
// =========================================================

router.get(
  "/jobs/:jobId/rounds",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const jobId =
        Number(
          req.params.jobId
        );

      const result =
        await pool.query(
          `
            SELECT *
            FROM ssl_check_rounds
            WHERE job_id = $1
            ORDER BY
              round_number ASC
          `,
          [jobId]
        );

      return res.json({
        rounds:
          result.rows,
      });
    } catch (error) {
      console.error(
        "Get SSL rounds error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "โหลดรอบการตรวจไม่สำเร็จ",
        });
    }
  }
);

// =========================================================
// GET ROUND RESULTS
// =========================================================

router.get(
  "/jobs/:jobId/rounds/:roundNumber/results",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const jobId =
        Number(
          req.params.jobId
        );

      const roundNumber =
        Number(
          req.params.roundNumber
        );

      const page =
        parsePositiveInt(
          req.query.page,
          1
        );

      const pageSize =
        Math.min(
          100,
          parsePositiveInt(
            req.query.pageSize,
            100
          )
        );

      const offset =
        (page - 1) *
        pageSize;

      const roundResult =
        await pool.query(
          `
            SELECT id
            FROM ssl_check_rounds
            WHERE
              job_id = $1
              AND round_number =
                $2
            LIMIT 1
          `,
          [
            jobId,
            roundNumber,
          ]
        );

      if (
        roundResult.rows.length ===
        0
      ) {
        return res
          .status(404)
          .json({
            message:
              "ไม่พบรอบที่ต้องการ",
          });
      }

      const roundId =
        roundResult.rows[0].id;

      const countResult =
        await pool.query(
          `
            SELECT
              COUNT(*)::int
                AS total
            FROM ssl_check_results
            WHERE round_id = $1
          `,
          [roundId]
        );

      const result =
        await pool.query(
          `
            SELECT
              id,
              file_order,
              url,
              hostname,
              port,
              valid_from,
              expiration_date,
              days_left,
              status,
              issuer,
              subject,
              checked_at,
              error
            FROM ssl_check_results
            WHERE round_id = $1
            ORDER BY
              file_order ASC
            LIMIT $2
            OFFSET $3
          `,
          [
            roundId,
            pageSize,
            offset,
          ]
        );

      return res.json({
        page,
        pageSize,
        total:
          Number(
            countResult.rows[0]
              ?.total ?? 0
          ),
        results:
          result.rows,
      });
    } catch (error) {
      console.error(
        "Get SSL round results error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "โหลดผลตรวจ SSL/TLS ไม่สำเร็จ",
        });
    }
  }
);

// =========================================================
// START / RESUME
// =========================================================

router.post(
  "/jobs/:jobId/start",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const jobId =
        Number(
          req.params.jobId
        );

      if (
        !Number.isInteger(
          jobId
        ) ||
        jobId <= 0
      ) {
        return res
          .status(400)
          .json({
            message:
              "jobId ไม่ถูกต้อง",
          });
      }

      if (
        activeJobs.has(
          jobId
        )
      ) {
        return res
          .status(409)
          .json({
            message:
              "งานนี้กำลังตรวจอยู่แล้ว",
          });
      }

      const mode =
        cleanText(
          req.body?.mode
        ).toLowerCase();

      let roundLimit:
        number | null;

      if (
        mode === "all"
      ) {
        roundLimit =
          null;
      } else {
        roundLimit =
          parsePositiveInt(
            req.body
              ?.roundCount,
            1
          );
      }

      const jobResult =
        await pool.query(
          `
            SELECT
              id,
              status
            FROM ssl_check_jobs
            WHERE id = $1
            LIMIT 1
          `,
          [jobId]
        );

      if (
        jobResult.rows.length ===
        0
      ) {
        return res
          .status(404)
          .json({
            message:
              "ไม่พบงานตรวจ SSL/TLS",
          });
      }

      if (
        jobResult.rows[0]
          .status ===
        "COMPLETED"
      ) {
        return res
          .status(400)
          .json({
            message:
              "งานนี้ตรวจครบทั้งหมดแล้ว",
          });
      }

      const started =
        startBackgroundJob(
          jobId,
          roundLimit
        );

      if (!started) {
        return res
          .status(409)
          .json({
            message:
              "งานนี้กำลังตรวจอยู่แล้ว",
          });
      }

      return res
        .status(202)
        .json({
          message:
            roundLimit === null
              ? "เริ่มตรวจต่อจนจบทั้งหมดแล้ว"
              : `เริ่มตรวจต่อ ${roundLimit} รอบแล้ว`,
          jobId,
          roundLimit,
        });
    } catch (error) {
      console.error(
        "Start SSL job error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "เริ่มงานตรวจ SSL/TLS ไม่สำเร็จ",
        });
    }
  }
);

router.post(
  "/jobs/:jobId/resume",
  async (
    req: Request,
    res: Response
  ) => {
    /*
      ใช้ logic เดียวกับ start
      Frontend จะเรียก route ไหนก็ได้
    */
    const jobId =
      Number(
        req.params.jobId
      );

    if (
      !Number.isInteger(
        jobId
      ) ||
      jobId <= 0
    ) {
      return res
        .status(400)
        .json({
          message:
            "jobId ไม่ถูกต้อง",
        });
    }

    if (
      activeJobs.has(
        jobId
      )
    ) {
      return res
        .status(409)
        .json({
          message:
            "งานนี้กำลังตรวจอยู่แล้ว",
        });
    }

    const mode =
      cleanText(
        req.body?.mode
      ).toLowerCase();

    const roundLimit =
      mode === "all"
        ? null
        : parsePositiveInt(
            req.body
              ?.roundCount,
            1
          );

    const started =
      startBackgroundJob(
        jobId,
        roundLimit
      );

    if (!started) {
      return res
        .status(409)
        .json({
          message:
            "งานนี้กำลังตรวจอยู่แล้ว",
        });
    }

    return res
      .status(202)
      .json({
        message:
          "ตรวจต่อแล้ว",
        jobId,
        roundLimit,
      });
  }
);

// =========================================================
// PAUSE
// =========================================================

router.post(
  "/jobs/:jobId/pause",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const jobId =
        Number(
          req.params.jobId
        );

      if (
        !Number.isInteger(
          jobId
        ) ||
        jobId <= 0
      ) {
        return res
          .status(400)
          .json({
            message:
              "jobId ไม่ถูกต้อง",
          });
      }

      pauseRequested.add(
        jobId
      );

      /*
        ถ้า job ไม่ได้ active ใน process นี้
        สามารถเปลี่ยน PAUSED ได้ทันที
      */
      if (
        !activeJobs.has(
          jobId
        )
      ) {
        await pool.query(
          `
            UPDATE ssl_check_jobs
            SET
              status = 'PAUSED',
              paused_at = NOW(),
              updated_at = NOW()
            WHERE
              id = $1
              AND status <>
                'COMPLETED'
          `,
          [jobId]
        );

        await pool.query(
          `
            UPDATE ssl_check_rounds
            SET
              status = 'PAUSED',
              paused_at = NOW(),
              updated_at = NOW()
            WHERE
              job_id = $1
              AND status =
                'RUNNING'
          `,
          [jobId]
        );
      }

      return res
        .status(202)
        .json({
          message:
            "รับคำสั่งหยุดพักแล้ว ระบบจะหยุดหลังชุด 100 URL ปัจจุบันเสร็จ",
          jobId,
        });
    } catch (error) {
      console.error(
        "Pause SSL job error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "หยุดพักงาน SSL/TLS ไม่สำเร็จ",
        });
    }
  }
);

// =========================================================\n// RESET JOB\n// ล้างผลตรวจทั้งหมดเพื่อเริ่มตรวจใหม่\n// เฉพาะ Super Admin\n// =========================================================\n\nrouter.post(\n  \"/jobs/:jobId/reset\",\n  async (\n    req: Request,\n    res: Response\n  ) => {\n    const client = await pool.connect();\n\n    try {\n      const adminId = await requireSuperAdmin(req, res);\n      if (!adminId) return;\n\n      const jobId = Number(req.params.jobId);\n      if (!Number.isInteger(jobId) || jobId <= 0) {\n        return res.status(400).json({ message: \"jobId ไม่ถูกต้อง\" });\n      }\n\n      if (activeJobs.has(jobId)) {\n        return res.status(409).json({ message: \"กรุณาหยุดพักงานก่อนล้างผลตรวจ\" });\n      }\n\n      const jobResult = await client.query(\n        `SELECT id, original_file_name, status FROM ssl_check_jobs WHERE id = $1 LIMIT 1`,\n        [jobId]\n      );\n\n      if (jobResult.rows.length === 0) {\n        return res.status(404).json({ message: \"ไม่พบงานตรวจ SSL/TLS\" });\n      }\n\n      if (jobResult.rows[0].status === \"RUNNING\") {\n        return res.status(409).json({ message: \"กรุณาหยุดพักงานก่อนล้างผลตรวจ\" });\n      }\n\n      await client.query(\"BEGIN\");\n      await client.query(`DELETE FROM ssl_check_results WHERE job_id = $1`, [jobId]);\n      await client.query(\n        `UPDATE ssl_check_rounds SET checked_count = 0, status = 'WAITING', started_at = NULL, paused_at = NULL, completed_at = NULL, updated_at = NOW() WHERE job_id = $1`,\n        [jobId]\n      );\n      await client.query(\n        `UPDATE ssl_check_jobs SET checked_count = 0, completed_rounds = 0, current_round = 0, status = 'READY', last_error = NULL, started_at = NULL, paused_at = NULL, completed_at = NULL, updated_at = NOW() WHERE id = $1`,\n        [jobId]\n      );\n      await client.query(\n        `INSERT INTO admin_logs (admin_id, action_type, message, ip_address) VALUES ($1, $2, $3, $4)`,\n        [adminId, \"ssl_job_reset\", `ล้างผลตรวจ SSL/TLS job #${jobId}`, req.ip ?? null]\n      );\n      await client.query(\"COMMIT\");\n      return res.json({ message: \"ล้างผลตรวจทั้งหมดแล้ว สามารถเริ่มตรวจใหม่ได้\", jobId });\n    } catch (error) {\n      await client.query(\"ROLLBACK\");\n      console.error(\"Reset SSL job error:\", error);\n      return res.status(500).json({ message: \"ล้างผลตรวจ SSL/TLS ไม่สำเร็จ\" });\n    } finally {\n      client.release();\n    }\n  }\n);\n\n// =========================================================
// PATCH RESULT
// สำหรับหน้าจอ Super Admin ในขั้นถัดไป
// ตอนนี้ยังไม่ใส่ auth middleware
// =========================================================

router.patch(
  "/results/:resultId",
  async (
    req: Request,
    res: Response
  ) => {
    try {
      const adminId =
        await requireSuperAdmin(
          req,
          res
        );

      if (!adminId) {
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

      const {
        status,
        validFrom,
        expirationDate,
        daysLeft,
        issuer,
        subject,
        error,
      } = req.body ?? {};

      const allowedStatuses:
        SSLStatus[] =
        [
          "VALID",
          "EXPIRING_SOON",
          "EXPIRED",
          "ERROR",
        ];

      if (
        status !== undefined &&
        !allowedStatuses.includes(
          status
        )
      ) {
        return res
          .status(400)
          .json({
            message:
              "status ไม่ถูกต้อง",
          });
      }

      const updated =
        await pool.query(
          `
            UPDATE ssl_check_results
            SET
              status =
                COALESCE(
                  $2,
                  status
                ),
              valid_from =
                COALESCE(
                  $3,
                  valid_from
                ),
              expiration_date =
                COALESCE(
                  $4,
                  expiration_date
                ),
              days_left =
                COALESCE(
                  $5,
                  days_left
                ),
              issuer =
                COALESCE(
                  $6,
                  issuer
                ),
              subject =
                COALESCE(
                  $7,
                  subject
                ),
              error =
                CASE
                  WHEN $8::boolean
                  THEN $9
                  ELSE error
                END,
              updated_at =
                NOW()
            WHERE id = $1
            RETURNING *
          `,
          [
            resultId,
            status ?? null,
            validFrom ?? null,
            expirationDate ??
              null,
            daysLeft ?? null,
            issuer ?? null,
            subject ?? null,
            Object.prototype
              .hasOwnProperty.call(
                req.body ?? {},
                "error"
              ),
            error ?? null,
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
              "ไม่พบผลตรวจที่ต้องการแก้ไข",
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
          VALUES ($1, $2, $3, $4)
        `,
        [
          adminId,
          "ssl_result_update",
          `แก้ไขผล SSL/TLS result #${resultId}`,
          req.ip ?? null,
        ]
      );

      return res.json({
        message:
          "แก้ไขผลตรวจสำเร็จ",
        result:
          updated.rows[0],
      });
    } catch (error) {
      console.error(
        "Patch SSL result error:",
        error
      );

      return res
        .status(500)
        .json({
          message:
            "แก้ไขผลตรวจ SSL/TLS ไม่สำเร็จ",
        });
    }
  }
);

export default router;

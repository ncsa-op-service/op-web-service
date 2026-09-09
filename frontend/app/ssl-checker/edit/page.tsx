"use client";

import {
  useEffect,
  useMemo,
  useState,
} from "react";
import {
  useRouter,
  useSearchParams,
} from "next/navigation";
import "../ssl-checker.css";
import "./edit.css";

type SSLStatus =
  | "VALID"
  | "EXPIRING_SOON"
  | "EXPIRED"
  | "ERROR";

type JobStatus =
  | "READY"
  | "RUNNING"
  | "PAUSED"
  | "COMPLETED"
  | "ERROR";

type UserInfo = {
  id: number;
  name: string;
  email: string;
  role:
    | "super_admin"
    | "editor"
    | "viewer";
};

type SSLJob = {
  id: number;
  original_file_name: string;
  total_urls: number;
  total_rounds: number;
  checked_count: number;
  completed_rounds: number;
  current_round: number;
  status: JobStatus;
};

type SSLRound = {
  id: number;
  job_id: number;
  round_number: number;
  start_order: number;
  end_order: number;
  total_urls: number;
  checked_count: number;
  status:
    | "WAITING"
    | "RUNNING"
    | "PAUSED"
    | "COMPLETED"
    | "ERROR";
};

type SSLResultRow = {
  id: number;
  file_order: number;
  url: string;
  hostname: string;
  port: number;
  valid_from: string | null;
  expiration_date: string | null;
  days_left: number | null;
  status: SSLStatus;
  issuer: string;
  subject: string;
  checked_at: string;
  error: string | null;
};

const API_URL =
  process.env.NEXT_PUBLIC_API_URL &&
  !process.env.NEXT_PUBLIC_API_URL.includes("backend:")
    ? process.env.NEXT_PUBLIC_API_URL.replace(/\/$/, "")
    : "http://localhost:4000";

const PAGE_SIZE = 100;

type SSLColumnKey =
  | "order"
  | "url"
  | "status"
  | "validFrom"
  | "expiration"
  | "daysLeft"
  | "issuer"
  | "error"
  | "checkedAt";

const SSL_COLUMNS: Array<{
  key: SSLColumnKey;
  label: string;
}> = [
  { key: "order", label: "ลำดับ" },
  { key: "url", label: "URL / HOSTNAME" },
  { key: "status", label: "STATUS" },
  { key: "validFrom", label: "VALID FROM" },
  { key: "expiration", label: "EXPIRATION" },
  { key: "daysLeft", label: "DAYS LEFT" },
  { key: "issuer", label: "ISSUER" },
  { key: "error", label: "ERROR" },
  { key: "checkedAt", label: "CHECKED AT" },
];

const DEFAULT_VISIBLE_COLUMNS: SSLColumnKey[] =
  SSL_COLUMNS.map((column) => column.key);

function readCurrentUser():
  UserInfo | null {
  if (
    typeof window ===
    "undefined"
  ) {
    return null;
  }

  const keys = [
    "user",
    "currentUser",
    "authUser",
    "op_user",
  ];

  for (
    const key of keys
  ) {
    const raw =
      window.localStorage.getItem(
        key
      );

    if (!raw) {
      continue;
    }

    try {
      const parsed =
        JSON.parse(raw);

      if (
        parsed &&
        typeof parsed.id ===
          "number" &&
        typeof parsed.role ===
          "string"
      ) {
        return parsed as
          UserInfo;
      }
    } catch {
      // ignore
    }
  }

  return null;
}

async function readJson<T>(
  response: Response
): Promise<T> {
  const text =
    await response.text();

  let data:
    T & {
      message?: string;
    };

  try {
    data =
      (
        text
          ? JSON.parse(text)
          : {}
      ) as T & {
        message?: string;
      };
  } catch {
    throw new Error(
      "Backend ส่งข้อมูลที่อ่านไม่ได้"
    );
  }

  if (!response.ok) {
    throw new Error(
      data.message ??
        `Request ไม่สำเร็จ (HTTP ${response.status})`
    );
  }

  return data;
}

function toDateTimeLocal(
  value: string | null
) {
  if (!value) {
    return "";
  }

  const date =
    new Date(value);

  if (
    Number.isNaN(
      date.getTime()
    )
  ) {
    return "";
  }

  const pad = (
    number: number
  ) =>
    String(number).padStart(
      2,
      "0"
    );

  return `${date.getFullYear()}-${pad(
    date.getMonth() + 1
  )}-${pad(
    date.getDate()
  )}T${pad(
    date.getHours()
  )}:${pad(
    date.getMinutes()
  )}`;
}

function toIsoOrNull(
  value: string
) {
  if (!value) {
    return null;
  }

  const date =
    new Date(value);

  return Number.isNaN(
    date.getTime()
  )
    ? null
    : date.toISOString();
}

function formatDateTime(
  value: string | null
) {
  if (!value) {
    return "-";
  }

  const date =
    new Date(value);

  if (
    Number.isNaN(
      date.getTime()
    )
  ) {
    return value;
  }

  return new Intl.DateTimeFormat(
    "th-TH",
    {
      dateStyle:
        "medium",
      timeStyle:
        "short",
    }
  ).format(date);
}

export default function SSLEditPage() {
  const router =
    useRouter();

  const searchParams =
    useSearchParams();

  const jobId =
    Number(
      searchParams.get(
        "jobId"
      )
    );

  const [
    currentUser,
    setCurrentUser,
  ] =
    useState<UserInfo | null>(
      null
    );

  const [
    job,
    setJob,
  ] =
    useState<SSLJob | null>(
      null
    );

  const [
    rounds,
    setRounds,
  ] =
    useState<SSLRound[]>(
      []
    );

  const [
    selectedRound,
    setSelectedRound,
  ] =
    useState(1);

  const [
    rows,
    setRows,
  ] =
    useState<SSLResultRow[]>(
      []
    );

  const [
    originalRows,
    setOriginalRows,
  ] =
    useState<SSLResultRow[]>(
      []
    );

  const [
    visibleColumns,
    setVisibleColumns,
  ] = useState<SSLColumnKey[]>(
    DEFAULT_VISIBLE_COLUMNS
  );

  const [
    columnMenuOpen,
    setColumnMenuOpen,
  ] = useState(false);


  const [
    loading,
    setLoading,
  ] =
    useState(true);

  const [
    saving,
    setSaving,
  ] =
    useState(false);

  const [
    message,
    setMessage,
  ] =
    useState("");

  const [
    error,
    setError,
  ] =
    useState("");

  const isSuperAdmin =
    currentUser?.role ===
    "super_admin";

  const editableRounds =
    useMemo(
      () =>
        rounds.filter(
          (
            round
          ) =>
            round.checked_count >
            0
        ),
      [
        rounds,
      ]
    );

  const visibleColumnCount =
    visibleColumns.length;

  function isColumnVisible(
    key: SSLColumnKey
  ) {
    return visibleColumns.includes(
      key
    );
  }

  function toggleColumn(
    key: SSLColumnKey
  ) {
    setVisibleColumns(
      (current) => {
        if (current.includes(key)) {
          if (current.length === 1) {
            return current;
          }

          return current.filter(
            (item) => item !== key
          );
        }

        return SSL_COLUMNS
          .map((column) => column.key)
          .filter(
            (columnKey) =>
              current.includes(
                columnKey
              ) ||
              columnKey === key
          );
      }
    );
  }

  function selectAllColumns() {
    setVisibleColumns(
      DEFAULT_VISIBLE_COLUMNS
    );
  }

  const changedRows =
    useMemo(() => {
      const originalMap =
        new Map(
          originalRows.map(
            (
              item
            ) => [
              item.id,
              JSON.stringify(
                item
              ),
            ]
          )
        );

      return rows.filter(
        (
          item
        ) =>
          originalMap.get(
            item.id
          ) !==
          JSON.stringify(
            item
          )
      );
    }, [
      rows,
      originalRows,
    ]);

  async function loadHeader() {
    if (
      !Number.isInteger(
        jobId
      ) ||
      jobId <= 0
    ) {
      throw new Error(
        "ไม่พบ Job ที่ต้องการแก้ไข"
      );
    }

    const [
      jobResponse,
      roundsResponse,
    ] =
      await Promise.all([
        fetch(
          `${API_URL}/api/ssl-checker/jobs/${jobId}`,
          {
            cache:
              "no-store",
          }
        ),
        fetch(
          `${API_URL}/api/ssl-checker/jobs/${jobId}/rounds`,
          {
            cache:
              "no-store",
          }
        ),
      ]);

    const jobData =
      await readJson<{
        job: SSLJob;
      }>(
        jobResponse
      );

    const roundsData =
      await readJson<{
        rounds: SSLRound[];
      }>(
        roundsResponse
      );

    setJob(
      jobData.job
    );

    setRounds(
      roundsData.rounds
    );

    const firstEditable =
      roundsData.rounds.find(
        (
          round
        ) =>
          round.checked_count >
          0
      );

    if (
      firstEditable
    ) {
      setSelectedRound(
        firstEditable
          .round_number
      );
    }
  }

  async function loadWholeRound(
    roundNumber: number
  ) {
    setLoading(true);
    setError("");
    setMessage("");

    try {
      const firstResponse =
        await fetch(
          `${API_URL}/api/ssl-checker/jobs/${jobId}/rounds/${roundNumber}/results?page=1&pageSize=${PAGE_SIZE}`,
          {
            cache:
              "no-store",
          }
        );

      const first =
        await readJson<{
          total: number;
          results:
            SSLResultRow[];
        }>(
          firstResponse
        );

      const allRows =
        [
          ...first.results,
        ];

      const totalPages =
        Math.max(
          1,
          Math.ceil(
            first.total /
              PAGE_SIZE
          )
        );

      for (
        let page = 2;
        page <=
        totalPages;
        page += 1
      ) {
        const response =
          await fetch(
            `${API_URL}/api/ssl-checker/jobs/${jobId}/rounds/${roundNumber}/results?page=${page}&pageSize=${PAGE_SIZE}`,
            {
              cache:
                "no-store",
            }
          );

        const data =
          await readJson<{
            results:
              SSLResultRow[];
          }>(
            response
          );

        allRows.push(
          ...data.results
        );
      }

      setRows(
        allRows
      );

      setOriginalRows(
        structuredClone(
          allRows
        )
      );
    } catch (
      loadError
    ) {
      setRows([]);
      setOriginalRows([]);

      setError(
        loadError instanceof
          Error
          ? loadError.message
          : "โหลดผลตรวจไม่สำเร็จ"
      );
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    const user =
      readCurrentUser();

    setCurrentUser(
      user
    );

    loadHeader()
      .catch(
        (
          loadError
        ) => {
          setError(
            loadError instanceof
              Error
              ? loadError.message
              : "โหลดงานไม่สำเร็จ"
          );

          setLoading(
            false
          );
        }
      );
  }, []);

  useEffect(() => {
    if (
      selectedRound > 0 &&
      rounds.length > 0
    ) {
      void loadWholeRound(
        selectedRound
      );
    }
  }, [
    selectedRound,
    rounds.length,
  ]);

  function updateRow(
    id: number,
    patch:
      Partial<SSLResultRow>
  ) {
    setRows(
      (
        current
      ) =>
        current.map(
          (
            row
          ) =>
            row.id === id
              ? {
                  ...row,
                  ...patch,
                }
              : row
        )
    );
  }

  async function saveChanges() {
    if (
      !currentUser ||
      !isSuperAdmin
    ) {
      setError(
        "เฉพาะ Super Admin เท่านั้นที่แก้ไขผลได้"
      );
      return;
    }

    if (
      job?.status ===
      "RUNNING"
    ) {
      setError(
        "กรุณาหยุดพักงานก่อนแก้ไขผล"
      );
      return;
    }

    if (
      changedRows.length ===
      0
    ) {
      setMessage(
        "ไม่มีข้อมูลที่เปลี่ยนแปลง"
      );
      return;
    }

    setSaving(true);
    setError("");
    setMessage("");

    try {
      /*
        ส่งทีละกลุ่มเล็ก ๆ เพื่อไม่ยิงพร้อมกันมากเกินไป
      */
      const chunkSize =
        10;

      for (
        let start = 0;
        start <
        changedRows.length;
        start +=
          chunkSize
      ) {
        const chunk =
          changedRows.slice(
            start,
            start +
              chunkSize
          );

        await Promise.all(
          chunk.map(
            async (
              row
            ) => {
              const response =
                await fetch(
                  `${API_URL}/api/ssl-checker/results/${row.id}`,
                  {
                    method:
                      "PATCH",
                    headers: {
                      "Content-Type":
                        "application/json",
                      "x-user-id":
                        String(
                          currentUser.id
                        ),
                    },
                    body:
                      JSON.stringify({
                        status:
                          row.status,
                        validFrom:
                          row.valid_from,
                        expirationDate:
                          row.expiration_date,
                        daysLeft:
                          row.days_left,
                        issuer:
                          row.issuer,
                        subject:
                          row.subject,
                        error:
                          row.error,
                      }),
                  }
                );

              await readJson(
                response
              );
            }
          )
        );
      }

      setOriginalRows(
        structuredClone(
          rows
        )
      );

      setMessage(
        `บันทึกสำเร็จ ${changedRows.length} รายการ`
      );
    } catch (
      saveError
    ) {
      setError(
        saveError instanceof
          Error
          ? saveError.message
          : "บันทึกข้อมูลไม่สำเร็จ"
      );
    } finally {
      setSaving(false);
    }
  }

  if (
    currentUser &&
    !isSuperAdmin
  ) {
    return (
      <main className="ssl-page">
        <div className="ssl-edit-shell">
          <section className="ssl-edit-denied">
            <h1>
              ไม่มีสิทธิ์แก้ไข
            </h1>

            <p>
              หน้านี้อนุญาตเฉพาะ Super Admin
            </p>

            <button
              type="button"
              onClick={() =>
                router.push(
                  "/ssl-checker"
                )
              }
            >
              กลับหน้า SSL/TLS Checker
            </button>
          </section>
        </div>
      </main>
    );
  }

  return (
    <main className="ssl-page">
      <div className="ssl-edit-shell">
        <section className="ssl-edit-topbar">
          <div>
            <button
              type="button"
              className="ssl-edit-back"
              onClick={() =>
                router.push(
                  "/ssl-checker"
                )
              }
            >
              ← กลับ
            </button>

            <span className="ssl-kicker">
              SUPER ADMIN / SSL TLS
            </span>

            <h1>
              แก้ไขผลการตรวจสอบ
            </h1>

            <p>
              เลือกรอบที่ต้องการแก้ไข แก้ข้อมูล แล้วกดบันทึก จากนั้นกลับไปกดตรวจต่อได้
            </p>
          </div>

          <div className="ssl-edit-job-meta">
            <span>
              Job #{job?.id ?? "-"}
            </span>

            <strong>
              {job?.original_file_name ??
                "-"}
            </strong>

            <small>
              สถานะ:{" "}
              {job?.status ??
                "-"}
            </small>
          </div>
        </section>

        {job?.status ===
          "RUNNING" && (
          <div className="ssl-edit-warning">
            งานนี้กำลังตรวจอยู่ กรุณากลับไปกด “หยุดพัก” ก่อน แล้วจึงกลับมาแก้ไข
          </div>
        )}

        <section className="ssl-edit-control-card">
          <div>
            <label htmlFor="round">
              รอบที่ต้องการแก้ไข
            </label>

            <select
              id="round"
              value={
                selectedRound
              }
              disabled={
                loading ||
                job?.status ===
                  "RUNNING"
              }
              onChange={(
                event
              ) =>
                setSelectedRound(
                  Number(
                    event
                      .target
                      .value
                  )
                )
              }
            >
              {editableRounds.map(
                (
                  round
                ) => (
                  <option
                    key={
                      round.id
                    }
                    value={
                      round.round_number
                    }
                  >
                    รอบที่{" "}
                    {round.round_number} • URL{" "}
                    {round.start_order.toLocaleString()}-
                    {round.end_order.toLocaleString()} • ตรวจแล้ว{" "}
                    {round.checked_count.toLocaleString()}/
                    {round.total_urls.toLocaleString()}
                  </option>
                )
              )}
            </select>
          </div>

          <div className="ssl-edit-control-summary">
            <span>
              รายการในรอบ
            </span>

            <strong>
              {rows.length.toLocaleString()}
            </strong>
          </div>

          <div className="ssl-edit-control-summary">
            <span>
              แก้ไขแล้ว
            </span>

            <strong>
              {changedRows.length.toLocaleString()}
            </strong>
          </div>

          <button
            type="button"
            className="ssl-edit-save"
            disabled={
              saving ||
              loading ||
              changedRows.length ===
                0 ||
              job?.status ===
                "RUNNING"
            }
            onClick={
              saveChanges
            }
          >
            {saving
              ? "กำลังบันทึก..."
              : "✓ บันทึกการแก้ไข"}
          </button>
        </section>

        {message && (
          <div className="ssl-edit-success">
            {message}
          </div>
        )}

        {error && (
          <div className="ssl-error-box">
            {error}
          </div>
        )}

        <section className="ssl-edit-table-card">
          <div className="ssl-edit-table-title">
            <div>
              <h2>
                รอบที่{" "}
                {selectedRound}
              </h2>

              <p>
                URL และลำดับไฟล์เป็นข้อมูลอ้างอิง ไม่แก้ไข เพื่อไม่ให้ลำดับงานตรวจเสีย
              </p>
            </div>

            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: "8px",
                flexWrap: "wrap",
                justifyContent: "flex-end",
              }}
            >
              <div
                style={{
                  position: "relative",
                }}
              >
                <button
                  type="button"
                  className="ssl-edit-return"
                  onClick={() =>
                    setColumnMenuOpen(
                      (current) =>
                        !current
                    )
                  }
                  aria-expanded={
                    columnMenuOpen
                  }
                >
                  ☷ เลือกคอลัมน์ (
                  {visibleColumnCount}/
                  {SSL_COLUMNS.length})
                </button>

                {columnMenuOpen && (
                  <div
                    style={{
                      position:
                        "absolute",
                      zIndex: 50,
                      top:
                        "calc(100% + 8px)",
                      right: 0,
                      width: "280px",
                      maxWidth:
                        "calc(100vw - 40px)",
                      padding: "10px",
                      border:
                        "1px solid #dce5f0",
                      borderRadius:
                        "10px",
                      background:
                        "#ffffff",
                      boxShadow:
                        "0 14px 34px rgba(18, 44, 85, 0.16)",
                    }}
                  >
                    <div
                      style={{
                        display:
                          "flex",
                        alignItems:
                          "center",
                        justifyContent:
                          "space-between",
                        gap: "8px",
                        padding:
                          "3px 4px 9px",
                        borderBottom:
                          "1px solid #edf1f5",
                      }}
                    >
                      <strong
                        style={{
                          color:
                            "#173151",
                          fontSize:
                            "calc(10px * var(--app-font-scale, 1))",
                        }}
                      >
                        คอลัมน์ที่แสดง
                      </strong>

                      <button
                        type="button"
                        onClick={
                          selectAllColumns
                        }
                        style={{
                          border: 0,
                          background:
                            "transparent",
                          color:
                            "#1768e6",
                          fontWeight: 800,
                          cursor:
                            "pointer",
                          fontSize:
                            "calc(9px * var(--app-font-scale, 1))",
                        }}
                      >
                        เลือกทั้งหมด
                      </button>
                    </div>

                    <div
                      style={{
                        display:
                          "grid",
                        gap: "4px",
                        paddingTop:
                          "7px",
                      }}
                    >
                      {SSL_COLUMNS.map(
                        (column) => {
                          const checked =
                            isColumnVisible(
                              column.key
                            );

                          const isLastVisible =
                            checked &&
                            visibleColumnCount ===
                              1;

                          return (
                            <label
                              key={
                                column.key
                              }
                              style={{
                                display:
                                  "flex",
                                alignItems:
                                  "center",
                                gap: "9px",
                                minHeight:
                                  "34px",
                                padding:
                                  "6px 8px",
                                borderRadius:
                                  "7px",
                                background:
                                  checked
                                    ? "#f4f8ff"
                                    : "#ffffff",
                                color:
                                  "#344a66",
                                cursor:
                                  isLastVisible
                                    ? "not-allowed"
                                    : "pointer",
                                fontSize:
                                  "calc(9px * var(--app-font-scale, 1))",
                              }}
                            >
                              <input
                                type="checkbox"
                                checked={
                                  checked
                                }
                                disabled={
                                  isLastVisible
                                }
                                onChange={() =>
                                  toggleColumn(
                                    column.key
                                  )
                                }
                                style={{
                                  accentColor:
                                    "#1768e6",
                                }}
                              />
                              <span>
                                {
                                  column.label
                                }
                              </span>
                            </label>
                          );
                        }
                      )}
                    </div>
                  </div>
                )}
              </div>

              <button
                type="button"
                className="ssl-edit-return"
                onClick={() =>
                  router.push(
                    "/ssl-checker"
                  )
                }
              >
                กลับไปตรวจต่อ →
              </button>
            </div>
          </div>

          <div className="ssl-table-wrapper">
            <table
              className="ssl-table ssl-edit-table"
              style={{
                minWidth: `${Math.max(
                  720,
                  visibleColumnCount *
                    170
                )}px`,
              }}
            >
              <thead>
                <tr>
                  {isColumnVisible(
                    "order"
                  ) && <th>ลำดับ</th>}

                  {isColumnVisible(
                    "url"
                  ) && (
                    <th>
                      URL / HOSTNAME
                    </th>
                  )}

                  {isColumnVisible(
                    "status"
                  ) && <th>STATUS</th>}

                  {isColumnVisible(
                    "validFrom"
                  ) && (
                    <th>VALID FROM</th>
                  )}

                  {isColumnVisible(
                    "expiration"
                  ) && (
                    <th>EXPIRATION</th>
                  )}

                  {isColumnVisible(
                    "daysLeft"
                  ) && (
                    <th>DAYS LEFT</th>
                  )}

                  {isColumnVisible(
                    "issuer"
                  ) && <th>ISSUER</th>}

                  {isColumnVisible(
                    "error"
                  ) && <th>ERROR</th>}

                  {isColumnVisible(
                    "checkedAt"
                  ) && (
                    <th>CHECKED AT</th>
                  )}
                </tr>
              </thead>

              <tbody>
                {loading ? (
                  <tr>
                    <td
                      colSpan={
                        visibleColumnCount
                      }
                      className="ssl-empty-cell"
                    >
                      กำลังโหลดข้อมูล...
                    </td>
                  </tr>
                ) : rows.length ===
                  0 ? (
                  <tr>
                    <td
                      colSpan={
                        visibleColumnCount
                      }
                      className="ssl-empty-cell"
                    >
                      รอบนี้ยังไม่มีผลให้แก้ไข
                    </td>
                  </tr>
                ) : (
                  rows.map(
                    (row) => (
                      <tr
                        key={row.id}
                      >
                        {isColumnVisible(
                          "order"
                        ) && (
                          <td className="ssl-order-cell">
                            {
                              row.file_order
                            }
                          </td>
                        )}

                        {isColumnVisible(
                          "url"
                        ) && (
                          <td className="ssl-url-cell">
                            <strong>
                              {row.url}
                            </strong>

                            <span>
                              {
                                row.hostname
                              }
                            </span>
                          </td>
                        )}

                        {isColumnVisible(
                          "status"
                        ) && (
                          <td>
                            <select
                              className="ssl-edit-input ssl-edit-status"
                              value={
                                row.status
                              }
                              disabled={
                                job?.status ===
                                "RUNNING"
                              }
                              onChange={(
                                event
                              ) =>
                                updateRow(
                                  row.id,
                                  {
                                    status:
                                      event
                                        .target
                                        .value as SSLStatus,
                                  }
                                )
                              }
                            >
                              <option value="VALID">
                                VALID
                              </option>

                              <option value="EXPIRING_SOON">
                                EXPIRING SOON
                              </option>

                              <option value="EXPIRED">
                                EXPIRED
                              </option>

                              <option value="ERROR">
                                ERROR
                              </option>
                            </select>
                          </td>
                        )}

                        {isColumnVisible(
                          "validFrom"
                        ) && (
                          <td>
                            <input
                              className="ssl-edit-input ssl-edit-date"
                              type="datetime-local"
                              value={toDateTimeLocal(
                                row.valid_from
                              )}
                              disabled={
                                job?.status ===
                                "RUNNING"
                              }
                              onChange={(
                                event
                              ) =>
                                updateRow(
                                  row.id,
                                  {
                                    valid_from:
                                      toIsoOrNull(
                                        event
                                          .target
                                          .value
                                      ),
                                  }
                                )
                              }
                            />
                          </td>
                        )}

                        {isColumnVisible(
                          "expiration"
                        ) && (
                          <td>
                            <input
                              className="ssl-edit-input ssl-edit-date"
                              type="datetime-local"
                              value={toDateTimeLocal(
                                row.expiration_date
                              )}
                              disabled={
                                job?.status ===
                                "RUNNING"
                              }
                              onChange={(
                                event
                              ) =>
                                updateRow(
                                  row.id,
                                  {
                                    expiration_date:
                                      toIsoOrNull(
                                        event
                                          .target
                                          .value
                                      ),
                                  }
                                )
                              }
                            />
                          </td>
                        )}

                        {isColumnVisible(
                          "daysLeft"
                        ) && (
                          <td>
                            <input
                              className="ssl-edit-input ssl-edit-days"
                              type="number"
                              value={
                                row.days_left ??
                                ""
                              }
                              disabled={
                                job?.status ===
                                "RUNNING"
                              }
                              onChange={(
                                event
                              ) =>
                                updateRow(
                                  row.id,
                                  {
                                    days_left:
                                      event
                                        .target
                                        .value ===
                                      ""
                                        ? null
                                        : Number(
                                            event
                                              .target
                                              .value
                                          ),
                                  }
                                )
                              }
                            />
                          </td>
                        )}

                        {isColumnVisible(
                          "issuer"
                        ) && (
                          <td>
                            <input
                              className="ssl-edit-input ssl-edit-text"
                              type="text"
                              value={
                                row.issuer ??
                                ""
                              }
                              disabled={
                                job?.status ===
                                "RUNNING"
                              }
                              onChange={(
                                event
                              ) =>
                                updateRow(
                                  row.id,
                                  {
                                    issuer:
                                      event
                                        .target
                                        .value,
                                  }
                                )
                              }
                            />
                          </td>
                        )}

                        {isColumnVisible(
                          "error"
                        ) && (
                          <td>
                            <input
                              className="ssl-edit-input ssl-edit-text"
                              type="text"
                              value={
                                row.error ??
                                ""
                              }
                              disabled={
                                job?.status ===
                                "RUNNING"
                              }
                              onChange={(
                                event
                              ) =>
                                updateRow(
                                  row.id,
                                  {
                                    error:
                                      event
                                        .target
                                        .value ||
                                      null,
                                  }
                                )
                              }
                            />
                          </td>
                        )}

                        {isColumnVisible(
                          "checkedAt"
                        ) && (
                          <td>
                            {formatDateTime(
                              row.checked_at
                            )}
                          </td>
                        )}
                      </tr>
                    )
                  )
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}

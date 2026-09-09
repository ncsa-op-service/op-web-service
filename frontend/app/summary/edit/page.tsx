"use client";

import {
  useCallback,
  useEffect,
  useMemo,
  useState,
} from "react";
import {
  useRouter,
} from "next/navigation";
import * as XLSX from "xlsx";

import "../summary.css";
import "./edit.css";

type UserInfo = {
  id: number;
  name: string;
  email: string;
  role:
    | "super_admin"
    | "editor"
    | "viewer";
};

type Batch = {
  id: number;
  original_file_name:
    | string
    | null;
};

type ResultRow = {
  id: number;
  batch_id: number;
  case_id: string;
  url_sms: string;
  ais_result: string | null;
  true_dtac_result:
    | string
    | null;
  nt_result: string | null;
  cloudflare_result:
    | string
    | null;
  created_at: string;
};

const API_URL =
  process.env.NEXT_PUBLIC_API_URL &&
  !process.env.NEXT_PUBLIC_API_URL.includes(
    "backend:"
  )
    ? process.env.NEXT_PUBLIC_API_URL.replace(
        /\/$/,
        ""
      )
    : "http://localhost:4000";

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

  for (const key of keys) {
    const raw =
      window.localStorage.getItem(
        key
      );

    if (!raw) continue;

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

  let data: unknown = {};

  if (text) {
    try {
      data =
        JSON.parse(text);
    } catch {
      data = {
        error: text,
      };
    }
  }

  if (!response.ok) {
    const message =
      typeof data ===
        "object" &&
      data !== null &&
      "error" in data
        ? String(
            (
              data as {
                error?: unknown;
              }
            ).error
          )
        : "เกิดข้อผิดพลาด";

    throw new Error(message);
  }

  return data as T;
}

function isUrl(
  value:
    | string
    | null
    | undefined
) {
  const text =
    String(value ?? "")
      .trim();

  return (
    text.startsWith(
      "http://"
    ) ||
    text.startsWith(
      "https://"
    )
  );
}

function statusKind(
  value:
    | string
    | null
    | undefined
) {
  const text =
    String(value ?? "")
      .trim();

  if (!text) {
    return "pending";
  }

  if (isUrl(text)) {
    return "reachable";
  }

  return "blocked";
}

export default function SummaryEditPage() {
  const router =
    useRouter();

  const [
    currentUser,
    setCurrentUser,
  ] =
    useState<UserInfo | null>(
      null
    );

  const [
    batch,
    setBatch,
  ] =
    useState<Batch | null>(
      null
    );

  const [
    rows,
    setRows,
  ] =
    useState<ResultRow[]>(
      []
    );

  const [
    originalRows,
    setOriginalRows,
  ] =
    useState<ResultRow[]>(
      []
    );

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
    editMode,
    setEditMode,
  ] =
    useState(false);

  const [
    search,
    setSearch,
  ] =
    useState("");

  const [
    error,
    setError,
  ] =
    useState("");

  const canEdit =
    currentUser?.role ===
      "super_admin" ||
    currentUser?.role ===
      "editor";

  const loadData =
    useCallback(
      async () => {
        setLoading(true);
        setError("");

        try {
          const response =
            await fetch(
              `${API_URL}/api/summary/edit/fake_domain`,
              {
                cache:
                  "no-store",
              }
            );

          const data =
            await readJson<{
              batch:
                Batch | null;
              results:
                ResultRow[];
            }>(response);

          setBatch(data.batch);
          setRows(data.results);
          setOriginalRows(
            data.results.map(
              (row) => ({
                ...row,
              })
            )
          );
        } catch (
          loadError
        ) {
          setBatch(null);
          setRows([]);
          setOriginalRows([]);

          setError(
            loadError instanceof
              Error
              ? loadError.message
              : "โหลดข้อมูลไม่สำเร็จ"
          );
        } finally {
          setLoading(false);
        }
      },
      []
    );

  useEffect(() => {
    setCurrentUser(
      readCurrentUser()
    );

    void loadData();
  }, [loadData]);

  const filteredRows =
    useMemo(() => {
      const keyword =
        search
          .trim()
          .toLowerCase();

      if (!keyword) {
        return rows;
      }

      return rows.filter(
        (row) =>
          row.case_id
            .toLowerCase()
            .includes(
              keyword
            ) ||
          row.url_sms
            .toLowerCase()
            .includes(
              keyword
            )
      );
    }, [rows, search]);

  function updateRow(
    id: number,
    field:
      | "case_id"
      | "url_sms"
      | "ais_result"
      | "true_dtac_result"
      | "nt_result"
      | "cloudflare_result",
    value: string
  ) {
    if (!editMode) {
      return;
    }

    setRows((current) =>
      current.map((row) =>
        row.id === id
          ? {
              ...row,
              [field]:
                field ===
                  "case_id" ||
                field ===
                  "url_sms"
                  ? value
                  : value || null,
            }
          : row
      )
    );
  }

  function cancelEdit() {
    setRows(
      originalRows.map(
        (row) => ({
          ...row,
        })
      )
    );

    setEditMode(false);
    setError("");
  }

  async function saveAll() {
    if (
      !currentUser ||
      !canEdit ||
      saving
    ) {
      return;
    }

    setSaving(true);
    setError("");

    try {
      /*
        ปุ่มเดียวบันทึกทุกแถว
        ใช้ PATCH เดิมของ backend ทีละรายการ
      */
      for (const row of rows) {
        const response =
          await fetch(
            `${API_URL}/api/summary/edit/results/${row.id}`,
            {
              method: "PATCH",
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
                  caseId:
                    row.case_id,
                  urlSms:
                    row.url_sms,
                  aisResult:
                    row.ais_result,
                  trueDtacResult:
                    row.true_dtac_result,
                  ntResult:
                    row.nt_result,
                  cloudflareResult:
                    row.cloudflare_result,
                }),
            }
          );

        await readJson(response);
      }

      setOriginalRows(
        rows.map((row) => ({
          ...row,
        }))
      );

      setEditMode(false);

      // บันทึกเสร็จกลับหน้า Main Summary
      router.push("/summary");
    } catch (
      saveError
    ) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "บันทึกข้อมูลไม่สำเร็จ"
      );
    } finally {
      setSaving(false);
    }
  }

  function downloadExcel() {
    if (rows.length === 0) {
      alert(
        "ไม่มีข้อมูลสำหรับดาวน์โหลด"
      );
      return;
    }

    const exportRows =
      rows.map((row) => ({
        "CASE ID":
          row.case_id,
        "URL SMS":
          row.url_sms,
        AIS:
          row.ais_result ??
          "",
        "TRUE / DTAC":
          row.true_dtac_result ??
          "",
        NT:
          row.nt_result ??
          "",
        CloudFlare:
          row.cloudflare_result ??
          "",
      }));

    const worksheet =
      XLSX.utils.json_to_sheet(
        exportRows
      );

    worksheet["!cols"] = [
      { wch: 18 },
      { wch: 45 },
      { wch: 45 },
      { wch: 45 },
      { wch: 45 },
      { wch: 45 },
    ];

    const workbook =
      XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "URL Fake Web"
    );

    const sourceName =
      batch
        ?.original_file_name
        ?.replace(
          /\.(xlsx|xls)$/i,
          ""
        ) ||
      "url-fake-web";

    XLSX.writeFile(
      workbook,
      `${sourceName}_summary.xlsx`
    );
  }

  if (
    currentUser &&
    !canEdit
  ) {
    return (
      <main className="summary-page">
        <div className="summary-edit-denied">
          <h1>
            ไม่มีสิทธิ์แก้ไข
          </h1>

          <p>
            หน้านี้อนุญาตเฉพาะ Super Admin และ Admin
          </p>

          <button
            type="button"
            onClick={() =>
              router.push(
                "/summary"
              )
            }
          >
            กลับ Main Summary
          </button>
        </div>
      </main>
    );
  }

  return (
    <main className="summary-page">
      <div className="summary-edit-container">
        <section className="summary-edit-heading summary-edit-heading-simple">
          <div>
            <button
              type="button"
              className="summary-edit-back"
              onClick={() =>
                router.push(
                  "/summary"
                )
              }
            >
              ← กลับ Main Summary
            </button>

            <h1>
              URL Fake Web
            </h1>
          </div>

          <div className="summary-edit-top-actions">
            <button
              type="button"
              className="summary-download-excel"
              onClick={
                downloadExcel
              }
            >
              ↓ ดาวน์โหลด Excel
            </button>

            {!editMode ? (
              <button
                type="button"
                className="summary-edit-all"
                onClick={() =>
                  setEditMode(true)
                }
              >
                ✎ แก้ไข
              </button>
            ) : (
              <>
                <button
                  type="button"
                  className="summary-cancel-all"
                  disabled={saving}
                  onClick={
                    cancelEdit
                  }
                >
                  ยกเลิก
                </button>

                <button
                  type="button"
                  className="summary-save-all"
                  disabled={saving}
                  onClick={() =>
                    void saveAll()
                  }
                >
                  {saving
                    ? "กำลังบันทึก..."
                    : "✓ บันทึกทั้งหมด"}
                </button>
              </>
            )}
          </div>
        </section>

        {error && (
          <div className="summary-edit-error">
            {error}
          </div>
        )}

        <section className="summary-edit-list-card">
          <div className="summary-edit-list-header">
            <div>
              <h2>
                URL Fake Web
              </h2>

              <p>
                {rows.length.toLocaleString(
                  "th-TH"
                )}{" "}
                รายการ
                {editMode
                  ? " • กำลังแก้ไข"
                  : ""}
              </p>
            </div>

            <input
              type="text"
              className="summary-edit-search"
              placeholder="ค้นหา Case ID / URL"
              value={search}
              onChange={(
                event
              ) =>
                setSearch(
                  event.target.value
                )
              }
            />
          </div>

          <div className="summary-edit-table-wrapper">
            <table className="summary-edit-table summary-edit-inline-table">
              <thead>
                <tr>
                  <th>
                    Case ID
                  </th>
                  <th>
                    URL SMS
                  </th>
                  <th>AIS</th>
                  <th>
                    TRUE / DTAC
                  </th>
                  <th>NT</th>
                  <th>
                    CloudFlare
                  </th>
                </tr>
              </thead>

              <tbody>
                {loading ? (
                  <tr>
                    <td
                      colSpan={6}
                      className="summary-edit-empty"
                    >
                      กำลังโหลดข้อมูล...
                    </td>
                  </tr>
                ) : filteredRows.length ===
                  0 ? (
                  <tr>
                    <td
                      colSpan={6}
                      className="summary-edit-empty"
                    >
                      ไม่พบข้อมูล
                    </td>
                  </tr>
                ) : (
                  filteredRows.map(
                    (row) => (
                      <tr
                        key={row.id}
                      >
                        <td>
                          <input
                            className="summary-inline-input case-id"
                            type="text"
                            disabled={
                              !editMode
                            }
                            value={
                              row.case_id
                            }
                            onChange={(
                              event
                            ) =>
                              updateRow(
                                row.id,
                                "case_id",
                                event
                                  .target
                                  .value
                              )
                            }
                          />
                        </td>

                        <td>
                          <input
                            className="summary-inline-input url"
                            type="text"
                            disabled={
                              !editMode
                            }
                            value={
                              row.url_sms
                            }
                            onChange={(
                              event
                            ) =>
                              updateRow(
                                row.id,
                                "url_sms",
                                event
                                  .target
                                  .value
                              )
                            }
                          />
                        </td>

                        {[
                          [
                            "ais_result",
                            row.ais_result,
                          ],
                          [
                            "true_dtac_result",
                            row.true_dtac_result,
                          ],
                          [
                            "nt_result",
                            row.nt_result,
                          ],
                          [
                            "cloudflare_result",
                            row.cloudflare_result,
                          ],
                        ].map(
                          ([
                            field,
                            value,
                          ]) => (
                            <td
                              key={
                                field
                              }
                            >
                              <input
                                className={`summary-inline-input provider ${statusKind(
                                  value
                                )}`}
                                type="text"
                                disabled={
                                  !editMode
                                }
                                value={
                                  value ??
                                  ""
                                }
                                placeholder="ยังไม่ได้ตรวจ"
                                onChange={(
                                  event
                                ) =>
                                  updateRow(
                                    row.id,
                                    field as
                                      | "ais_result"
                                      | "true_dtac_result"
                                      | "nt_result"
                                      | "cloudflare_result",
                                    event
                                      .target
                                      .value
                                  )
                                }
                              />
                            </td>
                          )
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

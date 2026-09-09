"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./tool-op02.css";

type Props = {
  worksheet: XLSX.WorkSheet;
  sheetName: string;
};

type CellValue = string | number | boolean | null | undefined;
type MatrixRow = CellValue[];

type ToolOP02Row = {
  sourceRow: number;
  organization: string;
  status: string;
  hasApp: string;
  appName: string;
  ios: boolean;
  android: boolean;
};

type ViewType = "table" | "bar" | "pie";

type ColumnKey =
  | "organization"
  | "status"
  | "hasApp"
  | "appName"
  | "ios"
  | "android";

const COLUMN_OPTIONS: { key: ColumnKey; label: string }[] = [
  { key: "organization", label: "หน่วยงาน" },
  { key: "status", label: "สถานะ OP02" },
  { key: "hasApp", label: "มีแอปพลิเคชัน" },
  { key: "appName", label: "ชื่อแอป" },
  { key: "ios", label: "iOS" },
  { key: "android", label: "Android" },
];

const DEFAULT_VISIBLE_COLUMNS: ColumnKey[] =
  COLUMN_OPTIONS.map((column) => column.key);


function cleanText(value: unknown): string {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalize(value: unknown): string {
  return cleanText(value).toLowerCase();
}

function fillMergedCells(
  worksheet: XLSX.WorkSheet
): MatrixRow[] {
  const ref = worksheet["!ref"];

  if (!ref) return [];

  const range = XLSX.utils.decode_range(ref);
  const result: MatrixRow[] = [];

  for (let r = range.s.r; r <= range.e.r; r += 1) {
    const row: MatrixRow = [];

    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const address = XLSX.utils.encode_cell({ r, c });
      const cell = worksheet[address];

      row[c] = cell?.w ?? cell?.v ?? "";
    }

    result[r] = row;
  }

  for (const merge of worksheet["!merges"] ?? []) {
    const sourceValue =
      result[merge.s.r]?.[merge.s.c] ?? "";

    for (let r = merge.s.r; r <= merge.e.r; r += 1) {
      if (!result[r]) result[r] = [];

      for (let c = merge.s.c; c <= merge.e.c; c += 1) {
        if (cleanText(result[r][c]) === "") {
          result[r][c] = sourceValue;
        }
      }
    }
  }

  return result;
}

function findHeaderRow(matrix: MatrixRow[]) {
  const maxRows = Math.min(matrix.length, 30);

  for (let r = 0; r < maxRows; r += 1) {
    const row = matrix[r] ?? [];
    const values = row.map(normalize);

    /*
      ใน Sheet จริงคอลัมน์ A ไม่มีชื่อ Header
      จึงใช้หัว B-F ในการหาแถว Header แทน
    */
    const hasStatus = values.some(
      (value) =>
        value.includes("เสร็จสิ้น 12") ||
        value.includes("op2")
    );

    const hasApp = values.some(
      (value) =>
        value.includes("มีแอปพลิเคชัน")
    );

    const hasAppName = values.some(
      (value) =>
        value.includes("ชื่อแอป")
    );

    const hasIOS = values.some(
      (value) => value === "ios"
    );

    const hasAndroid = values.some(
      (value) => value.includes("android")
    );

    if (
      hasStatus &&
      hasApp &&
      hasAppName &&
      (hasIOS || hasAndroid)
    ) {
      return r;
    }
  }

  return -1;
}

function findColumn(
  row: MatrixRow,
  candidates: string[]
) {
  for (let c = 0; c < row.length; c += 1) {
    const value = normalize(row[c]);

    if (
      candidates.some((candidate) =>
        value.includes(normalize(candidate))
      )
    ) {
      return c;
    }
  }

  return -1;
}

function parseBooleanCell(value: unknown) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value === 1;
  }

  const text = normalize(value);

  if (!text) return false;

  return [
    "true",
    "yes",
    "1",
    "✓",
    "✔",
    "☑",
    "checked",
  ].some((item) => text.includes(item));
}

function isCompleted(value: string) {
  const status = normalize(value);

  return (
    status.includes("ดำเนินการเสร็จสิ้น") ||
    status.includes("เสร็จสิ้น") ||
    status.includes("complete")
  );
}

function hasApplication(value: string) {
  const text = normalize(value);

  return (
    text === "มี" ||
    text.includes("มีแอป") ||
    text.includes("มี application")
  );
}

export default function ToolOP02({
  worksheet,
  sheetName,
}: Props) {
  const [search, setSearch] = useState("");
  const [viewType, setViewType] =
    useState<ViewType>("table");

  const [showColumnPicker, setShowColumnPicker] =
    useState(false);

  const [visibleColumns, setVisibleColumns] =
    useState<ColumnKey[]>(DEFAULT_VISIBLE_COLUMNS);

  function isColumnVisible(key: ColumnKey) {
    return visibleColumns.includes(key);
  }

  function toggleColumn(key: ColumnKey) {
    setVisibleColumns((current) =>
      current.includes(key)
        ? current.filter((item) => item !== key)
        : [...current, key]
    );
  }

  function setAllColumns(visible: boolean) {
    setVisibleColumns(
      visible ? DEFAULT_VISIBLE_COLUMNS : []
    );
  }

  const [showHasApp, setShowHasApp] =
    useState(true);

  const [showNoApp, setShowNoApp] =
    useState(true);

  const [showIOS, setShowIOS] =
    useState(true);

  const [showAndroid, setShowAndroid] =
    useState(true);

  const [showCompleted, setShowCompleted] =
    useState(true);

  const [showNotCompleted, setShowNotCompleted] =
    useState(true);

  const matrix = useMemo(
    () => fillMergedCells(worksheet),
    [worksheet]
  );

  const headerRowIndex = useMemo(
    () => findHeaderRow(matrix),
    [matrix]
  );

  const rows = useMemo<ToolOP02Row[]>(() => {
    if (headerRowIndex < 0) {
      return [];
    }

    const headerRow =
      matrix[headerRowIndex] ?? [];

    /*
      คอลัมน์ A คือชื่อหน่วยงาน แต่ A1 ไม่มีชื่อ Header
      ถ้าหา header ไม่เจอ ให้ใช้ column 0
    */
    const detectedOrganizationColumn =
      findColumn(headerRow, [
        "หน่วยงาน",
        "ชื่อหน่วยงาน",
      ]);

    const organizationColumn =
      detectedOrganizationColumn >= 0
        ? detectedOrganizationColumn
        : 0;

    const statusColumn =
      findColumn(headerRow, [
        "เสร็จสิ้น 12",
        "op2",
      ]);

    const hasAppColumn =
      findColumn(headerRow, [
        "มีแอปพลิเคชัน",
      ]);

    const appNameColumn =
      findColumn(headerRow, [
        "ชื่อแอป",
      ]);

    const iosColumn =
      findColumn(headerRow, [
        "ios",
      ]);

    const androidColumn =
      findColumn(headerRow, [
        "android",
      ]);

    const result: ToolOP02Row[] = [];

    let currentOrganization = "";

    for (
      let r = headerRowIndex + 1;
      r < matrix.length;
      r += 1
    ) {
      const source = matrix[r] ?? [];

      const rawOrganization =
        organizationColumn >= 0
          ? cleanText(source[organizationColumn])
          : "";

      if (rawOrganization) {
        currentOrganization =
          rawOrganization;
      }

      const status =
        statusColumn >= 0
          ? cleanText(source[statusColumn])
          : "";

      const hasApp =
        hasAppColumn >= 0
          ? cleanText(source[hasAppColumn])
          : "";

      const appName =
        appNameColumn >= 0
          ? cleanText(source[appNameColumn])
          : "";

      const ios =
        iosColumn >= 0
          ? parseBooleanCell(source[iosColumn])
          : false;

      const android =
        androidColumn >= 0
          ? parseBooleanCell(source[androidColumn])
          : false;

      const hasAnyData =
        currentOrganization ||
        status ||
        hasApp ||
        appName ||
        ios ||
        android;

      if (!hasAnyData) {
        continue;
      }

      if (
        !currentOrganization &&
        !appName
      ) {
        continue;
      }

      result.push({
        sourceRow: r + 1,
        organization: currentOrganization,
        status,
        hasApp,
        appName,
        ios,
        android,
      });
    }

    return result;
  }, [matrix, headerRowIndex]);

  const filteredRows = useMemo(() => {
    const keyword =
      normalize(search);

    return rows.filter((row) => {
      const matchesSearch =
        !keyword ||
        [
          row.organization,
          row.status,
          row.hasApp,
          row.appName,
        ].some((value) =>
          normalize(value).includes(keyword)
        );

      if (!matchesSearch) {
        return false;
      }

      const appState =
        hasApplication(row.hasApp);

      if (appState && !showHasApp) {
        return false;
      }

      if (!appState && !showNoApp) {
        return false;
      }

      if (row.ios && !showIOS) {
        return false;
      }

      if (row.android && !showAndroid) {
        return false;
      }

      const completed =
        isCompleted(row.status);

      if (
        completed &&
        !showCompleted
      ) {
        return false;
      }

      if (
        !completed &&
        !showNotCompleted
      ) {
        return false;
      }

      return true;
    });
  }, [
    rows,
    search,
    showHasApp,
    showNoApp,
    showIOS,
    showAndroid,
    showCompleted,
    showNotCompleted,
  ]);

  const summary = useMemo(() => {
    const organizations =
      new Set(
        rows
          .map((row) => row.organization)
          .filter(Boolean)
      );

    const apps =
      rows.filter((row) =>
        hasApplication(row.hasApp)
      );

    const noApps =
      rows.filter((row) =>
        !hasApplication(row.hasApp)
      );

    const ios =
      rows.filter((row) => row.ios);

    const android =
      rows.filter((row) => row.android);

    const completed =
      rows.filter((row) =>
        isCompleted(row.status)
      );

    return {
      organizations: organizations.size,
      apps: apps.length,
      noApps: noApps.length,
      ios: ios.length,
      android: android.length,
      completed: completed.length,
      total: rows.length,
    };
  }, [rows]);

  const barItems = [
    {
      label: "มีแอป",
      value: summary.apps,
    },
    {
      label: "ไม่มีแอป",
      value: summary.noApps,
    },
    {
      label: "iOS",
      value: summary.ios,
    },
    {
      label: "Android",
      value: summary.android,
    },
  ];

  const maxBarValue =
    Math.max(
      1,
      ...barItems.map((item) => item.value)
    );

  const appPercent =
    summary.total > 0
      ? Math.round(
          (summary.apps / summary.total) * 100
        )
      : 0;

  function openInExcel() {
    try {
      const workbook = XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        sheetName
      );

      const safeFileName =
        cleanText(sheetName)
          .replace(/[\\/:*?"<>|]/g, "_")
          .trim() || "tool-op02";

      XLSX.writeFile(
        workbook,
        `${safeFileName}.xlsx`
      );
    } catch (error) {
      console.error("Open Excel error:", error);
      alert("สร้างไฟล์ Excel ไม่สำเร็จ");
    }
  }

  return (
    <section className="tool-op02-page">
      <div className="tool-op02-heading">
        <div>
          <h2>{sheetName}</h2>

          <p>
            ติดตามสถานะ OP02 และแพลตฟอร์มของแอปพลิเคชัน
          </p>
        </div>

        <span className="tool-op02-count">
          {rows.length} รายการ
        </span>
      </div>

      <div className="tool-op02-toolbar">
        <div className="tool-op02-search">
          <span>⌕</span>

          <input
            type="text"
            value={search}
            placeholder="ค้นหาหน่วยงาน / ชื่อแอป..."
            onChange={(event) =>
              setSearch(event.target.value)
            }
          />
        </div>

        <div className="tool-op02-toolbar-actions">
          <button
            type="button"
            className="tool-op02-excel-button"
            onClick={openInExcel}
            title="สร้างไฟล์ .xlsx ของ Sheet นี้เพื่อเปิดด้วย Microsoft Excel"
          >
            ▣ เปิดใน Excel
          </button>

          <div className="tool-op02-column-wrapper">
            <button
              type="button"
              className="tool-op02-column-button"
              onClick={() =>
                setShowColumnPicker((current) => !current)
              }
            >
              ☷ เลือกคอลัมน์ ({visibleColumns.length}/{COLUMN_OPTIONS.length})
            </button>

            {showColumnPicker && (
              <div className="tool-op02-column-picker">
                <div className="tool-op02-column-picker-head">
                  <strong>เลือกคอลัมน์ที่ต้องการแสดง</strong>

                  <div>
                    <button
                      type="button"
                      onClick={() => setAllColumns(true)}
                    >
                      เลือกทั้งหมด
                    </button>

                    <button
                      type="button"
                      onClick={() => setAllColumns(false)}
                    >
                      ล้างทั้งหมด
                    </button>
                  </div>
                </div>

                <div className="tool-op02-column-picker-grid">
                  {COLUMN_OPTIONS.map((column) => (
                    <label key={column.key}>
                      <input
                        type="checkbox"
                        checked={isColumnVisible(column.key)}
                        onChange={() => toggleColumn(column.key)}
                      />
                      {column.label}
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>

          <span className="tool-op02-result-count">
            {filteredRows.length} รายการ
          </span>
        </div>
      </div>

      <div className="tool-op02-filters">
        <strong>
          ตัวกรอง
        </strong>

        <label>
          <input
            type="checkbox"
            checked={showHasApp}
            onChange={(event) =>
              setShowHasApp(
                event.target.checked
              )
            }
          />
          มีแอป
        </label>

        <label>
          <input
            type="checkbox"
            checked={showNoApp}
            onChange={(event) =>
              setShowNoApp(
                event.target.checked
              )
            }
          />
          ไม่มีแอป
        </label>

        <label>
          <input
            type="checkbox"
            checked={showIOS}
            onChange={(event) =>
              setShowIOS(
                event.target.checked
              )
            }
          />
          iOS
        </label>

        <label>
          <input
            type="checkbox"
            checked={showAndroid}
            onChange={(event) =>
              setShowAndroid(
                event.target.checked
              )
            }
          />
          Android
        </label>

        <label>
          <input
            type="checkbox"
            checked={showCompleted}
            onChange={(event) =>
              setShowCompleted(
                event.target.checked
              )
            }
          />
          เสร็จสิ้น
        </label>

        <label>
          <input
            type="checkbox"
            checked={showNotCompleted}
            onChange={(event) =>
              setShowNotCompleted(
                event.target.checked
              )
            }
          />
          ยังไม่เสร็จ
        </label>
      </div>

      <div className="tool-op02-summary-grid">
        <article>
          <span>หน่วยงาน</span>
          <strong>
            {summary.organizations}
          </strong>
          <small>หน่วยงาน</small>
        </article>

        <article>
          <span>มีแอป</span>
          <strong>
            {summary.apps}
          </strong>
          <small>รายการ</small>
        </article>

        <article>
          <span>ไม่มีแอป</span>
          <strong>
            {summary.noApps}
          </strong>
          <small>รายการ</small>
        </article>

        <article>
          <span>iOS</span>
          <strong>
            {summary.ios}
          </strong>
          <small>รายการ</small>
        </article>

        <article>
          <span>Android</span>
          <strong>
            {summary.android}
          </strong>
          <small>รายการ</small>
        </article>

        <article>
          <span>ดำเนินการเสร็จสิ้น</span>
          <strong>
            {summary.completed}
          </strong>
          <small>รายการ</small>
        </article>
      </div>

      <section className="tool-op02-main-card">
        <div className="tool-op02-card-header">
          <div>
            <h3>
              ภาพรวม Tool OP02
            </h3>

            <p>
              เลือกรูปแบบการแสดงผล
            </p>
          </div>

          <div className="tool-op02-tabs">
            {(
              [
                ["table", "TABLE"],
                ["bar", "BAR"],
                ["pie", "PIE"],
              ] as const
            ).map(([value, label]) => (
              <button
                key={value}
                type="button"
                className={
                  viewType === value
                    ? "active"
                    : ""
                }
                onClick={() =>
                  setViewType(value)
                }
              >
                {label}
              </button>
            ))}
          </div>
        </div>

        {viewType === "table" && (
          <div className="tool-op02-table-wrapper">
            {visibleColumns.length === 0 ? (
              <div className="tool-op02-no-columns">
                กรุณาเลือกอย่างน้อย 1 คอลัมน์
              </div>
            ) : (
              <table
                className="tool-op02-table"
                style={{
                  minWidth: `${Math.max(
                    520,
                    visibleColumns.length * 150
                  )}px`,
                }}
              >
                <thead>
                  <tr>
                    {isColumnVisible("organization") && (
                      <th>หน่วยงาน</th>
                    )}
                    {isColumnVisible("status") && (
                      <th>สถานะ OP02</th>
                    )}
                    {isColumnVisible("hasApp") && (
                      <th>มีแอปพลิเคชัน</th>
                    )}
                    {isColumnVisible("appName") && (
                      <th>ชื่อแอป</th>
                    )}
                    {isColumnVisible("ios") && (
                      <th>iOS</th>
                    )}
                    {isColumnVisible("android") && (
                      <th>Android</th>
                    )}
                  </tr>
                </thead>

                <tbody>
                  {filteredRows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={Math.max(1, visibleColumns.length)}
                        className="tool-op02-empty"
                      >
                        ไม่พบข้อมูล
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map((row, index) => {
                      const previousRow =
                        index > 0
                          ? filteredRows[index - 1]
                          : null;

                      const sameOrganization =
                        previousRow &&
                        normalize(previousRow.organization) ===
                          normalize(row.organization);

                      return (
                        <tr key={row.sourceRow}>
                          {isColumnVisible("organization") && (
                            <td className="tool-op02-org">
                              {sameOrganization
                                ? ""
                                : row.organization || "-"}
                            </td>
                          )}

                          {isColumnVisible("status") && (
                            <td>
                              <span
                                className={`tool-op02-status ${
                                  isCompleted(row.status)
                                    ? "done"
                                    : "not-done"
                                }`}
                              >
                                {row.status || "-"}
                              </span>
                            </td>
                          )}

                          {isColumnVisible("hasApp") && (
                            <td>
                              <span
                                className={`tool-op02-app-badge ${
                                  hasApplication(row.hasApp)
                                    ? "yes"
                                    : "no"
                                }`}
                              >
                                {row.hasApp || "-"}
                              </span>
                            </td>
                          )}

                          {isColumnVisible("appName") && (
                            <td>{row.appName || "-"}</td>
                          )}

                          {isColumnVisible("ios") && (
                            <td className="tool-op02-platform">
                              {row.ios ? (
                                <span className="checked">✓</span>
                              ) : (
                                <span className="unchecked">—</span>
                              )}
                            </td>
                          )}

                          {isColumnVisible("android") && (
                            <td className="tool-op02-platform">
                              {row.android ? (
                                <span className="checked">✓</span>
                              ) : (
                                <span className="unchecked">—</span>
                              )}
                            </td>
                          )}
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            )}
          </div>
        )}

        {viewType === "bar" && (
          <div className="tool-op02-chart-area">
            <div className="tool-op02-bar-chart">
              {barItems.map((item) => {
                const height =
                  (item.value /
                    maxBarValue) *
                  100;

                return (
                  <div
                    key={item.label}
                    className="tool-op02-bar-item"
                  >
                    <strong>
                      {item.value}
                    </strong>

                    <div className="tool-op02-bar-track">
                      <div
                        className="tool-op02-bar-fill"
                        style={{
                          height: `${height}%`,
                        }}
                      />
                    </div>

                    <span>
                      {item.label}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {viewType === "pie" && (
          <div className="tool-op02-pie-area">
            <div
              className="tool-op02-pie"
              style={{
                background: `conic-gradient(
                  #50ad63 0 ${appPercent}%,
                  #dfe4eb ${appPercent}% 100%
                )`,
              }}
            >
              <div className="tool-op02-pie-hole">
                <strong>
                  {appPercent}%
                </strong>

                <span>
                  มีแอป
                </span>
              </div>
            </div>

            <div className="tool-op02-pie-legend">
              <span>
                <i className="yes" />
                มีแอป {summary.apps}
              </span>

              <span>
                <i className="no" />
                ไม่มีแอป {summary.noApps}
              </span>
            </div>
          </div>
        )}
      </section>
    </section>
  );
}

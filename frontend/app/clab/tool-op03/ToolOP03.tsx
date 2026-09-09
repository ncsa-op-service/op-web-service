"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./tool-op03.css";

type Props = {
  worksheet: XLSX.WorkSheet;
  sheetName: string;
};

type CellValue =
  | string
  | number
  | boolean
  | null
  | undefined;

type MatrixRow = CellValue[];

type ToolOP03Row = {
  sourceRow: number;
  order: string;
  organization: string;

  allowDGA: boolean;
  allow365: boolean;
  submitUser: boolean;
  template: boolean;

  schedule: string;
  note: string;
  contact: string;
};

type ViewType =
  | "table"
  | "bar"
  | "pie";

type StatusFilter =
  | "done"
  | "sending"
  | "waiting"
  | "other";

type ColumnKey =
  | "order"
  | "organization"
  | "allowDGA"
  | "allow365"
  | "submitUser"
  | "template"
  | "schedule"
  | "note"
  | "contact";

const COLUMN_OPTIONS: { key: ColumnKey; label: string; checklist?: boolean }[] = [
  { key: "order", label: "ลำดับ" },
  { key: "organization", label: "ชื่อหน่วยงาน" },
  { key: "allowDGA", label: "DGA", checklist: true },
  { key: "allow365", label: "365", checklist: true },
  { key: "submitUser", label: "User ส่งเมล", checklist: true },
  { key: "template", label: "Template", checklist: true },
  { key: "schedule", label: "กำหนดการ" },
  { key: "note", label: "หมายเหตุ" },
  { key: "contact", label: "ข้อมูลติดต่อ" },
];

const DEFAULT_VISIBLE_COLUMNS: ColumnKey[] =
  COLUMN_OPTIONS.map((column) => column.key);

/* ==============================
   TEXT
================================ */

function cleanText(
  value: unknown
): string {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(
      /[\u200B-\u200D\uFEFF]/g,
      ""
    )
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalize(
  value: unknown
): string {
  return cleanText(
    value
  ).toLowerCase();
}

/* ==============================
   EXCEL -> MATRIX
================================ */

function buildMatrix(
  worksheet: XLSX.WorkSheet
): MatrixRow[] {
  const ref =
    worksheet["!ref"];

  if (!ref) {
    return [];
  }

  const range =
    XLSX.utils.decode_range(
      ref
    );

  const matrix: MatrixRow[] =
    [];

  for (
    let r = range.s.r;
    r <= range.e.r;
    r += 1
  ) {
    const row: MatrixRow =
      [];

    for (
      let c = range.s.c;
      c <= range.e.c;
      c += 1
    ) {
      const address =
        XLSX.utils.encode_cell({
          r,
          c,
        });

      const cell =
        worksheet[address];

      row[c] =
        cell?.w ??
        cell?.v ??
        "";
    }

    matrix[r] =
      row;
  }

  /*
    เติม merged cell เพื่อให้ชื่อหน่วยงาน
    หรือค่าที่ merge หลายแถวยังอ่านได้
  */
  for (
    const merge of
      worksheet["!merges"] ?? []
  ) {
    const sourceValue =
      matrix[
        merge.s.r
      ]?.[
        merge.s.c
      ] ?? "";

    for (
      let r = merge.s.r;
      r <= merge.e.r;
      r += 1
    ) {
      if (!matrix[r]) {
        matrix[r] = [];
      }

      for (
        let c = merge.s.c;
        c <= merge.e.c;
        c += 1
      ) {
        if (
          cleanText(
            matrix[r][c]
          ) === ""
        ) {
          matrix[r][c] =
            sourceValue;
        }
      }
    }
  }

  return matrix;
}

/* ==============================
   HEADER
================================ */

function findHeaderRows(
  matrix: MatrixRow[]
) {
  const limit =
    Math.min(
      matrix.length,
      30
    );

  for (
    let r = 0;
    r < limit;
    r += 1
  ) {
    const current =
      (matrix[r] ?? [])
        .map(normalize)
        .join(" ");

    const next =
      (matrix[r + 1] ?? [])
        .map(normalize)
        .join(" ");

    const combined =
      `${current} ${next}`;

    const hasOrg =
      combined.includes(
        "ชื่อหน่วยงาน"
      );

    const hasAllow =
      combined.includes(
        "allow"
      );

    const hasTemplate =
      combined.includes(
        "template"
      );

    const hasUser =
      combined.includes(
        "user"
      );

    if (
      hasOrg &&
      hasAllow &&
      hasTemplate &&
      hasUser
    ) {
      return {
        top: r,
        bottom:
          r + 1,
      };
    }
  }

  return null;
}

function combinedHeaderText(
  matrix: MatrixRow[],
  top: number,
  bottom: number,
  column: number
) {
  const values: string[] =
    [];

  for (
    let r = top;
    r <= bottom + 1;
    r += 1
  ) {
    values.push(
      cleanText(
        matrix[r]?.[
          column
        ]
      )
    );
  }

  return normalize(
    values.join(" ")
  );
}

function findColumn(
  matrix: MatrixRow[],
  top: number,
  bottom: number,
  keywords: string[]
) {
  const maxColumns =
    Math.max(
      ...matrix
        .slice(
          top,
          bottom + 2
        )
        .map(
          (row) =>
            row?.length ??
            0
        ),
      0
    );

  for (
    let c = 0;
    c < maxColumns;
    c += 1
  ) {
    const header =
      combinedHeaderText(
        matrix,
        top,
        bottom,
        c
      );

    if (
      keywords.every(
        (keyword) =>
          header.includes(
            normalize(
              keyword
            )
          )
      )
    ) {
      return c;
    }
  }

  return -1;
}

/* ==============================
   CHECKBOX
================================ */

function parseCheckbox(
  value: unknown
) {
  if (
    typeof value ===
    "boolean"
  ) {
    return value;
  }

  if (
    typeof value ===
    "number"
  ) {
    return value === 1;
  }

  const text =
    normalize(value);

  if (!text) {
    return false;
  }

  return [
    "true",
    "yes",
    "1",
    "✓",
    "✔",
    "☑",
    "checked",
  ].some(
    (token) =>
      text.includes(
        token
      )
  );
}

/* ==============================
   STATUS
================================ */

function getStatusType(
  note: string
): StatusFilter {
  const value =
    normalize(note);

  if (
    value.includes(
      "ดำเนินการเสร็จแล้ว"
    ) ||
    value.includes(
      "ดำเนินการเสร็จ"
    ) ||
    value.includes(
      "เสร็จแล้ว"
    )
  ) {
    return "done";
  }

  if (
    value.includes(
      "กำลังดำเนินการส่ง"
    ) ||
    value.includes(
      "กำลังดำเนินการ"
    )
  ) {
    return "sending";
  }

  if (
    value.includes("รอส่ง") ||
    value.includes("รอติดต่อ") ||
    value.includes("ติดต่อ") ||
    value.includes("ยังไม่ระบุ")
  ) {
    return "waiting";
  }

  return "other";
}

function statusLabel(
  type: StatusFilter
) {
  switch (type) {
    case "done":
      return "ดำเนินการเสร็จแล้ว";

    case "sending":
      return "กำลังดำเนินการ";

    case "waiting":
      return "รอดำเนินการ";

    default:
      return "อื่น ๆ";
  }
}

/* ==============================
   COMPONENT
================================ */

export default function ToolOP03({
  worksheet,
  sheetName,
}: Props) {
  const [
    search,
    setSearch,
  ] = useState("");

  const [
    viewType,
    setViewType,
  ] =
    useState<ViewType>(
      "table"
    );

  const [
    showColumnPicker,
    setShowColumnPicker,
  ] = useState(false);

  const [
    visibleColumns,
    setVisibleColumns,
  ] = useState<ColumnKey[]>(
    DEFAULT_VISIBLE_COLUMNS
  );

  function isColumnVisible(
    key: ColumnKey
  ) {
    return visibleColumns.includes(
      key
    );
  }

  function toggleColumn(
    key: ColumnKey
  ) {
    setVisibleColumns(
      (current) =>
        current.includes(key)
          ? current.filter(
              (item) =>
                item !== key
            )
          : [
              ...current,
              key,
            ]
    );
  }

  function setAllColumns(
    visible: boolean
  ) {
    setVisibleColumns(
      visible
        ? DEFAULT_VISIBLE_COLUMNS
        : []
    );
  }

  const visibleChecklistCount =
    COLUMN_OPTIONS.filter(
      ({ key, checklist }) =>
        checklist &&
        visibleColumns.includes(key)
    ).length;

  const [
    showDone,
    setShowDone,
  ] = useState(true);

  const [
    showSending,
    setShowSending,
  ] = useState(true);

  const [
    showWaiting,
    setShowWaiting,
  ] = useState(true);

  const [
    showOther,
    setShowOther,
  ] = useState(true);

  const [
    requireDGA,
    setRequireDGA,
  ] = useState(false);

  const [
    require365,
    setRequire365,
  ] = useState(false);

  const [
    requireSubmitUser,
    setRequireSubmitUser,
  ] = useState(false);

  const [
    requireTemplate,
    setRequireTemplate,
  ] = useState(false);

  /* ==============================
     MATRIX
  ============================== */

  const matrix =
    useMemo(
      () =>
        buildMatrix(
          worksheet
        ),
      [worksheet]
    );

  const headerInfo =
    useMemo(
      () =>
        findHeaderRows(
          matrix
        ),
      [matrix]
    );

  /* ==============================
     ROWS
  ============================== */

  const rows =
    useMemo<
      ToolOP03Row[]
    >(() => {
      if (
        !headerInfo
      ) {
        return [];
      }

      const {
        top,
        bottom,
      } =
        headerInfo;

      const orderColumn =
        findColumn(
          matrix,
          top,
          bottom,
          ["ลำดับ"]
        );

      const organizationColumn =
        findColumn(
          matrix,
          top,
          bottom,
          ["ชื่อหน่วยงาน"]
        );

      const dgaColumn =
        findColumn(
          matrix,
          top,
          bottom,
          [
            "allow",
            "dga",
          ]
        );

      const column365 =
        findColumn(
          matrix,
          top,
          bottom,
          [
            "allow",
            "365",
          ]
        );

      const userColumn =
        findColumn(
          matrix,
          top,
          bottom,
          ["user"]
        );

      const templateColumn =
        findColumn(
          matrix,
          top,
          bottom,
          ["template"]
        );

      const noteColumn =
        findColumn(
          matrix,
          top,
          bottom,
          ["หมายเหตุ"]
        );

      /*
        ช่องกำหนดการในไฟล์ไม่มีชื่อ Header ชัดเจน
        จากภาพอยู่ก่อน "หมายเหตุ" 1 คอลัมน์
      */
      const scheduleColumn =
        noteColumn > 0
          ? noteColumn - 1
          : -1;

      /*
        ข้อมูลติดต่ออยู่ถัดจากหมายเหตุด้านขวา
        เช่น กรมประมง / ชื่อผู้ประสานงาน / เบอร์ / อีเมล
      */
      const contactColumn =
        noteColumn >= 0
          ? noteColumn + 1
          : -1;

      const result:
        ToolOP03Row[] =
        [];

      let currentOrder = "";
      let currentOrganization =
        "";

      const firstDataRow =
        bottom + 2;

      for (
        let r =
          firstDataRow;
        r <
        matrix.length;
        r += 1
      ) {
        const source =
          matrix[r] ?? [];

        const rawOrder =
          orderColumn >= 0
            ? cleanText(
                source[
                  orderColumn
                ]
              )
            : "";

        const rawOrganization =
          organizationColumn >= 0
            ? cleanText(
                source[
                  organizationColumn
                ]
              )
            : "";

        if (
          rawOrder
        ) {
          currentOrder =
            rawOrder;
        }

        if (
          rawOrganization
        ) {
          currentOrganization =
            rawOrganization;
        }

        const allowDGA =
          dgaColumn >= 0
            ? parseCheckbox(
                source[
                  dgaColumn
                ]
              )
            : false;

        const allow365 =
          column365 >= 0
            ? parseCheckbox(
                source[
                  column365
                ]
              )
            : false;

        const submitUser =
          userColumn >= 0
            ? parseCheckbox(
                source[
                  userColumn
                ]
              )
            : false;

        const template =
          templateColumn >= 0
            ? parseCheckbox(
                source[
                  templateColumn
                ]
              )
            : false;

        const schedule =
          scheduleColumn >= 0
            ? cleanText(
                source[
                  scheduleColumn
                ]
              )
            : "";

        const note =
          noteColumn >= 0
            ? cleanText(
                source[
                  noteColumn
                ]
              )
            : "";

        const contact =
          contactColumn >= 0
            ? cleanText(
                source[
                  contactColumn
                ]
              )
            : "";

        const hasData =
          currentOrder ||
          currentOrganization ||
          allowDGA ||
          allow365 ||
          submitUser ||
          template ||
          schedule ||
          note ||
          contact;

        if (
          !hasData
        ) {
          continue;
        }

        if (
          !currentOrganization &&
          !note &&
          !contact
        ) {
          continue;
        }

        result.push({
          sourceRow:
            r + 1,

          order:
            currentOrder,

          organization:
            currentOrganization,

          allowDGA,
          allow365,
          submitUser,
          template,

          schedule,
          note,
          contact,
        });
      }

      return result;
    }, [
      matrix,
      headerInfo,
    ]);

  /* ==============================
     FILTER
  ============================== */

  const filteredRows =
    useMemo(() => {
      const keyword =
        normalize(
          search
        );

      return rows.filter(
        (row) => {
          const matchesSearch =
            !keyword ||
            [
              row.order,
              row.organization,
              row.schedule,
              row.note,
              row.contact,
            ].some(
              (value) =>
                normalize(
                  value
                ).includes(
                  keyword
                )
            );

          if (
            !matchesSearch
          ) {
            return false;
          }

          const status =
            getStatusType(
              row.note
            );

          if (
            status ===
              "done" &&
            !showDone
          ) {
            return false;
          }

          if (
            status ===
              "sending" &&
            !showSending
          ) {
            return false;
          }

          if (
            status ===
              "waiting" &&
            !showWaiting
          ) {
            return false;
          }

          if (
            status ===
              "other" &&
            !showOther
          ) {
            return false;
          }

          if (
            requireDGA &&
            !row.allowDGA
          ) {
            return false;
          }

          if (
            require365 &&
            !row.allow365
          ) {
            return false;
          }

          if (
            requireSubmitUser &&
            !row.submitUser
          ) {
            return false;
          }

          if (
            requireTemplate &&
            !row.template
          ) {
            return false;
          }

          return true;
        }
      );
    }, [
      rows,
      search,
      showDone,
      showSending,
      showWaiting,
      showOther,
      requireDGA,
      require365,
      requireSubmitUser,
      requireTemplate,
    ]);

  /* ==============================
     SUMMARY
  ============================== */

  const summary =
    useMemo(() => {
      const organizations =
        new Set(
          rows
            .map(
              (row) =>
                row.organization
            )
            .filter(
              Boolean
            )
        );

      const done =
        rows.filter(
          (row) =>
            getStatusType(
              row.note
            ) === "done"
        ).length;

      const sending =
        rows.filter(
          (row) =>
            getStatusType(
              row.note
            ) === "sending"
        ).length;

      const waiting =
        rows.filter(
          (row) =>
            getStatusType(
              row.note
            ) === "waiting"
        ).length;

      const dga =
        rows.filter(
          (row) =>
            row.allowDGA
        ).length;

      const microsoft365 =
        rows.filter(
          (row) =>
            row.allow365
        ).length;

      const submitUser =
        rows.filter(
          (row) =>
            row.submitUser
        ).length;

      const template =
        rows.filter(
          (row) =>
            row.template
        ).length;

      return {
        total:
          rows.length,

        organizations:
          organizations.size,

        done,
        sending,
        waiting,

        dga,
        microsoft365,
        submitUser,
        template,
      };
    }, [rows]);

  /* ==============================
     BAR
  ============================== */

  const barItems = [
    {
      label:
        "DGA",
      value:
        summary.dga,
    },
    {
      label:
        "365",
      value:
        summary.microsoft365,
    },
    {
      label:
        "User ส่งเมล",
      value:
        summary.submitUser,
    },
    {
      label:
        "Template",
      value:
        summary.template,
    },
  ];

  const maxBarValue =
    Math.max(
      1,
      ...barItems.map(
        (item) =>
          item.value
      )
    );

  /* ==============================
     PIE
  ============================== */

  const pieTotal =
    Math.max(
      1,
      summary.done +
        summary.sending +
        summary.waiting
    );

  const donePercent =
    (
      summary.done /
      pieTotal
    ) * 100;

  const sendingPercent =
    (
      summary.sending /
      pieTotal
    ) * 100;

  const pieBackground =
    `conic-gradient(
      #50ad63 0 ${donePercent}%,
      #f5b51d ${donePercent}% ${donePercent + sendingPercent}%,
      #d8dee8 ${donePercent + sendingPercent}% 100%
    )`;

  function openInExcel() {
    try {
      const outputWorkbook =
        XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        worksheet,
        sheetName
      );

      const safeFileName =
        cleanText(
          sheetName
        )
          .replace(
            /[\\/:*?"<>|]/g,
            "_"
          )
          .trim() ||
        "tool-op03";

      XLSX.writeFile(
        outputWorkbook,
        `${safeFileName}.xlsx`
      );
    } catch (excelError) {
      console.error(
        "Open Excel error:",
        excelError
      );

      alert(
        "สร้างไฟล์ Excel ไม่สำเร็จ"
      );
    }
  }

  /* ==============================
     RENDER
  ============================== */

  return (
    <section className="tool-op03-page">

      {/* HEADER */}

      <div className="tool-op03-heading">
        <div>
          <span className="tool-op03-kicker">
            CLAB / OP03
          </span>

          <h2>
            {sheetName}
          </h2>

          <p>
            ระบบติดตาม Checklist การเตรียมความพร้อมและการดำเนินงาน Tool OP03
          </p>
        </div>

        <div className="tool-op03-heading-meta">
          <span>
            {summary.organizations}
          </span>

          <small>
            หน่วยงาน
          </small>
        </div>
      </div>

      {/* SEARCH */}

      <div className="tool-op03-toolbar">
        <div className="tool-op03-search">
          <span>⌕</span>

          <input
            type="text"
            value={
              search
            }
            placeholder="ค้นหาหน่วยงาน / หมายเหตุ / ผู้ประสานงาน..."
            onChange={(
              event
            ) =>
              setSearch(
                event
                  .target
                  .value
              )
            }
          />
        </div>

        <div className="tool-op03-toolbar-actions">
          <button
            type="button"
            className="tool-op03-excel-button"
            onClick={openInExcel}
            title="สร้างไฟล์ .xlsx ของ Sheet นี้เพื่อเปิดด้วย Microsoft Excel"
          >
            ▣ เปิดใน Excel
          </button>

          <div className="tool-op03-column-wrapper">
            <button
              type="button"
              className="tool-op03-column-button"
              onClick={() =>
                setShowColumnPicker(
                  (current) =>
                    !current
                )
              }
            >
              ☷ เลือกคอลัมน์ ({visibleColumns.length}/{COLUMN_OPTIONS.length})
            </button>

            {showColumnPicker && (
              <div className="tool-op03-column-picker">
                <div className="tool-op03-column-picker-head">
                  <strong>
                    เลือกคอลัมน์ที่ต้องการแสดง
                  </strong>

                  <div>
                    <button
                      type="button"
                      onClick={() =>
                        setAllColumns(true)
                      }
                    >
                      เลือกทั้งหมด
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        setAllColumns(false)
                      }
                    >
                      ล้างทั้งหมด
                    </button>
                  </div>
                </div>

                <div className="tool-op03-column-picker-grid">
                  {COLUMN_OPTIONS.map(
                    (column) => (
                      <label
                        key={
                          column.key
                        }
                      >
                        <input
                          type="checkbox"
                          checked={
                            isColumnVisible(
                              column.key
                            )
                          }
                          onChange={() =>
                            toggleColumn(
                              column.key
                            )
                          }
                        />
                        {column.label}
                      </label>
                    )
                  )}
                </div>
              </div>
            )}
          </div>

          <span className="tool-op03-result-count">
            แสดง{" "}
            <strong>
              {
                filteredRows.length
              }
            </strong>{" "}
            จาก{" "}
            {
              rows.length
            }{" "}
            รายการ
          </span>
        </div>
      </div>

      {/* FILTERS */}

      <div className="tool-op03-filter-panel">
        <div className="tool-op03-filter-group">
          <strong>
            สถานะ
          </strong>

          <label>
            <input
              type="checkbox"
              checked={
                showDone
              }
              onChange={(
                event
              ) =>
                setShowDone(
                  event
                    .target
                    .checked
                )
              }
            />
            เสร็จแล้ว
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                showSending
              }
              onChange={(
                event
              ) =>
                setShowSending(
                  event
                    .target
                    .checked
                )
              }
            />
            กำลังดำเนินการ
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                showWaiting
              }
              onChange={(
                event
              ) =>
                setShowWaiting(
                  event
                    .target
                    .checked
                )
              }
            />
            รอดำเนินการ
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                showOther
              }
              onChange={(
                event
              ) =>
                setShowOther(
                  event
                    .target
                    .checked
                )
              }
            />
            อื่น ๆ
          </label>
        </div>

        <div className="tool-op03-filter-divider" />

        <div className="tool-op03-filter-group">
          <strong>
            Checklist
          </strong>

          <label>
            <input
              type="checkbox"
              checked={
                requireDGA
              }
              onChange={(
                event
              ) =>
                setRequireDGA(
                  event
                    .target
                    .checked
                )
              }
            />
            DGA
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                require365
              }
              onChange={(
                event
              ) =>
                setRequire365(
                  event
                    .target
                    .checked
                )
              }
            />
            365
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                requireSubmitUser
              }
              onChange={(
                event
              ) =>
                setRequireSubmitUser(
                  event
                    .target
                    .checked
                )
              }
            />
            User ส่งเมล
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                requireTemplate
              }
              onChange={(
                event
              ) =>
                setRequireTemplate(
                  event
                    .target
                    .checked
                )
              }
            />
            Template
          </label>
        </div>
      </div>

      {/* SUMMARY */}

      <div className="tool-op03-summary-grid">
        <article>
          <div className="tool-op03-summary-label">
            หน่วยงานทั้งหมด
          </div>

          <strong>
            {
              summary.organizations
            }
          </strong>

          <small>
            หน่วยงาน
          </small>
        </article>

        <article>
          <div className="tool-op03-summary-label">
            ดำเนินการเสร็จแล้ว
          </div>

          <strong>
            {
              summary.done
            }
          </strong>

          <small className="done">
            เสร็จสิ้น
          </small>
        </article>

        <article>
          <div className="tool-op03-summary-label">
            กำลังดำเนินการ
          </div>

          <strong>
            {
              summary.sending
            }
          </strong>

          <small className="sending">
            อยู่ระหว่างดำเนินงาน
          </small>
        </article>

        <article>
          <div className="tool-op03-summary-label">
            รอดำเนินการ
          </div>

          <strong>
            {
              summary.waiting
            }
          </strong>

          <small className="waiting">
            ต้องติดตาม
          </small>
        </article>

        <article>
          <div className="tool-op03-summary-label">
            DGA พร้อม
          </div>

          <strong>
            {
              summary.dga
            }
          </strong>

          <small>
            Checklist
          </small>
        </article>

        <article>
          <div className="tool-op03-summary-label">
            Template พร้อม
          </div>

          <strong>
            {
              summary.template
            }
          </strong>

          <small>
            Checklist
          </small>
        </article>
      </div>

      {/* MAIN */}

      <section className="tool-op03-main-card">
        <div className="tool-op03-card-header">
          <div>
            <h3>
              ภาพรวมการติดตาม OP03
            </h3>

            <p>
              ตรวจสอบ Checklist สถานะการดำเนินงาน และข้อมูลติดต่อ
            </p>
          </div>

          <div className="tool-op03-tabs">
            {(
              [
                [
                  "table",
                  "TABLE",
                ],
                [
                  "bar",
                  "BAR",
                ],
                [
                  "pie",
                  "PIE",
                ],
              ] as const
            ).map(
              ([
                key,
                label,
              ]) => (
                <button
                  key={
                    key
                  }
                  type="button"
                  className={
                    viewType ===
                    key
                      ? "active"
                      : ""
                  }
                  onClick={() =>
                    setViewType(
                      key
                    )
                  }
                >
                  {
                    label
                  }
                </button>
              )
            )}
          </div>
        </div>

        {/* TABLE */}

        {viewType ===
          "table" && (
          <div className="tool-op03-table-wrapper">
            {visibleColumns.length === 0 ? (
              <div className="tool-op03-no-columns">
                กรุณาเลือกอย่างน้อย 1 คอลัมน์
              </div>
            ) : (
              <table
                className="tool-op03-table"
                style={{
                  minWidth:
                    `${Math.max(
                      760,
                      visibleColumns.length *
                        145
                    )}px`,
                }}
              >
                <thead>
                  <tr>
                    {isColumnVisible(
                      "order"
                    ) && (
                      <th rowSpan={2}>
                        ลำดับ
                      </th>
                    )}

                    {isColumnVisible(
                      "organization"
                    ) && (
                      <th rowSpan={2}>
                        ชื่อหน่วยงาน
                      </th>
                    )}

                    {visibleChecklistCount >
                      0 && (
                      <th
                        colSpan={
                          visibleChecklistCount
                        }
                        className="checklist-head"
                      >
                        CHECKLIST
                      </th>
                    )}

                    {isColumnVisible(
                      "schedule"
                    ) && (
                      <th rowSpan={2}>
                        กำหนดการ
                      </th>
                    )}

                    {isColumnVisible(
                      "note"
                    ) && (
                      <th rowSpan={2}>
                        หมายเหตุ
                      </th>
                    )}

                    {isColumnVisible(
                      "contact"
                    ) && (
                      <th rowSpan={2}>
                        ข้อมูลติดต่อ
                      </th>
                    )}
                  </tr>

                  <tr>
                    {isColumnVisible(
                      "allowDGA"
                    ) && (
                      <th>DGA</th>
                    )}

                    {isColumnVisible(
                      "allow365"
                    ) && (
                      <th>365</th>
                    )}

                    {isColumnVisible(
                      "submitUser"
                    ) && (
                      <th>
                        User ส่งเมล
                      </th>
                    )}

                    {isColumnVisible(
                      "template"
                    ) && (
                      <th>
                        Template
                      </th>
                    )}
                  </tr>
                </thead>

                <tbody>
                  {filteredRows.length ===
                  0 ? (
                    <tr>
                      <td
                        colSpan={
                          Math.max(
                            1,
                            visibleColumns.length
                          )
                        }
                        className="tool-op03-empty"
                      >
                        ไม่พบข้อมูลที่ตรงกับตัวกรอง
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map(
                      (
                        row,
                        index
                      ) => {
                        const previous =
                          index > 0
                            ? filteredRows[
                                index -
                                  1
                              ]
                            : null;

                        const sameOrg =
                          previous &&
                          normalize(
                            previous.organization
                          ) ===
                            normalize(
                              row.organization
                            );

                        const status =
                          getStatusType(
                            row.note
                          );

                        return (
                          <tr
                            key={
                              row.sourceRow
                            }
                          >
                            {isColumnVisible(
                              "order"
                            ) && (
                              <td className="tool-op03-order">
                                {sameOrg
                                  ? ""
                                  : row.order ||
                                    "-"}
                              </td>
                            )}

                            {isColumnVisible(
                              "organization"
                            ) && (
                              <td className="tool-op03-org">
                                {sameOrg
                                  ? ""
                                  : row.organization ||
                                    "-"}
                              </td>
                            )}

                            {isColumnVisible(
                              "allowDGA"
                            ) && (
                              <td className="tool-op03-check">
                                <CheckIcon
                                  checked={
                                    row.allowDGA
                                  }
                                />
                              </td>
                            )}

                            {isColumnVisible(
                              "allow365"
                            ) && (
                              <td className="tool-op03-check">
                                <CheckIcon
                                  checked={
                                    row.allow365
                                  }
                                />
                              </td>
                            )}

                            {isColumnVisible(
                              "submitUser"
                            ) && (
                              <td className="tool-op03-check">
                                <CheckIcon
                                  checked={
                                    row.submitUser
                                  }
                                />
                              </td>
                            )}

                            {isColumnVisible(
                              "template"
                            ) && (
                              <td className="tool-op03-check">
                                <CheckIcon
                                  checked={
                                    row.template
                                  }
                                />
                              </td>
                            )}

                            {isColumnVisible(
                              "schedule"
                            ) && (
                              <td className="tool-op03-schedule">
                                {row.schedule ||
                                  "-"}
                              </td>
                            )}

                            {isColumnVisible(
                              "note"
                            ) && (
                              <td>
                                <span
                                  className={`tool-op03-status ${status}`}
                                >
                                  {row.note ||
                                    statusLabel(
                                      status
                                    )}
                                </span>
                              </td>
                            )}

                            {isColumnVisible(
                              "contact"
                            ) && (
                              <td className="tool-op03-contact">
                                {row.contact ? (
                                  <div>
                                    {row.contact
                                      .split(
                                        /\n|<br\s*\/?>/i
                                      )
                                      .map(
                                        (
                                          line,
                                          lineIndex
                                        ) => (
                                          <span
                                            key={
                                              lineIndex
                                            }
                                          >
                                            {
                                              line
                                            }
                                          </span>
                                        )
                                      )}
                                  </div>
                                ) : (
                                  "-"
                                )}
                              </td>
                            )}
                          </tr>
                        );
                      }
                    )
                  )}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* BAR */}

        {viewType ===
          "bar" && (
          <div className="tool-op03-chart-area">
            <div className="tool-op03-bar-chart">
              {barItems.map(
                (
                  item
                ) => {
                  const height =
                    (
                      item.value /
                      maxBarValue
                    ) *
                    100;

                  return (
                    <div
                      key={
                        item.label
                      }
                      className="tool-op03-bar-item"
                    >
                      <strong>
                        {
                          item.value
                        }
                      </strong>

                      <div className="tool-op03-bar-track">
                        <div
                          className="tool-op03-bar-fill"
                          style={{
                            height:
                              `${height}%`,
                          }}
                        />
                      </div>

                      <span>
                        {
                          item.label
                        }
                      </span>
                    </div>
                  );
                }
              )}
            </div>
          </div>
        )}

        {/* PIE */}

        {viewType ===
          "pie" && (
          <div className="tool-op03-pie-area">
            <div
              className="tool-op03-pie"
              style={{
                background:
                  pieBackground,
              }}
            >
              <div className="tool-op03-pie-hole">
                <strong>
                  {
                    summary.done
                  }
                </strong>

                <span>
                  เสร็จแล้ว
                </span>
              </div>
            </div>

            <div className="tool-op03-pie-legend">
              <span>
                <i className="done" />
                เสร็จแล้ว{" "}
                {
                  summary.done
                }
              </span>

              <span>
                <i className="sending" />
                กำลังดำเนินการ{" "}
                {
                  summary.sending
                }
              </span>

              <span>
                <i className="waiting" />
                รอดำเนินการ{" "}
                {
                  summary.waiting
                }
              </span>
            </div>
          </div>
        )}
      </section>
    </section>
  );
}

/* ==============================
   CHECK ICON
================================ */

function CheckIcon({
  checked,
}: {
  checked: boolean;
}) {
  return (
    <span
      className={`tool-op03-check-icon ${
        checked
          ? "checked"
          : "unchecked"
      }`}
    >
      {
        checked
          ? "✓"
          : ""
      }
    </span>
  );
}

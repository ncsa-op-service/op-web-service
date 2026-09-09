"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./email-service.css";

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

type EmailServiceRow = {
  sourceRow: number;

  order: string;
  organization: string;

  ncsaCoordinator: string;
  organizationCoordinator: string;

  email: string;
  phone: string;

  contactDate: string;
  details: string;

  letterTo: string;

  confirmedParticipation: boolean;
  responseLetter: boolean;
  sentLetterEmail: boolean;

  invitationSentDate: string;

  followUpEmail: boolean;
  requestFormReceived: boolean;

  kickOffDate: string;
  sentKickOffEmail: boolean;

  toolName: string;
  note: string;
};

type ViewType =
  | "table"
  | "progress"
  | "tool";

type ColumnKey =
  | "order" | "organization" | "ncsaCoordinator" | "organizationCoordinator"
  | "email" | "phone" | "contactDate" | "details" | "letterTo"
  | "confirmedParticipation" | "responseLetter" | "sentLetterEmail"
  | "invitationSentDate" | "followUpEmail" | "requestFormReceived"
  | "kickOffDate" | "sentKickOffEmail" | "toolName" | "note";

const COLUMN_OPTIONS: { key: ColumnKey; label: string; process?: boolean }[] = [
  { key: "order", label: "ลำดับ" },
  { key: "organization", label: "ชื่อหน่วยงาน" },
  { key: "ncsaCoordinator", label: "ผู้ประสานงาน สกมช." },
  { key: "organizationCoordinator", label: "ผู้ประสานงานหน่วยงาน" },
  { key: "email", label: "อีเมล" },
  { key: "phone", label: "เบอร์โทร" },
  { key: "contactDate", label: "วันที่ติดต่อ" },
  { key: "details", label: "รายละเอียดอื่นๆ" },
  { key: "letterTo", label: "หนังสือเรียนใคร" },
  { key: "confirmedParticipation", label: "ยืนยันเข้าร่วม", process: true },
  { key: "responseLetter", label: "หนังสือตอบรับ", process: true },
  { key: "sentLetterEmail", label: "ส่งหนังสือ", process: true },
  { key: "invitationSentDate", label: "วันส่งหนังสือ", process: true },
  { key: "followUpEmail", label: "ติดตามอีเมล", process: true },
  { key: "requestFormReceived", label: "ได้รับแบบคำขอ", process: true },
  { key: "kickOffDate", label: "วัน Kick-off", process: true },
  { key: "sentKickOffEmail", label: "ส่ง Kick-off", process: true },
  { key: "toolName", label: "Tool" },
  { key: "note", label: "หมายเหตุ" },
];

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

function buildMatrix(
  worksheet: XLSX.WorkSheet
): MatrixRow[] {
  const ref = worksheet["!ref"];

  if (!ref) {
    return [];
  }

  const range =
    XLSX.utils.decode_range(ref);

  const matrix: MatrixRow[] = [];

  for (
    let r = range.s.r;
    r <= range.e.r;
    r += 1
  ) {
    const row: MatrixRow = [];

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

    matrix[r] = row;
  }

  for (
    const merge of
      worksheet["!merges"] ?? []
  ) {
    const sourceValue =
      matrix[merge.s.r]?.[
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

function findHeaderRow(
  matrix: MatrixRow[]
) {
  const limit =
    Math.min(
      matrix.length,
      20
    );

  for (
    let r = 0;
    r < limit;
    r += 1
  ) {
    const text =
      (matrix[r] ?? [])
        .map(normalize)
        .join(" ");

    const hasOrder =
      text.includes("ลำดับ");

    const hasOrganization =
      text.includes("ชื่อหน่วยงาน");

    const hasEmail =
      text.includes("อีเมล");

    const hasPhone =
      text.includes("เบอร์โทรติดต่อ");

    if (
      hasOrder &&
      hasOrganization &&
      hasEmail &&
      hasPhone
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
  for (
    let c = 0;
    c < row.length;
    c += 1
  ) {
    const value =
      normalize(row[c]);

    if (
      candidates.some(
        (candidate) =>
          value.includes(
            normalize(candidate)
          )
      )
    ) {
      return c;
    }
  }

  return -1;
}

function fallbackColumn(
  detected: number,
  fixedIndex: number
) {
  return detected >= 0
    ? detected
    : fixedIndex;
}

function parseBooleanCell(
  value: unknown
) {
  if (
    typeof value === "boolean"
  ) {
    return value;
  }

  if (
    typeof value === "number"
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
      text.includes(token)
  );
}

function progressOf(
  row: EmailServiceRow
) {
  const steps = [
    row.confirmedParticipation,
    row.responseLetter,
    row.sentLetterEmail,
    row.followUpEmail,
    row.requestFormReceived,
    Boolean(row.kickOffDate),
    row.sentKickOffEmail,
  ];

  const completed =
    steps.filter(Boolean).length;

  return Math.round(
    (completed / steps.length) *
      100
  );
}

function CheckIcon({
  checked,
}: {
  checked: boolean;
}) {
  return (
    <span
      className={`email-service-check-icon ${
        checked
          ? "checked"
          : "unchecked"
      }`}
    >
      {checked ? "✓" : ""}
    </span>
  );
}

function Step({
  label,
  done,
}: {
  label: string;
  done: boolean;
}) {
  return (
    <span
      className={`email-service-step ${
        done ? "done" : ""
      }`}
    >
      <i>
        {done ? "✓" : ""}
      </i>
      {label}
    </span>
  );
}

export default function EmailService({
  worksheet,
  sheetName,
}: Props) {
  const [search, setSearch] =
    useState("");

  const [showColumnPicker, setShowColumnPicker] = useState(false);
  const [visibleColumns, setVisibleColumns] = useState<Record<ColumnKey, boolean>>(() =>
    Object.fromEntries(COLUMN_OPTIONS.map(({ key }) => [key, true])) as Record<ColumnKey, boolean>
  );

  const visibleColumnCount = COLUMN_OPTIONS.filter(({ key }) => visibleColumns[key]).length;
  const visibleProcessCount = COLUMN_OPTIONS.filter(
    ({ key, process }) => process && visibleColumns[key]
  ).length;

  function toggleColumn(key: ColumnKey) {
    setVisibleColumns((current) => ({ ...current, [key]: !current[key] }));
  }

  function setAllColumns(value: boolean) {
    setVisibleColumns(
      Object.fromEntries(COLUMN_OPTIONS.map(({ key }) => [key, value])) as Record<ColumnKey, boolean>
    );
  }

  const [
    viewType,
    setViewType,
  ] =
    useState<ViewType>(
      "table"
    );

  const [
    onlyConfirmed,
    setOnlyConfirmed,
  ] =
    useState(false);

  const [
    onlyResponse,
    setOnlyResponse,
  ] =
    useState(false);

  const [
    onlyRequestForm,
    setOnlyRequestForm,
  ] =
    useState(false);

  const [
    onlyKickOff,
    setOnlyKickOff,
  ] =
    useState(false);

  const [
    onlyPending,
    setOnlyPending,
  ] =
    useState(false);

  const matrix =
    useMemo(
      () =>
        buildMatrix(
          worksheet
        ),
      [worksheet]
    );

  const headerRowIndex =
    useMemo(() => {
      const detected =
        findHeaderRow(
          matrix
        );

      /*
        Sheet "จากพี่นุ บริการอีเมล"
        ใช้หัวตารางอยู่ช่วงแถว 1-2
        ถ้าหา Header จากข้อความไม่เจอ
        ให้ใช้แถวแรกเป็น fallback
      */
      return detected >= 0
        ? detected
        : 0;
    }, [matrix]);

  const rows =
    useMemo<
      EmailServiceRow[]
    >(() => {
      if (
        headerRowIndex < 0
      ) {
        return [];
      }

      const headerRow =
        matrix[
          headerRowIndex
        ] ?? [];

      /*
        โครงสร้างจริงของ Sheet จากพี่นุ บริการอีเมล

        A  = ลำดับ
        B  = ชื่อหน่วยงาน
        C  = ชื่อผู้ประสานงาน (สกมช.)
        D  = ชื่อผู้ประสานงาน (หน่วยงาน)
        E  = อีเมล
        F  = เบอร์โทรติดต่อ
        G  = วันที่ติดต่อประสานงาน
        H  = รายละเอียดอื่นๆ
        I  = หนังสือ เรียนใคร ?
        J  = ยืนยันการเข้าร่วมโครงการ
        K  = หนังสือตอบรับ
        L  = ส่งอีเมล หนังสือ
        M  = วันส่งหนังสือเชิญฯ ทางอีเมล
        N  = ติดตามจนกว่าจะได้รับอีเมล
        O  = ได้รับแบบคำขอรับการทดสอบ
        P  = วันนัดหมายประชุม kick-off
        Q  = ส่งอีเมล kick off
        R  = Tool ที่หน่วยงานขอทดสอบ
        S  = หมายเหตุ
      */

      const orderColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            ["ลำดับ"]
          ),
          0
        );

      const organizationColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            ["ชื่อหน่วยงาน"]
          ),
          1
        );

      const ncsaCoordinatorColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "ชื่อผู้ประสานงาน (สกมช.)",
              "ชื่อผู้ประสานงาน สกมช",
              "สกมช",
            ]
          ),
          2
        );

      const organizationCoordinatorColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "ชื่อผู้ประสานงาน (หน่วยงาน)",
              "ชื่อผู้ประสานงาน หน่วยงาน",
            ]
          ),
          3
        );

      const emailColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            ["อีเมล"]
          ),
          4
        );

      const phoneColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            ["เบอร์โทรติดต่อ"]
          ),
          5
        );

      const contactDateColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "วันที่ติดต่อประสานงาน",
            ]
          ),
          6
        );

      const detailsColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "รายละเอียดอื่นๆ",
              "รายละเอียดอื่น",
            ]
          ),
          7
        );

      const letterToColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "หนังสือ เรียนใคร",
              "หนังสือเรียนใคร",
            ]
          ),
          8
        );

      const confirmedColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "ยืนยันการเข้าร่วมโครงการ",
              "ยันยันการเข้าร่วมโครงการ",
            ]
          ),
          9
        );

      const responseColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            ["หนังสือตอบรับ"]
          ),
          10
        );

      const sentLetterEmailColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "ส่งอีเมล หนังสือ",
              "ส่งอีเมลหนังสือ",
            ]
          ),
          11
        );

      const invitationSentDateColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            ["วันส่งหนังสือเชิญ"]
          ),
          12
        );

      const followUpColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "ติดตามจนกว่าจะได้รับอีเมล",
            ]
          ),
          13
        );

      const requestFormColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "ได้รับแบบคำขอรับการทดสอบ",
            ]
          ),
          14
        );

      const kickOffDateColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "วันนัดหมายประชุม kick-off",
              "วันนัดหมายประชุม kick off",
            ]
          ),
          15
        );

      const sentKickOffColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "ส่งอีเมล kick off",
              "ส่งอีเมล kick-off",
            ]
          ),
          16
        );

      const toolColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            [
              "tool ที่หน่วยงานขอทดสอบ",
              "tool",
            ]
          ),
          17
        );

      const noteColumn =
        fallbackColumn(
          findColumn(
            headerRow,
            ["หมายเหตุ"]
          ),
          18
        );

      const result:
        EmailServiceRow[] =
        [];

      /*
        ในไฟล์จริงหัวตารางกิน 2 แถว
        ข้อมูลเริ่มที่ Excel row 3
      */
      const firstDataRow =
        Math.max(
          headerRowIndex + 1,
          2
        );

      for (
        let r =
          firstDataRow;
        r <
        matrix.length;
        r += 1
      ) {
        const source =
          matrix[r] ?? [];

        const row: EmailServiceRow = {
          sourceRow:
            r + 1,

          order:
            orderColumn >= 0
              ? cleanText(
                  source[
                    orderColumn
                  ]
                )
              : "",

          organization:
            organizationColumn >= 0
              ? cleanText(
                  source[
                    organizationColumn
                  ]
                )
              : "",

          ncsaCoordinator:
            ncsaCoordinatorColumn >= 0
              ? cleanText(
                  source[
                    ncsaCoordinatorColumn
                  ]
                )
              : "",

          organizationCoordinator:
            organizationCoordinatorColumn >= 0
              ? cleanText(
                  source[
                    organizationCoordinatorColumn
                  ]
                )
              : "",

          email:
            emailColumn >= 0
              ? cleanText(
                  source[
                    emailColumn
                  ]
                )
              : "",

          phone:
            phoneColumn >= 0
              ? cleanText(
                  source[
                    phoneColumn
                  ]
                )
              : "",

          contactDate:
            contactDateColumn >= 0
              ? cleanText(
                  source[
                    contactDateColumn
                  ]
                )
              : "",

          details:
            detailsColumn >= 0
              ? cleanText(
                  source[
                    detailsColumn
                  ]
                )
              : "",

          letterTo:
            letterToColumn >= 0
              ? cleanText(
                  source[
                    letterToColumn
                  ]
                )
              : "",

          confirmedParticipation:
            confirmedColumn >= 0
              ? parseBooleanCell(
                  source[
                    confirmedColumn
                  ]
                )
              : false,

          responseLetter:
            responseColumn >= 0
              ? parseBooleanCell(
                  source[
                    responseColumn
                  ]
                )
              : false,

          sentLetterEmail:
            sentLetterEmailColumn >= 0
              ? parseBooleanCell(
                  source[
                    sentLetterEmailColumn
                  ]
                )
              : false,

          invitationSentDate:
            invitationSentDateColumn >= 0
              ? cleanText(
                  source[
                    invitationSentDateColumn
                  ]
                )
              : "",

          followUpEmail:
            followUpColumn >= 0
              ? parseBooleanCell(
                  source[
                    followUpColumn
                  ]
                )
              : false,

          requestFormReceived:
            requestFormColumn >= 0
              ? parseBooleanCell(
                  source[
                    requestFormColumn
                  ]
                )
              : false,

          kickOffDate:
            kickOffDateColumn >= 0
              ? cleanText(
                  source[
                    kickOffDateColumn
                  ]
                )
              : "",

          sentKickOffEmail:
            sentKickOffColumn >= 0
              ? parseBooleanCell(
                  source[
                    sentKickOffColumn
                  ]
                )
              : false,

          toolName:
            toolColumn >= 0
              ? cleanText(
                  source[
                    toolColumn
                  ]
                )
              : "",

          note:
            noteColumn >= 0
              ? cleanText(
                  source[
                    noteColumn
                  ]
                )
              : "",
        };

        const hasAnyData =
          row.order ||
          row.organization ||
          row.ncsaCoordinator ||
          row.organizationCoordinator ||
          row.email ||
          row.phone ||
          row.contactDate ||
          row.details ||
          row.letterTo ||
          row.confirmedParticipation ||
          row.responseLetter ||
          row.sentLetterEmail ||
          row.invitationSentDate ||
          row.followUpEmail ||
          row.requestFormReceived ||
          row.kickOffDate ||
          row.sentKickOffEmail ||
          row.toolName ||
          row.note;

        if (
          !hasAnyData
        ) {
          continue;
        }

        if (
          !row.organization &&
          !row.email &&
          !row.note
        ) {
          continue;
        }

        result.push(row);
      }

      return result;
    }, [
      matrix,
      headerRowIndex,
    ]);

  const filteredRows =
    useMemo(() => {
      const keyword =
        normalize(search);

      return rows.filter(
        (row) => {
          const matchesSearch =
            !keyword ||
            [
              row.order,
              row.organization,
              row.ncsaCoordinator,
              row.organizationCoordinator,
              row.email,
              row.phone,
              row.details,
              row.letterTo,
              row.toolName,
              row.note,
            ].some(
              (value) =>
                normalize(value)
                  .includes(
                    keyword
                  )
            );

          if (
            !matchesSearch
          ) {
            return false;
          }

          if (
            onlyConfirmed &&
            !row.confirmedParticipation
          ) {
            return false;
          }

          if (
            onlyResponse &&
            !row.responseLetter
          ) {
            return false;
          }

          if (
            onlyRequestForm &&
            !row.requestFormReceived
          ) {
            return false;
          }

          if (
            onlyKickOff &&
            !row.kickOffDate
          ) {
            return false;
          }

          if (
            onlyPending &&
            progressOf(row) >= 100
          ) {
            return false;
          }

          return true;
        }
      );
    }, [
      rows,
      search,
      onlyConfirmed,
      onlyResponse,
      onlyRequestForm,
      onlyKickOff,
      onlyPending,
    ]);

  const summary =
    useMemo(() => {
      const confirmed =
        rows.filter(
          (row) =>
            row.confirmedParticipation
        ).length;

      const response =
        rows.filter(
          (row) =>
            row.responseLetter
        ).length;

      const requestForm =
        rows.filter(
          (row) =>
            row.requestFormReceived
        ).length;

      const kickOff =
        rows.filter(
          (row) =>
            Boolean(
              row.kickOffDate
            )
        ).length;

      const sentKickOff =
        rows.filter(
          (row) =>
            row.sentKickOffEmail
        ).length;

      const completed =
        rows.filter(
          (row) =>
            progressOf(row) >=
            100
        ).length;

      return {
        total:
          rows.length,
        confirmed,
        response,
        requestForm,
        kickOff,
        sentKickOff,
        completed,
        pending:
          rows.length -
          completed,
      };
    }, [rows]);

  const toolSummary =
    useMemo(() => {
      const counts =
        new Map<
          string,
          number
        >();

      rows.forEach(
        (row) => {
          const name =
            cleanText(
              row.toolName
            );

          if (!name) {
            return;
          }

          counts.set(
            name,
            (
              counts.get(
                name
              ) ?? 0
            ) + 1
          );
        }
      );

      return Array.from(
        counts.entries()
      )
        .map(
          ([
            name,
            value,
          ]) => ({
            name,
            value,
          })
        )
        .sort(
          (a, b) =>
            b.value -
            a.value
        );
    }, [rows]);

  const maxTool =
    Math.max(
      1,
      ...toolSummary.map(
        (item) =>
          item.value
      )
    );

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
        cleanText(sheetName)
          .replace(
            /[\\/:*?"<>|]/g,
            "_"
          )
          .trim() ||
        "email-service";

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

  return (
    <section className="email-service-page">

      <div className="email-service-heading">
        <div>
          <span className="email-service-kicker">
            CLAB / EMAIL SERVICE
          </span>

          <h2>
            {sheetName}
          </h2>

          <p>
            ติดตามขั้นตอนการประสานงาน หนังสือเชิญ แบบคำขอรับการทดสอบ และการนัดหมาย Kick-off
          </p>
        </div>

        <div className="email-service-heading-count">
          <strong>
            {summary.total}
          </strong>

          <span>
            หน่วยงาน
          </span>
        </div>
      </div>

      <div className="email-service-toolbar">
        <div className="email-service-search">
          <span>⌕</span>

          <input
            type="text"
            value={search}
            placeholder="ค้นหาหน่วยงาน / ผู้ประสานงาน / อีเมล / Tool..."
            onChange={(event) =>
              setSearch(
                event.target.value
              )
            }
          />
        </div>

        <span className="email-service-result-count">
          แสดง{" "}
          <strong>
            {filteredRows.length}
          </strong>{" "}
          จาก{" "}
          {rows.length}{" "}
          รายการ
        </span>
      </div>

      <div className="email-service-column-controls">
        <button
          type="button"
          className="email-service-excel-button"
          onClick={openInExcel}
          title="สร้างไฟล์ .xlsx ของ Sheet นี้เพื่อเปิดด้วย Microsoft Excel"
        >
          ▣ เปิดใน Excel
        </button>

        <button
          type="button"
          className="email-service-column-button"
          onClick={() => setShowColumnPicker((value) => !value)}
        >
          ☷ เลือกคอลัมน์ ({visibleColumnCount}/{COLUMN_OPTIONS.length})
        </button>

        {showColumnPicker && (
          <div className="email-service-column-picker">
            <div className="email-service-column-picker-actions">
              <strong>เลือกคอลัมน์ที่ต้องการแสดง</strong>
              <div>
                <button type="button" onClick={() => setAllColumns(true)}>เลือกทั้งหมด</button>
                <button type="button" onClick={() => setAllColumns(false)}>ล้างทั้งหมด</button>
              </div>
            </div>
            <div className="email-service-column-picker-grid">
              {COLUMN_OPTIONS.map((column) => (
                <label key={column.key}>
                  <input
                    type="checkbox"
                    checked={visibleColumns[column.key]}
                    onChange={() => toggleColumn(column.key)}
                  />
                  {column.label}
                </label>
              ))}
            </div>
          </div>
        )}
      </div>

      <div className="email-service-filter-panel">
        <strong>
          ตัวกรองขั้นตอน
        </strong>

        <label>
          <input
            type="checkbox"
            checked={onlyConfirmed}
            onChange={(event) =>
              setOnlyConfirmed(
                event.target.checked
              )
            }
          />
          ยืนยันเข้าร่วมแล้ว
        </label>

        <label>
          <input
            type="checkbox"
            checked={onlyResponse}
            onChange={(event) =>
              setOnlyResponse(
                event.target.checked
              )
            }
          />
          หนังสือตอบรับแล้ว
        </label>

        <label>
          <input
            type="checkbox"
            checked={onlyRequestForm}
            onChange={(event) =>
              setOnlyRequestForm(
                event.target.checked
              )
            }
          />
          ได้รับแบบคำขอทดสอบ
        </label>

        <label>
          <input
            type="checkbox"
            checked={onlyKickOff}
            onChange={(event) =>
              setOnlyKickOff(
                event.target.checked
              )
            }
          />
          นัด Kick-off แล้ว
        </label>

        <label>
          <input
            type="checkbox"
            checked={onlyPending}
            onChange={(event) =>
              setOnlyPending(
                event.target.checked
              )
            }
          />
          ยังไม่ครบทุกขั้นตอน
        </label>
      </div>

      <div className="email-service-summary-grid">
        <article>
          <span>
            หน่วยงานทั้งหมด
          </span>

          <strong>
            {summary.total}
          </strong>

          <small>
            รายการ
          </small>
        </article>

        <article>
          <span>
            ยืนยันเข้าร่วม
          </span>

          <strong>
            {summary.confirmed}
          </strong>

          <small className="positive">
            Confirmed
          </small>
        </article>

        <article>
          <span>
            หนังสือตอบรับ
          </span>

          <strong>
            {summary.response}
          </strong>

          <small>
            Received
          </small>
        </article>

        <article>
          <span>
            แบบคำขอทดสอบ
          </span>

          <strong>
            {summary.requestForm}
          </strong>

          <small>
            Request form
          </small>
        </article>

        <article>
          <span>
            นัด Kick-off
          </span>

          <strong>
            {summary.kickOff}
          </strong>

          <small className="positive">
            Scheduled
          </small>
        </article>

        <article>
          <span>
            รอดำเนินการ
          </span>

          <strong>
            {summary.pending}
          </strong>

          <small className="warning">
            Pending
          </small>
        </article>
      </div>

      <section className="email-service-main-card">
        <div className="email-service-card-header">
          <div>
            <h3>
              ภาพรวมบริการอีเมล
            </h3>

            <p>
              ตรวจสอบข้อมูล ขั้นตอนการดำเนินงาน และ Tool ที่หน่วยงานขอทดสอบ
            </p>
          </div>

          <div className="email-service-tabs">
            {(
              [
                [
                  "table",
                  "TABLE",
                ],
                [
                  "progress",
                  "PROGRESS",
                ],
                [
                  "tool",
                  "TOOL",
                ],
              ] as const
            ).map(
              ([
                key,
                label,
              ]) => (
                <button
                  key={key}
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
                  {label}
                </button>
              )
            )}
          </div>
        </div>

        {viewType ===
          "table" && (
          <div className="email-service-table-wrapper">
            <table className="email-service-table">
              <thead>
                <tr>
                  {visibleColumns.order && <th rowSpan={2}>ลำดับ</th>}
                  {visibleColumns.organization && <th rowSpan={2}>ชื่อหน่วยงาน</th>}
                  {visibleColumns.ncsaCoordinator && <th rowSpan={2}>ผู้ประสานงาน สกมช.</th>}
                  {visibleColumns.organizationCoordinator && <th rowSpan={2}>ผู้ประสานงานหน่วยงาน</th>}
                  {visibleColumns.email && <th rowSpan={2}>อีเมล</th>}
                  {visibleColumns.phone && <th rowSpan={2}>เบอร์โทร</th>}
                  {visibleColumns.contactDate && <th rowSpan={2}>วันที่ติดต่อ</th>}
                  {visibleColumns.details && <th rowSpan={2}>รายละเอียดอื่นๆ</th>}
                  {visibleColumns.letterTo && <th rowSpan={2}>หนังสือเรียนใคร</th>}
                  {visibleProcessCount > 0 && (
                    <th colSpan={visibleProcessCount} className="email-service-process-head">
                      ขั้นตอนดำเนินงาน
                    </th>
                  )}
                  {visibleColumns.toolName && <th rowSpan={2}>Tool</th>}
                  {visibleColumns.note && <th rowSpan={2}>หมายเหตุ</th>}
                </tr>

                <tr>
                  {visibleColumns.confirmedParticipation && <th>ยืนยันเข้าร่วม</th>}
                  {visibleColumns.responseLetter && <th>หนังสือตอบรับ</th>}
                  {visibleColumns.sentLetterEmail && <th>ส่งหนังสือ</th>}
                  {visibleColumns.invitationSentDate && <th>วันส่งหนังสือ</th>}
                  {visibleColumns.followUpEmail && <th>ติดตามอีเมล</th>}
                  {visibleColumns.requestFormReceived && <th>ได้รับแบบคำขอ</th>}
                  {visibleColumns.kickOffDate && <th>วัน Kick-off</th>}
                  {visibleColumns.sentKickOffEmail && <th>ส่ง Kick-off</th>}
                </tr>
              </thead>

              <tbody>
                {filteredRows.length === 0 ? (
                  <tr>
                    <td colSpan={Math.max(1, visibleColumnCount)} className="email-service-empty">
                      ไม่พบข้อมูลที่ตรงกับตัวกรอง
                    </td>
                  </tr>
                ) : visibleColumnCount === 0 ? (
                  <tr>
                    <td className="email-service-empty">กรุณาเลือกอย่างน้อย 1 คอลัมน์</td>
                  </tr>
                ) : (
                  filteredRows.map((row) => (
                    <tr key={row.sourceRow}>
                      {visibleColumns.order && <td className="email-service-order">{row.order || "-"}</td>}
                      {visibleColumns.organization && <td className="email-service-org">{row.organization || "-"}</td>}
                      {visibleColumns.ncsaCoordinator && <td>{row.ncsaCoordinator || "-"}</td>}
                      {visibleColumns.organizationCoordinator && <td>{row.organizationCoordinator || "-"}</td>}
                      {visibleColumns.email && <td className="email-service-email">{row.email || "-"}</td>}
                      {visibleColumns.phone && <td>{row.phone || "-"}</td>}
                      {visibleColumns.contactDate && <td className="email-service-date">{row.contactDate || "-"}</td>}
                      {visibleColumns.details && <td>{row.details || "-"}</td>}
                      {visibleColumns.letterTo && <td>{row.letterTo || "-"}</td>}
                      {visibleColumns.confirmedParticipation && <td className="email-service-check"><CheckIcon checked={row.confirmedParticipation} /></td>}
                      {visibleColumns.responseLetter && <td className="email-service-check"><CheckIcon checked={row.responseLetter} /></td>}
                      {visibleColumns.sentLetterEmail && <td className="email-service-check"><CheckIcon checked={row.sentLetterEmail} /></td>}
                      {visibleColumns.invitationSentDate && <td className="email-service-date">{row.invitationSentDate || "-"}</td>}
                      {visibleColumns.followUpEmail && <td className="email-service-check"><CheckIcon checked={row.followUpEmail} /></td>}
                      {visibleColumns.requestFormReceived && <td className="email-service-check"><CheckIcon checked={row.requestFormReceived} /></td>}
                      {visibleColumns.kickOffDate && <td className="email-service-date">{row.kickOffDate || "-"}</td>}
                      {visibleColumns.sentKickOffEmail && <td className="email-service-check"><CheckIcon checked={row.sentKickOffEmail} /></td>}
                      {visibleColumns.toolName && <td className="email-service-tool">{row.toolName || "-"}</td>}
                      {visibleColumns.note && <td>{row.note || "-"}</td>}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {viewType ===
          "progress" && (
          <div className="email-service-progress-list">
            {filteredRows.map(
              (row) => {
                const progress =
                  progressOf(
                    row
                  );

                return (
                  <article
                    key={
                      row.sourceRow
                    }
                    className="email-service-progress-item"
                  >
                    <div className="email-service-progress-heading">
                      <div>
                        <strong>
                          {row.organization ||
                            `รายการ ${row.sourceRow}`}
                        </strong>

                        <span>
                          {row.organizationCoordinator ||
                            row.email ||
                            "ไม่มีข้อมูลผู้ประสานงาน"}
                        </span>
                      </div>

                      <b>
                        {progress}%
                      </b>
                    </div>

                    <div className="email-service-progress-track">
                      <div
                        className="email-service-progress-fill"
                        style={{
                          width:
                            `${progress}%`,
                        }}
                      />
                    </div>

                    <div className="email-service-progress-steps">
                      <Step
                        label="ยืนยัน"
                        done={
                          row.confirmedParticipation
                        }
                      />

                      <Step
                        label="ตอบรับ"
                        done={
                          row.responseLetter
                        }
                      />

                      <Step
                        label="ส่งหนังสือ"
                        done={
                          row.sentLetterEmail
                        }
                      />

                      <Step
                        label="ติดตาม"
                        done={
                          row.followUpEmail
                        }
                      />

                      <Step
                        label="รับคำขอ"
                        done={
                          row.requestFormReceived
                        }
                      />

                      <Step
                        label="นัด Kick-off"
                        done={
                          Boolean(
                            row.kickOffDate
                          )
                        }
                      />

                      <Step
                        label="ส่ง Kick-off"
                        done={
                          row.sentKickOffEmail
                        }
                      />
                    </div>
                  </article>
                );
              }
            )}
          </div>
        )}

        {viewType ===
          "tool" && (
          <div className="email-service-tool-area">
            {toolSummary.length ===
            0 ? (
              <div className="email-service-empty-chart">
                ไม่พบข้อมูล Tool
              </div>
            ) : (
              <div className="email-service-tool-chart">
                {toolSummary.map(
                  (item) => {
                    const height =
                      (
                        item.value /
                        maxTool
                      ) *
                      100;

                    return (
                      <div
                        key={
                          item.name
                        }
                        className="email-service-tool-chart-item"
                      >
                        <strong>
                          {item.value}
                        </strong>

                        <div className="email-service-tool-chart-track">
                          <div
                            className="email-service-tool-chart-fill"
                            style={{
                              height:
                                `${height}%`,
                            }}
                          />
                        </div>

                        <span>
                          {item.name}
                        </span>
                      </div>
                    );
                  }
                )}
              </div>
            )}
          </div>
        )}
      </section>
    </section>
  );
}

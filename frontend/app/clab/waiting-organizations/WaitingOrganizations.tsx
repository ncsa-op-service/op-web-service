"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./waiting-organizations.css";

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

type WaitingRow = {
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

  kickOffDate: string;

  sentKickOffEmail: boolean;

  hcl: boolean;
  picus: boolean;
  proofpoint: boolean;
  seciron: boolean;

  note: string;
};

type ViewType =
  | "table"
  | "progress"
  | "tools";

type ToolFilter =
  | "hcl"
  | "picus"
  | "proofpoint"
  | "seciron";

type ColumnKey =
  | "order"
  | "organization"
  | "ncsaCoordinator"
  | "organizationCoordinator"
  | "email"
  | "phone"
  | "contactDate"
  | "details"
  | "letterTo"
  | "confirmedParticipation"
  | "responseLetter"
  | "sentLetterEmail"
  | "followUpEmail"
  | "kickOffDate"
  | "sentKickOffEmail"
  | "hcl"
  | "picus"
  | "proofpoint"
  | "seciron"
  | "note";

const COLUMN_OPTIONS: {
  key: ColumnKey;
  label: string;
  group?: "process" | "tool";
}[] = [
  { key: "order", label: "ลำดับ" },
  { key: "organization", label: "ชื่อหน่วยงาน" },
  { key: "ncsaCoordinator", label: "ผู้ประสานงาน สกมช." },
  { key: "organizationCoordinator", label: "ผู้ประสานงานหน่วยงาน" },
  { key: "email", label: "อีเมล" },
  { key: "phone", label: "เบอร์โทร" },
  { key: "contactDate", label: "วันที่ติดต่อ" },
  { key: "details", label: "รายละเอียดอื่นๆ" },
  { key: "letterTo", label: "หนังสือเรียนใคร" },

  { key: "confirmedParticipation", label: "ยืนยันเข้าร่วม", group: "process" },
  { key: "responseLetter", label: "หนังสือตอบรับ", group: "process" },
  { key: "sentLetterEmail", label: "ส่งหนังสือทางอีเมล", group: "process" },
  { key: "followUpEmail", label: "ติดตามอีเมล", group: "process" },
  { key: "kickOffDate", label: "วัน Kick-off", group: "process" },
  { key: "sentKickOffEmail", label: "ส่งอีเมล Kick-off", group: "process" },

  { key: "hcl", label: "HCL", group: "tool" },
  { key: "picus", label: "PICUS", group: "tool" },
  { key: "proofpoint", label: "PROOFPOINT", group: "tool" },
  { key: "seciron", label: "SECIRON", group: "tool" },

  { key: "note", label: "หมายเหตุ" },
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
   MATRIX
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
    เติม merged cell เพื่อให้ header ที่ merge
    เช่น Tool ที่หน่วยงานขอทดสอบ ยังหา column ได้
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
  const limit = Math.min(matrix.length, 20);

  for (let r = 0; r < limit; r += 1) {
    const current = (matrix[r] ?? []).map(normalize).join(" ");
    const next = (matrix[r + 1] ?? []).map(normalize).join(" ");
    const next2 = (matrix[r + 2] ?? []).map(normalize).join(" ");
    const combined = `${current} ${next} ${next2}`;

    const score = [
      "ลำดับ",
      "ชื่อหน่วยงาน",
      "อีเมล",
      "เบอร์โทรติดต่อ",
      "วันที่ติดต่อประสานงาน",
      "หนังสือตอบรับ",
      "kick-off",
    ].filter((label) => combined.includes(normalize(label))).length;

    if (score >= 4) {
      const subHeaderText = `${next} ${next2}`;
      const hasToolSubHeader =
        subHeaderText.includes("hcl") ||
        subHeaderText.includes("picus") ||
        subHeaderText.includes("proofpoint") ||
        subHeaderText.includes("seciron");

      return {
        top: r,
        bottom: hasToolSubHeader ? r + 1 : r,
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
    r <= bottom;
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
  candidates: string[]
) {
  const maxColumns =
    Math.max(
      ...matrix
        .slice(
          top,
          bottom + 1
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
    const text =
      combinedHeaderText(
        matrix,
        top,
        bottom,
        c
      );

    if (
      candidates.some(
        (candidate) =>
          text.includes(
            normalize(
              candidate
            )
          )
      )
    ) {
      return c;
    }
  }

  return -1;
}

function findColumnByAll(
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
          bottom + 1
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
    const text =
      combinedHeaderText(
        matrix,
        top,
        bottom,
        c
      );

    if (
      keywords.every(
        (keyword) =>
          text.includes(
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
   BOOLEAN
================================ */

function parseBooleanCell(
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
   COMPONENT
================================ */

export default function WaitingOrganizations({
  worksheet,
  sheetName,
}: Props) {
  const [
    search,
    setSearch,
  ] =
    useState("");

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

  const visibleProcessCount =
    COLUMN_OPTIONS.filter(
      ({ key, group }) =>
        group === "process" &&
        visibleColumns.includes(key)
    ).length;

  const visibleToolCount =
    COLUMN_OPTIONS.filter(
      ({ key, group }) =>
        group === "tool" &&
        visibleColumns.includes(key)
    ).length;

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
    onlyKickOff,
    setOnlyKickOff,
  ] =
    useState(false);

  const [
    onlyPending,
    setOnlyPending,
  ] =
    useState(false);

  const [
    toolFilters,
    setToolFilters,
  ] =
    useState<
      ToolFilter[]
    >([]);

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

  const rows =
    useMemo<
      WaitingRow[]
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

      const orderColumn = (() => {
        const found = findColumn(matrix, top, bottom, ["ลำดับ"]);
        return found >= 0 ? found : 0;
      })();

      const organizationColumn = (() => {
        const found = findColumn(matrix, top, bottom, ["ชื่อหน่วยงาน"]);
        return found >= 0 ? found : 1;
      })();

      const ncsaCoordinatorColumn = (() => {
        const found = findColumnByAll(matrix, top, bottom, ["ชื่อผู้ประสานงาน", "สกมช"]);
        return found >= 0 ? found : 2;
      })();

      const organizationCoordinatorColumn = (() => {
        const found = findColumnByAll(matrix, top, bottom, ["ชื่อผู้ประสานงาน", "หน่วยงาน"]);
        return found >= 0 ? found : 3;
      })();

      const emailColumn = (() => {
        const found = findColumn(matrix, top, bottom, ["อีเมล"]);
        return found >= 0 ? found : 4;
      })();

      const phoneColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "เบอร์โทรติดต่อ",
          ]);
        return found >= 0 ? found : 5;
      })();

      const contactDateColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "วันที่ติดต่อประสานงาน",
          ]);
        return found >= 0 ? found : 6;
      })();

      const detailsColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "รายละเอียดอื่นๆ",
            "รายละเอียดอื่น",
          ]);
        return found >= 0 ? found : 7;
      })();

      const letterToColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "หนังสือ เรียนใคร",
            "หนังสือเรียนใคร",
          ]);
        return found >= 0 ? found : 8;
      })();

      const confirmedColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "ยืนยันการเข้าร่วมโครงการ",
            "ยันยันการเข้าร่วมโครงการ",
          ]);
        return found >= 0 ? found : 9;
      })();

      const responseLetterColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "หนังสือตอบรับ",
          ]);
        return found >= 0 ? found : 10;
      })();

      const sentLetterEmailColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "ส่งอีเมล หนังสือ",
          ]);
        return found >= 0 ? found : 11;
      })();

      const invitationSentDateColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "วันส่งหนังสือเชิญ",
          ]);
        return found >= 0 ? found : 12;
      })();

      const followUpColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "ติดตามจนกว่าจะได้รับอีเมล",
          ]);
        return found >= 0 ? found : 13;
      })();

      const kickOffDateColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "วันนัดหมายประชุม kick-off",
            "วันนัดหมายประชุม kick off",
          ]);
        return found >= 0 ? found : 14;
      })();

      const sentKickOffColumn = (() => {
        const found = findColumn(matrix, top, bottom, [
            "ส่งอีเมล kick off",
            "ส่งอีเมล kick-off",
          ]);
        return found >= 0 ? found : 15;
      })();

      const hclColumn = (() => {
        const byAll = findColumnByAll(matrix, top, bottom, ["tool", "hcl"]);
        if (byAll >= 0) return byAll;
        const direct = findColumn(matrix, top, bottom, ["hcl"]);
        return direct >= 0 ? direct : 16;
      })();

      const picusColumn = (() => {
        const byAll = findColumnByAll(matrix, top, bottom, ["tool", "picus"]);
        if (byAll >= 0) return byAll;
        const direct = findColumn(matrix, top, bottom, ["picus"]);
        return direct >= 0 ? direct : 17;
      })();

      const proofpointColumn = (() => {
        const byAll = findColumnByAll(matrix, top, bottom, ["tool", "proofpoint"]);
        if (byAll >= 0) return byAll;
        const direct = findColumn(matrix, top, bottom, ["proofpoint"]);
        return direct >= 0 ? direct : 18;
      })();

      const secironColumn = (() => {
        const byAll = findColumnByAll(matrix, top, bottom, ["tool", "seciron"]);
        if (byAll >= 0) return byAll;
        const direct = findColumn(matrix, top, bottom, ["seciron"]);
        return direct >= 0 ? direct : 19;
      })();

      const noteColumn = (() => {
        const found = findColumn(matrix, top, bottom, ["หมายเหตุ"]);
        return found >= 0 ? found : 20;
      })();

      const result:
        WaitingRow[] =
        [];

      const firstDataRow =
        bottom + 1;

      for (
        let r =
          firstDataRow;
        r <
        matrix.length;
        r += 1
      ) {
        const source =
          matrix[r] ?? [];

        const organization =
          organizationColumn >=
          0
            ? cleanText(
                source[
                  organizationColumn
                ]
              )
            : "";

        const order =
          orderColumn >= 0
            ? cleanText(
                source[
                  orderColumn
                ]
              )
            : "";

        const ncsaCoordinator =
          ncsaCoordinatorColumn >=
          0
            ? cleanText(
                source[
                  ncsaCoordinatorColumn
                ]
              )
            : "";

        const organizationCoordinator =
          organizationCoordinatorColumn >=
          0
            ? cleanText(
                source[
                  organizationCoordinatorColumn
                ]
              )
            : "";

        const email =
          emailColumn >= 0
            ? cleanText(
                source[
                  emailColumn
                ]
              )
            : "";

        const phone =
          phoneColumn >= 0
            ? cleanText(
                source[
                  phoneColumn
                ]
              )
            : "";

        const contactDate =
          contactDateColumn >= 0
            ? cleanText(
                source[
                  contactDateColumn
                ]
              )
            : "";

        const details =
          detailsColumn >= 0
            ? cleanText(
                source[
                  detailsColumn
                ]
              )
            : "";

        const letterTo =
          letterToColumn >= 0
            ? cleanText(
                source[
                  letterToColumn
                ]
              )
            : "";

        const confirmedParticipation =
          confirmedColumn >= 0
            ? parseBooleanCell(
                source[
                  confirmedColumn
                ]
              )
            : false;

        const responseLetter =
          responseLetterColumn >= 0
            ? parseBooleanCell(
                source[
                  responseLetterColumn
                ]
              )
            : false;

        const sentLetterEmail =
          sentLetterEmailColumn >= 0
            ? parseBooleanCell(
                source[
                  sentLetterEmailColumn
                ]
              )
            : false;

        const invitationSentDate =
          invitationSentDateColumn >=
          0
            ? cleanText(
                source[
                  invitationSentDateColumn
                ]
              )
            : "";

        const followUpEmail =
          followUpColumn >= 0
            ? parseBooleanCell(
                source[
                  followUpColumn
                ]
              )
            : false;

        const kickOffDate =
          kickOffDateColumn >= 0
            ? cleanText(
                source[
                  kickOffDateColumn
                ]
              )
            : "";

        const sentKickOffEmail =
          sentKickOffColumn >= 0
            ? parseBooleanCell(
                source[
                  sentKickOffColumn
                ]
              )
            : false;

        const hcl =
          hclColumn >= 0
            ? parseBooleanCell(
                source[
                  hclColumn
                ]
              )
            : false;

        const picus =
          picusColumn >= 0
            ? parseBooleanCell(
                source[
                  picusColumn
                ]
              )
            : false;

        const proofpoint =
          proofpointColumn >=
          0
            ? parseBooleanCell(
                source[
                  proofpointColumn
                ]
              )
            : false;

        const seciron =
          secironColumn >= 0
            ? parseBooleanCell(
                source[
                  secironColumn
                ]
              )
            : false;

        const note =
          noteColumn >= 0
            ? cleanText(
                source[
                  noteColumn
                ]
              )
            : "";

        const hasAnyData =
          order ||
          organization ||
          ncsaCoordinator ||
          organizationCoordinator ||
          email ||
          phone ||
          contactDate ||
          details ||
          letterTo ||
          confirmedParticipation ||
          responseLetter ||
          sentLetterEmail ||
          invitationSentDate ||
          followUpEmail ||
          kickOffDate ||
          sentKickOffEmail ||
          hcl ||
          picus ||
          proofpoint ||
          seciron ||
          note;

        if (
          !hasAnyData
        ) {
          continue;
        }

        /*
          ป้องกันแถวหัว/แถวว่างท้าย Sheet
        */
        if (
          !organization &&
          !email &&
          !phone &&
          !note
        ) {
          continue;
        }

        result.push({
          sourceRow:
            r + 1,

          order,
          organization,

          ncsaCoordinator,
          organizationCoordinator,

          email,
          phone,

          contactDate,
          details,

          letterTo,

          confirmedParticipation,
          responseLetter,
          sentLetterEmail,

          invitationSentDate,

          followUpEmail,

          kickOffDate,

          sentKickOffEmail,

          hcl,
          picus,
          proofpoint,
          seciron,

          note,
        });
      }

      return result;
    }, [
      matrix,
      headerInfo,
    ]);

  function toggleToolFilter(
    tool: ToolFilter
  ) {
    setToolFilters(
      (current) =>
        current.includes(
          tool
        )
          ? current.filter(
              (item) =>
                item !==
                tool
            )
          : [
              ...current,
              tool,
            ]
    );
  }

  function progressOf(
    row: WaitingRow
  ) {
    const steps = [
      row.confirmedParticipation,
      row.responseLetter,
      row.sentLetterEmail,
      row.followUpEmail,
      Boolean(
        row.kickOffDate
      ),
      row.sentKickOffEmail,
    ];

    const completed =
      steps.filter(
        Boolean
      ).length;

    return Math.round(
      (
        completed /
        steps.length
      ) * 100
    );
  }

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
              row.ncsaCoordinator,
              row.organizationCoordinator,
              row.email,
              row.phone,
              row.details,
              row.letterTo,
              row.note,
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
            onlyKickOff &&
            !row.kickOffDate
          ) {
            return false;
          }

          if (
            onlyPending &&
            progressOf(
              row
            ) >= 100
          ) {
            return false;
          }

          if (
            toolFilters.includes(
              "hcl"
            ) &&
            !row.hcl
          ) {
            return false;
          }

          if (
            toolFilters.includes(
              "picus"
            ) &&
            !row.picus
          ) {
            return false;
          }

          if (
            toolFilters.includes(
              "proofpoint"
            ) &&
            !row.proofpoint
          ) {
            return false;
          }

          if (
            toolFilters.includes(
              "seciron"
            ) &&
            !row.seciron
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
      onlyKickOff,
      onlyPending,
      toolFilters,
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

      const kickOff =
        rows.filter(
          (row) =>
            Boolean(
              row.kickOffDate
            )
        ).length;

      const completed =
        rows.filter(
          (row) =>
            progressOf(
              row
            ) >= 100
        ).length;

      return {
        total:
          rows.length,

        confirmed,
        response,
        kickOff,
        completed,

        pending:
          rows.length -
          completed,

        hcl:
          rows.filter(
            (row) =>
              row.hcl
          ).length,

        picus:
          rows.filter(
            (row) =>
              row.picus
          ).length,

        proofpoint:
          rows.filter(
            (row) =>
              row.proofpoint
          ).length,

        seciron:
          rows.filter(
            (row) =>
              row.seciron
          ).length,
      };
    }, [rows]);

  const toolSummary = [
    {
      key:
        "hcl",
      name:
        "HCL",
      value:
        summary.hcl,
    },
    {
      key:
        "picus",
      name:
        "PICUS",
      value:
        summary.picus,
    },
    {
      key:
        "proofpoint",
      name:
        "PROOFPOINT",
      value:
        summary.proofpoint,
    },
    {
      key:
        "seciron",
      name:
        "SECIRON",
      value:
        summary.seciron,
    },
  ];

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
        cleanText(
          sheetName
        )
          .replace(
            /[\\/:*?"<>|]/g,
            "_"
          )
          .trim() ||
        "waiting-organizations";

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
    <section className="waiting-page">

      {/* HEADER */}

      <div className="waiting-heading">
        <div>
          <span className="waiting-kicker">
            CLAB / PENDING ORGANIZATIONS
          </span>

          <h2>
            {sheetName}
          </h2>

          <p>
            ติดตามขั้นตอนการประสานงาน หนังสือเชิญ การตอบรับ และการนัดหมาย Kick-off
          </p>
        </div>

        <div className="waiting-heading-count">
          <strong>
            {
              summary.total
            }
          </strong>

          <span>
            หน่วยงาน
          </span>
        </div>
      </div>

      {/* SEARCH */}

      <div className="waiting-toolbar">
        <div className="waiting-search">
          <span>
            ⌕
          </span>

          <input
            type="text"
            value={
              search
            }
            placeholder="ค้นหาหน่วยงาน / ผู้ประสานงาน / อีเมล / หมายเหตุ..."
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

        <div className="waiting-toolbar-actions">
          <button
            type="button"
            className="waiting-excel-button"
            onClick={openInExcel}
            title="สร้างไฟล์ .xlsx ของ Sheet นี้เพื่อเปิดด้วย Microsoft Excel"
          >
            ▣ เปิดใน Excel
          </button>

          <div className="waiting-column-wrapper">
            <button
              type="button"
              className="waiting-column-button"
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
              <div className="waiting-column-picker">
                <div className="waiting-column-picker-head">
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

                <div className="waiting-column-picker-grid">
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

          <span className="waiting-result-count">
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

      {/* FILTER */}

      <div className="waiting-filter-panel">
        <div className="waiting-filter-group">
          <strong>
            ขั้นตอน
          </strong>

          <label>
            <input
              type="checkbox"
              checked={
                onlyConfirmed
              }
              onChange={(
                event
              ) =>
                setOnlyConfirmed(
                  event
                    .target
                    .checked
                )
              }
            />
            ยืนยันเข้าร่วมแล้ว
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                onlyResponse
              }
              onChange={(
                event
              ) =>
                setOnlyResponse(
                  event
                    .target
                    .checked
                )
              }
            />
            มีหนังสือตอบรับ
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                onlyKickOff
              }
              onChange={(
                event
              ) =>
                setOnlyKickOff(
                  event
                    .target
                    .checked
                )
              }
            />
            นัด Kick-off แล้ว
          </label>

          <label>
            <input
              type="checkbox"
              checked={
                onlyPending
              }
              onChange={(
                event
              ) =>
                setOnlyPending(
                  event
                    .target
                    .checked
                )
              }
            />
            ยังไม่ครบทุกขั้นตอน
          </label>
        </div>

        <div className="waiting-filter-divider" />

        <div className="waiting-filter-group">
          <strong>
            Tool
          </strong>

          {(
            [
              [
                "hcl",
                "HCL",
              ],
              [
                "picus",
                "PICUS",
              ],
              [
                "proofpoint",
                "PROOFPOINT",
              ],
              [
                "seciron",
                "SECIRON",
              ],
            ] as const
          ).map(
            ([
              key,
              label,
            ]) => (
              <label
                key={
                  key
                }
              >
                <input
                  type="checkbox"
                  checked={
                    toolFilters.includes(
                      key
                    )
                  }
                  onChange={() =>
                    toggleToolFilter(
                      key
                    )
                  }
                />
                {
                  label
                }
              </label>
            )
          )}
        </div>
      </div>

      {/* SUMMARY */}

      <div className="waiting-summary-grid">
        <article>
          <span>
            หน่วยงานทั้งหมด
          </span>

          <strong>
            {
              summary.total
            }
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
            {
              summary.confirmed
            }
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
            {
              summary.response
            }
          </strong>

          <small>
            Received
          </small>
        </article>

        <article>
          <span>
            นัด Kick-off
          </span>

          <strong>
            {
              summary.kickOff
            }
          </strong>

          <small className="positive">
            Scheduled
          </small>
        </article>

        <article>
          <span>
            ครบทุกขั้นตอน
          </span>

          <strong>
            {
              summary.completed
            }
          </strong>

          <small className="positive">
            Completed
          </small>
        </article>

        <article>
          <span>
            รอดำเนินการ
          </span>

          <strong>
            {
              summary.pending
            }
          </strong>

          <small className="warning">
            Pending
          </small>
        </article>
      </div>

      {/* MAIN */}

      <section className="waiting-main-card">
        <div className="waiting-card-header">
          <div>
            <h3>
              ภาพรวมการติดตามหน่วยงาน
            </h3>

            <p>
              แสดงรายละเอียด ขั้นตอนดำเนินงาน และ Tool ที่หน่วยงานขอทดสอบ
            </p>
          </div>

          <div className="waiting-tabs">
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
                  "tools",
                  "TOOLS",
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
          <div className="waiting-table-wrapper">
            {visibleColumns.length === 0 ? (
              <div className="waiting-no-columns">
                กรุณาเลือกอย่างน้อย 1 คอลัมน์
              </div>
            ) : (
              <table
                className="waiting-table"
                style={{
                  minWidth:
                    `${Math.max(
                      900,
                      visibleColumns.length *
                        145
                    )}px`,
                }}
              >
                <thead>
                  <tr>
                    {isColumnVisible("order") && (
                      <th rowSpan={2}>
                        ลำดับ
                      </th>
                    )}

                    {isColumnVisible("organization") && (
                      <th rowSpan={2}>
                        ชื่อหน่วยงาน
                      </th>
                    )}

                    {isColumnVisible("ncsaCoordinator") && (
                      <th rowSpan={2}>
                        ผู้ประสานงาน สกมช.
                      </th>
                    )}

                    {isColumnVisible("organizationCoordinator") && (
                      <th rowSpan={2}>
                        ผู้ประสานงานหน่วยงาน
                      </th>
                    )}

                    {isColumnVisible("email") && (
                      <th rowSpan={2}>
                        อีเมล
                      </th>
                    )}

                    {isColumnVisible("phone") && (
                      <th rowSpan={2}>
                        เบอร์โทร
                      </th>
                    )}

                    {isColumnVisible("contactDate") && (
                      <th rowSpan={2}>
                        วันที่ติดต่อ
                      </th>
                    )}

                    {isColumnVisible("details") && (
                      <th rowSpan={2}>
                        รายละเอียดอื่นๆ
                      </th>
                    )}

                    {isColumnVisible("letterTo") && (
                      <th rowSpan={2}>
                        หนังสือเรียนใคร
                      </th>
                    )}

                    {visibleProcessCount > 0 && (
                      <th
                        colSpan={visibleProcessCount}
                        className="waiting-process-head"
                      >
                        ขั้นตอนดำเนินงาน
                      </th>
                    )}

                    {visibleToolCount > 0 && (
                      <th
                        colSpan={visibleToolCount}
                        className="waiting-tool-head"
                      >
                        Tool ที่หน่วยงานขอทดสอบ
                      </th>
                    )}

                    {isColumnVisible("note") && (
                      <th rowSpan={2}>
                        หมายเหตุ
                      </th>
                    )}
                  </tr>

                  <tr>
                    {isColumnVisible("confirmedParticipation") && (
                      <th>ยืนยันเข้าร่วม</th>
                    )}

                    {isColumnVisible("responseLetter") && (
                      <th>หนังสือตอบรับ</th>
                    )}

                    {isColumnVisible("sentLetterEmail") && (
                      <th>ส่งหนังสือทางอีเมล</th>
                    )}

                    {isColumnVisible("followUpEmail") && (
                      <th>ติดตามอีเมล</th>
                    )}

                    {isColumnVisible("kickOffDate") && (
                      <th>วัน Kick-off</th>
                    )}

                    {isColumnVisible("sentKickOffEmail") && (
                      <th>ส่งอีเมล Kick-off</th>
                    )}

                    {isColumnVisible("hcl") && (
                      <th>HCL</th>
                    )}

                    {isColumnVisible("picus") && (
                      <th>PICUS</th>
                    )}

                    {isColumnVisible("proofpoint") && (
                      <th>PROOFPOINT</th>
                    )}

                    {isColumnVisible("seciron") && (
                      <th>SECIRON</th>
                    )}
                  </tr>
                </thead>

                <tbody>
                  {filteredRows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={Math.max(
                          1,
                          visibleColumns.length
                        )}
                        className="waiting-empty"
                      >
                        ไม่พบข้อมูลที่ตรงกับตัวกรอง
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map(
                      (row) => (
                        <tr
                          key={row.sourceRow}
                        >
                          {isColumnVisible("order") && (
                            <td className="waiting-order">
                              {row.order || "-"}
                            </td>
                          )}

                          {isColumnVisible("organization") && (
                            <td className="waiting-org">
                              {row.organization || "-"}
                            </td>
                          )}

                          {isColumnVisible("ncsaCoordinator") && (
                            <td>
                              {row.ncsaCoordinator || "-"}
                            </td>
                          )}

                          {isColumnVisible("organizationCoordinator") && (
                            <td>
                              {row.organizationCoordinator || "-"}
                            </td>
                          )}

                          {isColumnVisible("email") && (
                            <td className="waiting-email">
                              {row.email || "-"}
                            </td>
                          )}

                          {isColumnVisible("phone") && (
                            <td>
                              {row.phone || "-"}
                            </td>
                          )}

                          {isColumnVisible("contactDate") && (
                            <td className="waiting-date">
                              {row.contactDate || "-"}
                            </td>
                          )}

                          {isColumnVisible("details") && (
                            <td>
                              {row.details || "-"}
                            </td>
                          )}

                          {isColumnVisible("letterTo") && (
                            <td>
                              {row.letterTo || "-"}
                            </td>
                          )}

                          {isColumnVisible("confirmedParticipation") && (
                            <td className="waiting-check">
                              <CheckIcon
                                checked={
                                  row.confirmedParticipation
                                }
                              />
                            </td>
                          )}

                          {isColumnVisible("responseLetter") && (
                            <td className="waiting-check">
                              <CheckIcon
                                checked={
                                  row.responseLetter
                                }
                              />
                            </td>
                          )}

                          {isColumnVisible("sentLetterEmail") && (
                            <td className="waiting-check">
                              <CheckIcon
                                checked={
                                  row.sentLetterEmail
                                }
                              />
                            </td>
                          )}

                          {isColumnVisible("followUpEmail") && (
                            <td className="waiting-check">
                              <CheckIcon
                                checked={
                                  row.followUpEmail
                                }
                              />
                            </td>
                          )}

                          {isColumnVisible("kickOffDate") && (
                            <td className="waiting-date">
                              {row.kickOffDate || "-"}
                            </td>
                          )}

                          {isColumnVisible("sentKickOffEmail") && (
                            <td className="waiting-check">
                              <CheckIcon
                                checked={
                                  row.sentKickOffEmail
                                }
                              />
                            </td>
                          )}

                          {isColumnVisible("hcl") && (
                            <td className="waiting-check tool hcl">
                              <CheckIcon
                                checked={row.hcl}
                              />
                            </td>
                          )}

                          {isColumnVisible("picus") && (
                            <td className="waiting-check tool picus">
                              <CheckIcon
                                checked={row.picus}
                              />
                            </td>
                          )}

                          {isColumnVisible("proofpoint") && (
                            <td className="waiting-check tool proofpoint">
                              <CheckIcon
                                checked={row.proofpoint}
                              />
                            </td>
                          )}

                          {isColumnVisible("seciron") && (
                            <td className="waiting-check tool seciron">
                              <CheckIcon
                                checked={row.seciron}
                              />
                            </td>
                          )}

                          {isColumnVisible("note") && (
                            <td>
                              {row.note || "-"}
                            </td>
                          )}
                        </tr>
                      )
                    )
                  )}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* PROGRESS */}

        {viewType ===
          "progress" && (
          <div className="waiting-progress-list">
            {filteredRows.map(
              (
                row
              ) => {
                const progress =
                  progressOf(
                    row
                  );

                return (
                  <article
                    key={
                      row.sourceRow
                    }
                    className="waiting-progress-item"
                  >
                    <div className="waiting-progress-heading">
                      <div>
                        <strong>
                          {
                            row.organization ||
                            `รายการ ${row.sourceRow}`
                          }
                        </strong>

                        <span>
                          {
                            row.organizationCoordinator ||
                            row.email ||
                            "ไม่มีข้อมูลผู้ประสานงาน"
                          }
                        </span>
                      </div>

                      <b>
                        {
                          progress
                        }
                        %
                      </b>
                    </div>

                    <div className="waiting-progress-track">
                      <div
                        className="waiting-progress-fill"
                        style={{
                          width:
                            `${progress}%`,
                        }}
                      />
                    </div>

                    <div className="waiting-progress-steps">
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

        {/* TOOLS */}

        {viewType ===
          "tools" && (
          <div className="waiting-tools-area">
            <div className="waiting-tool-chart">
              {toolSummary.map(
                (
                  item
                ) => {
                  const height =
                    (
                      item.value /
                      maxTool
                    ) *
                    100;

                  return (
                    <div
                      key={
                        item.key
                      }
                      className="waiting-tool-chart-item"
                    >
                      <strong>
                        {
                          item.value
                        }
                      </strong>

                      <div className="waiting-tool-chart-track">
                        <div
                          className={`waiting-tool-chart-fill ${item.key}`}
                          style={{
                            height:
                              `${height}%`,
                          }}
                        />
                      </div>

                      <span>
                        {
                          item.name
                        }
                      </span>
                    </div>
                  );
                }
              )}
            </div>
          </div>
        )}
      </section>
    </section>
  );
}

/* ==============================
   SMALL COMPONENTS
================================ */

function CheckIcon({
  checked,
}: {
  checked: boolean;
}) {
  return (
    <span
      className={`waiting-check-icon ${
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

function Step({
  label,
  done,
}: {
  label: string;
  done: boolean;
}) {
  return (
    <span
      className={`waiting-step ${
        done
          ? "done"
          : ""
      }`}
    >
      <i>
        {
          done
            ? "✓"
            : ""
        }
      </i>

      {
        label
      }
    </span>
  );
}

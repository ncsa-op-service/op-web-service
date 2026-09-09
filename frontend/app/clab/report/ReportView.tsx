"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./report.css";

type Props = {
  worksheet: XLSX.WorkSheet;
  sheetName: string;
};

type CellValue = string | number | boolean | null | undefined;
type MatrixRow = CellValue[];
type ViewType = "table" | "bar" | "pie" | "progress";

type BarMetric =
  | "status"
  | "project"
  | "personal";

type ReportRow = {
  sourceRow: number;
  order: string;
  organization: string;
  organizationType: string;
  progress: string;
  note: string;
  caseId: string;
  password: string;
  evidence: string;

  hclReport: string;
  hclStatus: string;

  picusReport: string;
  picusStatus: string;

  proofpointReport: string;
  proofpointStatus: string;

  secironReport: string;
  secironStatus: string;

  closingDate: string;
  satisfaction: string;
  score: string;
  provider: string;
  totalReports: string;
  completedReports: string;
  incompleteReports: string;
};

type ColumnKey =
  | "order"
  | "organization"
  | "organizationType"
  | "progress"
  | "note"
  | "caseId"
  | "password"
  | "evidence"
  | "hclReport"
  | "hclStatus"
  | "picusReport"
  | "picusStatus"
  | "proofpointReport"
  | "proofpointStatus"
  | "secironReport"
  | "secironStatus"
  | "closingDate"
  | "satisfaction"
  | "score"
  | "provider"
  | "totalReports"
  | "completedReports"
  | "incompleteReports";

type ColumnDef = {
  key: ColumnKey;
  label: string;
};

type ColumnMap = Record<ColumnKey, number>;

type ToolKey = "hcl" | "picus" | "proofpoint" | "seciron";

type ToolSummary = {
  key: ToolKey;
  name: string;
  completed: number;
  incomplete: number;
  total: number;

  // ค่าจากตารางสรุปใน Excel
  projectTotal: number;
  personalTotal: number;
};

const COLUMNS: ColumnDef[] = [
  { key: "order", label: "ลำดับ" },
  { key: "organization", label: "ชื่อหน่วยงาน" },
  { key: "organizationType", label: "ประเภทหน่วยงาน" },
  { key: "progress", label: "ระดับความคืบหน้า" },
  { key: "note", label: "NOTE" },
  { key: "caseId", label: "เลขเคส" },
  { key: "password", label: "Password" },
  { key: "evidence", label: "หลักฐาน Report" },

  { key: "hclReport", label: "จำนวน Report" },
  { key: "hclStatus", label: "สถานะ" },

  { key: "picusReport", label: "จำนวน Report" },
  { key: "picusStatus", label: "สถานะ" },

  { key: "proofpointReport", label: "จำนวน Report" },
  { key: "proofpointStatus", label: "สถานะ" },

  { key: "secironReport", label: "จำนวน Report" },
  { key: "secironStatus", label: "สถานะ" },

  { key: "closingDate", label: "วันนัดหมายสรุปปิดเคส" },
  { key: "satisfaction", label: "แบบประเมินความพึงพอใจ" },
  { key: "score", label: "คะแนน" },
  { key: "provider", label: "Provider" },
  { key: "totalReports", label: "รวมทั้งหมด" },
  { key: "completedReports", label: "จำนวนรายงานดำเนินการเสร็จสิ้น" },
  { key: "incompleteReports", label: "จำนวนรายงานยังไม่เสร็จสิ้น" },
];

const DEFAULT_VISIBLE_COLUMNS: ColumnKey[] = [
  "order",
  "organization",
  "organizationType",
  "progress",
  "note",
  "caseId",
  "password",
  "evidence",
  "hclReport",
  "hclStatus",
  "picusReport",
  "picusStatus",
  "proofpointReport",
  "proofpointStatus",
  "secironReport",
  "secironStatus",
  "closingDate",
  "satisfaction",
  "score",
  "provider",
  "totalReports",
  "completedReports",
  "incompleteReports",
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

function fillMergedCells(
  worksheet: XLSX.WorkSheet
): MatrixRow[] {
  /*
    อ่าน cell โดยตรงจาก worksheet["!ref"]
    เพื่อไม่ให้ header หลายชั้น / merged cells
    ถูกแปลงจนตำแหน่งเพี้ยน
  */
  const ref = worksheet["!ref"];

  if (!ref) {
    return [];
  }

  const range = XLSX.utils.decode_range(ref);
  const result: MatrixRow[] = [];

  for (let r = range.s.r; r <= range.e.r; r += 1) {
    const row: MatrixRow = [];

    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const address = XLSX.utils.encode_cell({ r, c });
      const cell = worksheet[address];

      row[c] =
        cell?.w ??
        cell?.v ??
        "";
    }

    result[r] = row;
  }

  /*
    เติมค่าของ merged cell ให้ทุกตำแหน่งในช่วง merge
    เช่น HCL(OP01) ที่ครอบ "จำนวน Report" + "สถานะ"
  */
  for (const merge of worksheet["!merges"] ?? []) {
    const sourceValue =
      result[merge.s.r]?.[merge.s.c] ?? "";

    for (let r = merge.s.r; r <= merge.e.r; r += 1) {
      if (!result[r]) {
        result[r] = [];
      }

      for (let c = merge.s.c; c <= merge.e.c; c += 1) {
        if (cleanText(result[r][c]) === "") {
          result[r][c] = sourceValue;
        }
      }
    }
  }

  return result;
}

function getCell(row: MatrixRow, columnIndex: number) {
  if (columnIndex < 0) return "";
  return cleanText(row[columnIndex]);
}

function findHeaderPosition(
  matrix: MatrixRow[],
  candidates: string[]
) {
  const maxRows = Math.min(matrix.length, 60);

  for (let r = 0; r < maxRows; r += 1) {
    const row = matrix[r] ?? [];

    for (let c = 0; c < row.length; c += 1) {
      const value = normalize(row[c]);

      if (
        candidates.some((candidate) =>
          value.includes(normalize(candidate))
        )
      ) {
        return {
          row: r,
          column: c,
        };
      }
    }
  }

  return null;
}

function findSubHeaderRow(
  matrix: MatrixRow[]
) {
  const maxRows = Math.min(matrix.length, 60);

  for (let r = 0; r < maxRows; r += 1) {
    const values = (matrix[r] ?? []).map(normalize);

    const reportCount = values.filter((value) =>
      value.includes("จำนวน report")
    ).length;

    const statusCount = values.filter((value) =>
      value === "สถานะ" ||
      value.includes("สถานะ")
    ).length;

    if (reportCount >= 2 && statusCount >= 2) {
      return r;
    }
  }

  return -1;
}

function findColumnByKeywords(
  matrix: MatrixRow[],
  headerStartRow: number,
  required: string[]
) {
  const startRow = Math.max(headerStartRow, 0);
  const endRow = Math.min(
    matrix.length - 1,
    startRow + 4
  );

  const maxColumns = Math.max(
    0,
    ...matrix
      .slice(startRow, endRow + 1)
      .map((row) => row?.length ?? 0)
  );

  for (let c = 0; c < maxColumns; c += 1) {
    const combined: string[] = [];

    for (let r = startRow; r <= endRow; r += 1) {
      const value = normalize(matrix[r]?.[c]);

      if (value) {
        combined.push(value);
      }
    }

    const joined = combined.join(" ");

    if (
      required.every((keyword) =>
        joined.includes(normalize(keyword))
      )
    ) {
      return c;
    }
  }

  return -1;
}

function findColumnByAny(
  matrix: MatrixRow[],
  headerStartRow: number,
  candidates: string[]
) {
  const startRow = Math.max(headerStartRow, 0);
  const endRow = Math.min(
    matrix.length - 1,
    startRow + 4
  );

  const maxColumns = Math.max(
    0,
    ...matrix
      .slice(startRow, endRow + 1)
      .map((row) => row?.length ?? 0)
  );

  for (let c = 0; c < maxColumns; c += 1) {
    const combined: string[] = [];

    for (let r = startRow; r <= endRow; r += 1) {
      const value = normalize(matrix[r]?.[c]);

      if (value) {
        combined.push(value);
      }
    }

    const joined = combined.join(" ");

    if (
      candidates.some((candidate) =>
        joined.includes(normalize(candidate))
      )
    ) {
      return c;
    }
  }

  return -1;
}

function buildColumnMap(
  matrix: MatrixRow[],
  headerRowIndex: number
): ColumnMap {
  return {
    order: findColumnByAny(matrix, headerRowIndex, ["ลำดับ"]),

    organization: findColumnByAny(matrix, headerRowIndex, [
      "ชื่อหน่วยงาน",
    ]),

    organizationType: findColumnByAny(matrix, headerRowIndex, [
      "ประเภทหน่วยงาน",
      "ประเภทหน่วงาน",
    ]),

    progress: findColumnByAny(matrix, headerRowIndex, [
      "ระดับความคืบหน้า",
    ]),

    note: findColumnByAny(matrix, headerRowIndex, ["note"]),

    caseId: findColumnByAny(matrix, headerRowIndex, [
      "เลขเคส",
      "case id",
    ]),

    password: findColumnByAny(matrix, headerRowIndex, ["password"]),

    evidence: findColumnByAny(matrix, headerRowIndex, [
      "หลักฐาน report",
    ]),

    hclReport: findColumnByKeywords(matrix, headerRowIndex, [
      "hcl",
      "จำนวน report",
    ]),

    hclStatus: findColumnByKeywords(matrix, headerRowIndex, [
      "hcl",
      "สถานะ",
    ]),

    picusReport: findColumnByKeywords(matrix, headerRowIndex, [
      "picus",
      "จำนวน report",
    ]),

    picusStatus: findColumnByKeywords(matrix, headerRowIndex, [
      "picus",
      "สถานะ",
    ]),

    proofpointReport: findColumnByKeywords(matrix, headerRowIndex, [
      "proofpoint",
      "จำนวน report",
    ]),

    proofpointStatus: findColumnByKeywords(matrix, headerRowIndex, [
      "proofpoint",
      "สถานะ",
    ]),

    secironReport: findColumnByKeywords(matrix, headerRowIndex, [
      "seciron",
      "จำนวน report",
    ]),

    secironStatus: findColumnByKeywords(matrix, headerRowIndex, [
      "seciron",
      "สถานะ",
    ]),

    closingDate: findColumnByAny(matrix, headerRowIndex, [
      "วันนัดหมายสรุปปิดเคส",
    ]),

    satisfaction: findColumnByAny(matrix, headerRowIndex, [
      "แบบประเมินความพึงพอใจ",
    ]),

    score: findColumnByAny(
  matrix,
  headerRowIndex,
  ["คะแนน"]
),

provider: findColumnByAny(
  matrix,
  headerRowIndex,
  ["provider"]
),

/*
  3 ตัวนี้เป็นตาราง Summary ด้านขวาของ Excel
  ไม่ใช่ column ของข้อมูลรายหน่วยงาน
  จึงไม่ต้องหาใน Header หลัก
*/
totalReports: -1,
completedReports: -1,
incompleteReports: -1,
  };
  
}

function isDataOrder(value: string) {
  return /^\d+(?:\.\d+)?$/.test(value.trim());
}

function parseReportCount(value: string) {
  if (!value) return 0;

  const cleaned = value
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");

  const number = Number(cleaned);

  return Number.isFinite(number) && number > 0 ? number : 0;
}

function isFinishedStatus(value: string) {
  const status = normalize(value);

  if (!status) return false;

  return (
    status.includes("เสร็จ") ||
    status.includes("complete") ||
    status === "done" ||
    status === "completed"
  );
}

function statusClass(value: string) {
  const status = normalize(value);

  if (!status) return "empty";

  if (isFinishedStatus(value)) {
    return "done";
  }

  if (
    status.includes("กำลัง") ||
    status.includes("ดำเนิน") ||
    status.includes("รอ")
  ) {
    return "working";
  }

  return "default";
}

function renderMaybeLink(value: string) {
  if (!value) return "-";

  if (/^https?:\/\//i.test(value)) {
    return (
      <a
        href={value}
        target="_blank"
        rel="noreferrer"
        className="report-link"
      >
        เปิดลิงก์
      </a>
    );
  }

  return value;
}

function getToolValues(row: ReportRow, tool: ToolKey) {
  switch (tool) {
    case "hcl":
      return {
        count: row.hclReport,
        status: row.hclStatus,
      };

    case "picus":
      return {
        count: row.picusReport,
        status: row.picusStatus,
      };

    case "proofpoint":
      return {
        count: row.proofpointReport,
        status: row.proofpointStatus,
      };

    case "seciron":
      return {
        count: row.secironReport,
        status: row.secironStatus,
      };
  }
}

export default function ReportView({
  worksheet,
  sheetName,
}: Props) {
  const [search, setSearch] = useState("");
  const [viewType, setViewType] = useState<ViewType>("table");
  const [showColumnPicker, setShowColumnPicker] = useState(false);

  const [visibleBarMetrics, setVisibleBarMetrics] =
    useState<BarMetric[]>([
      "status",
      "project",
      "personal",
    ]);

  function toggleBarMetric(metric: BarMetric) {
    setVisibleBarMetrics((current) =>
      current.includes(metric)
        ? current.filter((item) => item !== metric)
        : [...current, metric]
    );
  }

  const [visibleColumns, setVisibleColumns] =
    useState<ColumnKey[]>(DEFAULT_VISIBLE_COLUMNS);

  const matrix = useMemo(
    () => fillMergedCells(worksheet),
    [worksheet]
  );

  const headerInfo = useMemo(() => {
    const orderHeader =
      findHeaderPosition(
        matrix,
        ["ลำดับ"]
      );

    const organizationHeader =
      findHeaderPosition(
        matrix,
        ["ชื่อหน่วยงาน"]
      );

    const subHeaderRow =
      findSubHeaderRow(matrix);

    if (
      !orderHeader ||
      !organizationHeader
    ) {
      return null;
    }

    const headerStartRow =
      Math.min(
        orderHeader.row,
        organizationHeader.row
      );

    const dataStartRow =
      Math.max(
        orderHeader.row,
        organizationHeader.row,
        subHeaderRow
      ) + 1;

    return {
      headerStartRow,
      dataStartRow,
    };
  }, [matrix]);

  const columnMap = useMemo(
    () =>
      headerInfo
        ? buildColumnMap(
            matrix,
            headerInfo.headerStartRow
          )
        : null,
    [matrix, headerInfo]
  );

  const rows = useMemo<ReportRow[]>(() => {
    if (!columnMap || !headerInfo) {
      return [];
    }

    const result: ReportRow[] = [];

    const dataStartRow =
      headerInfo.dataStartRow;

    for (
      let rowIndex = dataStartRow;
      rowIndex < matrix.length;
      rowIndex += 1
    ) {
      const source = matrix[rowIndex];

      if (!source) continue;

      const order = getCell(source, columnMap.order);
      const organization = getCell(
        source,
        columnMap.organization
      );

      /*
        สำคัญ:
        ใช้ "ลำดับ + ชื่อหน่วยงาน" เป็นตัวตัดสินว่าเป็นข้อมูลจริง
        จึงไม่ดึงหัวตาราง / แถวสรุป / สูตรท้าย Sheet มาปน
      */
      if (!isDataOrder(order) || !organization) {
        continue;
      }

      result.push({
        sourceRow: rowIndex + 1,
        order,
        organization,

        organizationType: getCell(
          source,
          columnMap.organizationType
        ),

        progress: getCell(source, columnMap.progress),
        note: getCell(source, columnMap.note),
        caseId: getCell(source, columnMap.caseId),
        password: getCell(source, columnMap.password),
        evidence: getCell(source, columnMap.evidence),

        hclReport: getCell(source, columnMap.hclReport),
        hclStatus: getCell(source, columnMap.hclStatus),

        picusReport: getCell(source, columnMap.picusReport),
        picusStatus: getCell(source, columnMap.picusStatus),

        proofpointReport: getCell(
          source,
          columnMap.proofpointReport
        ),

        proofpointStatus: getCell(
          source,
          columnMap.proofpointStatus
        ),

        secironReport: getCell(source, columnMap.secironReport),
        secironStatus: getCell(source, columnMap.secironStatus),

        closingDate: getCell(source, columnMap.closingDate),

        satisfaction: getCell(
          source,
          columnMap.satisfaction
        ),

        score: getCell(source, columnMap.score),
        provider: getCell(source, columnMap.provider),

        totalReports: "",
        completedReports: "",
        incompleteReports: "",
      });
    }

    /*
      ไม่ sort ใหม่เด็ดขาด
      เรียงตามตำแหน่งจริงใน Excel
      รวมถึงเลขซ้ำ/เลขข้าม เช่น 29 -> 31 -> 32 -> 32 -> 33
    */
    return result;
  }, [matrix, columnMap, headerInfo]);

  const filteredRows = useMemo(() => {
    const keyword = normalize(search);

    if (!keyword) {
      return rows;
    }

    return rows.filter((row) =>
      [
        row.order,
        row.organization,
        row.organizationType,
        row.progress,
        row.note,
        row.caseId,
        row.provider,
      ].some((value) =>
        normalize(value).includes(keyword)
      )
    );
  }, [rows, search]);

  const toolSummary = useMemo<ToolSummary[]>(() => {
    const definitions: {
      key: ToolKey;
      name: string;
    }[] = [
      { key: "hcl", name: "HCL" },
      { key: "picus", name: "PICUS" },
      { key: "proofpoint", name: "PROOFPOINT" },
      { key: "seciron", name: "SECIRON" },
    ];

    /*
      ตารางสรุปด้านล่างของ Excel มีแถว:
      - จำนวนรายงานทั้งหมด (ตัวชี้วัด โครงการ)
      - จำนวนรายงานทั้งหมด (ตัวชี้วัด รายบุคคล)
      และมีแถว Provider ที่บอกว่าแต่ละคอลัมน์เป็น
      HCL / PICUS / PROOFPOINT / SECIRON
    */

    function findSummaryLabelRow(
      label: string
    ) {
      const wanted = normalize(label);

      return matrix.findIndex(
        (row) =>
          (row ?? []).some(
            (cell) =>
              normalize(cell).includes(
                wanted
              )
          )
      );
    }

    function findProviderRow() {
      /*
        ต้องหา "แถวสรุป Provider" โดยเฉพาะ
        ห้ามหา HCL / PICUS จากทั้ง Sheet เพราะในข้อมูลรายหน่วยงาน
        ก็มีคำว่า HCL / PICUS / PROOFPOINT / SECIRON อยู่เช่นกัน
      */
      for (
        let r = 0;
        r < matrix.length;
        r += 1
      ) {
        const values =
          (matrix[r] ?? []).map(
            (cell) => normalize(cell)
          );

        const hasProvider =
          values.some(
            (value) =>
              value === "provider"
          );

        const toolCount = [
          "hcl",
          "picus",
          "proofpoint",
          "seciron",
        ].filter(
          (tool) =>
            values.includes(tool)
        ).length;

        if (
          hasProvider &&
          toolCount >= 2
        ) {
          return r;
        }
      }

      return -1;
    }

    function findProviderColumn(
      providerRow: number,
      toolName: string
    ) {
      if (providerRow < 0) {
        return -1;
      }

      const wanted =
        normalize(toolName);

      const row =
        matrix[providerRow] ?? [];

      for (
        let c = 0;
        c < row.length;
        c += 1
      ) {
        if (
          normalize(
            row[c]
          ) === wanted
        ) {
          return c;
        }
      }

      return -1;
    }

    function readSummaryNumber(
      rowIndex: number,
      columnIndex: number
    ) {
      if (
        rowIndex < 0 ||
        columnIndex < 0
      ) {
        return 0;
      }

      const raw =
        cleanText(
          matrix[
            rowIndex
          ]?.[
            columnIndex
          ]
        );

      if (!raw) {
        return 0;
      }

      const value =
        Number(
          raw
            .replace(/,/g, "")
            .replace(
              /[^\d.-]/g,
              ""
            )
        );

      return Number.isFinite(
        value
      )
        ? value
        : 0;
    }

    const projectRow =
      findSummaryLabelRow(
        "จำนวนรายงานทั้งหมด (ตัวชี้วัด โครงการ)"
      );

    const personalRow =
      findSummaryLabelRow(
        "จำนวนรายงานทั้งหมด (ตัวชี้วัด รายบุคคล)"
      );

    const providerRow =
      findProviderRow();

    return definitions.map(
      ({ key, name }) => {
        /*
          เสร็จ / ยังไม่เสร็จ
          ยังใช้ logic เดิมจากรายการจริง
          เพื่อให้การ์ดด้านบนเหมือนเดิม
        */
        let completed = 0;
        let incomplete = 0;

        rows.forEach(
          (row) => {
            const tool =
              getToolValues(
                row,
                key
              );

            const reportCount =
              parseReportCount(
                tool.count
              );

            if (
              reportCount <= 0
            ) {
              return;
            }

            if (
              isFinishedStatus(
                tool.status
              )
            ) {
              completed +=
                reportCount;
            } else {
              incomplete +=
                reportCount;
            }
          }
        );

        const providerColumn =
          findProviderColumn(
            providerRow,
            name
          );

        const projectTotal =
          readSummaryNumber(
            projectRow,
            providerColumn
          );

        const personalTotal =
          readSummaryNumber(
            personalRow,
            providerColumn
          );

        return {
          key,
          name,
          completed,
          incomplete,

          // ค่ารวมเดิมของการ์ดด้านบน
          total:
            completed +
            incomplete,

          // ค่าที่อ่านจากตารางสรุป Excel โดยตรง
          projectTotal,
          personalTotal,
        };
      }
    );
  }, [rows, matrix]);

  function toggleColumn(key: ColumnKey) {
    setVisibleColumns((current) =>
      current.includes(key)
        ? current.filter((item) => item !== key)
        : [...current, key]
    );
  }

  function renderCell(row: ReportRow, key: ColumnKey) {
    switch (key) {
      case "order":
        return row.order;

      case "organization":
        return (
          <strong className="report-org-name">
            {row.organization}
          </strong>
        );

      case "organizationType":
        return row.organizationType || "-";

      case "progress":
        return (
          <span
            className={`report-status-badge ${statusClass(
              row.progress
            )}`}
          >
            {row.progress || "-"}
          </span>
        );

      case "note":
        return row.note || "-";

      case "caseId":
        return row.caseId || "-";

      case "password":
        return row.password || "-";

      case "evidence":
        return renderMaybeLink(row.evidence);

      case "hclReport":
        return row.hclReport || "-";

      case "hclStatus":
        return (
          <span
            className={`report-status-badge ${statusClass(
              row.hclStatus
            )}`}
          >
            {row.hclStatus || "-"}
          </span>
        );

      case "picusReport":
        return row.picusReport || "-";

      case "picusStatus":
        return (
          <span
            className={`report-status-badge ${statusClass(
              row.picusStatus
            )}`}
          >
            {row.picusStatus || "-"}
          </span>
        );

      case "proofpointReport":
        return row.proofpointReport || "-";

      case "proofpointStatus":
        return (
          <span
            className={`report-status-badge ${statusClass(
              row.proofpointStatus
            )}`}
          >
            {row.proofpointStatus || "-"}
          </span>
        );

      case "secironReport":
        return row.secironReport || "-";

      case "secironStatus":
        return (
          <span
            className={`report-status-badge ${statusClass(
              row.secironStatus
            )}`}
          >
            {row.secironStatus || "-"}
          </span>
        );

      case "closingDate":
        return row.closingDate || "-";

      case "satisfaction":
        return renderMaybeLink(row.satisfaction);

      case "score":
        return row.score || "-";

      case "provider":
        return row.provider || "-";

 case "totalReports": {
  const summary = toolSummary.find(
    (tool) =>
      normalize(tool.name) ===
      normalize(row.provider)
  );

  return summary?.total ?? "-";
}

case "completedReports": {
  const summary = toolSummary.find(
    (tool) =>
      normalize(tool.name) ===
      normalize(row.provider)
  );

  return summary?.completed ?? "-";
}

case "incompleteReports": {
  const summary = toolSummary.find(
    (tool) =>
      normalize(tool.name) ===
      normalize(row.provider)
  );

  return summary?.incomplete ?? "-";
}

      default:
        return "-";
    }
  }

  function rowProgress(
    row: ReportRow
  ) {
    const value = row.progress.trim();

    if (!value) {
      return 0;
    }

    // กรณี Excel แสดงเป็นเปอร์เซ็นต์ เช่น 100% หรือ 65%
    if (value.includes("%")) {
      const number = Number(
        value
          .replace("%", "")
          .trim()
      );

      return Number.isFinite(number)
        ? Math.min(
            Math.max(number, 0),
            100
          )
        : 0;
    }

    // กรณี Excel เก็บเป็น 0.8 = 80%
    const number = Number(value);

    if (Number.isFinite(number)) {
      if (
        number >= 0 &&
        number <= 1
      ) {
        return Math.round(
          number * 100
        );
      }

      return Math.min(
        Math.max(number, 0),
        100
      );
    }

    // กรณีคอลัมน์ระดับความคืบหน้าเป็นข้อความสถานะ
    const status = normalize(value);

    if (
      status.includes(
        "ดำเนินการเสร็จสิ้น"
      ) ||
      status.includes(
        "เสร็จสิ้น"
      )
    ) {
      return 100;
    }

    if (
      status.includes(
        "กำลังดำเนินการ"
      )
    ) {
      return 50;
    }

    if (
      status.includes(
        "ยังไม่ดำเนินการ"
      ) ||
      status.includes(
        "ไม่ดำเนินการ"
      )
    ) {
      return 0;
    }

    return 0;
  }

  function openInExcel() {
    try {
      const outputWorkbook = XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        worksheet,
        sheetName
      );

      const safeFileName =
        cleanText(sheetName)
          .replace(/[\\/:*?"<>|]/g, "_")
          .trim() || "report";

      XLSX.writeFile(
        outputWorkbook,
        `${safeFileName}.xlsx`
      );
    } catch (excelError) {
      console.error("Open Excel error:", excelError);
      alert("สร้างไฟล์ Excel ไม่สำเร็จ");
    }
  }

  if (!headerInfo || !columnMap) {
    return (
      <section className="report-page">
        <div className="report-heading">
          <div>
            <h2>{sheetName}</h2>
            <p>
              ไม่พบหัวตาราง "ลำดับ" และ "ชื่อหน่วยงาน"
              ใน Sheet นี้
            </p>
          </div>
        </div>
      </section>
    );
  }

  return (
    <section className="report-page">
      <div className="report-heading">
        <div>
          <h2>{sheetName}</h2>

          <p>
            ดึงข้อมูลตามลำดับจริงจาก Excel
            และแยก HCL / PICUS / PROOFPOINT / SECIRON
          </p>
        </div>

        <span className="report-total-badge">
          {rows.length} รายการ
        </span>
      </div>

      <div className="report-toolbar">
        <div className="report-search">
          <span>⌕</span>

          <input
            type="text"
            placeholder="ค้นหาลำดับ / หน่วยงาน / เลขเคส / Provider..."
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
        </div>

        <div className="report-toolbar-right">
          <button
            type="button"
            className="report-excel-button"
            onClick={openInExcel}
            title="สร้างไฟล์ .xlsx ของ Sheet นี้เพื่อเปิดด้วย Microsoft Excel"
          >
            ▣ เปิดใน Excel
          </button>

          <span className="report-result-count">
            {filteredRows.length} รายการ
          </span>

          <div className="report-column-wrapper">
            <button
              type="button"
              className="report-column-button"
              onClick={() =>
                setShowColumnPicker((current) => !current)
              }
            >
              ☷ เลือกคอลัมน์
            </button>

            {showColumnPicker && (
              <div className="report-column-menu">
                <div className="report-column-menu-heading">
                  <strong>เลือกหัวข้อที่ต้องการแสดง</strong>

                  <button
                    type="button"
                    onClick={() => setShowColumnPicker(false)}
                  >
                    ×
                  </button>
                </div>

                <div className="report-column-actions">
                  <button
                    type="button"
                    onClick={() =>
                      setVisibleColumns(
                        COLUMNS.map((column) => column.key)
                      )
                    }
                  >
                    เลือกทั้งหมด
                  </button>

                  <button
                    type="button"
                    onClick={() => setVisibleColumns([])}
                  >
                    ล้างทั้งหมด
                  </button>
                </div>

                <div className="report-column-list">
                  {COLUMNS.map((column) => (
                    <label
                      key={column.key}
                      className="report-column-option"
                    >
                      <input
                        type="checkbox"
                        checked={visibleColumns.includes(
                          column.key
                        )}
                        onChange={() =>
                          toggleColumn(column.key)
                        }
                      />

                      <span>{column.label}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      <section className="report-summary-grid">
        {toolSummary.map((tool) => {
          const percent =
            tool.total > 0
              ? Math.round(
                  (tool.completed / tool.total) * 1000
                ) / 10
              : 0;

          return (
            <article
              key={tool.key}
              className="report-summary-card"
            >
              <span>{tool.name}</span>

              <strong>
                {tool.completed}/{tool.total}
              </strong>

              <p>
                เสร็จ {tool.completed} · ยังไม่เสร็จ{" "}
                {tool.incomplete}
              </p>

              <div className="report-summary-progress">
                <div
                  className="report-summary-progress-fill"
                  style={{ width: `${percent}%` }}
                />
              </div>
            </article>
          );
        })}
      </section>

      <section className="report-main-card">
        <div className="report-card-header">
          <div>
            <h3>ภาพรวม Report</h3>

            <p>
              แถวข้อมูลเรียงตามตำแหน่งจริงใน Sheet
            </p>
          </div>

          <div className="report-tabs">
            {(
              [
                ["table", "TABLE"],
                ["bar", "BAR"],
                ["pie", "PIE"],
                ["progress", "PROGRESS"],
              ] as const
            ).map(([value, label]) => (
              <button
                key={value}
                type="button"
                className={
                  viewType === value ? "active" : ""
                }
                onClick={() => setViewType(value)}
              >
                {label}
              </button>
            ))}
          </div>
        </div>

        {viewType === "table" && (
          <div className="report-table-wrapper">
            {visibleColumns.length === 0 ? (
              <div className="report-empty-columns">
                กรุณาเลือกอย่างน้อย 1 คอลัมน์
              </div>
            ) : (
              <table className="report-table">
                <thead>
                  <tr className="report-group-header">
                    {visibleColumns.includes("order") && (
                      <th rowSpan={2}>ลำดับ</th>
                    )}

                    {visibleColumns.includes("organization") && (
                      <th rowSpan={2}>ชื่อหน่วยงาน</th>
                    )}

                    {visibleColumns.includes(
                      "organizationType"
                    ) && (
                      <th rowSpan={2}>ประเภทหน่วยงาน</th>
                    )}

                    {visibleColumns.includes("progress") && (
                      <th rowSpan={2}>
                        ระดับความคืบหน้า
                      </th>
                    )}

                    {visibleColumns.includes("note") && (
                      <th rowSpan={2}>NOTE</th>
                    )}

                    {visibleColumns.includes("caseId") && (
                      <th rowSpan={2}>เลขเคส</th>
                    )}

                    {visibleColumns.includes("password") && (
                      <th rowSpan={2}>Password</th>
                    )}

                    {visibleColumns.includes("evidence") && (
                      <th rowSpan={2}>หลักฐาน Report</th>
                    )}

                    {(visibleColumns.includes("hclReport") ||
                      visibleColumns.includes("hclStatus")) && (
                      <th
                        className="tool-head hcl"
                        colSpan={
                          Number(
                            visibleColumns.includes("hclReport")
                          ) +
                          Number(
                            visibleColumns.includes("hclStatus")
                          )
                        }
                      >
                        HCL(OP01)
                      </th>
                    )}

                    {(visibleColumns.includes("picusReport") ||
                      visibleColumns.includes("picusStatus")) && (
                      <th
                        className="tool-head picus"
                        colSpan={
                          Number(
                            visibleColumns.includes(
                              "picusReport"
                            )
                          ) +
                          Number(
                            visibleColumns.includes(
                              "picusStatus"
                            )
                          )
                        }
                      >
                        PICUS(OP04)
                      </th>
                    )}

                    {(visibleColumns.includes(
                      "proofpointReport"
                    ) ||
                      visibleColumns.includes(
                        "proofpointStatus"
                      )) && (
                      <th
                        className="tool-head proofpoint"
                        colSpan={
                          Number(
                            visibleColumns.includes(
                              "proofpointReport"
                            )
                          ) +
                          Number(
                            visibleColumns.includes(
                              "proofpointStatus"
                            )
                          )
                        }
                      >
                        PROOFPOINT(OP03)
                      </th>
                    )}

                    {(visibleColumns.includes("secironReport") ||
                      visibleColumns.includes(
                        "secironStatus"
                      )) && (
                      <th
                        className="tool-head seciron"
                        colSpan={
                          Number(
                            visibleColumns.includes(
                              "secironReport"
                            )
                          ) +
                          Number(
                            visibleColumns.includes(
                              "secironStatus"
                            )
                          )
                        }
                      >
                        SECIRON(OP02)
                      </th>
                    )}

                    {visibleColumns.includes("closingDate") && (
                      <th rowSpan={2}>
                        วันนัดหมายสรุปปิดเคส
                      </th>
                    )}

                    {visibleColumns.includes(
                      "satisfaction"
                    ) && (
                      <th rowSpan={2}>
                        แบบประเมินความพึงพอใจ
                      </th>
                    )}

                    {visibleColumns.includes("score") && (
                      <th rowSpan={2}>คะแนน</th>
                    )}

                    {visibleColumns.includes("provider") && (
                      <th rowSpan={2}>Provider</th>
                    )}

                    {visibleColumns.includes("totalReports") && (
                      <th rowSpan={2}>รวมทั้งหมด</th>
                    )}

                    {visibleColumns.includes("completedReports") && (
                      <th rowSpan={2}>
                        จำนวนรายงานดำเนินการเสร็จสิ้น
                      </th>
                    )}

                    {visibleColumns.includes("incompleteReports") && (
                      <th rowSpan={2}>
                        จำนวนรายงานยังไม่เสร็จสิ้น
                      </th>
                    )}
                  </tr>

                  <tr className="report-sub-header">
                    {visibleColumns.includes("hclReport") && (
                      <th className="hcl">จำนวน Report</th>
                    )}
                    {visibleColumns.includes("hclStatus") && (
                      <th className="hcl">สถานะ</th>
                    )}

                    {visibleColumns.includes("picusReport") && (
                      <th className="picus">จำนวน Report</th>
                    )}
                    {visibleColumns.includes("picusStatus") && (
                      <th className="picus">สถานะ</th>
                    )}

                    {visibleColumns.includes(
                      "proofpointReport"
                    ) && (
                      <th className="proofpoint">
                        จำนวน Report
                      </th>
                    )}
                    {visibleColumns.includes(
                      "proofpointStatus"
                    ) && (
                      <th className="proofpoint">สถานะ</th>
                    )}

                    {visibleColumns.includes("secironReport") && (
                      <th className="seciron">
                        จำนวน Report
                      </th>
                    )}
                    {visibleColumns.includes("secironStatus") && (
                      <th className="seciron">สถานะ</th>
                    )}
                  </tr>
                </thead>

                <tbody>
                  {filteredRows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={Math.max(
                          visibleColumns.length,
                          1
                        )}
                        className="report-empty-cell"
                      >
                        ไม่พบข้อมูล
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map((row, rowIndex) => {
                      /*
                        ตารางสรุป Provider ใน Excel เป็นคนละชุดกับ
                        Provider ของข้อมูลรายหน่วยงาน

                        จึงเอา HCL / PICUS / PROOFPOINT / SECIRON
                        มาเรียงใน 4 แถวแรกโดยตรง เพื่อให้ตรงกับ
                        ตาราง Summary ด้านขวาของ Excel
                      */
                      const summary =
                        rowIndex < toolSummary.length
                          ? toolSummary[rowIndex]
                          : null;

                      return (
                        <tr key={row.sourceRow}>
                          {COLUMNS.filter((column) =>
                            visibleColumns.includes(column.key)
                          ).map((column) => {
                            if (column.key === "provider") {
                              return (
                                <td key={column.key}>
                                  {summary?.name ?? "-"}
                                </td>
                              );
                            }

                            if (column.key === "totalReports") {
                              return (
                                <td key={column.key}>
                                  {summary?.total ?? "-"}
                                </td>
                              );
                            }

                            if (column.key === "completedReports") {
                              return (
                                <td key={column.key}>
                                  {summary?.completed ?? "-"}
                                </td>
                              );
                            }

                            if (column.key === "incompleteReports") {
                              return (
                                <td key={column.key}>
                                  {summary?.incomplete ?? "-"}
                                </td>
                              );
                            }

                            return (
                              <td key={column.key}>
                                {renderCell(row, column.key)}
                              </td>
                            );
                          })}
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
          <div className="report-chart-area">
            <div className="report-bar-filter">
              <span>เลือกข้อมูลที่ต้องการแสดง</span>

              <label>
                <input
                  type="checkbox"
                  checked={visibleBarMetrics.includes("status")}
                  onChange={() => toggleBarMetric("status")}
                />
                สถานะเสร็จ / ยังไม่เสร็จ
              </label>

              <label>
                <input
                  type="checkbox"
                  checked={visibleBarMetrics.includes("project")}
                  onChange={() => toggleBarMetric("project")}
                />
                ตัวชี้วัดโครงการ
              </label>

              <label>
                <input
                  type="checkbox"
                  checked={visibleBarMetrics.includes("personal")}
                  onChange={() => toggleBarMetric("personal")}
                />
                ตัวชี้วัดรายบุคคล
              </label>
            </div>

            {visibleBarMetrics.length === 0 ? (
              <div className="report-bar-empty">
                กรุณาเลือกอย่างน้อย 1 รายการ
              </div>
            ) : (
              (() => {
                const globalMax = Math.max(
                  1,
                  ...toolSummary.flatMap((tool) => {
                    const values: number[] = [];

                    if (visibleBarMetrics.includes("status")) {
                      values.push(tool.total);
                    }

                    if (visibleBarMetrics.includes("project")) {
                      values.push(tool.projectTotal);
                    }

                    if (visibleBarMetrics.includes("personal")) {
                      values.push(tool.personalTotal);
                    }

                    return values;
                  })
                );

                return (
                  <>
                    <div className="report-stack-chart">
                      {toolSummary.map((tool) => {
                        const completedHeight =
                          (tool.completed / globalMax) * 100;

                        const incompleteHeight =
                          (tool.incomplete / globalMax) * 100;

                        const projectHeight =
                          (tool.projectTotal / globalMax) * 100;

                        const personalHeight =
                          (tool.personalTotal / globalMax) * 100;

                        return (
                          <div
                            key={tool.key}
                            className="report-stack-item"
                          >
                            <div className="report-bar-values">
                              {visibleBarMetrics.includes("status") && (
                                <strong>
                                  {tool.completed} / {tool.incomplete}
                                </strong>
                              )}

                              {visibleBarMetrics.includes("project") && (
                                <span>
                                  โครงการ: {tool.projectTotal}
                                </span>
                              )}

                              {visibleBarMetrics.includes("personal") && (
                                <span>
                                  รายบุคคล: {tool.personalTotal}
                                </span>
                              )}
                            </div>

                            <div className="report-bar-group">
                              {visibleBarMetrics.includes("status") && (
                                <div className="report-bar-series">
                                  <div className="report-bar-track">
                                    <div
                                      className="report-stack-incomplete"
                                      style={{
                                        height: `${incompleteHeight}%`,
                                      }}
                                    />

                                    <div
                                      className="report-stack-completed"
                                      style={{
                                        height: `${completedHeight}%`,
                                      }}
                                    />
                                  </div>

                                  <small>สถานะ</small>
                                </div>
                              )}

                              {visibleBarMetrics.includes("project") && (
                                <div className="report-bar-series">
                                  <div className="report-bar-track">
                                    <div
                                      className="report-project-total"
                                      style={{
                                        height: `${projectHeight}%`,
                                      }}
                                    />
                                  </div>

                                  <small>โครงการ</small>
                                </div>
                              )}

                              {visibleBarMetrics.includes("personal") && (
                                <div className="report-bar-series">
                                  <div className="report-bar-track">
                                    <div
                                      className="report-personal-total"
                                      style={{
                                        height: `${personalHeight}%`,
                                      }}
                                    />
                                  </div>

                                  <small>รายบุคคล</small>
                                </div>
                              )}
                            </div>

                            <strong className="report-tool-name">
                              {tool.name}
                            </strong>
                          </div>
                        );
                      })}
                    </div>

                    <div className="report-chart-legend">
                      {visibleBarMetrics.includes("status") && (
                        <>
                          <span>
                            <i className="report-legend-completed" />
                            ดำเนินการเสร็จสิ้น
                          </span>

                          <span>
                            <i className="report-legend-incomplete" />
                            ยังไม่เสร็จสิ้น
                          </span>
                        </>
                      )}

                      {visibleBarMetrics.includes("project") && (
                        <span>
                          <i className="report-legend-project" />
                          ตัวชี้วัดโครงการ
                        </span>
                      )}

                      {visibleBarMetrics.includes("personal") && (
                        <span>
                          <i className="report-legend-personal" />
                          ตัวชี้วัดรายบุคคล
                        </span>
                      )}
                    </div>
                  </>
                );
              })()
            )}
          </div>
        )}

        {viewType === "pie" && (
          <div className="report-pie-grid">
            {toolSummary.map((tool) => {
              const percent =
                tool.total > 0
                  ? (tool.completed / tool.total) * 100
                  : 0;

              return (
                <article
                  key={tool.key}
                  className="report-pie-card"
                >
                  <div
                    className="report-pie"
                    style={{
                      background: `conic-gradient(
                        #2fa866 0 ${percent}%,
                        #f2b522 ${percent}% 100%
                      )`,
                    }}
                  >
                    <div className="report-pie-hole">
                      <strong>
                        {Math.round(percent)}%
                      </strong>
                    </div>
                  </div>

                  <h3>{tool.name}</h3>

                  <p>
                    เสร็จ {tool.completed} / ยังไม่เสร็จ{" "}
                    {tool.incomplete}
                  </p>
                </article>
              );
            })}
          </div>
        )}

        {viewType === "progress" && (
          <div className="report-progress-list">
            {filteredRows.map((row) => {
              const percent = rowProgress(row);

              return (
                <article
                  key={row.sourceRow}
                  className="report-progress-item"
                >
                  <div className="report-progress-heading">
                    <div>
                      <strong>
                        {row.order}. {row.organization}
                      </strong>

                      <span>
                        เลขเคส: {row.caseId || "-"}
                      </span>
                    </div>

                    <b>{percent}%</b>
                  </div>

                  <div className="report-progress-track">
                    <div
                      className="report-progress-fill"
                      style={{
                        width: `${percent}%`,
                      }}
                    />
                  </div>
                </article>
              );
            })}
          </div>
        )}
      </section>
    </section>
  );
}

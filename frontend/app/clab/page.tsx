"use client";

import {
  ChangeEvent,
  useCallback,
  useEffect,
  useMemo,
  useState,
} from "react";

import * as XLSX from "xlsx";

import "./clab.css";
import OverviewVA from "./overview-va/OverviewVA";
import Report from "./report/ReportView";
import ToolOP02 from "./tool-op02/ToolOP02";
import ToolOP03 from "./tool-op03/ToolOP03";
import WaitingOrganizations from "./waiting-organizations/WaitingOrganizations";
import EmailService from "./email-service/EmailService";

type RawExcelRow =
  Record<string, unknown>;

type UserInfo = {
  id: number;
  name: string;
  email: string;
  role:
    | "super_admin"
    | "editor"
    | "viewer";
};

type ClabType =
  | "overview_va"
  | "report"
  | "tool_op02"
  | "tool_op03"
  | "waiting_organizations"
  | "email_service";

type DetectedSheet = {
  sheetName: string;
  type: ClabType;
  label: string;
  confidence: number;
};

type LatestSheet = {
  id: number;
  snapshot_id: number;
  original_sheet_name:
    string;
  detected_type: ClabType;
  detected_label: string;
  confidence: number;
  headers: string[];
  rows: unknown[][];
  created_at: string;
  updated_at?: string | null;
  original_file_name?: string | null;
  saved_by?: number | null;
  saved_by_name?: string | null;
};

type LatestResponse = {
  snapshot: {
    id: number;
    original_file_name:
      | string
      | null;
    saved_by: number | null;
    saved_by_name:
      | string
      | null;
    created_at: string;
  } | null;

  sheets: LatestSheet[];
};

type PendingSave = {
  sheetName: string;
  detected: DetectedSheet;
  rows: unknown[][];
  headers: string[];
};

type EditSavedSheet = {
  id: number;
  originalSheetName: string;
  detectedLabel: string;
  confidence: number;
  rows: unknown[][];
  headers: string[];
};

const API_URL =
  process.env
    .NEXT_PUBLIC_API_URL &&
  !process.env
    .NEXT_PUBLIC_API_URL.includes(
      "backend:"
    )
    ? process.env
        .NEXT_PUBLIC_API_URL
        .replace(
          /\/$/,
          ""
        )
    : "http://localhost:4000";

const TYPE_LABELS:
  Record<
    ClabType,
    string
  > = {
  overview_va:
    "Overview VA",
  report:
    "ติดตาม Report",
  tool_op02:
    "Tool OP02",
  tool_op03:
    "Tool OP03",
  waiting_organizations:
    "หน่วยงานที่รอดำเนินการ",
  email_service:
    "บริการอีเมล",
};

function readCurrentUser():
  UserInfo | null {
  if (
    typeof window ===
    "undefined"
  ) {
    return null;
  }

  const raw =
    window.localStorage.getItem(
      "user"
    );

  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(
      raw
    ) as UserInfo;
  } catch {
    return null;
  }
}

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
    .replace(
      /\u00A0/g,
      " "
    )
    .replace(
      /\s+/g,
      " "
    )
    .trim();
}

function normalize(
  value: unknown
) {
  return cleanText(
    value
  ).toLowerCase();
}

function worksheetToMatrix(
  worksheet:
    XLSX.WorkSheet
): unknown[][] {
  const ref =
    worksheet["!ref"];

  if (!ref) {
    return [];
  }

  const range =
    XLSX.utils.decode_range(
      ref
    );

  const matrix: unknown[][] = [];

  /*
    อ่านค่าตามตำแหน่ง Cell จริงทั้งหมด
    ห้ามใช้ sheet_to_json(... blankrows:false) ตรงนี้
    เพราะมันอาจทำให้โครงสร้าง Header หลายชั้นของ Report
    เปลี่ยนตำแหน่งตอนบันทึก/เปิดใหม่
  */
  for (
    let r = range.s.r;
    r <= range.e.r;
    r += 1
  ) {
    const row: unknown[] = [];

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

  /*
    สำคัญมากสำหรับ Sheet "ติดตาม Report"

    Header HCL / PICUS / PROOFPOINT / SECIRON
    เป็น merged cell ครอบ 2 คอลัมน์:
      - จำนวน Report
      - สถานะ

    ตอนเก็บ JSONB จะไม่มี !merges ติดไปด้วย
    จึงต้องกระจายค่าของ merged cell ลงทุกช่องก่อนบันทึก
    ไม่อย่างนั้นตอนเปิด Saved Data จะหา column "สถานะ" ไม่เจอ
    และ Summary จะกลายเป็น 0/...
  */
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

/*
  ซ่อมข้อมูล Saved เก่าที่ถูกบันทึกก่อนมีการกระจาย merged cell

  ตัวอย่างเดิม:
    HCL(OP01) | ""
    จำนวน Report | สถานะ

  แปลงเป็น:
    HCL(OP01) | HCL(OP01)
    จำนวน Report | สถานะ

  เพื่อให้ ReportView หา HCL + สถานะ ได้เหมือนตอน Import Excel ต้นฉบับ
*/
function repairSavedReportHeaders(
  inputRows: unknown[][]
): unknown[][] {
  const rows = inputRows.map(
    (row) =>
      Array.isArray(row)
        ? [...row]
        : []
  );

  const toolPattern =
    /(hcl|picus|proofpoint|seciron)/i;

  for (
    let r = 0;
    r < rows.length - 1;
    r += 1
  ) {
    const topRow =
      rows[r] ?? [];
    const subRow =
      rows[r + 1] ?? [];

    const hasReportHeader =
      subRow.some((value) =>
        normalize(value).includes(
          "จำนวน report"
        )
      );

    const hasStatusHeader =
      subRow.some((value) =>
        normalize(value).includes(
          "สถานะ"
        )
      );

    if (
      !hasReportHeader ||
      !hasStatusHeader
    ) {
      continue;
    }

    for (
      let c = 0;
      c < topRow.length;
      c += 1
    ) {
      const parent =
        cleanText(topRow[c]);

      if (
        !toolPattern.test(
          parent
        )
      ) {
        continue;
      }

      const nextSubHeader =
        normalize(
          subRow[c + 1]
        );

      const nextParent =
        cleanText(
          topRow[c + 1]
        );

      if (
        nextSubHeader.includes(
          "สถานะ"
        ) &&
        !nextParent
      ) {
        topRow[c + 1] =
          parent;
      }
    }

    rows[r] = topRow;
  }

  return rows;
}

function scoreIncludes(
  text: string,
  keywords: string[]
) {
  return keywords.reduce(
    (
      score,
      keyword
    ) =>
      text.includes(
        normalize(keyword)
      )
        ? score + 1
        : score,
    0
  );
}

function detectSheetType(
  sheetName: string,
  worksheet:
    XLSX.WorkSheet
): DetectedSheet | null {
  const matrix =
    worksheetToMatrix(
      worksheet
    );

  const sampleText =
    matrix
      .slice(0, 35)
      .flat()
      .map(normalize)
      .join(" ");

  const normalizedName =
    normalize(
      sheetName
    );

  const candidates:
    Array<{
      type: ClabType;
      score: number;
      max: number;
    }> = [];

  /* Overview VA */
  {
    let score =
      scoreIncludes(
        sampleText,
        [
          "descriptions",
          "progress",
          "note",
          "ติดตาม report",
        ]
      );

    if (
      sampleText.includes(
        "desciptions"
      )
    ) {
      score += 1;
    }

    if (
      normalizedName.includes(
        "overview"
      )
    ) {
      score += 1;
    }

    candidates.push({
      type:
        "overview_va",
      score,
      max: 6,
    });
  }

  /* Report */
  {
    let score =
      scoreIncludes(
        sampleText,
        [
          "ลำดับ",
          "ชื่อหน่วยงาน",
          "ระดับความคืบหน้า",
          "เลขเคส",
          "hcl",
          "picus",
          "proofpoint",
          "seciron",
        ]
      );

    if (
      normalizedName.includes(
        "report"
      )
    ) {
      score += 1;
    }

    candidates.push({
      type: "report",
      score,
      max: 9,
    });
  }

  /* OP02 */
  {
    let score =
      scoreIncludes(
        sampleText,
        [
          "มีแอปพลิเคชัน",
          "ชื่อแอป",
          "ios",
          "android",
          "op2",
        ]
      );

    if (
      normalizedName.includes(
        "op02"
      ) ||
      normalizedName.includes(
        "op2"
      )
    ) {
      score += 1;
    }

    candidates.push({
      type:
        "tool_op02",
      score,
      max: 6,
    });
  }

  /* OP03 */
  {
    let score =
      scoreIncludes(
        sampleText,
        [
          "ชื่อหน่วยงาน",
          "allow",
          "dga",
          "365",
          "user",
          "template",
        ]
      );

    if (
      normalizedName.includes(
        "op03"
      ) ||
      normalizedName.includes(
        "op3"
      )
    ) {
      score += 1;
    }

    candidates.push({
      type:
        "tool_op03",
      score,
      max: 7,
    });
  }

  /* Waiting Organizations */
  {
    let score =
      scoreIncludes(
        sampleText,
        [
          "ชื่อหน่วยงาน",
          "หนังสือตอบรับ",
          "kick-off",
          "hcl",
          "picus",
          "proofpoint",
          "seciron",
        ]
      );

    if (
      normalizedName.includes(
        "รอดำเนินการ"
      )
    ) {
      score += 2;
    }

    candidates.push({
      type:
        "waiting_organizations",
      score,
      max: 9,
    });
  }

  /* Email Service */
  {
    let score =
      scoreIncludes(
        sampleText,
        [
          "ชื่อหน่วยงาน",
          "อีเมล",
          "เบอร์โทรติดต่อ",
          "วันที่ติดต่อประสานงาน",
          "หนังสือตอบรับ",
          "ได้รับแบบคำขอรับการทดสอบ",
          "ส่งอีเมล หนังสือ",
          "วันนัดหมายประชุม",
        ]
      );

    if (
      normalizedName.includes(
        "อีเมล"
      ) ||
      normalizedName.includes(
        "email"
      )
    ) {
      score += 1;
    }

    candidates.push({
      type:
        "email_service",
      score,
      max: 9,
    });
  }

  candidates.sort(
    (a, b) =>
      b.score / b.max -
      a.score / a.max
  );

  const best =
    candidates[0];

  if (
    !best ||
    best.score < 3
  ) {
    return null;
  }

  const confidence =
    Math.min(
      99,
      Math.max(
        50,
        Math.round(
          (best.score /
            best.max) *
            100
        )
      )
    );

  return {
    sheetName,
    type:
      best.type,
    label:
      TYPE_LABELS[
        best.type
      ],
    confidence,
  };
}

function headersFromMatrix(
  rows: unknown[][]
) {
  const firstUsefulRow =
    rows.find(
      (row) =>
        row.some(
          (value) =>
            cleanText(
              value
            ) !== ""
        )
    );

  return (
    firstUsefulRow?.map(
      (value) =>
        cleanText(value)
    ) ?? []
  );
}

function formatDate(
  value:
    | string
    | null
    | undefined
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

  return date.toLocaleString(
    "th-TH",
    {
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }
  );
}

export default function ClabPage() {
  const [
    currentUser,
    setCurrentUser,
  ] =
    useState<UserInfo | null>(
      null
    );

  const canManage =
    currentUser?.role ===
      "super_admin" ||
    currentUser?.role ===
      "editor";

  const [
    rawRows,
    setRawRows,
  ] =
    useState<
      RawExcelRow[]
    >([]);

  const [
    rawHeaders,
    setRawHeaders,
  ] =
    useState<string[]>(
      []
    );

  const [
    fileName,
    setFileName,
  ] =
    useState("");

  const [
    search,
    setSearch,
  ] =
    useState("");

  const [
    showColumnPicker,
    setShowColumnPicker,
  ] =
    useState(false);

  const [
    visibleRawColumns,
    setVisibleRawColumns,
  ] =
    useState<string[]>(
      []
    );

  const [
    workbook,
    setWorkbook,
  ] =
    useState<
      XLSX.WorkBook | null
    >(null);

  const [
    sheetNames,
    setSheetNames,
  ] =
    useState<string[]>(
      []
    );

  const [
    detectedSheets,
    setDetectedSheets,
  ] =
    useState<
      Record<
        string,
        DetectedSheet
      >
    >({});

  const [
    selectedSheet,
    setSelectedSheet,
  ] =
    useState("");

  const [
    loadedSheet,
    setLoadedSheet,
  ] =
    useState("");

  const [
    sourceMode,
    setSourceMode,
  ] =
    useState<
      "saved" | "local"
    >("saved");

  const [
    latestSnapshot,
    setLatestSnapshot,
  ] =
    useState<
      LatestResponse["snapshot"]
    >(null);

  const [
    savedSheets,
    setSavedSheets,
  ] =
    useState<LatestSheet[]>(
      []
    );

  const [
    loadingLatest,
    setLoadingLatest,
  ] =
    useState(true);

  const [
    pendingSave,
    setPendingSave,
  ] =
    useState<PendingSave | null>(
      null
    );

  const [
    editingSavedSheet,
    setEditingSavedSheet,
  ] =
    useState<EditSavedSheet | null>(
      null
    );

  const [
    savingEdit,
    setSavingEdit,
  ] =
    useState(false);

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

  const loadLatest =
    useCallback(
      async () => {
        try {
          setLoadingLatest(
            true
          );
          setError("");

          const response =
            await fetch(
              `${API_URL}/api/clab/latest-by-type`,
              {
                cache:
                  "no-store",
              }
            );

          if (!response.ok) {
            throw new Error(
              "โหลดข้อมูล CLAB ที่บันทึกไว้ไม่สำเร็จ"
            );
          }

          const data =
            (await response.json()) as {
              sheets: LatestSheet[];
            };

          const nextSheets =
            Array.isArray(
              data.sheets
            )
              ? data.sheets
              : [];

          setSavedSheets(
            nextSheets
          );

          const newest =
            [...nextSheets].sort(
              (a, b) =>
                new Date(
                  b.updated_at ??
                    b.created_at
                ).getTime() -
                new Date(
                  a.updated_at ??
                    a.created_at
                ).getTime()
            )[0];

          setLatestSnapshot(
            newest
              ? {
                  id:
                    newest.snapshot_id,
                  original_file_name:
                    newest.original_file_name ??
                    null,
                  saved_by:
                    newest.saved_by ??
                    null,
                  saved_by_name:
                    newest.saved_by_name ??
                    null,
                  created_at:
                    newest.updated_at ??
                    newest.created_at,
                }
              : null
          );
        } catch (
          loadError
        ) {
          console.error(
            "Load saved CLAB error:",
            loadError
          );

          setError(
            loadError instanceof
              Error
              ? loadError.message
              : "โหลดข้อมูลไม่สำเร็จ"
          );
        } finally {
          setLoadingLatest(
            false
          );
        }
      },
      []
    );

  useEffect(() => {
    setCurrentUser(
      readCurrentUser()
    );

    void loadLatest();
  }, [loadLatest]);

  async function handleFileChange(
    event:
      ChangeEvent<HTMLInputElement>
  ) {
    const file =
      event.target.files?.[0];

    if (!file) {
      return;
    }

    try {
      setError("");
      setMessage("");

      const buffer =
        await file.arrayBuffer();

      const loadedWorkbook =
        XLSX.read(
          buffer,
          {
            type: "array",
          }
        );

      const detected:
        Record<
          string,
          DetectedSheet
        > = {};

      for (
        const sheetName
        of loadedWorkbook
          .SheetNames
      ) {
        const worksheet =
          loadedWorkbook
            .Sheets[
              sheetName
            ];

        if (!worksheet) {
          continue;
        }

        const match =
          detectSheetType(
            sheetName,
            worksheet
          );

        if (match) {
          detected[
            sheetName
          ] = match;
        }
      }

      const available =
        loadedWorkbook
          .SheetNames
          .filter(
            (sheetName) =>
              Boolean(
                detected[
                  sheetName
                ]
              )
          );

      if (
        available.length ===
        0
      ) {
        alert(
          "ระบบยังจำแนกประเภท Sheet ในไฟล์นี้ไม่ได้"
        );

        return;
      }

      setWorkbook(
        loadedWorkbook
      );
      setFileName(
        file.name
      );
      setDetectedSheets(
        detected
      );
      setSheetNames(
        available
      );

      setSourceMode(
        "local"
      );

      setSelectedSheet(
        ""
      );
      setLoadedSheet(
        ""
      );
      setRawRows([]);
      setRawHeaders([]);
      setVisibleRawColumns(
        []
      );
      setShowColumnPicker(
        false
      );
      setSearch("");

      setMessage(
        `ตรวจพบ ${available.length} Sheet ที่รองรับ`
      );
    } catch (importError) {
      console.error(
        "CLAB import error:",
        importError
      );

      alert(
        "อ่านไฟล์ Excel ไม่สำเร็จ"
      );
    } finally {
      event.target.value =
        "";
    }
  }

  function handleLoadSheet() {
    if (
      !workbook ||
      !selectedSheet
    ) {
      return;
    }

    const worksheet =
      workbook.Sheets[
        selectedSheet
      ];

    if (!worksheet) {
      alert(
        "ไม่พบ Sheet ที่เลือก"
      );
      return;
    }

    try {
      const excelRows =
        XLSX.utils
          .sheet_to_json<
            RawExcelRow
          >(worksheet, {
            defval: "",
            raw: false,
          });

      const headers =
        excelRows.length >
        0
          ? Object.keys(
              excelRows[0]
            )
          : [];

      setLoadedSheet(
        selectedSheet
      );
      setRawRows(
        excelRows
      );
      setRawHeaders(
        headers
      );
      setVisibleRawColumns(
        headers
      );
      setShowColumnPicker(
        false
      );
      setSearch("");
    } catch (loadError) {
      console.error(
        "Load sheet error:",
        loadError
      );

      alert(
        "โหลดข้อมูลจาก Sheet ไม่สำเร็จ"
      );
    }
  }

  function requestSaveSelectedSheet() {
    if (
      !canManage ||
      !workbook ||
      sourceMode !== "local" ||
      !selectedSheet
    ) {
      return;
    }

    const worksheet =
      workbook.Sheets[
        selectedSheet
      ];

    const detected =
      detectedSheets[
        selectedSheet
      ];

    if (
      !worksheet ||
      !detected
    ) {
      alert(
        "ไม่พบข้อมูล Sheet ที่เลือก"
      );
      return;
    }

    const rows =
      worksheetToMatrix(
        worksheet
      );

    setPendingSave({
      sheetName:
        selectedSheet,
      detected,
      rows,
      headers:
        headersFromMatrix(
          rows
        ),
    });
  }

  async function confirmSaveSelectedSheet() {
    if (
      !canManage ||
      !currentUser ||
      !pendingSave
    ) {
      return;
    }

    try {
      setSaving(true);
      setError("");
      setMessage("");

      const response =
        await fetch(
          `${API_URL}/api/clab/sheets`,
          {
            method: "POST",
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
                fileName,
                sheet: {
                  originalSheetName:
                    pendingSave.sheetName,
                  detectedType:
                    pendingSave.detected.type,
                  detectedLabel:
                    pendingSave.detected.label,
                  confidence:
                    pendingSave.detected.confidence,
                  headers:
                    pendingSave.headers,
                  rows:
                    pendingSave.rows,
                },
              }),
          }
        );

      const result =
        await response.json();

      if (!response.ok) {
        throw new Error(
          result.error ??
            "บันทึก CLAB ไม่สำเร็จ"
        );
      }

      setMessage(
        `บันทึก ${pendingSave.detected.label} สำเร็จ`
      );

      setPendingSave(
        null
      );

      await loadLatest();
    } catch (
      saveError
    ) {
      setError(
        saveError instanceof
          Error
          ? saveError.message
          : "บันทึก CLAB ไม่สำเร็จ"
      );
    } finally {
      setSaving(false);
    }
  }

  function openSavedSheet(
    sheet: LatestSheet
  ) {
    const nextWorkbook =
      XLSX.utils.book_new();

    /*
      Saved Data ต้องถูกสร้างกลับจาก matrix ทั้งก้อนที่บันทึกไว้
      ไม่ใช้ headers หรือ sheet_to_json เป็น source เพราะ Excel บาง Sheet
      มีหัวตารางหลายชั้น / merged cells / แถวว่างคั่นอยู่
    */
    const savedRows: unknown[][] = Array.isArray(sheet.rows)
      ? sheet.rows.map((row) =>
          Array.isArray(row) ? [...row] : []
        )
      : [];

    const restoredRows =
      sheet.detected_type === "report"
        ? repairSavedReportHeaders(savedRows)
        : savedRows;

    const worksheet = XLSX.utils.aoa_to_sheet(
      restoredRows,
      { cellDates: true }
    );

    /*
      Report เดิมมีหัวกลุ่ม 2 คอลัมน์ เช่น HCL -> จำนวน Report/สถานะ
      ตอน Save เรากระจายชื่อหัวกลุ่มไว้ครบทุก cell แล้ว จึงไม่จำเป็นต้อง
      merge กลับเพื่ออ่านข้อมูล และช่วยให้ ReportView หาแต่ละ column ได้ครบ
    */

    XLSX.utils.book_append_sheet(
      nextWorkbook,
      worksheet,
      sheet.original_sheet_name
    );

    setWorkbook(
      nextWorkbook
    );

    setFileName(
      sheet.original_file_name ??
        "CLAB Saved"
    );

    setDetectedSheets({
      [sheet.original_sheet_name]:
        {
          sheetName:
            sheet.original_sheet_name,
          type:
            sheet.detected_type,
          label:
            sheet.detected_label,
          confidence:
            sheet.confidence,
        },
    });

    setSheetNames([
      sheet.original_sheet_name,
    ]);

    setSelectedSheet(
      sheet.original_sheet_name
    );

    setLoadedSheet(
      sheet.original_sheet_name
    );

    setSourceMode(
      "saved"
    );

    const rows =
      XLSX.utils.sheet_to_json<
        RawExcelRow
      >(worksheet, {
        defval: "",
        raw: false,
      });

    const headers =
      rows.length > 0
        ? Object.keys(
            rows[0]
          )
        : [];

    setRawRows(
      rows
    );
    setRawHeaders(
      headers
    );
    setVisibleRawColumns(
      headers
    );

    window.scrollTo({
      top: 0,
      behavior:
        "smooth",
    });
  }

  function startEditSavedSheet(
    sheet: LatestSheet
  ) {
    if (!canManage) {
      return;
    }

    setEditingSavedSheet({
      id:
        sheet.id,
      originalSheetName:
        sheet.original_sheet_name,
      detectedLabel:
        sheet.detected_label,
      confidence:
        sheet.confidence,
      headers:
        Array.isArray(
          sheet.headers
        )
          ? sheet.headers
          : [],
      rows:
        Array.isArray(
          sheet.rows
        )
          ? sheet.rows.map(
              (row) =>
                Array.isArray(
                  row
                )
                  ? [...row]
                  : []
            )
          : [],
    });
  }

  function updateSavedCell(
    rowIndex: number,
    columnIndex: number,
    value: string
  ) {
    setEditingSavedSheet(
      (current) => {
        if (!current) {
          return current;
        }

        const rows =
          current.rows.map(
            (row) => [
              ...row,
            ]
          );

        if (!rows[rowIndex]) {
          rows[rowIndex] =
            [];
        }

        rows[rowIndex][
          columnIndex
        ] = value;

        return {
          ...current,
          rows,
        };
      }
    );
  }

  async function saveEditedSheet() {
    if (
      !currentUser ||
      !editingSavedSheet ||
      !canManage
    ) {
      return;
    }

    try {
      setSavingEdit(
        true
      );
      setError("");

      const response =
        await fetch(
          `${API_URL}/api/clab/sheets/${editingSavedSheet.id}`,
          {
            method: "PUT",
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
                originalSheetName:
                  editingSavedSheet.originalSheetName,
                detectedLabel:
                  editingSavedSheet.detectedLabel,
                confidence:
                  editingSavedSheet.confidence,
                headers:
                  editingSavedSheet.headers,
                rows:
                  editingSavedSheet.rows,
              }),
          }
        );

      const result =
        await response.json();

      if (!response.ok) {
        throw new Error(
          result.error ??
            "แก้ไขข้อมูล CLAB ไม่สำเร็จ"
        );
      }

      setMessage(
        `แก้ไข ${editingSavedSheet.detectedLabel} สำเร็จ`
      );

      setEditingSavedSheet(
        null
      );

      await loadLatest();
    } catch (
      editError
    ) {
      setError(
        editError instanceof
          Error
          ? editError.message
          : "แก้ไขข้อมูล CLAB ไม่สำเร็จ"
      );
    } finally {
      setSavingEdit(
        false
      );
    }
  }

  const loadedType =
    loadedSheet
      ? detectedSheets[
          loadedSheet
        ]?.type
      : undefined;

  const filteredRawRows =
    useMemo(() => {
      const keyword =
        search
          .trim()
          .toLowerCase();

      if (!keyword) {
        return rawRows;
      }

      return rawRows.filter(
        (row) =>
          Object.values(
            row
          ).some(
            (value) =>
              String(
                value ?? ""
              )
                .toLowerCase()
                .includes(
                  keyword
                )
          )
      );
    }, [
      rawRows,
      search,
    ]);

  function toggleRawColumn(
    column: string
  ) {
    setVisibleRawColumns(
      (current) =>
        current.includes(
          column
        )
          ? current.filter(
              (item) =>
                item !== column
            )
          : [
              ...current,
              column,
            ]
    );
  }

  function selectAllRawColumns() {
    setVisibleRawColumns(
      rawHeaders
    );
  }

  function clearRawColumns() {
    setVisibleRawColumns(
      []
    );
  }

  const selectedMeta =
    selectedSheet
      ? detectedSheets[
          selectedSheet
        ]
      : undefined;

  return (
    <main className="clab-page">
      <div className="clab-container">
        <section className="clab-heading">
          <div>
            <h1>
              CLAB
            </h1>

            <p className="clab-subtitle">
              {loadedSheet
                ? detectedSheets[
                    loadedSheet
                  ]?.label ??
                  loadedSheet
                : latestSnapshot
                  ? "ข้อมูล CLAB ที่บันทึกล่าสุด"
                  : "เลือกข้อมูลจาก Excel"}
            </p>

            <p className="clab-description">
              {canManage
                ? "ระบบตรวจประเภท Sheet อัตโนมัติจากโครงสร้างข้อมูล ไม่ยึดชื่อไฟล์หรือชื่อ Sheet"
                : "ดูข้อมูล CLAB ล่าสุดที่ผู้ดูแลบันทึกไว้"}
            </p>
          </div>

          <div className="clab-heading-actions">
            {latestSnapshot && (
              <div className="clab-snapshot-badge">
                <span>
                  บันทึกล่าสุด
                </span>

                <strong>
                  {formatDate(
                    latestSnapshot
                      .created_at
                  )}
                </strong>
              </div>
            )}

            {canManage && (
              <label className="clab-import-button">
                ↑ Import Excel

                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={
                    handleFileChange
                  }
                  hidden
                />
              </label>
            )}

            {canManage &&
              sourceMode ===
                "local" &&
              selectedSheet && (
                <button
                  type="button"
                  className="clab-save-button"
                  disabled={
                    saving
                  }
                  onClick={
                    requestSaveSelectedSheet
                  }
                >
                  💾 บันทึกหัวข้อนี้
                </button>
              )}
          </div>
        </section>

        {message && (
          <div className="clab-message success">
            {message}
          </div>
        )}

        {error && (
          <div className="clab-message error">
            {error}
          </div>
        )}

        {/* ข้อมูลที่บันทึกไว้ — ย้ายขึ้นมาไว้ด้านบน */}
        <section className="clab-saved-section clab-saved-section-top">
          <div className="clab-saved-heading">
            <div>
              <h2>ข้อมูลที่บันทึกไว้</h2>

              <p>
                เลือกหัวข้อเพื่อเปิดดูข้อมูลล่าสุด หรือแก้ไขข้อมูลที่บันทึกไว้
              </p>
            </div>

            <span>
              {savedSheets.length} หัวข้อ
            </span>
          </div>

          {loadingLatest ? (
            <div className="clab-saved-empty">
              กำลังโหลดข้อมูล CLAB...
            </div>
          ) : savedSheets.length ===
            0 ? (
            <div className="clab-saved-empty">
              ยังไม่มีข้อมูล CLAB ที่บันทึกไว้
            </div>
          ) : (
            <div className="clab-saved-grid">
              {savedSheets.map(
                (sheet) => (
                  <article
                    key={sheet.id}
                    className="clab-saved-card"
                  >
                    <div className="clab-saved-card-top">
                      <span>
                        {sheet.detected_label}
                      </span>

                      <b>
                        {sheet.confidence}%
                      </b>
                    </div>

                    <h3>
                      {sheet.original_sheet_name}
                    </h3>

                    <p>
                      ไฟล์:{" "}
                      {sheet.original_file_name ??
                        "-"}
                    </p>

                    <small>
                      อัปเดต{" "}
                      {formatDate(
                        sheet.updated_at ??
                          sheet.created_at
                      )}
                      {sheet.saved_by_name
                        ? ` • ${sheet.saved_by_name}`
                        : ""}
                    </small>

                    <div className="clab-saved-actions">
                      <button
                        type="button"
                        className="view"
                        onClick={() =>
                          openSavedSheet(
                            sheet
                          )
                        }
                      >
                        เปิดดู
                      </button>

                      {canManage && (
                        <button
                          type="button"
                          className="edit"
                          onClick={() =>
                            startEditSavedSheet(
                              sheet
                            )
                          }
                        >
                          แก้ไข
                        </button>
                      )}
                    </div>
                  </article>
                )
              )}
            </div>
          )}
        </section>

        {!loadingLatest &&
          sheetNames.length >
            0 && (
            <section className="clab-sheet-card">
              <div className="clab-sheet-file">
                <span>
                  {sourceMode ===
                  "saved"
                    ? "SAVED SNAPSHOT"
                    : "FILE"}
                </span>

                <strong>
                  {fileName}
                </strong>
              </div>

              <div className="clab-sheet-actions">
                <select
                  className="clab-sheet-select"
                  value={
                    selectedSheet
                  }
                  onChange={(
                    event
                  ) =>
                    setSelectedSheet(
                      event
                        .target
                        .value
                    )
                  }
                >
                  <option value="">
                    -- เลือกหัวข้อ --
                  </option>

                  {sheetNames.map(
                    (sheet) => {
                      const info =
                        detectedSheets[
                          sheet
                        ];

                      return (
                        <option
                          key={
                            sheet
                          }
                          value={
                            sheet
                          }
                        >
                          {info
                            ? `${info.label} — ${sheet} (${info.confidence}%)`
                            : sheet}
                        </option>
                      );
                    }
                  )}
                </select>

                <button
                  type="button"
                  className="clab-load-sheet-button"
                  disabled={
                    !selectedSheet
                  }
                  onClick={
                    handleLoadSheet
                  }
                >
                  เปิดดูข้อมูล
                </button>
              </div>
            </section>
          )}

        {selectedMeta &&
          !loadedSheet && (
          <section className="clab-detect-card">
            <div>
              <span>
                ระบบตรวจพบ
              </span>

              <strong>
                {
                  selectedMeta.label
                }
              </strong>
            </div>

            <b>
              ความมั่นใจ{" "}
              {
                selectedMeta.confidence
              }
              %
            </b>
          </section>
        )}

        {!loadingLatest &&
          sheetNames.length ===
            0 && (
            <section className="clab-welcome-card">
              <div className="clab-welcome-icon">
                ▤
              </div>

              <h2>
                อัปโหลดไฟล์
              </h2>

              <p>
                {canManage
                  ? "อัปโหลดไฟล์ Excel เพื่อเริ่มใช้งานข้อมูล CLAB"
                  : "รอ Super Admin หรือ Admin อัปโหลดไฟล์ข้อมูล CLAB"}
              </p>
            </section>
          )}

        {sheetNames.length >
          0 &&
          !loadedSheet && (
          <section className="clab-welcome-card">
            <div className="clab-welcome-icon">
              ▤
            </div>

            <h2>
              เลือกหัวข้อที่ต้องการดู
            </h2>

            <p>
              ระบบแยกหัวข้อจากโครงสร้างข้อมูลใน Excel ให้อัตโนมัติ
            </p>
          </section>
        )}

        {loadedSheet &&
          workbook && (
          <>
            {loadedType ===
              "overview_va" && (
              <OverviewVA
                worksheet={
                  workbook
                    .Sheets[
                      loadedSheet
                    ]
                }
                sheetName={
                  loadedSheet
                }
              />
            )}

            {loadedType ===
              "report" && (
              <Report
                worksheet={
                  workbook
                    .Sheets[
                      loadedSheet
                    ]
                }
                sheetName={
                  loadedSheet
                }
              />
            )}

            {loadedType ===
              "tool_op02" && (
              <ToolOP02
                worksheet={
                  workbook
                    .Sheets[
                      loadedSheet
                    ]
                }
                sheetName={
                  loadedSheet
                }
              />
            )}

            {loadedType ===
              "tool_op03" && (
              <ToolOP03
                worksheet={
                  workbook
                    .Sheets[
                      loadedSheet
                    ]
                }
                sheetName={
                  loadedSheet
                }
              />
            )}

            {loadedType ===
              "waiting_organizations" && (
              <WaitingOrganizations
                worksheet={
                  workbook
                    .Sheets[
                      loadedSheet
                    ]
                }
                sheetName={
                  loadedSheet
                }
              />
            )}

            {loadedType ===
              "email_service" && (
              <EmailService
                worksheet={
                  workbook
                    .Sheets[
                      loadedSheet
                    ]
                }
                sheetName={
                  loadedSheet
                }
              />
            )}

            {!loadedType && (
              <>
                <section className="clab-toolbar">
                  <div className="clab-search">
                    <span>
                      ⌕
                    </span>

                    <input
                      type="text"
                      placeholder="Search..."
                      value={
                        search
                      }
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

                  <div className="clab-toolbar-right">
                    <span className="clab-total">
                      {
                        rawRows.length
                      }{" "}
                      รายการ
                    </span>

                    <div className="clab-column-picker-wrapper">
                      <button
                        type="button"
                        className="clab-column-button"
                        onClick={() =>
                          setShowColumnPicker(
                            !showColumnPicker
                          )
                        }
                      >
                        ☷ เลือกคอลัมน์
                      </button>

                      {showColumnPicker && (
                        <div className="clab-column-menu">
                          <div className="clab-column-menu-header">
                            <strong>
                              เลือกหัวข้อที่ต้องการแสดง
                            </strong>

                            <button
                              type="button"
                              onClick={() =>
                                setShowColumnPicker(
                                  false
                                )
                              }
                            >
                              ×
                            </button>
                          </div>

                          <div className="clab-column-actions">
                            <button
                              type="button"
                              onClick={
                                selectAllRawColumns
                              }
                            >
                              เลือกทั้งหมด
                            </button>

                            <button
                              type="button"
                              onClick={
                                clearRawColumns
                              }
                            >
                              ล้างทั้งหมด
                            </button>
                          </div>

                          <div className="clab-column-list">
                            {rawHeaders.map(
                              (
                                header
                              ) => (
                                <label
                                  key={
                                    header
                                  }
                                  className="clab-column-option"
                                >
                                  <input
                                    type="checkbox"
                                    checked={visibleRawColumns.includes(
                                      header
                                    )}
                                    onChange={() =>
                                      toggleRawColumn(
                                        header
                                      )
                                    }
                                  />

                                  <span>
                                    {
                                      header
                                    }
                                  </span>
                                </label>
                              )
                            )}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                </section>

                <section className="clab-main-card">
                  <div className="clab-card-header">
                    <div>
                      <h2>
                        {
                          loadedSheet
                        }
                      </h2>

                      <p>
                        ข้อมูลจาก Sheet ที่เลือก
                      </p>
                    </div>

                    <span className="clab-total">
                      {
                        filteredRawRows.length
                      }{" "}
                      รายการ
                    </span>
                  </div>

                  <div className="clab-table-wrapper">
                    {visibleRawColumns.length ===
                    0 ? (
                      <div className="clab-no-columns">
                        กรุณาเลือกอย่างน้อย 1 คอลัมน์
                      </div>
                    ) : (
                      <table className="clab-table clab-generic-table">
                        <thead>
                          <tr>
                            {rawHeaders
                              .filter(
                                (
                                  header
                                ) =>
                                  visibleRawColumns.includes(
                                    header
                                  )
                              )
                              .map(
                                (
                                  header
                                ) => (
                                  <th
                                    key={
                                      header
                                    }
                                  >
                                    {
                                      header
                                    }
                                  </th>
                                )
                              )}
                          </tr>
                        </thead>

                        <tbody>
                          {filteredRawRows.length ===
                          0 ? (
                            <tr>
                              <td
                                colSpan={Math.max(
                                  visibleRawColumns.length,
                                  1
                                )}
                                className="clab-empty"
                              >
                                ไม่มีข้อมูล
                              </td>
                            </tr>
                          ) : (
                            filteredRawRows.map(
                              (
                                row,
                                index
                              ) => (
                                <tr
                                  key={
                                    index
                                  }
                                >
                                  {rawHeaders
                                    .filter(
                                      (
                                        header
                                      ) =>
                                        visibleRawColumns.includes(
                                          header
                                        )
                                    )
                                    .map(
                                      (
                                        header
                                      ) => (
                                        <td
                                          key={
                                            header
                                          }
                                        >
                                          {String(
                                            row[
                                              header
                                            ] ??
                                              "-"
                                          )}
                                        </td>
                                      )
                                    )}
                                </tr>
                              )
                            )
                          )}
                        </tbody>
                      </table>
                    )}
                  </div>
                </section>
              </>
            )}
          </>
        )}

        {pendingSave && (
          <div className="clab-modal-backdrop">
            <section className="clab-confirm-modal">
              <div className="clab-modal-header">
                <div>
                  <span>
                    ยืนยันการบันทึก
                  </span>

                  <h2>
                    {
                      pendingSave.detected.label
                    }
                  </h2>
                </div>

                <button
                  type="button"
                  onClick={() =>
                    setPendingSave(
                      null
                    )
                  }
                >
                  ×
                </button>
              </div>

              <div className="clab-confirm-summary">
                <div>
                  <span>
                    ไฟล์
                  </span>

                  <strong>
                    {fileName}
                  </strong>
                </div>

                <div>
                  <span>
                    Sheet
                  </span>

                  <strong>
                    {
                      pendingSave.sheetName
                    }
                  </strong>
                </div>

                <div>
                  <span>
                    ประเภทที่ตรวจพบ
                  </span>

                  <strong>
                    {
                      pendingSave.detected.label
                    }
                  </strong>
                </div>

                <div>
                  <span>
                    ความมั่นใจ
                  </span>

                  <strong>
                    {
                      pendingSave.detected.confidence
                    }
                    %
                  </strong>
                </div>

                <div>
                  <span>
                    จำนวนแถว
                  </span>

                  <strong>
                    {
                      pendingSave.rows.length
                    }
                  </strong>
                </div>
              </div>

              <p className="clab-confirm-note">
                ระบบจะบันทึกเฉพาะหัวข้อนี้เป็นข้อมูลล่าสุด โดยหัวข้ออื่นที่เคยบันทึกไว้จะไม่ถูกลบ
              </p>

              <div className="clab-modal-actions">
                <button
                  type="button"
                  className="cancel"
                  onClick={() =>
                    setPendingSave(
                      null
                    )
                  }
                >
                  ยกเลิก
                </button>

                <button
                  type="button"
                  className="confirm"
                  disabled={
                    saving
                  }
                  onClick={() =>
                    void confirmSaveSelectedSheet()
                  }
                >
                  {saving
                    ? "กำลังบันทึก..."
                    : "ยืนยันบันทึก"}
                </button>
              </div>
            </section>
          </div>
        )}

        {editingSavedSheet && (
          <div className="clab-modal-backdrop">
            <section className="clab-edit-modal">
              <div className="clab-modal-header">
                <div>
                  <span>
                    แก้ไขข้อมูลที่บันทึกแล้ว
                  </span>

                  <h2>
                    {
                      editingSavedSheet.detectedLabel
                    }
                  </h2>
                </div>

                <button
                  type="button"
                  onClick={() =>
                    setEditingSavedSheet(
                      null
                    )
                  }
                >
                  ×
                </button>
              </div>

              <div className="clab-edit-table-wrapper">
                <table className="clab-edit-table">
                  <tbody>
                    {editingSavedSheet.rows.map(
                      (
                        row,
                        rowIndex
                      ) => (
                        <tr
                          key={
                            rowIndex
                          }
                        >
                          {row.map(
                            (
                              value,
                              columnIndex
                            ) => (
                              <td
                                key={
                                  columnIndex
                                }
                              >
                                <input
                                  value={String(
                                    value ??
                                      ""
                                  )}
                                  onChange={(
                                    event
                                  ) =>
                                    updateSavedCell(
                                      rowIndex,
                                      columnIndex,
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
                    )}
                  </tbody>
                </table>
              </div>

              <div className="clab-modal-actions">
                <button
                  type="button"
                  className="cancel"
                  onClick={() =>
                    setEditingSavedSheet(
                      null
                    )
                  }
                >
                  ยกเลิก
                </button>

                <button
                  type="button"
                  className="confirm"
                  disabled={
                    savingEdit
                  }
                  onClick={() =>
                    void saveEditedSheet()
                  }
                >
                  {savingEdit
                    ? "กำลังบันทึก..."
                    : "บันทึกการแก้ไข"}
                </button>
              </div>
            </section>
          </div>
        )}
      </div>
    </main>
  );
}

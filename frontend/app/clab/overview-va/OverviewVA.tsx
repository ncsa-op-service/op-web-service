"use client";

import {
  useMemo,
  useState,
} from "react";

import * as XLSX from "xlsx";

import "./overview-va.css";

type Props = {
  worksheet: XLSX.WorkSheet;
  sheetName: string;
};

type MatrixValue =
  | string
  | number
  | boolean
  | null
  | undefined;

type TimelineCell = {
  columnIndex: number;
  day: number;
  month: string;
};

type SubRow = {
  excelRow: number;

  description: string;
  progress: string;
  note: string;
  report: string;

  timeline: string[];
};

type OrganizationGroup = {
  id: number;
  excelRow: number;

  description: string;
  progress: string;
  note: string;
  report: string;

  timeline: string[];

  children: SubRow[];
};

type ViewType =
  | "table"
  | "progress";

const BASE_COLUMNS = [
  {
    key: "description",
    label: "DESCRIPTIONS",
  },
  {
    key: "progress",
    label: "Progress",
  },
  {
    key: "note",
    label: "NOTE",
  },
  {
    key: "report",
    label: "ติดตาม Report",
  },
] as const;

type BaseColumnKey =
  (typeof BASE_COLUMNS)[number]["key"];

const THAI_MONTHS = [
  "มกราคม",
  "กุมภาพันธ์",
  "มีนาคม",
  "เมษายน",
  "พฤษภาคม",
  "มิถุนายน",
  "กรกฎาคม",
  "สิงหาคม",
  "กันยายน",
  "ตุลาคม",
  "พฤศจิกายน",
  "ธันวาคม",
];

export default function OverviewVA({
  worksheet,
  sheetName,
}: Props) {
  const [search, setSearch] =
    useState("");

  const [viewType, setViewType] =
    useState<ViewType>("table");

  const [
    showColumnPicker,
    setShowColumnPicker,
  ] = useState(false);

  const [
    visibleColumns,
    setVisibleColumns,
  ] = useState<BaseColumnKey[]>(
    BASE_COLUMNS.map(
      (column) => column.key
    )
  );

  const [
    showTimeline,
    setShowTimeline,
  ] = useState(true);

  const [
    expandedGroups,
    setExpandedGroups,
  ] = useState<number[]>([]);

  /* ==============================
     READ EXCEL AS MATRIX
  ============================== */

  const matrix =
    useMemo(() => {
      return XLSX.utils.sheet_to_json<
        MatrixValue[]
      >(worksheet, {
        header: 1,
        defval: "",
        raw: false,
      });
    }, [worksheet]);

  /* ==============================
     FIND HEADER ROW
  ============================== */

  const headerRowIndex =
    useMemo(() => {
      return matrix.findIndex(
        (row) =>
          row.some(
            (cell) => {
              const value =
                String(
                  cell ?? ""
                )
                  .trim()
                  .toUpperCase();

              return (
                value ===
                  "DESCRIPTIONS" ||
                value ===
                  "DESCIPTIONS"
              );
            }
          )
      );
    }, [matrix]);

  /* ==============================
     FIND MAIN COLUMNS
  ============================== */

  const headerInfo =
    useMemo(() => {
      if (
        headerRowIndex < 0
      ) {
        return {
          descriptionIndex: -1,
          progressIndex: -1,
          noteIndex: -1,
          reportIndex: -1,
        };
      }

      const headerRow =
        matrix[
          headerRowIndex
        ] ?? [];

      function findHeader(
        keywords: string[]
      ) {
        return headerRow.findIndex(
          (cell) => {
            const value =
              String(
                cell ?? ""
              )
                .trim()
                .toLowerCase();

            return keywords.some(
              (keyword) =>
                value.includes(
                  keyword.toLowerCase()
                )
            );
          }
        );
      }

      return {
        descriptionIndex:
          findHeader([
            "descriptions",
            "desciptions",
          ]),

        progressIndex:
          findHeader([
            "progress",
          ]),

        noteIndex:
          findHeader([
            "note",
          ]),

        reportIndex:
          findHeader([
            "ติดตาม report",
          ]),
      };
    }, [
      headerRowIndex,
      matrix,
    ]);

  /* ==============================
   FIND DAY ROW

   Overview VA (New):
   วันที่ 1-31 อยู่แถวเดียวกับ
   DESCRIPTIONS / Progress / NOTE / ติดตาม Report
================================ */

const dayRowIndex =
  headerRowIndex;

  const timelineColumns =
    useMemo<
      TimelineCell[]
    >(() => {
      if (
        dayRowIndex < 0
      ) {
        return [];
      }

      const dayRow =
        matrix[
          dayRowIndex
        ] ?? [];

      /*
        หาเดือนจากแถวเหนือวัน
        รองรับ merged cells:
        เดือนจะอยู่ที่ cell แรก
        แล้วให้ carry forward
      */

      const monthRowCandidates =
        [
          dayRowIndex - 1,
          dayRowIndex - 2,
          dayRowIndex - 3,
        ].filter(
          (index) =>
            index >= 0
        );

      const monthByColumn =
        new Map<
          number,
          string
        >();

      for (
        const rowIndex of
        monthRowCandidates
      ) {
        const row =
          matrix[rowIndex];

        if (!row) {
          continue;
        }

        row.forEach(
          (
            cell,
            columnIndex
          ) => {
            const text =
              String(
                cell ?? ""
              ).trim();

            const month =
              THAI_MONTHS.find(
                (item) =>
                  text.includes(
                    item
                  )
              );

            if (
              month &&
              !monthByColumn.has(
                columnIndex
              )
            ) {
              monthByColumn.set(
                columnIndex,
                month
              );
            }
          }
        );
      }

      let currentMonth = "";

      const result:
        TimelineCell[] = [];

      dayRow.forEach(
        (
          cell,
          columnIndex
        ) => {
          /*
            ถ้าคอลัมน์นี้เป็นจุดเริ่มเดือน
            เปลี่ยน currentMonth
          */

          if (
            monthByColumn.has(
              columnIndex
            )
          ) {
            currentMonth =
              monthByColumn.get(
                columnIndex
              ) ?? "";
          }

          const text =
            String(
              cell ?? ""
            ).trim();

          if (
            !/^\d{1,2}$/.test(
              text
            )
          ) {
            return;
          }

          const day =
            Number(text);

          if (
            day < 1 ||
            day > 31
          ) {
            return;
          }

          result.push({
            columnIndex,
            day,
            month:
              currentMonth,
          });
        }
      );

      return result;
    }, [
      matrix,
      dayRowIndex,
    ]);

  /* ==============================
     MONTH GROUPS
  ============================== */

  const monthGroups =
    useMemo(() => {
      const groups: {
        month: string;
        count: number;
      }[] = [];

      timelineColumns.forEach(
        (item) => {
          const month =
            item.month ||
            "Timeline";

          const last =
            groups[
              groups.length -
                1
            ];

          if (
            last &&
            last.month === month
          ) {
            last.count += 1;
          } else {
            groups.push({
              month,
              count: 1,
            });
          }
        }
      );

      return groups;
    }, [
      timelineColumns,
    ]);

  /* ==============================
     DATA START
  ============================== */

const dataStartRow =
  headerRowIndex + 1;
  /* ==============================
     READ CELL HELPER
  ============================== */

  function getCell(
    row: MatrixValue[],
    columnIndex: number
  ) {
    if (
      columnIndex < 0
    ) {
      return "";
    }

    return String(
      row[
        columnIndex
      ] ?? ""
    ).trim();
  }
/* ==============================
   ORGANIZATION CHECK
================================ */

function normalizeExcelText(
  value: string
) {
  return value
    .normalize("NFKC")

    // ลบตัวอักษรซ่อนที่มักติดมาจาก Excel
    .replace(
      /[\u200B-\u200D\uFEFF]/g,
      ""
    )

    // non-breaking space
    .replace(
      /\u00A0/g,
      " "
    )

    // รวมช่องว่างหลายตัว
    .replace(
      /\s+/g,
      " "
    )

    .trim();
}

function isOrganizationRow(
  description: string
) {
  const text =
    normalizeExcelText(
      description
    );

  if (!text) {
    return false;
  }

  /*
    รองรับรูปแบบ เช่น

    หน่วยงาน : สกมช.
    หน่วยงาน: สกมช.
    หน่วยงาน ： สกมช.
    หน่วยงาน สกมช.

    โดยไม่สนช่องว่างที่ติดมาจาก Excel
  */

  return /^หน่วยงาน\s*[:：]?/i.test(
    text
  );
}

/* ==============================
   BUILD ORGANIZATION GROUPS
================================ */

const groups =
  useMemo(() => {
    const result:
      OrganizationGroup[] =
      [];

    let currentGroup:
      OrganizationGroup |
      null = null;

    /*
      เดินตามลำดับแถวใน Excel จากบนลงล่าง
      ดังนั้นลำดับบนหน้าเว็บ = ลำดับใน Sheet
    */

    for (
      let rowIndex =
        dataStartRow;

      rowIndex <
        matrix.length;

      rowIndex += 1
    ) {
      const excelRow =
        matrix[rowIndex];

      if (!excelRow) {
        continue;
      }

      /* ------------------------------
         READ MAIN COLUMNS
      ------------------------------ */

      const description =
        normalizeExcelText(
          getCell(
            excelRow,
            headerInfo
              .descriptionIndex
          )
        );

      const progress =
        normalizeExcelText(
          getCell(
            excelRow,
            headerInfo
              .progressIndex
          )
        );

      const note =
        normalizeExcelText(
          getCell(
            excelRow,
            headerInfo
              .noteIndex
          )
        );

      const report =
        normalizeExcelText(
          getCell(
            excelRow,
            headerInfo
              .reportIndex
          )
        );

      /* ------------------------------
         TIMELINE
      ------------------------------ */

      const timeline =
        timelineColumns.map(
          ({
            columnIndex,
          }) =>
            normalizeExcelText(
              getCell(
                excelRow,
                columnIndex
              )
            )
        );

      const hasTimeline =
        timeline.some(
          (value) =>
            value !== ""
        );

      /* ------------------------------
         SKIP EMPTY ROW
      ------------------------------ */

      const isCompletelyEmpty =
        !description &&
        !progress &&
        !note &&
        !report &&
        !hasTimeline;

      if (
        isCompletelyEmpty
      ) {
        continue;
      }

      /* ==============================
         NEW ORGANIZATION

         ถ้า DESCRIPTION ขึ้นต้น
         ด้วยคำว่า "หน่วยงาน"
         ให้ถือว่าเป็นหน่วยงานใหม่
      ============================== */

      if (
        isOrganizationRow(
          description
        )
      ) {
        const newGroup:
          OrganizationGroup =
          {
            /*
              id ใช้เฉพาะ React/UI

              ลำดับจริงยังอิง
              excelRow และตำแหน่ง
              ที่พบใน Sheet
            */

            id:
              result.length +
              1,

            /*
              หมายเลขแถวจริงใน Excel
            */

            excelRow:
              rowIndex + 1,

            description,
            progress,
            note,
            report,
            timeline,

            children: [],
          };

        result.push(
          newGroup
        );

        currentGroup =
          newGroup;

        continue;
      }

      /* ==============================
         CHILD ROW

         เช่น
         1
         2
         3
         5.1
         5.1.1
         ฯลฯ

         ให้อยู่ใต้หน่วยงานล่าสุด
      ============================== */

      if (
        currentGroup
      ) {
        currentGroup.children.push(
          {
            excelRow:
              rowIndex + 1,

            description,
            progress,
            note,
            report,
            timeline,
          }
        );
      }
    }

    /*
      ยืนยันให้เรียงตามแถวจริง
      ใน Excel เสมอ
    */

    result.sort(
      (
        first,
        second
      ) =>
        first.excelRow -
        second.excelRow
    );

    return result;
  }, [
    matrix,
    dataStartRow,
    headerInfo,
    timelineColumns,
  ]);

  const filteredGroups =
    useMemo(() => {
      const keyword =
        search
          .trim()
          .toLowerCase();

      if (!keyword) {
        return groups;
      }

      return groups.filter(
        (group) => {
          const parentMatch =
            [
              group.description,
              group.progress,
              group.note,
              group.report,
            ].some(
              (value) =>
                value
                  .toLowerCase()
                  .includes(
                    keyword
                  )
            );

          const childMatch =
            group.children.some(
              (child) =>
                [
                  child.description,
                  child.progress,
                  child.note,
                  child.report,
                ].some(
                  (value) =>
                    value
                      .toLowerCase()
                      .includes(
                        keyword
                      )
                )
            );

          return (
            parentMatch ||
            childMatch
          );
        }
      );
    }, [
      groups,
      search,
    ]);

  /* ==============================
     EXPAND
  ============================== */

  function toggleGroup(
    id: number
  ) {
    setExpandedGroups(
      (current) =>
        current.includes(id)
          ? current.filter(
              (item) =>
                item !== id
            )
          : [
              ...current,
              id,
            ]
    );
  }

  function expandAll() {
    setExpandedGroups(
      filteredGroups.map(
        (group) =>
          group.id
      )
    );
  }

  function collapseAll() {
    setExpandedGroups(
      []
    );
  }

  /* ==============================
     COLUMNS
  ============================== */

  function toggleColumn(
    key: BaseColumnKey
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

  /* ==============================
     PROGRESS
  ============================== */

  function getProgressPercent(
    progress: string
  ) {
    const text =
      progress.trim();

    if (!text) {
      return 0;
    }

    if (
      text.includes("%")
    ) {
      const value =
        Number(
          text.replace(
            "%",
            ""
          )
        );

      return Number.isFinite(
        value
      )
        ? Math.min(
            Math.max(
              value,
              0
            ),
            100
          )
        : 0;
    }

    const number =
      Number(text);

    if (
      Number.isFinite(
        number
      )
    ) {
      if (
        number >= 0 &&
        number <= 1
      ) {
        return Math.round(
          number * 100
        );
      }

      return Math.min(
        Math.max(
          number,
          0
        ),
        100
      );
    }

    const lower =
      text.toLowerCase();

    if (
      lower.includes(
        "ดำเนินการเสร็จ"
      ) ||
      lower.includes(
        "เสร็จสิ้น"
      ) ||
      lower.includes(
        "จบเคส"
      )
    ) {
      return 100;
    }

    if (
      lower.includes(
        "กำลังดำเนินการ"
      )
    ) {
      return 50;
    }

    return 0;
  }

  /* ==============================
     STATUS CLASS
  ============================== */

  function getProgressClass(
    progress: string
  ) {
    const lower =
      progress
        .trim()
        .toLowerCase();

    if (
      lower.includes(
        "เสร็จ"
      )
    ) {
      return "done";
    }

    if (
      lower.includes(
        "กำลัง"
      )
    ) {
      return "working";
    }

    return "default";
  }

  /* ==============================
     LINKS
  ============================== */

  function renderTextOrLink(
    value: string
  ) {
    if (!value) {
      return "-";
    }

    if (
      value.startsWith(
        "http://"
      ) ||
      value.startsWith(
        "https://"
      )
    ) {
      return (
        <a
          href={value}
          target="_blank"
          rel="noreferrer"
          className="overview-report-link"
        >
          เปิด Report
        </a>
      );
    }

    return value;
  }

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
        String(sheetName || "overview-va")
          .replace(
            /[\\/:*?"<>|]/g,
            "_"
          )
          .trim() ||
        "overview-va";

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

  if (
    headerRowIndex < 0
  ) {
    return (
      <section className="overview-va">
        <div className="overview-va-empty">
          <h3>
            ไม่พบ Header
            DESCRIPTIONS
          </h3>

          <p>
            ไม่สามารถอ่าน
            Overview VA (New)
            ได้
          </p>
        </div>
      </section>
    );
  }

  return (
    <section className="overview-va">

      {/* =====================
          HEADING
      ====================== */}

      <div className="overview-va-heading">
        <div>
          <h2>
            {sheetName}
          </h2>

          <p>
            ภาพรวมการทดสอบและประเมิน
            จุดอ่อนช่องโหว่ของแต่ละหน่วยงาน
          </p>
        </div>

        <span className="overview-va-count">
          {groups.length} หน่วยงาน
        </span>
      </div>

      {/* =====================
          TOOLBAR
      ====================== */}

      <div className="overview-va-toolbar">

        <div className="overview-va-search">
          <span>
            ⌕
          </span>

          <input
            type="text"
            placeholder="ค้นหาหน่วยงาน / NOTE / Report..."
            value={
              search
            }
            onChange={(
              event
            ) =>
              setSearch(
                event.target
                  .value
              )
            }
          />
        </div>

        <div className="overview-va-toolbar-right">

          <button
            type="button"
            className="overview-excel-button"
            onClick={openInExcel}
            title="สร้างไฟล์ .xlsx ของ Sheet นี้เพื่อเปิดด้วย Microsoft Excel"
          >
            ▣ เปิดใน Excel
          </button>

          {/* EXPAND */}

          {viewType ===
            "table" && (
            <div className="overview-expand-actions">
              <button
                type="button"
                onClick={
                  expandAll
                }
              >
                เปิดทั้งหมด
              </button>

              <button
                type="button"
                onClick={
                  collapseAll
                }
              >
                พับทั้งหมด
              </button>
            </div>
          )}

          {/* COLUMN */}

          <div className="overview-column-wrapper">

            <button
              type="button"
              className="overview-column-button"
              onClick={() =>
                setShowColumnPicker(
                  !showColumnPicker
                )
              }
            >
              ☷ เลือกหัวข้อ
            </button>

            {showColumnPicker && (
              <div className="overview-column-menu">

                <div className="overview-column-header">
                  <strong>
                    หัวข้อที่แสดง
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

                {BASE_COLUMNS.map(
                  (column) => (
                    <label
                      key={
                        column.key
                      }
                      className="overview-column-option"
                    >
                      <input
                        type="checkbox"
                        checked={
                          visibleColumns.includes(
                            column.key
                          )
                        }
                        onChange={() =>
                          toggleColumn(
                            column.key
                          )
                        }
                      />

                      <span>
                        {
                          column.label
                        }
                      </span>
                    </label>
                  )
                )}

                <label className="overview-column-option">
                  <input
                    type="checkbox"
                    checked={
                      showTimeline
                    }
                    onChange={() =>
                      setShowTimeline(
                        !showTimeline
                      )
                    }
                  />

                  <span>
                    Timeline / วันที่
                  </span>
                </label>

              </div>
            )}
          </div>

          {/* VIEW */}

          <div className="overview-va-tabs">

            <button
              type="button"
              className={
                viewType ===
                "table"
                  ? "active"
                  : ""
              }
              onClick={() =>
                setViewType(
                  "table"
                )
              }
            >
              TABLE
            </button>

            <button
              type="button"
              className={
                viewType ===
                "progress"
                  ? "active"
                  : ""
              }
              onClick={() =>
                setViewType(
                  "progress"
                )
              }
            >
              PROGRESS
            </button>

          </div>
        </div>
      </div>

      {/* =====================
          TABLE
      ====================== */}

      {viewType ===
        "table" && (
        <div className="overview-table-wrapper">

          <table className="overview-table">

            <thead>

              {/* MONTH */}

              {showTimeline &&
              timelineColumns.length >
                0 ? (
                <>
                  <tr className="overview-month-row">

                    {visibleColumns.map(
                      (key) => (
                        <th
                          key={
                            key
                          }
                          rowSpan={
                            2
                          }
                        >
                          {
                            BASE_COLUMNS.find(
                              (
                                column
                              ) =>
                                column.key ===
                                key
                            )?.label
                          }
                        </th>
                      )
                    )}

                    {monthGroups.map(
                      (
                        group,
                        index
                      ) => (
                        <th
                          key={`${group.month}-${index}`}
                          colSpan={
                            group.count
                          }
                        >
                          {
                            group.month
                          }
                        </th>
                      )
                    )}

                  </tr>

                  {/* DAYS */}

                  <tr className="overview-day-row">

                    {timelineColumns.map(
                      (
                        item,
                        index
                      ) => (
                        <th
                          key={`${item.columnIndex}-${index}`}
                        >
                          {
                            item.day
                          }
                        </th>
                      )
                    )}

                  </tr>
                </>
              ) : (
                <tr>
                  {visibleColumns.map(
                    (key) => (
                      <th
                        key={
                          key
                        }
                      >
                        {
                          BASE_COLUMNS.find(
                            (
                              column
                            ) =>
                              column.key ===
                              key
                          )?.label
                        }
                      </th>
                    )
                  )}
                </tr>
              )}

            </thead>

            <tbody>

              {filteredGroups.length ===
              0 ? (
                <tr>
                  <td
                    className="overview-empty-cell"
                    colSpan={
                      Math.max(
                        visibleColumns.length +
                          (showTimeline
                            ? timelineColumns.length
                            : 0),
                        1
                      )
                    }
                  >
                    ไม่พบข้อมูล
                  </td>
                </tr>
              ) : (
                filteredGroups.map(
                  (group) => {
                    const expanded =
                      expandedGroups.includes(
                        group.id
                      );

                    return (
                      <FragmentGroup
                        key={
                          group.id
                        }
                      >

                        {/* PARENT */}

                        <tr className="overview-organization-row">

                          {visibleColumns.includes(
                            "description"
                          ) && (
                            <td className="overview-description-cell">

                              <div className="overview-description-main">

                                <button
                                  type="button"
                                  className="overview-expand-button"
                                  onClick={() =>
                                    toggleGroup(
                                      group.id
                                    )
                                  }
                                >
                                  {expanded
                                    ? "▼"
                                    : "▶"}
                                </button>

                                <strong>
                                  {
                                    group.description
                                  }
                                </strong>

                                <span className="overview-child-count">
                                  {
                                    group.children.length
                                  }{" "}
                                  รายการ
                                </span>

                              </div>

                            </td>
                          )}

                          {visibleColumns.includes(
                            "progress"
                          ) && (
                            <td>
                              <span
                                className={`overview-progress-badge ${getProgressClass(
                                  group.progress
                                )}`}
                              >
                                {group.progress ||
                                  "-"}
                              </span>
                            </td>
                          )}

                          {visibleColumns.includes(
                            "note"
                          ) && (
                            <td>
                              {group.note ||
                                "-"}
                            </td>
                          )}

                          {visibleColumns.includes(
                            "report"
                          ) && (
                            <td>
                              {renderTextOrLink(
                                group.report
                              )}
                            </td>
                          )}

                          {showTimeline &&
                            group.timeline.map(
                              (
                                value,
                                index
                              ) => (
                                <td
                                  key={
                                    index
                                  }
                                  className={
                                    value
                                      ? "overview-timeline-cell active"
                                      : "overview-timeline-cell"
                                  }
                                  title={
                                    value
                                  }
                                >
                                  {value
                                    ? "●"
                                    : ""}
                                </td>
                              )
                            )}

                        </tr>

                        {/* CHILDREN */}

                        {expanded &&
                          group.children.map(
                            (
                              child,
                              childIndex
                            ) => (
                              <tr
                                key={`${group.id}-${child.excelRow}-${childIndex}`}
                                className="overview-child-row"
                              >

                                {visibleColumns.includes(
                                  "description"
                                ) && (
                                  <td className="overview-description-cell child">
                                    <span className="overview-child-indent">
                                      ↳
                                    </span>

                                    {child.description ||
                                      "-"}
                                  </td>
                                )}

                                {visibleColumns.includes(
                                  "progress"
                                ) && (
                                  <td>
                                    {child.progress ? (
                                      <span
                                        className={`overview-progress-badge ${getProgressClass(
                                          child.progress
                                        )}`}
                                      >
                                        {
                                          child.progress
                                        }
                                      </span>
                                    ) : (
                                      "-"
                                    )}
                                  </td>
                                )}

                                {visibleColumns.includes(
                                  "note"
                                ) && (
                                  <td>
                                    {child.note ||
                                      "-"}
                                  </td>
                                )}

                                {visibleColumns.includes(
                                  "report"
                                ) && (
                                  <td>
                                    {renderTextOrLink(
                                      child.report
                                    )}
                                  </td>
                                )}

                                {showTimeline &&
                                  child.timeline.map(
                                    (
                                      value,
                                      index
                                    ) => (
                                      <td
                                        key={
                                          index
                                        }
                                        className={
                                          value
                                            ? "overview-timeline-cell active"
                                            : "overview-timeline-cell"
                                        }
                                        title={
                                          value
                                        }
                                      >
                                        {value
                                          ? "●"
                                          : ""}
                                      </td>
                                    )
                                  )}

                              </tr>
                            )
                          )}

                      </FragmentGroup>
                    );
                  }
                )
              )}

            </tbody>

          </table>

        </div>
      )}

      {/* =====================
          PROGRESS
      ====================== */}

      {viewType ===
        "progress" && (
        <div className="overview-progress-list">

          {filteredGroups.map(
            (group) => {
              const percent =
                getProgressPercent(
                  group.progress
                );

              return (
                <article
                  key={
                    group.id
                  }
                  className="overview-progress-item"
                >

                  <div className="overview-progress-heading">

                    <div>
                      <strong>
                        {
                          group.description
                        }
                      </strong>

                      {group.note && (
                        <span>
                          {
                            group.note
                          }
                        </span>
                      )}
                    </div>

                    <b>
                      {
                        percent
                      }
                      %
                    </b>

                  </div>

                  <div className="overview-progress-track">

                    <div
                      className="overview-progress-fill"
                      style={{
                        width: `${percent}%`,
                      }}
                    />

                  </div>

                </article>
              );
            }
          )}

        </div>
      )}

    </section>
  );
}

/*
  ใช้ component เล็ก ๆ แทน Fragment
  เพื่อไม่ต้อง import Fragment เพิ่ม
*/

function FragmentGroup({
  children,
}: {
  children:
    React.ReactNode;
}) {
  return <>{children}</>;
}
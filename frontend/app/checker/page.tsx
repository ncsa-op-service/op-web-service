"use client";

import {
  ChangeEvent,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";

import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import "./checker.css";

type ProviderKey =
  | "ais"
  | "trueDtac"
  | "nt"
  | "cloudflare";

type ViewMode =
  | "original"
  | "checked"
  | "line_similarity";

type UserInfo = {
  id: number;
  name: string;
  email: string;
  role:
    | "super_admin"
    | "editor"
    | "viewer";
};

type ProviderData = {
  ais: string;
  trueDtac: string;
  nt: string;
  cloudflare: string;
};

type UrlItem = {
  id: number;
  dbResultId?: number;
  caseId: string;
  urlSms: string;

  // ตำแหน่งแถวจริงใน Excel ต้นฉบับ (0-based)
  sourceRowIndex: number;

  original: ProviderData;
  checked: ProviderData;
};

type SourceColumnMap = {
  headerRowIndex: number;
  urlColumn: number;
  caseColumn: number;
  aisColumn: number;
  trueDtacColumn: number;
  ntColumn: number;
  cloudflareColumn: number;
};

type CheckErrorType =
  | "nxdomain"
  | "timeout"
  | "connection"
  | "tls"
  | "invalid_url"
  | "network"
  | null;

type ProviderCheckState =
  | "idle"
  | "success"
  | "nxdomain"
  | "failed";

type ApiCheckResult = {
  url: string;
  reachable: boolean;
  statusCode: number | null;
  finalUrl: string | null;
  error: string | null;
  errorType: CheckErrorType;
};

type ApiCheckResponse = {
  total: number;
  results: ApiCheckResult[];
};

type LineSimilarityRow = {
  item: UrlItem;
  lineUrl: string;
  score: number;
  label: string;
  token: string;
  structure: string;
  prefix: string;
  suffix: string;
  matchedReference: number;
  matchedCount: number;
  comparedLength: number;
};

const OPEN_LINE_REFERENCE_URLS = [
  "https://line.me/ti/p/uEh4NMpZuc",
  "https://line.me/ti/p/zXfEZpWjGB",
  "https://line.me/ti/p/-A_-pnuPY2",
  "https://line.me/ti/p/WBH9eGmg_E",
] as const;

/*
  จุดแบ่งที่ผู้ใช้ยืนยันจาก 4 ลิงก์ที่เปิดได้จริง

  1) uEh4NMp | Zuc
  2) zXfEZpW | jGB
  3) -A_-pnu | PY2
  4) WBH9eG  | mg_E
*/
const OPEN_LINE_REFERENCE_PARTS = [
  {
    prefix: "uEh4NMp",
    suffix: "Zuc",
    token: "uEh4NMpZuc",
  },
  {
    prefix: "zXfEZpW",
    suffix: "jGB",
    token: "zXfEZpWjGB",
  },
  {
    prefix: "-A_-pnu",
    suffix: "PY2",
    token: "-A_-pnuPY2",
  },
  {
    prefix: "WBH9eG",
    suffix: "mg_E",
    token: "WBH9eGmg_E",
  },
] as const;

function countExactPositionMatches(
  candidate: string,
  reference: string
) {
  const compareLength =
    Math.max(
      candidate.length,
      reference.length
    );

  let matchedCount = 0;

  for (
    let index = 0;
    index < compareLength;
    index += 1
  ) {
    if (
      candidate[index] !== undefined &&
      candidate[index] ===
        reference[index]
    ) {
      matchedCount += 1;
    }
  }

  return {
    matchedCount,
    comparedLength:
      compareLength,
  };
}

function renderMatchedCharacters(
  value: string,
  reference: string,
  offset = 0
) {
  return value
    .split("")
    .map((char, index) => {
      const referenceChar =
        reference[
          offset + index
        ];

      const isMatch =
        referenceChar !== undefined &&
        char === referenceChar;

      return (
        <span
          key={`${offset + index}-${char}`}
          className={
            isMatch
              ? "checker-line-char-match"
              : "checker-line-char-normal"
          }
          title={
            isMatch
              ? `ตรงตำแหน่ง ${offset + index + 1}: ${char}`
              : undefined
          }
        >
          {char}
        </span>
      );
    });
}

function analyzeLineUrl(
  rawUrl: string
): {
  score: number;
  label: string;
  token: string;
  structure: string;
  prefix: string;
  suffix: string;
  matchedReference: number;
  matchedCount: number;
  comparedLength: number;
} {
  try {
    const parsed =
      new URL(rawUrl.trim());

    const host =
      parsed.hostname
        .toLowerCase();

    const directMatch =
      parsed.pathname.match(
        /^\/ti\/p\/([^/?#]+)\/?$/i
      );

    if (
      parsed.protocol !== "https:" ||
      !(
        host === "line.me" ||
        host === "www.line.me"
      ) ||
      !directMatch
    ) {
      return {
        score: 0,
        label: "ไม่ตรงรูปแบบ",
        token: "",
        structure: "OTHER",
        prefix: "",
        suffix: "",
        matchedReference: 0,
        matchedCount: 0,
        comparedLength: 0,
      };
    }

    const token =
      directMatch[1] ?? "";

    if (
      !token ||
      !/^[A-Za-z0-9_-]+$/.test(
        token
      )
    ) {
      return {
        score: 0,
        label: "Token ไม่ถูกต้อง",
        token,
        structure:
          "line.me/ti/p/{token}",
        prefix: "",
        suffix: "",
        matchedReference: 0,
        matchedCount: 0,
        comparedLength: 0,
      };
    }

    let bestReference = 0;
    let bestMatchedCount = -1;
    let bestComparedLength = 0;
    let bestScore = 0;
    let bestPrefix = "";
    let bestSuffix = "";

    OPEN_LINE_REFERENCE_PARTS.forEach(
      (reference, index) => {
        const {
          matchedCount,
          comparedLength,
        } =
          countExactPositionMatches(
            token,
            reference.token
          );

        const score =
          comparedLength > 0
            ? Math.round(
                (matchedCount /
                  comparedLength) *
                  100
              )
            : 0;

        if (
          matchedCount >
            bestMatchedCount ||
          (
            matchedCount ===
              bestMatchedCount &&
            score > bestScore
          )
        ) {
          bestReference =
            index + 1;
          bestMatchedCount =
            matchedCount;
          bestComparedLength =
            comparedLength;
          bestScore =
            score;

          bestPrefix =
            token.slice(
              0,
              reference.prefix.length
            );

          bestSuffix =
            token.slice(
              -reference.suffix.length
            );
        }
      }
    );

    let label =
      "ตรงน้อย";

    if (bestScore >= 70) {
      label =
        "ตรงมาก";
    } else if (bestScore >= 40) {
      label =
        "ค่อนข้างตรง";
    } else if (bestScore >= 20) {
      label =
        "ตรงบางส่วน";
    }

    return {
      score: bestScore,
      label,
      token,
      structure:
        "line.me/ti/p/{token}",
      prefix: bestPrefix,
      suffix: bestSuffix,
      matchedReference:
        bestReference,
      matchedCount:
        Math.max(
          0,
          bestMatchedCount
        ),
      comparedLength:
        bestComparedLength,
    };
  } catch {
    return {
      score: 0,
      label: "URL ไม่ถูกต้อง",
      token: "",
      structure: "INVALID",
      prefix: "",
      suffix: "",
      matchedReference: 0,
      matchedCount: 0,
      comparedLength: 0,
    };
  }
}

type NetworkInfo = {
  ip: string | null;
  asn: string | null;
  org: string | null;
  country: string | null;

  provider:
    | ProviderKey
    | "unknown";
};

function isCloudflareNetwork(
  info: NetworkInfo | null | undefined
) {
  if (!info) {
    return false;
  }

  if (info.provider === "cloudflare") {
    return true;
  }

  const text = [
    info.org ?? "",
    info.asn ?? "",
  ]
    .join(" ")
    .toLowerCase();

  return (
    text.includes("cloudflare") ||
    text.includes("as13335") ||
    text.includes("13335")
  );
}

type SaveResponse = {
  message?: string;
  batchId?: number;
  insertedCount?: number;
  createdAt?: string;
};

type SavedBatch = {
  id: number;
  original_file_name: string | null;
  selected_provider: string | null;
  public_ip: string | null;
  isp: string | null;
  detected_provider: string | null;
  case_type: string;
  detected_date: string | null;
  case_status: string | null;
  created_at: string;
};

type SavedResult = {
  id: number;
  batch_id: number;
  case_id: string;
  url_sms: string;
  ais_result: string | null;
  true_dtac_result: string | null;
  nt_result: string | null;
  cloudflare_result: string | null;
  ais_status: ProviderCheckState | null;
  true_dtac_status: ProviderCheckState | null;
  nt_status: ProviderCheckState | null;
  cloudflare_status: ProviderCheckState | null;
  created_at: string;
};

type LatestSavedResponse = {
  batch: SavedBatch | null;
  results: SavedResult[];
};

type SaveMode = "all" | "provider" | "history";

type HistoryBatch = SavedBatch & {
  result_count?: number;
  save_type?: SaveMode | string | null;
  saved_provider?: ProviderKey | string | null;
  publish_to_summary?: boolean | null;
  saved_at?: string | null;
};

type HistoryResponse = {
  batches?: HistoryBatch[];
  history?: HistoryBatch[];
};

function getToday() {
  const now = new Date();

  const year =
    now.getFullYear();

  const month =
    String(
      now.getMonth() + 1
    ).padStart(2, "0");

  const day =
    String(
      now.getDate()
    ).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function readCurrentUser():
  UserInfo | null {
  if (
    typeof window === "undefined"
  ) {
    return null;
  }

  const candidates = [
    "user",
    "currentUser",
    "authUser",
    "op_user",
  ];

  for (const key of candidates) {
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
        return parsed as UserInfo;
      }
    } catch {
      // ignore invalid localStorage
    }
  }

  return null;
}

export default function CheckerPage() {
  const [fileName, setFileName] =
    useState("");

  const [
    currentUser,
    setCurrentUser,
  ] =
    useState<UserInfo | null>(
      null
    );

  const [
    latestSavedBatch,
    setLatestSavedBatch,
  ] = useState<SavedBatch | null>(null);

  const [
    latestSavedCount,
    setLatestSavedCount,
  ] = useState(0);

  const [
    latestSavedResults,
    setLatestSavedResults,
  ] = useState<SavedResult[]>([]);

  const [
    isLoadingSaved,
    setIsLoadingSaved,
  ] = useState(false);

  const [
    openedSavedBatchId,
    setOpenedSavedBatchId,
  ] = useState<number | null>(null);

  const [
    isEditingSaved,
    setIsEditingSaved,
  ] = useState(false);

  const canManage =
    currentUser?.role ===
      "super_admin" ||
    currentUser?.role ===
      "editor";

  useEffect(() => {
    setCurrentUser(
      readCurrentUser()
    );

    void loadLatestSavedData();
    void loadHistoryData();
  }, []);

  const [items, setItems] =
    useState<UrlItem[]>([]);

  const [search, setSearch] =
    useState("");

  const [provider, setProvider] =
    useState<ProviderKey>("ais");

  const [viewMode, setViewMode] =
    useState<ViewMode>(
      "original"
    );

  const [
    isProcessing,
    setIsProcessing,
  ] = useState(false);

  const [
    isDetectingNetwork,
    setIsDetectingNetwork,
  ] = useState(false);

  const [isSaving, setIsSaving] =
    useState(false);

  const [showSaveModal, setShowSaveModal] =
    useState(false);

  const [showHistoryModal, setShowHistoryModal] =
    useState(false);

  const [saveMode, setSaveMode] =
    useState<SaveMode>("all");

  const [saveProvider, setSaveProvider] =
    useState<ProviderKey>("ais");

  const [historyBatches, setHistoryBatches] =
    useState<HistoryBatch[]>([]);

  const [isLoadingHistory, setIsLoadingHistory] =
    useState(false);

  const [calendarMonth, setCalendarMonth] =
    useState(() => {
      const now = new Date();
      return new Date(now.getFullYear(), now.getMonth(), 1);
    });

  const [selectedHistoryDate, setSelectedHistoryDate] =
    useState(getToday());

  const [
    lineSimilarityCount,
    setLineSimilarityCount,
  ] = useState("3");

  const [
    networkInfo,
    setNetworkInfo,
  ] =
    useState<NetworkInfo | null>(
      null
    );

  const [
    checkStatuses,
    setCheckStatuses,
  ] = useState<
    Record<
      number,
      Partial<Record<ProviderKey, ProviderCheckState>>
    >
  >({});

  const cloudflareConnected =
    isCloudflareNetwork(networkInfo);

  const originalWorkbookRef =
    useRef<XLSX.WorkBook | null>(
      null
    );

  // เก็บ byte ของไฟล์ Excel ต้นฉบับจริง
  // ใช้ ExcelJS ตอน Export เพื่อรักษา style / สี / border / merge / column width
  const originalExcelBufferRef =
    useRef<ArrayBuffer | null>(
      null
    );

  const [
    sourceSheetName,
    setSourceSheetName,
  ] =
    useState("");

  const [
    sourceColumnMap,
    setSourceColumnMap,
  ] =
    useState<SourceColumnMap | null>(
      null
    );

  const apiUrl =
    process.env
      .NEXT_PUBLIC_API_URL ??
    "http://localhost:4000";

  async function loadLatestSavedData() {
    try {
      setIsLoadingSaved(true);

      const response = await fetch(
        `${apiUrl}/api/summary/edit/fake_domain`,
        { cache: "no-store" }
      );

      const data: LatestSavedResponse =
        await response.json();

      if (!response.ok) {
        throw new Error(
          (data as { message?: string }).message ??
            "โหลดข้อมูลล่าสุดไม่สำเร็จ"
        );
      }

      setLatestSavedBatch(data.batch);
      setLatestSavedResults(
        Array.isArray(data.results)
          ? data.results
          : []
      );
      setLatestSavedCount(
        Array.isArray(data.results)
          ? data.results.length
          : 0
      );
    } catch (error) {
      console.error(
        "Load latest saved checker data error:",
        error
      );
      setLatestSavedBatch(null);
      setLatestSavedResults([]);
      setLatestSavedCount(0);
    } finally {
      setIsLoadingSaved(false);
    }
  }

  async function loadHistoryData() {
    try {
      setIsLoadingHistory(true);

      const response = await fetch(
        `${apiUrl}/api/checker-results/history`,
        { cache: "no-store" }
      );

      if (!response.ok) {
        throw new Error("history endpoint unavailable");
      }

      const data: HistoryResponse | HistoryBatch[] =
        await response.json();

      const rows = Array.isArray(data)
        ? data
        : Array.isArray(data.batches)
          ? data.batches
          : Array.isArray(data.history)
            ? data.history
            : [];

      setHistoryBatches(rows);
    } catch (error) {
      console.warn(
        "Load checker history fallback to latest batch:",
        error
      );

      setHistoryBatches((current) => {
        if (current.length > 0) return current;
        return latestSavedBatch
          ? [{ ...latestSavedBatch, result_count: latestSavedCount }]
          : [];
      });
    } finally {
      setIsLoadingHistory(false);
    }
  }

  function formatThaiDate(value: string | null | undefined) {
    if (!value) return "-";

    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return value;

    return parsed.toLocaleDateString("th-TH", {
      day: "numeric",
      month: "long",
      year: "numeric",
    });
  }

  function formatThaiTime(value: string | null | undefined) {
    if (!value) return "-";

    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return "-";

    return parsed.toLocaleTimeString("th-TH", {
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function getBatchDateKey(batch: HistoryBatch | SavedBatch) {
    const raw =
      ("saved_at" in batch ? batch.saved_at : null) ??
      batch.detected_date ??
      batch.created_at;

    if (!raw) return "";
    return String(raw).slice(0, 10);
  }

  function getProviderLabel(value: string | null | undefined) {
    if (value === "ais") return "AIS";
    if (value === "trueDtac") return "TRUE / DTAC";
    if (value === "nt") return "NT";
    if (value === "cloudflare") return "Cloudflare";
    return "ทั้งหมด";
  }

  const calendarDays = useMemo(() => {
    const year = calendarMonth.getFullYear();
    const month = calendarMonth.getMonth();
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const leading = (firstDay.getDay() + 6) % 7;

    const cells: Array<{
      date: Date | null;
      key: string;
      count: number;
    }> = [];

    for (let i = 0; i < leading; i += 1) {
      cells.push({ date: null, key: `empty-${i}`, count: 0 });
    }

    for (let day = 1; day <= lastDay.getDate(); day += 1) {
      const date = new Date(year, month, day);
      const key = [
        year,
        String(month + 1).padStart(2, "0"),
        String(day).padStart(2, "0"),
      ].join("-");

      const count = historyBatches.filter(
        (batch) => getBatchDateKey(batch) === key
      ).length;

      cells.push({ date, key, count });
    }

    return cells;
  }, [calendarMonth, historyBatches]);

  const selectedHistoryBatches = useMemo(() => {
    return historyBatches
      .filter((batch) => getBatchDateKey(batch) === selectedHistoryDate)
      .sort((a, b) =>
        String(b.saved_at ?? b.created_at).localeCompare(
          String(a.saved_at ?? a.created_at)
        )
      );
  }, [historyBatches, selectedHistoryDate]);

  function applySavedData(
    batch: SavedBatch,
    results: SavedResult[],
    scrollToTable = true
  ) {
    if (!batch || results.length === 0) return;

    const mappedItems: UrlItem[] = results.map((row, index) => {
      const providerData: ProviderData = {
        ais: row.ais_result ?? "",
        trueDtac: row.true_dtac_result ?? "",
        nt: row.nt_result ?? "",
        cloudflare: row.cloudflare_result ?? "",
      };

      return {
        id: index + 1,
        dbResultId: row.id,
        caseId: row.case_id,
        urlSms: row.url_sms,
        sourceRowIndex: index,
        original: { ...providerData },
        checked: { ...providerData },
      };
    });

    setItems(mappedItems);

    const restoredStatuses: Record<
      number,
      Partial<Record<ProviderKey, ProviderCheckState>>
    > = {};

    results.forEach((row, index) => {
      const itemId = index + 1;
      restoredStatuses[itemId] = {
        ais: row.ais_status ??
          (row.ais_result === "DNS_PROBE_FINISHED_NXDOMAIN" ? "nxdomain" : "idle"),
        trueDtac: row.true_dtac_status ??
          (row.true_dtac_result === "DNS_PROBE_FINISHED_NXDOMAIN" ? "nxdomain" : "idle"),
        nt: row.nt_status ??
          (row.nt_result === "DNS_PROBE_FINISHED_NXDOMAIN" ? "nxdomain" : "idle"),
        cloudflare: row.cloudflare_status ??
          (row.cloudflare_result === "DNS_PROBE_FINISHED_NXDOMAIN" ? "nxdomain" : "idle"),
      };
    });

    setCheckStatuses(restoredStatuses);
    setFileName(batch.original_file_name ?? `saved-batch-${batch.id}`);

    const savedProvider = batch.selected_provider;
    if (
      savedProvider === "ais" ||
      savedProvider === "trueDtac" ||
      savedProvider === "nt" ||
      savedProvider === "cloudflare"
    ) {
      setProvider(savedProvider);
    }

    setNetworkInfo({
      ip: batch.public_ip ?? null,
      asn: null,
      org: batch.isp ?? null,
      country: null,
      provider:
        batch.detected_provider === "ais" ||
        batch.detected_provider === "trueDtac" ||
        batch.detected_provider === "nt" ||
        batch.detected_provider === "cloudflare"
          ? batch.detected_provider
          : "unknown",
    });

    originalWorkbookRef.current = null;
    originalExcelBufferRef.current = null;
    setSourceSheetName("");
    setSourceColumnMap(null);
    setOpenedSavedBatchId(batch.id);
    setIsEditingSaved(false);
    setViewMode("original");
    setSearch("");

    if (scrollToTable) {
      window.setTimeout(() => {
        document.querySelector(".checker-table-card")?.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
      }, 50);
    }
  }

  async function openHistoryBatch(batch: HistoryBatch) {
    try {
      setIsLoadingSaved(true);

      const response = await fetch(
        `${apiUrl}/api/checker-results/${batch.id}`,
        { cache: "no-store" }
      );

      if (!response.ok) {
        if (latestSavedBatch?.id === batch.id) {
          applySavedData(latestSavedBatch, latestSavedResults, true);
          return;
        }
        throw new Error("โหลด Batch นี้ไม่สำเร็จ");
      }

      const data: LatestSavedResponse = await response.json();
      if (!data.batch || !Array.isArray(data.results)) {
        throw new Error("รูปแบบข้อมูล Batch ไม่ถูกต้อง");
      }

      applySavedData(data.batch, data.results, true);
    } catch (error) {
      alert(
        error instanceof Error
          ? error.message
          : "โหลดประวัติไม่สำเร็จ"
      );
    } finally {
      setIsLoadingSaved(false);
    }
  }

  function openLatestSavedData(
    scrollToTable = true
  ) {
    if (
      !latestSavedBatch ||
      latestSavedResults.length === 0
    ) {
      return;
    }

    applySavedData(
      latestSavedBatch,
      latestSavedResults,
      scrollToTable
    );
  }

  useEffect(() => {
    if (
      latestSavedBatch &&
      latestSavedResults.length > 0 &&
      openedSavedBatchId === null &&
      items.length === 0
    ) {
      openLatestSavedData(false);
    }
  }, [
    latestSavedBatch?.id,
    latestSavedResults.length,
  ]);


  /* =========================
     SEARCH
  ========================= */

  const filteredItems =
    useMemo(() => {
      const keyword =
        search
          .toLowerCase()
          .trim();

      if (!keyword) {
        return items;
      }

      return items.filter(
        (item) =>
          item.caseId
            .toLowerCase()
            .includes(keyword) ||
          item.urlSms
            .toLowerCase()
            .includes(keyword)
      );
    }, [items, search]);

  const lineSimilarityRows =
    useMemo<LineSimilarityRow[]>(() => {
      const rows: LineSimilarityRow[] =
        [];

      const seen =
        new Set<string>();

      items.forEach((item) => {
        const candidates = [
          item.urlSms,
          item.checked.ais,
          item.checked.trueDtac,
          item.checked.nt,
          item.checked.cloudflare,
        ]
          .map((value) =>
            String(
              value ?? ""
            ).trim()
          )
          .filter(Boolean);

        candidates.forEach(
          (lineUrl) => {
            const key =
              `${item.id}::${lineUrl}`;

            if (seen.has(key)) {
              return;
            }

            seen.add(key);

            const analysis =
              analyzeLineUrl(
                lineUrl
              );

            if (
              analysis.structure !==
                "line.me/ti/p/{token}" ||
              !OPEN_LINE_REFERENCE_PARTS.some(
                (reference) =>
                  reference.token === analysis.token
              )
            ) {
              return;
            }

            rows.push({
              item,
              lineUrl,
              ...analysis,
              score: 100,
              label: "ตรงเป๊ะ",
              matchedCount: analysis.token.length,
              comparedLength: analysis.token.length,
            });
          }
        );
      });

      return rows.sort(
        (a, b) =>
          b.matchedCount -
            a.matchedCount ||
          b.score -
            a.score ||
          a.item.id -
            b.item.id
      );
    }, [items]);

  const requestedSimilarityCount =
    Math.max(
      1,
      Number.parseInt(
        lineSimilarityCount,
        10
      ) || 1
    );

  const selectedSimilarityRows =
    useMemo(() => {
      return lineSimilarityRows.slice(
        0,
        requestedSimilarityCount
      );
    }, [
      lineSimilarityRows,
      requestedSimilarityCount,
    ]);


  function handleExportLineSimilarity() {
    if (
      selectedSimilarityRows.length ===
      0
    ) {
      alert(
        "ยังไม่มี LINE URL สำหรับดาวน์โหลด"
      );
      return;
    }

    const rows =
      selectedSimilarityRows.map(
        (row, index) => ({
          Rank: index + 1,
          "CASE ID":
            row.item.caseId,
          URL:
            row.lineUrl,
          "ช่วงต้น":
            row.prefix,
          "ช่วงท้าย":
            row.suffix,
          "ตรงกี่ตัว":
            `${row.matchedCount}/${row.comparedLength}`,
          "ตรงกับตัวหลัก":
            row.matchedReference
              ? `#${row.matchedReference}`
              : "-",
          "Exact Match (%)":
            row.score,
          "ระดับ":
            row.label,
          "โครงสร้าง":
            row.structure,
        })
      );

    const worksheet =
      XLSX.utils.json_to_sheet(
        rows
      );

    const workbook =
      XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "LINE Similarity"
    );

    XLSX.writeFile(
      workbook,
      `line-exact-match-${getToday()}.xlsx`
    );
  }

  /* =========================
     EXCEL
  ========================= */

  function normalizeExcelText(value: unknown) {
    return String(value ?? "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeHeader(value: unknown) {
    return normalizeExcelText(value)
      .toLowerCase()
      .replace(/[\s_\-\/\\().:]+/g, "");
  }

  function isUrlLike(value: unknown) {
    const text = normalizeExcelText(value);

    if (!text) {
      return false;
    }

    return (
      /^https?:\/\/\S+/i.test(text) ||
      /^www\.\S+/i.test(text) ||
      /^[a-z0-9.-]+\.[a-z]{2,}(?:[/?#]\S*)?$/i.test(text)
    );
  }

  function normalizeUrl(value: unknown) {
    const text = normalizeExcelText(value);

    if (!text) {
      return "";
    }

    if (/^https?:\/\//i.test(text)) {
      return text;
    }

    if (
      /^www\./i.test(text) ||
      /^[a-z0-9.-]+\.[a-z]{2,}(?:[/?#]\S*)?$/i.test(text)
    ) {
      return `https://${text}`;
    }

    return text;
  }

  const urlHeaderAliases = [
    "URL SMS",
    "URL",
    "URLSMS",
    "URL SMS/Domain",
    "URL/Domain",
    "Domain",
    "Domain Name",
    "Website",
    "Web Site",
    "Link",
    "URL Link",
    "ลิงก์",
    "ลิงค์",
    "เว็บไซต์",
    "โดเมน",
    "URL ที่ตรวจพบ",
    "URL ที่ต้องการตรวจ",
  ].map(normalizeHeader);

  const caseHeaderAliases = [
    "CASE ID",
    "CASE_ID",
    "CASEID",
    "Case",
    "Case No",
    "Case Number",
    "เลขเคส",
    "เลขที่เคส",
    "รหัสเคส",
  ].map(normalizeHeader);

  const providerAliases: Record<
    ProviderKey,
    string[]
  > = {
    ais: ["AIS"].map(normalizeHeader),

    trueDtac: [
      "TRUE/DTAC",
      "TRUE / DTAC",
      "TRUE-DTAC",
      "TRUE DTAC",
      "TRUE",
      "DTAC",
    ].map(normalizeHeader),

    nt: ["NT"].map(normalizeHeader),

    cloudflare: [
      "CloudFlare",
      "Cloudflare",
      "CLOUDFLARE",
      "CLOUD FLARE",
      "CloundFlare",
    ].map(normalizeHeader),
  };

  function findColumnIndex(
    headerRow: unknown[],
    aliases: string[]
  ) {
    return headerRow.findIndex((cell) =>
      aliases.includes(normalizeHeader(cell))
    );
  }

  function findHeaderRowIndex(
    matrix: unknown[][]
  ) {
    const maxRows = Math.min(
      matrix.length,
      50
    );

    for (
      let rowIndex = 0;
      rowIndex < maxRows;
      rowIndex += 1
    ) {
      const row =
        matrix[rowIndex] ?? [];

      const urlColumn =
        findColumnIndex(
          row,
          urlHeaderAliases
        );

      if (urlColumn !== -1) {
        return rowIndex;
      }
    }

    return -1;
  }

  function findLikelyUrlColumn(
    matrix: unknown[][],
    startRow = 0
  ) {
    let bestColumn = -1;
    let bestCount = 0;

    const maxColumns = matrix.reduce(
      (max, row) =>
        Math.max(max, row.length),
      0
    );

    for (
      let columnIndex = 0;
      columnIndex < maxColumns;
      columnIndex += 1
    ) {
      let count = 0;

      for (
        let rowIndex = startRow;
        rowIndex < matrix.length;
        rowIndex += 1
      ) {
        if (
          isUrlLike(
            matrix[rowIndex]?.[
              columnIndex
            ]
          )
        ) {
          count += 1;
        }
      }

      if (count > bestCount) {
        bestCount = count;
        bestColumn = columnIndex;
      }
    }

    return bestCount > 0
      ? bestColumn
      : -1;
  }

  function buildItemsFromSheet(
    worksheet: XLSX.WorkSheet
  ) {
    const matrix =
      XLSX.utils.sheet_to_json<
        unknown[]
      >(worksheet, {
        header: 1,
        defval: "",
        raw: false,
        blankrows: false,
      });

    if (matrix.length === 0) {
      return [] as UrlItem[];
    }

    const headerRowIndex =
      findHeaderRowIndex(matrix);

    const hasHeader =
      headerRowIndex !== -1;

    const headerRow =
      hasHeader
        ? matrix[headerRowIndex] ?? []
        : [];

    const dataStartRow =
      hasHeader
        ? headerRowIndex + 1
        : 0;

    let urlColumn =
      hasHeader
        ? findColumnIndex(
            headerRow,
            urlHeaderAliases
          )
        : -1;

    if (urlColumn === -1) {
      urlColumn =
        findLikelyUrlColumn(
          matrix,
          dataStartRow
        );
    }

    if (urlColumn === -1) {
      return [] as UrlItem[];
    }

    const caseColumn =
      hasHeader
        ? findColumnIndex(
            headerRow,
            caseHeaderAliases
          )
        : -1;

    const aisColumn =
      hasHeader
        ? findColumnIndex(
            headerRow,
            providerAliases.ais
          )
        : -1;

    const trueDtacColumn =
      hasHeader
        ? findColumnIndex(
            headerRow,
            providerAliases.trueDtac
          )
        : -1;

    const ntColumn =
      hasHeader
        ? findColumnIndex(
            headerRow,
            providerAliases.nt
          )
        : -1;

    const cloudflareColumn =
      hasHeader
        ? findColumnIndex(
            headerRow,
            providerAliases.cloudflare
          )
        : -1;

    const result: UrlItem[] = [];

    for (
      let rowIndex = dataStartRow;
      rowIndex < matrix.length;
      rowIndex += 1
    ) {
      const row =
        matrix[rowIndex] ?? [];

      const rawUrl =
        row[urlColumn];

      if (!isUrlLike(rawUrl)) {
        continue;
      }

      const urlSms =
        normalizeUrl(rawUrl);

      const caseValue =
        caseColumn !== -1
          ? normalizeExcelText(
              row[caseColumn]
            )
          : "";

      result.push({
        id: result.length + 1,

        sourceRowIndex:
          rowIndex,

        caseId:
          caseValue ||
          `CASE-${String(
            result.length + 1
          ).padStart(3, "0")}`,

        urlSms,

        original: {
          ais:
            aisColumn !== -1
              ? normalizeExcelText(
                  row[aisColumn]
                )
              : "",

          trueDtac:
            trueDtacColumn !== -1
              ? normalizeExcelText(
                  row[trueDtacColumn]
                )
              : "",

          nt:
            ntColumn !== -1
              ? normalizeExcelText(
                  row[ntColumn]
                )
              : "",

          cloudflare:
            cloudflareColumn !== -1
              ? normalizeExcelText(
                  row[
                    cloudflareColumn
                  ]
                )
              : "",
        },

        checked: {
          ais: "",
          trueDtac: "",
          nt: "",
          cloudflare: "",
        },
      });
    }

    return result;
  }

  function getSourceColumnMap(
    worksheet: XLSX.WorkSheet
  ): SourceColumnMap | null {
    const matrix =
      XLSX.utils.sheet_to_json<
        unknown[]
      >(worksheet, {
        header: 1,
        defval: "",
        raw: false,
        blankrows: false,
      });

    if (matrix.length === 0) {
      return null;
    }

    const headerRowIndex =
      findHeaderRowIndex(
        matrix
      );

    if (headerRowIndex === -1) {
      return null;
    }

    const headerRow =
      matrix[
        headerRowIndex
      ] ?? [];

    const urlColumn =
      findColumnIndex(
        headerRow,
        urlHeaderAliases
      );

    if (urlColumn === -1) {
      return null;
    }

    return {
      headerRowIndex,
      urlColumn,

      caseColumn:
        findColumnIndex(
          headerRow,
          caseHeaderAliases
        ),

      aisColumn:
        findColumnIndex(
          headerRow,
          providerAliases.ais
        ),

      trueDtacColumn:
        findColumnIndex(
          headerRow,
          providerAliases.trueDtac
        ),

      ntColumn:
        findColumnIndex(
          headerRow,
          providerAliases.nt
        ),

      cloudflareColumn:
        findColumnIndex(
          headerRow,
          providerAliases.cloudflare
        ),
    };
  }

  async function handleFileChange(
    event:
      ChangeEvent<
        HTMLInputElement
      >
  ) {
    const file =
      event.target
        .files?.[0];

    if (!file) {
      setFileName("");
      setItems([]);
      setSourceSheetName("");
      setSourceColumnMap(null);
      originalWorkbookRef.current =
        null;
      originalExcelBufferRef.current =
        null;
      return;
    }

    setFileName(file.name);
    setOpenedSavedBatchId(null);
    setCheckStatuses({});

    try {
      const buffer =
        await file.arrayBuffer();

      // เก็บไฟล์ต้นฉบับไว้ทั้งก้อน
      // เพื่อ Export กลับด้วยรูปแบบเดิม
      originalExcelBufferRef.current =
        buffer.slice(0);

      const workbook =
        XLSX.read(
          buffer,
          {
            type: "array",

            // เก็บข้อมูลของ Workbook ต้นฉบับไว้
            // เพื่อใช้ตอน Download กลับเป็นไฟล์เดิม
            cellStyles: true,
            cellDates: true,
          }
        );

      originalWorkbookRef.current =
        workbook;

      if (
        workbook.SheetNames.length ===
        0
      ) {
        alert(
          "ไม่พบ Sheet ภายในไฟล์ Excel"
        );

        setItems([]);
        setSourceSheetName("");
        setSourceColumnMap(null);
        originalWorkbookRef.current =
          null;
        originalExcelBufferRef.current =
          null;
        return;
      }

      let importedItems:
        UrlItem[] = [];

      let selectedSheetName = "";

      for (
        const sheetName
        of workbook.SheetNames
      ) {
        const worksheet =
          workbook.Sheets[
            sheetName
          ];

        if (!worksheet) {
          continue;
        }

        const sheetItems =
          buildItemsFromSheet(
            worksheet
          );

        if (
          sheetItems.length >
          importedItems.length
        ) {
          importedItems =
            sheetItems;

          selectedSheetName =
            sheetName;
        }
      }

      if (
        importedItems.length === 0
      ) {
        setItems([]);
        setViewMode("original");
        setNetworkInfo(null);
        setSourceSheetName("");
        setSourceColumnMap(null);

        alert(
          "ไม่พบ URL ในไฟล์ Excel\n\n" +
            "ระบบลองค้นหาแล้วทั้งชื่อคอลัมน์ URL / URL SMS / Domain / Website / Link " +
            "และลองค้นหาคอลัมน์ที่มีข้อมูลลักษณะ URL อัตโนมัติแล้ว\n\n" +
            "กรุณาตรวจสอบว่าในไฟล์มี URL หรือ Domain จริง"
        );

        return;
      }

      const selectedWorksheet =
        workbook.Sheets[
          selectedSheetName
        ];

      const selectedColumnMap =
        selectedWorksheet
          ? getSourceColumnMap(
              selectedWorksheet
            )
          : null;

      setItems(importedItems);
      setViewMode("original");
      setNetworkInfo(null);
      setSourceSheetName(
        selectedSheetName
      );
      setSourceColumnMap(
        selectedColumnMap
      );

      alert(
        `อ่านไฟล์สำเร็จ ${importedItems.length} รายการ\n` +
          `Sheet: ${selectedSheetName}`
      );
    } catch (error) {
      console.error(
        "Excel import error:",
        error
      );

      setItems([]);
      setSourceSheetName("");
      setSourceColumnMap(null);
      originalWorkbookRef.current =
        null;
      originalExcelBufferRef.current =
        null;

      alert(
        "อ่านไฟล์ Excel ไม่สำเร็จ กรุณาตรวจสอบรูปแบบไฟล์"
      );
    } finally {
      event.target.value = "";
    }
  }

  /* =========================
     DISPLAY NAME
  ========================= */

  function getProviderName(
    value:
      ProviderKey =
        provider
  ) {
    if (value === "ais") {
      return "AIS";
    }

    if (
      value ===
      "trueDtac"
    ) {
      return "TRUE / DTAC";
    }

    if (value === "nt") {
      return "NT";
    }

    return "CloudFlare";
  }

  /* =========================
     NETWORK
  ========================= */

  async function handleDetectNetwork():
    Promise<
      NetworkInfo | null
    > {
    setIsDetectingNetwork(
      true
    );

    try {
      const response =
        await fetch(
          `${apiUrl}/api/network-info`
        );

      if (!response.ok) {
        throw new Error(
          "ตรวจ Network ไม่สำเร็จ"
        );
      }

      const data: NetworkInfo =
        await response.json();

      setNetworkInfo(data);

      return data;
    } catch (error) {
      console.error(
        "Network detection error:",
        error
      );

      alert(
        "ไม่สามารถตรวจสอบ Network ปัจจุบันได้"
      );

      return null;
    } finally {
      setIsDetectingNetwork(
        false
      );
    }
  }

  function handleOpenCloudflareWarp() {
    window.open(
      "https://one.one.one.one/",
      "_blank",
      "noopener,noreferrer"
    );
  }

  async function handleVerifyCloudflare() {
    const detected =
      await handleDetectNetwork();

    if (!detected) {
      return;
    }

    if (isCloudflareNetwork(detected)) {
      alert(
        `ตรวจพบ Cloudflare / WARP แล้ว ✅\n\n` +
          `Public IP: ${detected.ip ?? "-"}\n` +
          `ISP: ${detected.org ?? "-"}`
      );
      return;
    }

    alert(
      `ยังไม่ตรวจพบ Cloudflare / WARP\n\n` +
        `Public IP: ${detected.ip ?? "-"}\n` +
        `ISP: ${detected.org ?? "-"}\n\n` +
        `กรุณาเปิด Cloudflare One Client / WARP ให้ขึ้น Connected แล้วกด “ตรวจ WARP” อีกครั้ง`
    );
  }

  /* =========================
     RUN CHECK
  ========================= */

  function getCheckState(
    result: ApiCheckResult
  ): ProviderCheckState {
    if (result.reachable) {
      return "success";
    }

    if (
      result.errorType === "nxdomain" ||
      result.error
        ?.toUpperCase()
        .includes("ENOTFOUND")
    ) {
      return "nxdomain";
    }

    return "failed";
  }

  async function handleRunProcess() {
    if (
      items.length === 0
    ) {
      alert(
        "ยังไม่มี URL สำหรับตรวจ"
      );
      return;
    }

    const detected =
      await handleDetectNetwork();

    if (!detected) {
      return;
    }

    // Cloudflare ต้องตรวจพบจริงก่อนจึงจะบันทึกผลลงคอลัมน์ Cloudflare
    // เพื่อป้องกันการสแกนผ่าน AIS / TRUE / NT แล้วติดป้ายว่าเป็น Cloudflare
    if (
      provider === "cloudflare" &&
      !isCloudflareNetwork(detected)
    ) {
      alert(
        `ยังไม่ตรวจพบ Cloudflare One Client / WARP\n\n` +
          `Public IP: ${detected.ip ?? "-"}\n` +
          `ISP: ${detected.org ?? "-"}\n\n` +
          `1. เปิด Cloudflare One Client\n` +
          `2. กด Connect ให้ขึ้น Connected\n` +
          `3. กลับมากด “ตรวจ WARP”\n` +
          `4. เมื่อสถานะเป็น Connected แล้วจึง Run check`
      );
      return;
    }

    if (
      detected.provider !==
        "unknown" &&
      detected.provider !==
        provider
    ) {
      alert(
        `Network ไม่ตรงกัน\n\n` +
          `คุณเลือก: ${getProviderName()}\n` +
          `ระบบตรวจพบ: ${getProviderName(
            detected.provider
          )}\n` +
          `ISP: ${detected.org ?? "-"}\n` +
          `Public IP: ${detected.ip ?? "-"}`
      );

      return;
    }

    if (
      detected.provider ===
      "unknown"
    ) {
      const confirmed =
        window.confirm(
          `ระบบไม่สามารถระบุผู้ให้บริการได้\n\n` +
            `Public IP: ${detected.ip ?? "-"}\n` +
            `ISP: ${detected.org ?? "-"}\n\n` +
            `ยืนยันว่ากำลังใช้ ${getProviderName()} หรือไม่?`
        );

      if (!confirmed) {
        return;
      }
    }

    setIsProcessing(true);

    try {
      const response =
        await fetch(
          `${apiUrl}/api/check-url`,
          {
            method: "POST",

            headers: {
              "Content-Type":
                "application/json",
            },

            body:
              JSON.stringify({
                urls:
                  items.map(
                    (item) =>
                      item.urlSms
                  ),
              }),
          }
        );

      if (!response.ok) {
        throw new Error(
          "ตรวจ URL ไม่สำเร็จ"
        );
      }

      const data: ApiCheckResponse =
        await response.json();

      setItems(
        (current) =>
          current.map(
            (item) => {
              const result =
                data.results
                  .find(
                    (
                      resultItem
                    ) =>
                      resultItem
                        .url ===
                      item.urlSms
                  );

              if (!result) {
                return item;
              }

              const state =
                getCheckState(result);

              const checkText =
                state === "success"
                  ? result.finalUrl ||
                    result.url ||
                    item.urlSms
                  : state === "nxdomain"
                  ? "DNS_PROBE_FINISHED_NXDOMAIN"
                  : result.url ||
                    item.urlSms;

              setCheckStatuses(
                (currentStatuses) => ({
                  ...currentStatuses,
                  [item.id]: {
                    ...currentStatuses[item.id],
                    [provider]: state,
                  },
                })
              );

              return {
                ...item,

                original: {
                  ...item.original,

                  [provider]:
                    checkText,
                },

                checked: {
                  ...item.checked,

                  [provider]:
                    checkText,
                },
              };
            }
          )
      );

      setViewMode(
        "checked"
      );

      alert(
        `ตรวจเสร็จ ${data.total} รายการ\n` +
          `Network: ${getProviderName()}\n` +
          `IP: ${detected.ip ?? "-"}`
      );
    } catch (error) {
      console.error(
        "Run check error:",
        error
      );

      alert(
        "ตรวจ URL ไม่สำเร็จ กรุณาตรวจสอบ Backend"
      );
    } finally {
      setIsProcessing(
        false
      );
    }
  }

  /* =========================
     EDIT EXCEL CELL
  ========================= */

  function handleEditExcelCell(
    id: number,
    providerKey: ProviderKey,
    value: string
  ) {
    if (!canManage) {
      return;
    }

    setCheckStatuses(
      (currentStatuses) => ({
        ...currentStatuses,
        [id]: {
          ...currentStatuses[id],
          [providerKey]: "idle",
        },
      })
    );

    setItems((current) =>
      current.map((item) => {
        if (item.id !== id) {
          return item;
        }

        return {
          ...item,
          original: {
            ...item.original,
            [providerKey]: value,
          },
          checked: {
            ...item.checked,
            [providerKey]: value,
          },
        };
      })
    );
  }

  /* =========================
     SAVE DATABASE
  ========================= */

  async function handleSaveAll(
    requestedSaveMode: SaveMode = "all",
    requestedProvider: ProviderKey = provider
  ) {
    if (!canManage) {
      alert(
        "บัญชี Viewer ไม่มีสิทธิ์บันทึกหรือแก้ไขข้อมูล"
      );
      return;
    }

    if (
      items.length === 0
    ) {
      alert(
        "ไม่มีข้อมูลสำหรับบันทึก"
      );
      return;
    }

    if (
      openedSavedBatchId &&
      items.every(
        (item) => item.dbResultId
      )
    ) {
      if (!currentUser) {
        alert("ไม่พบข้อมูลผู้ใช้งาน");
        return;
      }

      try {
        setIsSaving(true);

        await Promise.all(
          items.map(async (item) => {
            const response = await fetch(
              `${apiUrl}/api/summary/edit/results/${item.dbResultId}`,
              {
                method: "PATCH",
                headers: {
                  "Content-Type":
                    "application/json",
                  "x-user-id":
                    String(currentUser.id),
                },
                body: JSON.stringify({
                  caseId: item.caseId,
                  urlSms: item.urlSms,
                  aisResult:
                    item.original.ais,
                  trueDtacResult:
                    item.original.trueDtac,
                  ntResult:
                    item.original.nt,
                  cloudflareResult:
                    item.original.cloudflare,
                }),
              }
            );

            const result =
              await response.json();

            if (!response.ok) {
              throw new Error(
                result?.message ??
                  "บันทึกการแก้ไขไม่สำเร็จ"
              );
            }
          })
        );

        alert(
          `บันทึกการแก้ไขข้อมูลล่าสุดสำเร็จ\n${items.length} รายการ`
        );

        setIsEditingSaved(false);
        await loadLatestSavedData();
      } catch (error) {
        console.error(
          "Update saved checker data error:",
          error
        );

        alert(
          error instanceof Error
            ? error.message
            : "บันทึกการแก้ไขไม่สำเร็จ"
        );
      } finally {
        setIsSaving(false);
      }

      return;
    }

    const hasCheckedData =
      items.some(
        (item) =>
          item.checked.ais
            .trim() !== "" ||
          item.checked.trueDtac
            .trim() !== "" ||
          item.checked.nt
            .trim() !== "" ||
          item.checked.cloudflare
            .trim() !== ""
      );

    if (!hasCheckedData) {
      alert(
        "ยังไม่มีผลการตรวจ กรุณากด Run check หรือระบุผล QR ก่อน"
      );
      return;
    }

    const publishToSummary = requestedSaveMode === "all";

    const buildCheckedForSave = (item: UrlItem): ProviderData => {
      if (requestedSaveMode !== "provider") {
        return { ...item.checked };
      }

      return {
        ais: requestedProvider === "ais" ? item.checked.ais : "",
        trueDtac:
          requestedProvider === "trueDtac" ? item.checked.trueDtac : "",
        nt: requestedProvider === "nt" ? item.checked.nt : "",
        cloudflare:
          requestedProvider === "cloudflare" ? item.checked.cloudflare : "",
      };
    };

    try {
      setIsSaving(true);

      const response =
        await fetch(
          `${apiUrl}/api/checker-results`,
          {
            method:
              "POST",

            headers: {
              "Content-Type":
                "application/json",
            },

            body:
              JSON.stringify({
                fileName,
                caseType: "fake_domain",
                caseStatus: "active",
                detectedDate: getToday(),
                saveType: requestedSaveMode,
                savedProvider:
                  requestedSaveMode === "provider"
                    ? requestedProvider
                    : null,
                publishToSummary,

                selectedProvider:
                  requestedSaveMode === "provider"
                    ? requestedProvider
                    : provider,

                networkInfo:
                  networkInfo
                    ? {
                        ip:
                          networkInfo.ip,

                        org:
                          networkInfo.org,

                        provider:
                          networkInfo.provider,
                      }
                    : null,

                items:
                  items.map(
                    (item) => ({
                      caseId:
                        item.caseId,

                      urlSms:
                        item.urlSms,

                      checked:
                        buildCheckedForSave(item),

                      checkState: {
                        ais:
                          checkStatuses[item.id]?.ais ?? null,

                        trueDtac:
                          checkStatuses[item.id]?.trueDtac ?? null,

                        nt:
                          checkStatuses[item.id]?.nt ?? null,

                        cloudflare:
                          checkStatuses[item.id]?.cloudflare ?? null,
                      },
                    })
                  ),
              }),
          }
        );

      const data: SaveResponse =
        await response.json();

      if (!response.ok) {
        alert(
          data.message ??
            "บันทึกข้อมูลไม่สำเร็จ"
        );
        return;
      }

      alert(
        `บันทึกข้อมูลสำเร็จ\n` +
          `${data.insertedCount ?? 0} รายการ`
      );

      setOpenedSavedBatchId(
        data.batchId ?? null
      );
      setShowSaveModal(false);
      await loadLatestSavedData();
      await loadHistoryData();
    } catch (error) {
      console.error(
        "Save error:",
        error
      );

      alert(
        "ไม่สามารถเชื่อมต่อ Backend ได้"
      );
    } finally {
      setIsSaving(false);
    }
  }

  /* =========================
     EXPORT EXCEL
  ========================= */

  async function handleExportExcel() {
    if (!canManage) {
      alert(
        "เฉพาะ Super Admin และ Admin เท่านั้นที่ดาวน์โหลดไฟล์ Excel ได้"
      );
      return;
    }

    if (items.length === 0) {
      alert(
        "ไม่มีข้อมูลสำหรับดาวน์โหลด"
      );
      return;
    }

    const sourceBuffer =
      originalExcelBufferRef.current;

    if (
      !sourceBuffer ||
      !sourceSheetName
    ) {
      alert(
        "ไม่พบไฟล์ Excel ต้นฉบับ กรุณา Import ใหม่อีกครั้ง"
      );
      return;
    }

    let columnMap =
      sourceColumnMap;

    const xlsxWorkbook =
      originalWorkbookRef.current;

    const xlsxWorksheet =
      xlsxWorkbook?.Sheets[
        sourceSheetName
      ];

    if (
      !columnMap &&
      xlsxWorksheet
    ) {
      columnMap =
        getSourceColumnMap(
          xlsxWorksheet
        );
    }

    if (!columnMap) {
      alert(
        "ไม่พบหัวตารางของ Excel ต้นฉบับ"
      );
      return;
    }

    try {
      /*
        สำคัญ:
        โหลดจาก byte ของไฟล์ Excel ต้นฉบับโดยตรงด้วย ExcelJS
        แล้วแก้เฉพาะ value ของ cell
        จึงรักษา style เดิม เช่น:
        - สีพื้นหลัง
        - สีตัวอักษร
        - border
        - merge cell
        - ความกว้างคอลัมน์
        - ความสูงแถว
        - alignment
      */
      const workbook =
        new ExcelJS.Workbook();

      await workbook.xlsx.load(
        sourceBuffer
      );

      const worksheet =
        workbook.getWorksheet(
          sourceSheetName
        );

      if (!worksheet) {
        alert(
          "ไม่พบ Sheet ต้นฉบับ กรุณา Import ใหม่อีกครั้ง"
        );
        return;
      }

      /*
        ExcelJS ใช้ row/column แบบ 1-based
        แต่ค่าที่เราเก็บจาก XLSX เป็น 0-based
        จึง +1 ตอนเข้าถึง Cell
      */

      const headerRowNumber =
        columnMap.headerRowIndex +
        1;

      let nextColumnNumber =
        Math.max(
          worksheet.columnCount,
          columnMap.urlColumn + 1,
          columnMap.caseColumn + 1,
          columnMap.aisColumn + 1,
          columnMap.trueDtacColumn + 1,
          columnMap.ntColumn + 1,
          columnMap.cloudflareColumn + 1
        ) + 1;

      const ensureColumn = (
        zeroBasedColumn: number,
        headerText: string
      ) => {
        if (
          zeroBasedColumn !== -1
        ) {
          return (
            zeroBasedColumn + 1
          );
        }

        const newColumnNumber =
          nextColumnNumber;

        nextColumnNumber += 1;

        worksheet.getCell(
          headerRowNumber,
          newColumnNumber
        ).value =
          headerText;

        return newColumnNumber;
      };

      const aisColumn =
        ensureColumn(
          columnMap.aisColumn,
          "AIS"
        );

      const trueDtacColumn =
        ensureColumn(
          columnMap.trueDtacColumn,
          "TRUE/DTAC"
        );

      const ntColumn =
        ensureColumn(
          columnMap.ntColumn,
          "NT"
        );

      const cloudflareColumn =
        ensureColumn(
          columnMap.cloudflareColumn,
          "CloundFlare"
        );

      for (const item of items) {
        const excelRow =
          item.sourceRowIndex + 1;

        /*
          แก้เฉพาะข้อความใน Cell เดิม
          ไม่แตะ style ของ Cell
        */
        worksheet.getCell(
          excelRow,
          aisColumn
        ).value =
          item.original.ais;

        worksheet.getCell(
          excelRow,
          trueDtacColumn
        ).value =
          item.original.trueDtac;

        worksheet.getCell(
          excelRow,
          ntColumn
        ).value =
          item.original.nt;

        worksheet.getCell(
          excelRow,
          cloudflareColumn
        ).value =
          item.original.cloudflare;
      }

      const outputBuffer =
        await workbook.xlsx.writeBuffer();

      const blob =
        new Blob(
          [outputBuffer],
          {
            type:
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          }
        );

      const lastDot =
        fileName.lastIndexOf(
          "."
        );

      const baseName =
        lastDot > 0
          ? fileName.slice(
              0,
              lastDot
            )
          : fileName ||
            "checker";

      const downloadName =
        `${baseName}_checked.xlsx`;

      const url =
        URL.createObjectURL(
          blob
        );

      const link =
        document.createElement(
          "a"
        );

      link.href = url;
      link.download =
        downloadName;

      document.body.appendChild(
        link
      );

      link.click();

      document.body.removeChild(
        link
      );

      URL.revokeObjectURL(
        url
      );
    } catch (error) {
      console.error(
        "Excel export error:",
        error
      );

      alert(
        "สร้างไฟล์ Excel ไม่สำเร็จ กรุณาตรวจสอบไฟล์ต้นฉบับ"
      );
    }
  }

  /* =========================
     RESULT DISPLAY
  ========================= */

  function renderProviderValue(
    value: string,
    originalUrl?: string,
    state: ProviderCheckState = "idle"
  ) {
    const cleanValue =
      String(value ?? "").trim();

    if (!cleanValue) {
      return (
        <span className="provider-empty">
          -
        </span>
      );
    }

    if (
      state === "nxdomain" ||
      cleanValue ===
        "DNS_PROBE_FINISHED_NXDOMAIN"
    ) {
      const href =
        originalUrl || "#";

      return (
        <a
          href={href}
          target={
            originalUrl
              ? "_blank"
              : undefined
          }
          rel={
            originalUrl
              ? "noreferrer"
              : undefined
          }
          className="provider-unreachable provider-result-link provider-nxdomain-link"
          title={
            originalUrl
              ? `เปิด URL ต้นฉบับ: ${originalUrl}`
              : "DNS_PROBE_FINISHED_NXDOMAIN"
          }
          onClick={
            originalUrl
              ? undefined
              : (event) =>
                  event.preventDefault()
          }
        >
          DNS_PROBE_FINISHED_NXDOMAIN
        </a>
      );
    }

    const isUrl =
      cleanValue.startsWith("http://") ||
      cleanValue.startsWith("https://");

    if (state === "failed") {
      const href =
        isUrl
          ? cleanValue
          : originalUrl || cleanValue;

      return (
        <a
          href={href}
          target="_blank"
          rel="noreferrer"
          className="provider-unreachable provider-result-link"
          title={`เปิดไม่ได้: ${href}`}
        >
          {href}
        </a>
      );
    }

    if (isUrl) {
      return (
        <a
          href={cleanValue}
          target="_blank"
          rel="noreferrer"
          className="provider-found provider-result-link"
          title={cleanValue}
        >
          {cleanValue}
        </a>
      );
    }

    return (
      <span className="provider-text">
        {cleanValue}
      </span>
    );
  }

  function renderEditableExcelCell(
    item: UrlItem,
    providerKey: ProviderKey
  ) {
    const value =
      item.original[providerKey];

    const state =
      checkStatuses[item.id]?.[
        providerKey
      ] ?? "idle";

    if (
      !canManage ||
      (openedSavedBatchId !== null &&
        !isEditingSaved)
    ) {
      return renderProviderValue(
        value,
        item.urlSms,
        state
      );
    }

    return (
      <input
        type="text"
        className={
          `checker-excel-cell-input ${
            state === "failed" ||
            state === "nxdomain" ||
            value ===
              "DNS_PROBE_FINISHED_NXDOMAIN"
              ? "is-not-found"
              : value.startsWith("http://") ||
                value.startsWith("https://")
              ? "is-found"
              : ""
          }`
        }
        value={value}
        placeholder="-"
        onChange={(event) =>
          handleEditExcelCell(
            item.id,
            providerKey,
            event.target.value
          )
        }
      />
    );
  }

  function renderCheckedProviderValue(
    value: string,
    originalUrl: string,
    state: ProviderCheckState = "idle"
  ) {
    const cleanValue =
      String(value ?? "").trim();

    if (!cleanValue) {
      return (
        <span className="provider-empty">
          -
        </span>
      );
    }

    if (
      state === "nxdomain" ||
      cleanValue ===
        "DNS_PROBE_FINISHED_NXDOMAIN"
    ) {
      return (
        <a
          href={originalUrl}
          target="_blank"
          rel="noreferrer"
          className="provider-unreachable provider-result-link provider-nxdomain-link"
          title={`เปิด URL ต้นฉบับ: ${originalUrl}`}
        >
          DNS_PROBE_FINISHED_NXDOMAIN
        </a>
      );
    }

    const linkUrl =
      cleanValue.startsWith("http://") ||
      cleanValue.startsWith("https://")
        ? cleanValue
        : originalUrl;

    if (state === "failed") {
      return (
        <a
          href={linkUrl}
          target="_blank"
          rel="noreferrer"
          className="provider-unreachable provider-result-link"
          title={`เปิดไม่ได้: ${linkUrl}`}
        >
          {linkUrl}
        </a>
      );
    }

    return (
      <a
        href={linkUrl}
        target="_blank"
        rel="noreferrer"
        className="provider-found provider-result-link"
        title={`เปิดได้: ${linkUrl}`}
      >
        ✓ {linkUrl}
      </a>
    );
  }

  function getSelectedProviderValue(
    item: UrlItem
  ) {
    return item.checked[provider];
  }

  return (
    <main className="checker-page">
      <div className="checker-container">
        <section className="checker-heading">
          <div>
            <h1>URL Fake Web</h1>

            <p>
              ตรวจสอบ URL จากไฟล์ Excel
              ผ่านเครือข่ายที่เลือก
              และบันทึกผลลงฐานข้อมูล
            </p>
          </div>

          <button
            type="button"
            className="checker-history-compact"
            onClick={() => {
              setSelectedHistoryDate(getToday());
              setCalendarMonth(() => {
                const now = new Date();
                return new Date(now.getFullYear(), now.getMonth(), 1);
              });
              setShowHistoryModal(true);
              void loadHistoryData();
            }}
          >
            <span className="checker-history-compact-icon">▣</span>
            <span className="checker-history-compact-text">
              <small>ประวัติการบันทึก</small>
              <strong>ดูประวัติย้อนหลัง</strong>
            </span>
            <span className="checker-history-compact-arrow">›</span>
          </button>

          <div className="checker-latest-saved">
            <div className="checker-latest-saved-top">
              <span>ข้อมูลที่บันทึกล่าสุด</span>

              {latestSavedBatch && (
                <span className="checker-latest-batch-id">
                  Batch #{latestSavedBatch.id}
                </span>
              )}
            </div>

            {isLoadingSaved ? (
              <strong>กำลังโหลด...</strong>
            ) : latestSavedBatch ? (
              <>
                <strong
                  title={
                    latestSavedBatch.original_file_name ??
                    ""
                  }
                >
                  {latestSavedBatch.original_file_name ??
                    "ไม่ระบุชื่อไฟล์"}
                </strong>

                <div className="checker-latest-saved-meta">
                  <span>
                    {latestSavedCount.toLocaleString(
                      "th-TH"
                    )}{" "}
                    รายการ
                  </span>

                  <span>
                    {latestSavedBatch.detected_date ??
                      latestSavedBatch.created_at?.slice(
                        0,
                        10
                      ) ??
                      "-"}
                  </span>
                </div>

              </>
            ) : (
              <strong>ยังไม่มีข้อมูลที่บันทึก</strong>
            )}

          </div>
        </section>

        {!canManage && currentUser && (
          <div className="checker-permission-note">
            บัญชี Viewer ดูข้อมูลได้อย่างเดียว — การ Import, ตรวจ, แก้ผล และดาวน์โหลด Excel
            ใช้ได้เฉพาะ Super Admin และ Admin
          </div>
        )}


        <section className="checker-toolbar">
          <div className="checker-toolbar-item">
            <label className="checker-button checker-import-button">
              ↑{" "}
              {fileName ||
                "Import spreadsheet"}

              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={
                  handleFileChange
                }
                disabled={!canManage}
                hidden
              />
            </label>

            <span>
              รองรับ .xlsx, .xls
            </span>
          </div>

          <div className="checker-toolbar-item">
            <select
              className="checker-network-select"
              value={provider}
              onChange={(event) => {
                const value = event.target.value;

                if (
                  value === "ais" ||
                  value === "trueDtac" ||
                  value === "nt" ||
                  value === "cloudflare"
                ) {
                  setProvider(value);
                  setNetworkInfo(null);
                }
              }}
            >
              <option value="ais">
                AIS
              </option>

              <option value="trueDtac">
                TRUE / DTAC
              </option>

              <option value="nt">
                NT
              </option>

              <option value="cloudflare">
                CloudFlare
              </option>
            </select>

            <span>
              เครือข่ายที่ทดสอบ
            </span>
          </div>

          <div className="checker-toolbar-item">
            <button
              type="button"
              className="checker-button checker-network-button"
              onClick={
                handleDetectNetwork
              }
              disabled={
                !canManage ||
                isDetectingNetwork
              }
            >
              {isDetectingNetwork
                ? "กำลังตรวจ..."
                : "ตรวจ Network"}
            </button>

            <span>
              ตรวจ Public IP / ISP
            </span>
          </div>

          <div className="checker-toolbar-item">
            <button
              type="button"
              className="checker-button checker-run-button"
              onClick={
                handleRunProcess
              }
              disabled={
                !canManage ||
                isProcessing ||
                isDetectingNetwork ||
                isSaving ||
                items.length === 0
              }
            >
              ▶{" "}
              {isProcessing
                ? "กำลังตรวจ..."
                : provider === "cloudflare"
                  ? "Run Cloudflare check"
                  : "Run check"}
            </button>

            <span>
              ตรวจ URL
            </span>
          </div>

          <div className="checker-toolbar-item checker-search-item">
            <div className="checker-search">
              <span>⌕</span>

              <input
                type="text"
                placeholder="Search CASE ID, URL..."
                value={search}
                onChange={
                  (event) =>
                    setSearch(
                      event.target
                        .value
                    )
                }
              />
            </div>

            <span>
              ค้นหารายการ
            </span>
          </div>

          <div className="checker-toolbar-item">
            <button
              type="button"
              className="checker-button checker-save-button"
              onClick={() => {
                if (
                  openedSavedBatchId &&
                  items.every((item) => item.dbResultId)
                ) {
                  void handleSaveAll("all", provider);
                  return;
                }

                setSaveProvider(provider);
                setSaveMode("all");
                setShowSaveModal(true);
              }}
              disabled={
                !canManage ||
                items.length === 0 ||
                isSaving ||
                isProcessing
              }
            >
              {isSaving
                ? "กำลังบันทึก..."
                : "บันทึกข้อมูล"}
            </button>

            <span>
              เลือกรูปแบบการบันทึก
            </span>
          </div>
        </section>

        {provider === "cloudflare" && (
          <section
            className={`checker-warp-panel ${
              cloudflareConnected
                ? "is-connected"
                : "is-disconnected"
            }`}
          >
            <div className="checker-warp-icon">
              ☁
            </div>

            <div className="checker-warp-content">
              <div className="checker-warp-title-row">
                <div>
                  <span className="checker-warp-eyebrow">
                    CLOUDFLARE ONE CLIENT / WARP
                  </span>
                  <h3>
                    ตรวจ URL ผ่านเครือข่าย Cloudflare
                  </h3>
                </div>

                <span
                  className={`checker-warp-status ${
                    cloudflareConnected
                      ? "connected"
                      : "waiting"
                  }`}
                >
                  {cloudflareConnected
                    ? "✓ ตรวจพบ Cloudflare"
                    : "• รอเชื่อมต่อ WARP"}
                </span>
              </div>

              <p>
                1. เปิดหรือติดตั้ง Cloudflare One Client / WARP → 2. กด Connect ให้ขึ้น Connected → 3. กลับมากด “ตรวจ WARP” → 4. จึงกด Run Cloudflare check
              </p>

              <div className="checker-warp-actions">
                <button
                  type="button"
                  className="checker-warp-open"
                  onClick={
                    handleOpenCloudflareWarp
                  }
                  disabled={!canManage}
                >
                  ↗ เปิด / ดาวน์โหลด Cloudflare WARP
                </button>

                <button
                  type="button"
                  className="checker-warp-verify"
                  onClick={
                    handleVerifyCloudflare
                  }
                  disabled={
                    !canManage ||
                    isDetectingNetwork
                  }
                >
                  {isDetectingNetwork
                    ? "กำลังตรวจ WARP..."
                    : "✓ ตรวจ WARP"}
                </button>
              </div>

              <div className="checker-warp-network">
                <span>Public IP: <strong>{networkInfo?.ip ?? "-"}</strong></span>
                <span>ISP: <strong>{networkInfo?.org ?? "-"}</strong></span>
              </div>

              <small>
                หมายเหตุ: Browser ไม่สามารถกด Connect ในแอปแทนผู้ใช้ได้ ระบบจึงตรวจ Network ก่อน Run เพื่อยืนยันว่าเส้นทางที่ Backend เห็นเป็น Cloudflare จริง
              </small>
            </div>
          </section>
        )}

        <section className="checker-network-grid">

          <article className="checker-info-card">
            <span>
              Network ที่เลือก
            </span>

            <strong>
              {getProviderName()}
            </strong>
          </article>

          <article className="checker-info-card">
            <span>Public IP</span>

            <strong>
              {networkInfo?.ip ??
                "-"}
            </strong>
          </article>

          <article className="checker-info-card">
            <span>ISP</span>

            <strong>
              {networkInfo?.org ??
                "-"}
            </strong>
          </article>
        </section>

        <section className="checker-line-daily-panel">
          <div className="checker-line-daily-head">
            <div>
              <span>LINE DAILY WATCH</span>
              <h3>4 URL หลักที่ใช้ตรวจแบบตรงเป๊ะ</h3>
              <p>ระบบคัดเฉพาะ URL ที่ token ตรงกับ 4 ตัวนี้ 100% เท่านั้น ไม่ใช้ Network และไม่คิดเปอร์เซ็นต์ใกล้เคียง</p>
            </div>
            <div className="checker-line-daily-status">
              <strong>{lineSimilarityRows.length}</strong>
              <span>รายการที่พบวันนี้</span>
            </div>
          </div>

          <div className="checker-line-reference-grid">
            {OPEN_LINE_REFERENCE_URLS.map((url, index) => {
              const foundCount = lineSimilarityRows.filter(
                (row) => row.token === OPEN_LINE_REFERENCE_PARTS[index].token
              ).length;

              return (
                <article key={url} className={foundCount > 0 ? "is-found" : ""}>
                  <div>
                    <span>ตัวหลัก #{index + 1}</span>
                    <a href={url} target="_blank" rel="noreferrer">{url}</a>
                  </div>
                  <strong>{foundCount > 0 ? `พบ ${foundCount}` : "ไม่พบ"}</strong>
                </article>
              );
            })}
          </div>

          <small>วันที่ตรวจบนหน้านี้: {getToday()} • เมื่อเปิดข้อมูล/นำเข้าไฟล์ ระบบจะคัดผลจากรายการปัจจุบันให้อัตโนมัติ</small>
        </section>

        <section className="checker-table-card">
          <div className="checker-table-header">
            <div>
              <h2>
                Verification Queue
              </h2>

              <p>
                {viewMode === "line_similarity"
                  ? `พบตรงเป๊ะ ${lineSimilarityRows.length} รายการ`
                  : `${filteredItems.length} cases loaded`}
              </p>
            </div>

            <div className="checker-result-tabs">
              <button
                type="button"
                className={
                  viewMode ===
                  "original"
                    ? "active"
                    : ""
                }
                onClick={
                  () =>
                    setViewMode(
                      "original"
                    )
                }
              >
                หน้า Excel (แก้ไขได้)
              </button>

              <button
                type="button"
                className={
                  viewMode ===
                  "checked"
                    ? "active"
                    : ""
                }
                onClick={
                  () =>
                    setViewMode(
                      "checked"
                    )
                }
              >
                ผลตรวจ Network
              </button>

              <button
                type="button"
                className={
                  viewMode ===
                  "line_similarity"
                    ? "active"
                    : ""
                }
                onClick={
                  () =>
                    setViewMode(
                      "line_similarity"
                    )
                }
              >
                LINE ตรง 4 ตัวหลัก
              </button>
            </div>

            {canManage && (
              <div className="checker-table-actions">
                {viewMode === "line_similarity" && (
                  <div className="checker-line-view-controls">

                    <select
                      value={
                        lineSimilarityCount
                      }
                      onChange={(event) =>
                        setLineSimilarityCount(
                          event.target.value
                        )
                      }
                    >
                      <option value="3">
                        แสดง 3
                      </option>
                      <option value="5">
                        แสดง 5
                      </option>
                      <option value="10">
                        แสดง 10
                      </option>
                      <option value="20">
                        แสดง 20
                      </option>
                    </select>

                    <button
                      type="button"
                      onClick={
                        handleExportLineSimilarity
                      }
                      disabled={
                        selectedSimilarityRows.length ===
                        0
                      }
                    >
                      ↓ ดาวน์โหลดรายการที่ตรง
                    </button>
                  </div>
                )}

                <button
                  type="button"
                  className="checker-export-button"
                  onClick={
                    handleExportExcel
                  }
                  disabled={
                    items.length === 0
                  }
                >
                  ↓ ดาวน์โหลด Excel เดิม
                </button>

                {openedSavedBatchId !== null && (
                  <button
                    type="button"
                    className={`checker-edit-latest-button ${
                      isEditingSaved
                        ? "active"
                        : ""
                    }`}
                    onClick={() => {
                      setViewMode("original");
                      setIsEditingSaved(
                        (current) => !current
                      );
                    }}
                  >
                    {isEditingSaved
                      ? "✓ กำลังแก้ไข"
                      : "✎ แก้ไข"}
                  </button>
                )}
              </div>
            )}
          </div>

          <div className="checker-table-wrapper">
            <table className="checker-table">
              <thead>
                {viewMode === "line_similarity" ? (
                  <tr>
                    <th>อันดับ</th>
                    <th>CASE ID</th>
                    <th>URL LINE</th>
                    <th>ช่วงต้น</th>
                    <th>ช่วงท้าย</th>
                    <th>ตรงกี่ตัว</th>
                    <th>ตรงกับตัวหลัก</th>
                    <th>ผล</th>
                  </tr>
                ) : viewMode === "checked" ? (
                  <tr>
                    <th>CASE ID</th>
                    <th>URL SMS</th>
                    <th>
                      {getProviderName()}
                    </th>
                  </tr>
                ) : (
                  <tr>
                    <th>CASE ID</th>
                    <th>URL SMS</th>
                    <th>AIS</th>
                    <th>
                      TRUE / DTAC
                    </th>
                    <th>NT</th>
                    <th>
                      CloudFlare
                    </th>
                  </tr>
                )}
              </thead>

              <tbody>
                {viewMode === "line_similarity" ? (
                  selectedSimilarityRows.length ===
                  0 ? (
                    <tr>
                      <td
                        colSpan={8}
                        className="checker-empty-table"
                      >
                        วันนี้ยังไม่พบ LINE URL ที่ตรงเป๊ะกับ 4 ตัวหลัก
                      </td>
                    </tr>
                  ) : (
                    selectedSimilarityRows.map(
                      (row, index) => {
                        const checkedUrl =
                          row.lineUrl;

                        const matchedReference =
                          row.matchedReference
                            ? OPEN_LINE_REFERENCE_PARTS[
                                row.matchedReference -
                                  1
                              ]
                            : null;

                        return (
                          <tr
                            key={`line-${row.item.id}-${row.lineUrl}`}
                          >
                            <td>
                              <strong>
                                {index + 1}
                              </strong>
                            </td>

                            <td>
                              <strong>
                                {
                                  row.item
                                    .caseId
                                }
                              </strong>
                            </td>

                            <td>
                              <a
                                href={
                                  checkedUrl
                                }
                                target="_blank"
                                rel="noreferrer"
                                className="checker-url-link checker-line-result-url"
                                title={
                                  checkedUrl
                                }
                              >
                                {
                                  checkedUrl
                                }
                              </a>
                            </td>

                            <td>
                              <code
                                className="checker-line-token-part"
                                title={
                                  matchedReference
                                    ? `เทียบกับ ${matchedReference.prefix}`
                                    : ""
                                }
                              >
                                {matchedReference
                                  ? renderMatchedCharacters(
                                      row.prefix,
                                      matchedReference.token,
                                      0
                                    )
                                  : row.prefix}
                              </code>
                            </td>

                            <td>
                              <code
                                className="checker-line-token-part"
                                title={
                                  matchedReference
                                    ? `เทียบกับ ${matchedReference.suffix}`
                                    : ""
                                }
                              >
                                {matchedReference
                                  ? renderMatchedCharacters(
                                      row.suffix,
                                      matchedReference.token,
                                      matchedReference.token.length -
                                        matchedReference.suffix.length
                                    )
                                  : row.suffix}
                              </code>
                            </td>

                            <td>
                              <span className="checker-line-match-count">
                                {row.matchedCount}/{row.comparedLength}
                              </span>
                            </td>

                            <td>
                              <span
                                className="checker-line-reference-badge"
                                title={
                                  matchedReference
                                    ? `${matchedReference.prefix} | ${matchedReference.suffix}`
                                    : ""
                                }
                              >
                                {row.matchedReference
                                  ? `#${row.matchedReference}`
                                  : "-"}
                              </span>
                            </td>

                            <td>
                              <div className="checker-line-score">
                                <div className="checker-line-score-track">
                                  <span
                                    style={{
                                      width: `${row.score}%`,
                                    }}
                                  />
                                </div>

                                <strong>
                                  {row.score}%
                                </strong>

                                <span
                                  className={`checker-line-score-label ${
                                    row.score >=
                                    90
                                      ? "is-high"
                                      : row.score >=
                                          75
                                        ? "is-medium"
                                        : "is-low"
                                  }`}
                                >
                                  {
                                    row.label
                                  }
                                </span>
                              </div>
                            </td>
                          </tr>
                        );
                      }
                    )
                  )
                ) : filteredItems.length ===
                  0 ? (
                  <tr>
                    <td
                      colSpan={
                        viewMode ===
                        "checked"
                          ? 3
                          : 6
                      }
                      className="checker-empty-table"
                    >
                      ยังไม่มีข้อมูล
                      กรุณา Import
                      ไฟล์ Excel
                    </td>
                  </tr>
                ) : (
                  filteredItems.map(
                    (item) => {
                      return (
                        <tr
                          key={
                            item.id
                          }
                        >
                          <td>
                            <strong>
                              {
                                item.caseId
                              }
                            </strong>
                          </td>

                          <td>
                            <a
                              href={
                                item.urlSms
                              }
                              target="_blank"
                              rel="noreferrer"
                              className="checker-url-link"
                            >
                              {
                                item.urlSms
                              }
                            </a>
                          </td>

                          {viewMode ===
                          "checked" ? (
                            <td>
                              {renderCheckedProviderValue(
                                getSelectedProviderValue(
                                  item
                                ),
                                item.urlSms,
                                checkStatuses[
                                  item.id
                                ]?.[
                                  provider
                                ] ??
                                  (getSelectedProviderValue(
                                    item
                                  ) ===
                                  "DNS_PROBE_FINISHED_NXDOMAIN"
                                    ? "nxdomain"
                                    : "idle")
                              )}
                            </td>
                          ) : (
                            <>
                              <td>
                                {renderEditableExcelCell(
                                  item,
                                  "ais"
                                )}
                              </td>

                              <td>
                                {renderEditableExcelCell(
                                  item,
                                  "trueDtac"
                                )}
                              </td>

                              <td>
                                {renderEditableExcelCell(
                                  item,
                                  "nt"
                                )}
                              </td>

                              <td>
                                {renderEditableExcelCell(
                                  item,
                                  "cloudflare"
                                )}
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    }
                  )
                )}
              </tbody>
            </table>
          </div>
        </section>

        {showHistoryModal && (
          <div
            className="checker-history-modal-backdrop"
            role="presentation"
            onMouseDown={(event) => {
              if (event.target === event.currentTarget) {
                setShowHistoryModal(false);
              }
            }}
          >
            <section
              className="checker-history-modal"
              role="dialog"
              aria-modal="true"
              aria-labelledby="checker-history-modal-title"
            >
              <div className="checker-history-modal-head">
                <div>
                  <span>HISTORY CALENDAR</span>
                  <h2 id="checker-history-modal-title">เลือกวันที่ที่ต้องการดู</h2>
                  <p>วันที่ที่มีจุดสีน้ำเงินคือวันที่มีข้อมูลบันทึกไว้</p>
                </div>

                <button
                  type="button"
                  className="checker-history-modal-close"
                  onClick={() => setShowHistoryModal(false)}
                  aria-label="ปิด"
                >
                  ×
                </button>
              </div>

              <div className="checker-history-modal-body">
                <div className="checker-history-popup-calendar">
                  <div className="checker-calendar-header">
                    <button
                      type="button"
                      onClick={() =>
                        setCalendarMonth(
                          new Date(
                            calendarMonth.getFullYear(),
                            calendarMonth.getMonth() - 1,
                            1
                          )
                        )
                      }
                      aria-label="เดือนก่อนหน้า"
                    >
                      ‹
                    </button>

                    <strong>
                      {calendarMonth.toLocaleDateString("th-TH", {
                        month: "long",
                        year: "numeric",
                      })}
                    </strong>

                    <button
                      type="button"
                      onClick={() =>
                        setCalendarMonth(
                          new Date(
                            calendarMonth.getFullYear(),
                            calendarMonth.getMonth() + 1,
                            1
                          )
                        )
                      }
                      aria-label="เดือนถัดไป"
                    >
                      ›
                    </button>
                  </div>

                  <div className="checker-calendar-weekdays">
                    {["จ", "อ", "พ", "พฤ", "ศ", "ส", "อา"].map((label) => (
                      <span key={label}>{label}</span>
                    ))}
                  </div>

                  <div className="checker-calendar-days">
                    {calendarDays.map((cell) => {
                      if (!cell.date) {
                        return (
                          <span
                            key={cell.key}
                            className="checker-calendar-empty"
                          />
                        );
                      }

                      const isSelected = cell.key === selectedHistoryDate;
                      const isToday = cell.key === getToday();

                      return (
                        <button
                          key={cell.key}
                          type="button"
                          className={[
                            "checker-calendar-day",
                            isSelected ? "is-selected" : "",
                            isToday ? "is-today" : "",
                            cell.count > 0 ? "has-data" : "",
                          ]
                            .filter(Boolean)
                            .join(" ")}
                          onClick={() => setSelectedHistoryDate(cell.key)}
                        >
                          <span>{cell.date.getDate()}</span>
                          {cell.count > 0 && <small>{cell.count}</small>}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className="checker-history-popup-result">
                  <div className="checker-history-popup-selected">
                    <div>
                      <span>วันที่เลือก</span>
                      <strong>{formatThaiDate(selectedHistoryDate)}</strong>
                    </div>
                    <span className="checker-history-count">
                      {selectedHistoryBatches.length} รอบ
                    </span>
                  </div>

                  <div className="checker-history-popup-list">
                    {isLoadingHistory ? (
                      <div className="checker-history-empty">
                        กำลังโหลดประวัติ...
                      </div>
                    ) : selectedHistoryBatches.length === 0 ? (
                      <div className="checker-history-empty">
                        วันที่นี้ยังไม่มีข้อมูลที่บันทึก
                      </div>
                    ) : (
                      selectedHistoryBatches.map((batch) => (
                        <article
                          key={batch.id}
                          className="checker-history-popup-item"
                        >
                          <div>
                            <div className="checker-history-item-title">
                              <strong>
                                {batch.original_file_name ?? `Batch #${batch.id}`}
                              </strong>
                              <span>Batch #{batch.id}</span>
                            </div>

                            <div className="checker-history-meta">
                              <span>🕒 {formatThaiTime(batch.saved_at ?? batch.created_at)}</span>
                              <span>
                                📄 {(batch.result_count ?? 0).toLocaleString("th-TH")} รายการ
                              </span>
                              <span>
                                {batch.save_type === "provider"
                                  ? `เฉพาะ ${getProviderLabel(batch.saved_provider)}`
                                  : batch.save_type === "history"
                                    ? "ประวัติเท่านั้น"
                                    : "ทุกเครือข่าย"}
                              </span>
                            </div>
                          </div>

                          <button
                            type="button"
                            className="checker-history-open"
                            onClick={() => {
                              setShowHistoryModal(false);
                              void openHistoryBatch(batch);
                            }}
                          >
                            เปิดดูข้อมูล
                          </button>
                        </article>
                      ))
                    )}
                  </div>
                </div>
              </div>

              <div className="checker-history-modal-footer">
                <button
                  type="button"
                  className="checker-history-refresh"
                  onClick={() => void loadHistoryData()}
                  disabled={isLoadingHistory}
                >
                  {isLoadingHistory ? "กำลังโหลด..." : "↻ รีเฟรชประวัติ"}
                </button>
                <button
                  type="button"
                  className="checker-history-modal-done"
                  onClick={() => setShowHistoryModal(false)}
                >
                  ปิด
                </button>
              </div>
            </section>
          </div>
        )}

        {showSaveModal && (
          <div
            className="checker-save-modal-backdrop"
            role="presentation"
            onMouseDown={(event) => {
              if (event.target === event.currentTarget && !isSaving) {
                setShowSaveModal(false);
              }
            }}
          >
            <section
              className="checker-save-modal"
              role="dialog"
              aria-modal="true"
              aria-labelledby="checker-save-title"
            >
              <div className="checker-save-modal-head">
                <div>
                  <span>SAVE CHECKER RESULT</span>
                  <h2 id="checker-save-title">เลือกวิธีบันทึกข้อมูล</h2>
                  <p>ข้อมูลวันที่ {formatThaiDate(getToday())}</p>
                </div>

                <button
                  type="button"
                  className="checker-save-close"
                  onClick={() => setShowSaveModal(false)}
                  disabled={isSaving}
                  aria-label="ปิด"
                >
                  ×
                </button>
              </div>

              <div className="checker-save-options">
                <button
                  type="button"
                  className={`checker-save-option ${saveMode === "all" ? "active" : ""}`}
                  onClick={() => setSaveMode("all")}
                >
                  <span className="checker-save-option-icon">▦</span>
                  <span>
                    <strong>บันทึกทั้งหมด + Main Summary</strong>
                    <small>AIS, TRUE / DTAC, NT และ Cloudflare พร้อมส่งขึ้นหน้าสรุป</small>
                  </span>
                </button>

                <button
                  type="button"
                  className={`checker-save-option ${saveMode === "provider" ? "active" : ""}`}
                  onClick={() => setSaveMode("provider")}
                >
                  <span className="checker-save-option-icon">⌁</span>
                  <span>
                    <strong>บันทึกเฉพาะเครือข่าย</strong>
                    <small>เก็บเฉพาะผลของเครือข่ายที่เลือก และไม่ส่งขึ้น Main Summary</small>
                  </span>
                </button>

                <button
                  type="button"
                  className={`checker-save-option ${saveMode === "history" ? "active" : ""}`}
                  onClick={() => setSaveMode("history")}
                >
                  <span className="checker-save-option-icon">◷</span>
                  <span>
                    <strong>บันทึกเป็นประวัติเท่านั้น</strong>
                    <small>เก็บผลทุกเครือข่ายไว้ดูย้อนหลัง แต่ไม่กระทบ Main Summary</small>
                  </span>
                </button>
              </div>

              {saveMode === "provider" && (
                <div className="checker-save-provider-box">
                  <label htmlFor="checker-save-provider">เลือกเครือข่ายที่จะบันทึก</label>
                  <select
                    id="checker-save-provider"
                    value={saveProvider}
                    onChange={(event) =>
                      setSaveProvider(event.target.value as ProviderKey)
                    }
                  >
                    <option value="ais">AIS</option>
                    <option value="trueDtac">TRUE / DTAC</option>
                    <option value="nt">NT</option>
                    <option value="cloudflare">Cloudflare</option>
                  </select>
                </div>
              )}

              <div className="checker-save-summary-box">
                <div>
                  <span>ไฟล์</span>
                  <strong>{fileName || "ไม่ระบุชื่อไฟล์"}</strong>
                </div>
                <div>
                  <span>จำนวน</span>
                  <strong>{items.length.toLocaleString("th-TH")} รายการ</strong>
                </div>
                <div>
                  <span>Main Summary</span>
                  <strong className={saveMode === "all" ? "is-publish" : "is-history"}>
                    {saveMode === "all" ? "อัปเดต" : "ไม่อัปเดต"}
                  </strong>
                </div>
              </div>

              <div className="checker-save-modal-actions">
                <button
                  type="button"
                  className="checker-save-cancel"
                  onClick={() => setShowSaveModal(false)}
                  disabled={isSaving}
                >
                  ยกเลิก
                </button>
                <button
                  type="button"
                  className="checker-save-confirm"
                  onClick={() => void handleSaveAll(saveMode, saveProvider)}
                  disabled={isSaving}
                >
                  {isSaving ? "กำลังบันทึก..." : "ยืนยันการบันทึก"}
                </button>
              </div>
            </section>
          </div>
        )}
      </div>
    </main>
  );
}

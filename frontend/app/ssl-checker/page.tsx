"use client";

import {
  ChangeEvent,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import * as XLSX from "xlsx";
import { useRouter } from "next/navigation";
import "./ssl-checker.css";

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

type RoundStatus =
  | "WAITING"
  | "RUNNING"
  | "PAUSED"
  | "COMPLETED"
  | "ERROR";

type SSLJob = {
  id: number;
  original_file_name: string;
  total_urls: number;
  round_size: number;
  request_size: number;
  total_rounds: number;
  checked_count: number;
  completed_rounds: number;
  current_round: number;
  status: JobStatus;
  created_by: number | null;
  last_error: string | null;
  created_at: string;
  updated_at: string;
  started_at: string | null;
  paused_at: string | null;
  completed_at: string | null;
  valid_count?: number;
  expiring_count?: number;
  expired_count?: number;
  error_count?: number;
};

type SSLRound = {
  id: number;
  job_id: number;
  round_number: number;
  start_order: number;
  end_order: number;
  total_urls: number;
  checked_count: number;
  status: RoundStatus;
  started_at: string | null;
  paused_at: string | null;
  completed_at: string | null;
  updated_at: string;
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

type StatusFilter =
  | "ALL"
  | SSLStatus;

type ResultColumnKey =
  | "order"
  | "url"
  | "validFrom"
  | "expirationDate"
  | "daysLeft"
  | "status"
  | "issuer"
  | "checkedAt"
  | "error";

const RESULT_COLUMNS: Array<{
  key: ResultColumnKey;
  label: string;
}> = [
  { key: "order", label: "ลำดับในไฟล์" },
  { key: "url", label: "URL / Hostname" },
  { key: "validFrom", label: "วันเริ่ม" },
  { key: "expirationDate", label: "วันหมด" },
  { key: "daysLeft", label: "วันที่เหลือ" },
  { key: "status", label: "สถานะ" },
  { key: "issuer", label: "Issuer" },
  { key: "checkedAt", label: "เวลาที่ตรวจ" },
  { key: "error", label: "Error" },
];

type UserInfo = {
  id: number;
  name: string;
  email: string;
  role:
    | "super_admin"
    | "editor"
    | "viewer";
};

const API_URL =
  process.env.NEXT_PUBLIC_API_URL &&
  !process.env.NEXT_PUBLIC_API_URL.includes("backend:")
    ? process.env.NEXT_PUBLIC_API_URL.replace(/\/$/, "")
    : "http://localhost:4000";

const ROUND_SIZE = 500;
const TABLE_PAGE_SIZE = 100;
const STORAGE_JOB_KEY =
  "ssl_checker_current_job_id";
const STORAGE_HISTORY_KEY =
  "ssl_checker_saved_job_ids";
const MAX_SAVED_JOBS = 20;

const URL_HEADER_CANDIDATES = [
  "url",
  "website",
  "web",
  "domain",
  "link",
  "เว็บไซต์",
  "ลิงก์",
  "ลิงค์",
  "โดเมน",
];

function cleanText(value: unknown) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(
      /[\u200B-\u200D\uFEFF]/g,
      ""
    )
    .replace(/\u00A0/g, " ")
    .trim();
}

function normalizeHeader(value: unknown) {
  return cleanText(value)
    .toLowerCase()
    .replace(/\s+/g, "");
}

function normalizeInputUrl(value: unknown) {
  const text = cleanText(value);

  if (!text) {
    return "";
  }

  const withoutQuotes =
    text.replace(
      /^["']|["']$/g,
      ""
    );

  if (
    /^https?:\/\//i.test(
      withoutQuotes
    )
  ) {
    return withoutQuotes;
  }

  if (
    /^[a-z0-9.-]+\.[a-z]{2,}(:\d+)?(\/.*)?$/i.test(
      withoutQuotes
    )
  ) {
    return `https://${withoutQuotes}`;
  }

  return "";
}

function uniqueUrls(values: string[]) {
  const seen =
    new Set<string>();

  const result:
    string[] = [];

  for (const value of values) {
    const normalized =
      normalizeInputUrl(value);

    if (!normalized) {
      continue;
    }

    let key =
      normalized;

    try {
      const parsed =
        new URL(normalized);

      key =
        `${parsed.protocol}//${parsed.hostname.toLowerCase()}:${
          parsed.port ||
          (
            parsed.protocol ===
            "https:"
              ? "443"
              : "80"
          )
        }${
          parsed.pathname === "/"
            ? ""
            : parsed.pathname
        }`;
    } catch {
      key =
        normalized.toLowerCase();
    }

    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    result.push(normalized);
  }

  return result;
}

function extractUrlsFromWorkbook(
  workbook: XLSX.WorkBook
) {
  const collected:
    string[] = [];

  for (
    const sheetName of
    workbook.SheetNames
  ) {
    const worksheet =
      workbook.Sheets[
        sheetName
      ];

    if (!worksheet) {
      continue;
    }

    const matrix =
      XLSX.utils.sheet_to_json<
        unknown[]
      >(
        worksheet,
        {
          header: 1,
          defval: "",
          raw: false,
        }
      );

    if (
      matrix.length === 0
    ) {
      continue;
    }

    const headerSearchLimit =
      Math.min(
        matrix.length,
        12
      );

    let urlColumn = -1;
    let headerRowIndex = -1;

    for (
      let r = 0;
      r <
      headerSearchLimit;
      r += 1
    ) {
      const row =
        matrix[r] ?? [];

      for (
        let c = 0;
        c < row.length;
        c += 1
      ) {
        const header =
          normalizeHeader(
            row[c]
          );

        if (
          URL_HEADER_CANDIDATES.some(
            (
              candidate
            ) =>
              header ===
                normalizeHeader(
                  candidate
                ) ||
              header.includes(
                normalizeHeader(
                  candidate
                )
              )
          )
        ) {
          urlColumn = c;
          headerRowIndex = r;
          break;
        }
      }

      if (
        urlColumn >= 0
      ) {
        break;
      }
    }

    if (urlColumn >= 0) {
      for (
        let r =
          headerRowIndex + 1;
        r < matrix.length;
        r += 1
      ) {
        const url =
          normalizeInputUrl(
            matrix[r]?.[
              urlColumn
            ]
          );

        if (url) {
          collected.push(url);
        }
      }

      continue;
    }

    for (
      const row of matrix
    ) {
      for (
        const value of
        row ?? []
      ) {
        const url =
          normalizeInputUrl(
            value
          );

        if (url) {
          collected.push(url);
        }
      }
    }
  }

  return uniqueUrls(
    collected
  );
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
        "medium",
    }
  ).format(date);
}

function statusLabel(
  status: SSLStatus
) {
  switch (status) {
    case "VALID":
      return "VALID";

    case "EXPIRING_SOON":
      return "EXPIRING SOON";

    case "EXPIRED":
      return "EXPIRED";

    case "ERROR":
      return "ERROR";
  }
}

function jobStatusLabel(
  status: JobStatus
) {
  switch (status) {
    case "READY":
      return "พร้อมตรวจ";

    case "RUNNING":
      return "กำลังตรวจ";

    case "PAUSED":
      return "หยุดพัก";

    case "COMPLETED":
      return "เสร็จสิ้น";

    case "ERROR":
      return "เกิดข้อผิดพลาด";
  }
}

async function readJsonResponse<T>(
  response: Response
): Promise<T> {
  const contentType =
    response.headers.get(
      "content-type"
    ) ?? "";

  const raw =
    await response.text();

  if (
    !contentType.includes(
      "application/json"
    )
  ) {
    throw new Error(
      `API ไม่ได้ส่ง JSON กลับมา (HTTP ${response.status})`
    );
  }

  let data:
    T & {
      message?: string;
    };

  try {
    data =
      (
        raw
          ? JSON.parse(raw)
          : {}
      ) as T & {
        message?: string;
      };
  } catch {
    throw new Error(
      "Backend ส่ง JSON ที่อ่านไม่ได้"
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

function readCurrentUser():
  UserInfo | null {
  if (
    typeof window ===
    "undefined"
  ) {
    return null;
  }

  const candidates = [
    "user",
    "currentUser",
    "authUser",
    "op_user",
  ];

  for (
    const key of candidates
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

export default function SSLCheckerPage() {
  const router = useRouter();

  const inputRef =
    useRef<HTMLInputElement | null>(
      null
    );

  const [
    fileName,
    setFileName,
  ] =
    useState("");

  const [
    urls,
    setUrls,
  ] =
    useState<string[]>(
      []
    );

  const [
    job,
    setJob,
  ] =
    useState<SSLJob | null>(
      null
    );

  const [
    savedJobs,
    setSavedJobs,
  ] =
    useState<SSLJob[]>([]);

  const [
    historyLoading,
    setHistoryLoading,
  ] =
    useState(false);

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
    roundResults,
    setRoundResults,
  ] =
    useState<SSLResultRow[]>(
      []
    );

  const [
    roundResultTotal,
    setRoundResultTotal,
  ] =
    useState(0);

  const [
    search,
    setSearch,
  ] =
    useState("");

  const [
    filter,
    setFilter,
  ] =
    useState<StatusFilter>(
      "ALL"
    );

  const [
    selectedColumns,
    setSelectedColumns,
  ] = useState<ResultColumnKey[]>([
    "order",
    "url",
    "validFrom",
    "expirationDate",
    "daysLeft",
    "status",
    "issuer",
    "checkedAt",
    "error",
  ]);

  const [
    tablePage,
    setTablePage,
  ] =
    useState(1);

  const [
    roundCount,
    setRoundCount,
  ] =
    useState(1);

  const [
    runAll,
    setRunAll,
  ] =
    useState(false);

  const [
    busy,
    setBusy,
  ] =
    useState(false);

  const [
    roundsOpen,
    setRoundsOpen,
  ] =
    useState(false);

  const [
    errorMessage,
    setErrorMessage,
  ] =
    useState("");

  const [
    currentUser,
    setCurrentUser,
  ] =
    useState<UserInfo | null>(
      null
    );

  const isViewer =
    currentUser?.role ===
    "viewer";

  const canControl =
    currentUser?.role ===
      "super_admin" ||
    currentUser?.role ===
      "editor";

  const canEdit =
    currentUser?.role ===
      "super_admin";

  function readSavedJobIds() {
    if (typeof window === "undefined") {
      return [] as number[];
    }

    try {
      const raw = window.localStorage.getItem(STORAGE_HISTORY_KEY);
      const parsed = raw ? JSON.parse(raw) : [];

      if (!Array.isArray(parsed)) {
        return [] as number[];
      }

      return parsed
        .map((value) => Number(value))
        .filter((value) => Number.isInteger(value) && value > 0)
        .slice(0, MAX_SAVED_JOBS);
    } catch {
      return [] as number[];
    }
  }

  function rememberJobId(jobId: number) {
    if (typeof window === "undefined") {
      return;
    }

    const nextIds = [
      jobId,
      ...readSavedJobIds().filter((id) => id !== jobId),
    ].slice(0, MAX_SAVED_JOBS);

    window.localStorage.setItem(
      STORAGE_HISTORY_KEY,
      JSON.stringify(nextIds)
    );
  }

  async function loadSavedJobs() {
    const ids = readSavedJobIds();

    if (ids.length === 0) {
      setSavedJobs([]);
      return;
    }

    setHistoryLoading(true);

    try {
      const results = await Promise.all(
        ids.map(async (id) => {
          try {
            const response = await fetch(
              `${API_URL}/api/ssl-checker/jobs/${id}`,
              { cache: "no-store" }
            );

            if (!response.ok) {
              return null;
            }

            const data = await readJsonResponse<{ job: SSLJob }>(response);
            return data.job;
          } catch {
            return null;
          }
        })
      );

      const jobs = results
        .filter((item): item is SSLJob => item !== null)
        .sort(
          (a, b) =>
            new Date(b.updated_at).getTime() -
            new Date(a.updated_at).getTime()
        );

      setSavedJobs(jobs);

      window.localStorage.setItem(
        STORAGE_HISTORY_KEY,
        JSON.stringify(jobs.map((item) => item.id))
      );
    } finally {
      setHistoryLoading(false);
    }
  }

  async function openSavedJob(jobId: number) {
    if (job?.status === "RUNNING" && job.id !== jobId) {
      setErrorMessage(
        "กรุณาหยุดพักงานที่กำลังตรวจอยู่ก่อนเปิดงานอื่น"
      );
      return;
    }

    setBusy(true);
    setErrorMessage("");

    try {
      window.localStorage.setItem(STORAGE_JOB_KEY, String(jobId));
      rememberJobId(jobId);
      await loadJob(jobId);
      await loadSavedJobs();
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "เปิดงานที่บันทึกไว้ไม่สำเร็จ"
      );
    } finally {
      setBusy(false);
    }
  }

  async function loadJob(
    jobId: number,
    options?: {
      keepSelectedRound?: boolean;
    }
  ) {
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
      await readJsonResponse<{
        job: SSLJob;
      }>(
        jobResponse
      );

    const roundsData =
      await readJsonResponse<{
        rounds: SSLRound[];
      }>(
        roundsResponse
      );

    setJob(
      jobData.job
    );

    rememberJobId(jobData.job.id);
    setSavedJobs((current) => {
      const next = [
        jobData.job,
        ...current.filter((item) => item.id !== jobData.job.id),
      ];

      return next
        .sort(
          (a, b) =>
            new Date(b.updated_at).getTime() -
            new Date(a.updated_at).getTime()
        )
        .slice(0, MAX_SAVED_JOBS);
    });

    setRounds(
      roundsData.rounds
    );

    setFileName(
      jobData.job
        .original_file_name
    );

    if (
      !options
        ?.keepSelectedRound
    ) {
      const suggestedRound =
        jobData.job
          .current_round > 0
          ? jobData.job
              .current_round
          : Math.max(
              1,
              jobData.job
                .completed_rounds
            );

      setSelectedRound(
        Math.min(
          Math.max(
            1,
            suggestedRound
          ),
          Math.max(
            1,
            jobData.job
              .total_rounds
          )
        )
      );

      setTablePage(1);
    }
  }

  async function loadRoundResults(
    jobId: number,
    roundNumber: number,
    page = 1
  ) {
    const response =
      await fetch(
        `${API_URL}/api/ssl-checker/jobs/${jobId}/rounds/${roundNumber}/results?page=${page}&pageSize=${TABLE_PAGE_SIZE}`,
        {
          cache:
            "no-store",
        }
      );

    const data =
      await readJsonResponse<{
        page: number;
        pageSize: number;
        total: number;
        results:
          SSLResultRow[];
      }>(
        response
      );

    setRoundResults(
      data.results
    );

    setRoundResultTotal(
      data.total
    );
  }

  useEffect(() => {
    setCurrentUser(
      readCurrentUser()
    );

    loadSavedJobs().catch((error) => {
      console.error("Load saved SSL jobs error:", error);
    });

    const stored =
      window.localStorage.getItem(
        STORAGE_JOB_KEY
      );

    const jobId =
      Number(stored);

    if (
      Number.isInteger(
        jobId
      ) &&
      jobId > 0
    ) {
      loadJob(jobId)
        .catch(
          (
            error
          ) => {
            console.error(
              "Load saved SSL job error:",
              error
            );

            window.localStorage.removeItem(
              STORAGE_JOB_KEY
            );
          }
        );
    }
  }, []);

  useEffect(() => {
    if (!job) {
      return;
    }

    loadRoundResults(
      job.id,
      selectedRound,
      tablePage
    ).catch(
      (
        error
      ) => {
        console.error(
          "Load round results error:",
          error
        );
      }
    );
  }, [
    job?.id,
    selectedRound,
    tablePage,
  ]);

  useEffect(() => {
    if (
      !job ||
      job.status !==
        "RUNNING"
    ) {
      return;
    }

    const timer =
      window.setInterval(
        () => {
          loadJob(
            job.id,
            {
              keepSelectedRound:
                true,
            }
          )
            .then(() =>
              loadRoundResults(
                job.id,
                selectedRound,
                tablePage
              )
            )
            .catch(
              (
                error
              ) => {
                console.error(
                  "Poll SSL job error:",
                  error
                );
              }
            );
        },
        2000
      );

    return () => {
      window.clearInterval(
        timer
      );
    };
  }, [
    job?.id,
    job?.status,
    selectedRound,
    tablePage,
  ]);

  async function handleFileChange(
    event:
      ChangeEvent<HTMLInputElement>
  ) {
    if (isViewer) {
      event.target.value = "";
      setErrorMessage(
        "บัญชี Viewer สามารถดูข้อมูลได้อย่างเดียว ไม่สามารถอัปโหลดไฟล์ได้"
      );
      return;
    }

    const file =
      event.target.files?.[
        0
      ];

    if (!file) {
      return;
    }

    setErrorMessage("");
    setJob(null);
    setRounds([]);
    setRoundResults([]);
    setRoundResultTotal(0);
    setSelectedRound(1);
    setTablePage(1);

    try {
      const extension =
        file.name
          .split(".")
          .pop()
          ?.toLowerCase();

      let extracted:
        string[] = [];

      if (
        extension ===
        "txt"
      ) {
        const text =
          await file.text();

        extracted =
          uniqueUrls(
            text
              .split(
                /\r?\n/
              )
              .map(
                (
                  line
                ) =>
                  line.trim()
              )
          );
      } else if (
        extension ===
          "xlsx" ||
        extension ===
          "xls"
      ) {
        const buffer =
          await file.arrayBuffer();

        const workbook =
          XLSX.read(
            buffer,
            {
              type:
                "array",
            }
          );

        extracted =
          extractUrlsFromWorkbook(
            workbook
          );
      } else {
        throw new Error(
          "รองรับเฉพาะไฟล์ .txt, .xlsx และ .xls"
        );
      }

      setFileName(
        file.name
      );

      if (
        extracted.length ===
        0
      ) {
        setUrls([]);
        setErrorMessage(
          "ไม่พบ URL หรือ Domain ที่สามารถนำไปตรวจสอบได้ในไฟล์"
        );
        return;
      }

      setUrls(
        extracted
      );
    } catch (
      error
    ) {
      console.error(
        "Read SSL file error:",
        error
      );

      setFileName("");
      setUrls([]);

      setErrorMessage(
        error instanceof
          Error
          ? error.message
          : "อ่านไฟล์ไม่สำเร็จ"
      );
    } finally {
      event.target.value =
        "";
    }
  }

  async function createJob():
    Promise<SSLJob> {
    if (isViewer) {
      throw new Error(
        "บัญชี Viewer ไม่มีสิทธิ์สร้างงานตรวจใหม่"
      );
    }

    if (
      urls.length === 0 ||
      !fileName
    ) {
      throw new Error(
        "กรุณาอัปโหลดไฟล์ก่อน"
      );
    }

    const response =
      await fetch(
        `${API_URL}/api/ssl-checker/jobs`,
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
              urls,
              createdBy:
                currentUser?.id ??
                null,
            }),
        }
      );

    const data =
      await readJsonResponse<{
        job: SSLJob;
      }>(
        response
      );

    window.localStorage.setItem(
      STORAGE_JOB_KEY,
      String(
        data.job.id
      )
    );

    rememberJobId(data.job.id);

    setJob(
      data.job
    );

    await loadJob(
      data.job.id
    );
    await loadSavedJobs();

    return data.job;
  }

  async function startOrResume() {
    if (!canControl) {
      setErrorMessage(
        "บัญชีนี้ไม่มีสิทธิ์เริ่มหรือดำเนินการตรวจสอบ"
      );
      return;
    }

    setBusy(true);
    setErrorMessage("");

    try {
      let targetJob =
        job;

      if (!targetJob) {
        targetJob =
          await createJob();
      }

      const response =
        await fetch(
          `${API_URL}/api/ssl-checker/jobs/${targetJob.id}/start`,
          {
            method:
              "POST",
            headers: {
              "Content-Type":
                "application/json",
            },
            body:
              JSON.stringify(
                runAll
                  ? {
                      mode:
                        "all",
                    }
                  : {
                      mode:
                        "rounds",
                      roundCount:
                        Math.max(
                          1,
                          roundCount
                        ),
                    }
              ),
          }
        );

      await readJsonResponse<{
        message:
          string;
      }>(
        response
      );

      await loadJob(
        targetJob.id,
        {
          keepSelectedRound:
            false,
        }
      );
    } catch (
      error
    ) {
      setErrorMessage(
        error instanceof
          Error
          ? error.message
          : "เริ่มตรวจไม่สำเร็จ"
      );
    } finally {
      setBusy(false);
    }
  }

  async function pauseJob() {
    if (
      !job ||
      !canControl
    ) {
      return;
    }

    setBusy(true);
    setErrorMessage("");

    try {
      const response =
        await fetch(
          `${API_URL}/api/ssl-checker/jobs/${job.id}/pause`,
          {
            method:
              "POST",
          }
        );

      await readJsonResponse<{
        message:
          string;
      }>(
        response
      );

      await loadJob(
        job.id,
        {
          keepSelectedRound:
            true,
        }
      );
    } catch (
      error
    ) {
      setErrorMessage(
        error instanceof
          Error
          ? error.message
          : "หยุดพักไม่สำเร็จ"
      );
    } finally {
      setBusy(false);
    }
  }

  function chooseRound(
    roundNumber:
      number
  ) {
    setSelectedRound(
      roundNumber
    );

    setSearch("");
    setFilter("ALL");
    setTablePage(1);
    setRoundsOpen(false);
  }

  async function resetJob() {
    if (
      !job ||
      !currentUser ||
      currentUser.role !==
        "super_admin"
    ) {
      return;
    }

    if (
      job.status ===
      "RUNNING"
    ) {
      setErrorMessage(
        "กรุณาหยุดพักก่อนล้างผลตรวจ"
      );
      return;
    }

    const confirmed =
      window.confirm(
        "ต้องการล้างผลตรวจทั้งหมดของงานนี้ใช่หรือไม่?\nผลทุก URL จะถูกลบ และรอบทั้งหมดจะกลับเป็น WAITING"
      );

    if (!confirmed) {
      return;
    }

    setBusy(true);
    setErrorMessage("");

    try {
      const response =
        await fetch(
          `${API_URL}/api/ssl-checker/jobs/${job.id}/reset`,
          {
            method:
              "POST",
            headers: {
              "x-user-id":
                String(
                  currentUser.id
                ),
            },
          }
        );

      await readJsonResponse<{
        message:
          string;
      }>(
        response
      );

      setSelectedRound(1);
      setTablePage(1);
      setRoundResults([]);
      setRoundResultTotal(0);

      await loadJob(
        job.id,
        {
          keepSelectedRound:
            false,
        }
      );
    } catch (
      error
    ) {
      setErrorMessage(
        error instanceof
          Error
          ? error.message
          : "ล้างผลตรวจไม่สำเร็จ"
      );
    } finally {
      setBusy(false);
    }
  }

  function startNewFile() {
    if (isViewer) {
      setErrorMessage(
        "บัญชี Viewer สามารถดูข้อมูลได้อย่างเดียว ไม่สามารถอัปโหลดไฟล์ใหม่ได้"
      );
      return;
    }

    if (
      job?.status ===
      "RUNNING"
    ) {
      setErrorMessage(
        "กรุณาหยุดพักงานปัจจุบันก่อนอัปโหลดไฟล์ใหม่"
      );
      return;
    }

    window.localStorage.removeItem(
      STORAGE_JOB_KEY
    );

    setJob(null);
    setRounds([]);
    setRoundResults([]);
    setRoundResultTotal(0);
    setUrls([]);
    setFileName("");
    setSelectedRound(1);
    setTablePage(1);
    setErrorMessage("");

    inputRef.current?.click();
  }

  const selectedRoundData =
    rounds.find(
      (
        item
      ) =>
        item.round_number ===
        selectedRound
    ) ?? null;

  const filteredResults =
    useMemo(() => {
      const keyword =
        search
          .trim()
          .toLowerCase();

      return roundResults.filter(
        (
          item
        ) => {
          const matchesStatus =
            filter ===
              "ALL" ||
            item.status ===
              filter;

          if (
            !matchesStatus
          ) {
            return false;
          }

          if (!keyword) {
            return true;
          }

          return [
            item.url,
            item.hostname,
            item.issuer,
            item.subject,
            item.error ??
              "",
          ].some(
            (
              value
            ) =>
              value
                .toLowerCase()
                .includes(
                  keyword
                )
          );
        }
      );
    }, [
      roundResults,
      search,
      filter,
    ]);

  const roundPageSummary =
    useMemo(
      () => ({
        total:
          filteredResults.length,
        valid:
          filteredResults.filter(
            (
              item
            ) =>
              item.status ===
              "VALID"
          ).length,
        expiring:
          filteredResults.filter(
            (
              item
            ) =>
              item.status ===
              "EXPIRING_SOON"
          ).length,
        expired:
          filteredResults.filter(
            (
              item
            ) =>
              item.status ===
              "EXPIRED"
          ).length,
        error:
          filteredResults.filter(
            (
              item
            ) =>
              item.status ===
              "ERROR"
          ).length,
      }),
      [
        filteredResults,
      ]
    );

  const progressPercent =
    job &&
    job.total_urls > 0
      ? Math.min(
          100,
          Math.round(
            (
              job.checked_count /
              job.total_urls
            ) *
              100
          )
        )
      : 0;

  const isColumnVisible = (
    key: ResultColumnKey
  ) => selectedColumns.includes(key);

  const visibleColumnCount =
    Math.max(1, selectedColumns.length);

  function toggleResultColumn(
    key: ResultColumnKey
  ) {
    setSelectedColumns((current) => {
      if (current.includes(key)) {
        // ต้องเหลืออย่างน้อย 1 คอลัมน์
        if (current.length === 1) {
          return current;
        }

        return current.filter(
          (column) => column !== key
        );
      }

      const next = [...current, key];

      return RESULT_COLUMNS
        .map((column) => column.key)
        .filter((column) =>
          next.includes(column)
        );
    });
  }

  const totalPages =
    Math.max(
      1,
      Math.ceil(
        roundResultTotal /
          TABLE_PAGE_SIZE
      )
    );

  const safePage =
    Math.min(
      tablePage,
      totalPages
    );

  function downloadCurrentPage() {
    if (
      filteredResults.length ===
      0
    ) {
      return;
    }

    const rows =
      filteredResults.map(
        (
          item
        ) => ({
          ORDER:
            item.file_order,
          URL:
            item.url,
          HOSTNAME:
            item.hostname,
          PORT:
            item.port,
          VALID_FROM:
            item.valid_from
              ? formatDateTime(
                  item.valid_from
                )
              : "",
          EXPIRATION_DATE:
            item.expiration_date
              ? formatDateTime(
                  item.expiration_date
                )
              : "",
          DAYS_LEFT:
            item.days_left ??
            "",
          STATUS:
            item.status,
          ISSUER:
            item.issuer,
          SUBJECT:
            item.subject,
          CHECKED_AT:
            formatDateTime(
              item.checked_at
            ),
          ERROR:
            item.error ??
            "",
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
      `Round ${selectedRound}`
    );

    XLSX.writeFile(
      workbook,
      `ssl-round-${selectedRound}.xlsx`
    );
  }

  const totalUrlsDisplay =
    job?.total_urls ??
    urls.length;

  return (
    <main className="ssl-page">
      <div className="ssl-container">
        <section className="ssl-heading">
          <div>
            <span className="ssl-kicker">
              OP WEB SERVICE / SECURITY
            </span>

            <h1>
              SSL / TLS Checker
            </h1>

            <p>
              บันทึกผลการตรวจลงฐานข้อมูล สามารถหยุดพักและกลับมาตรวจต่อได้
            </p>
          </div>

          {!isViewer && (
            <>
              <button
                type="button"
                className="ssl-import-button"
                onClick={
                  startNewFile
                }
                disabled={
                  job?.status ===
                  "RUNNING"
                }
              >
                ↑ อัปโหลดไฟล์ใหม่
              </button>

              <input
                ref={inputRef}
                type="file"
                accept=".txt,.xlsx,.xls"
                hidden
                onChange={
                  handleFileChange
                }
              />
            </>
          )}
        </section>

        <section className="ssl-saved-section">
          <div className="ssl-saved-heading">
            <div>
              <span className="ssl-saved-kicker">งานที่บันทึกไว้</span>
              <h2>เปิดงานเดิมและตรวจต่อจากรอบที่ค้างไว้</h2>
              <p>แต่ละ Job เก็บผลลงฐานข้อมูลอยู่แล้ว จึงไม่ต้องอัปโหลดไฟล์เดิมใหม่</p>
            </div>

            <button
              type="button"
              className="ssl-saved-refresh"
              disabled={historyLoading || busy}
              onClick={() => loadSavedJobs()}
            >
              {historyLoading ? "กำลังโหลด..." : "↻ รีเฟรช"}
            </button>
          </div>

          {savedJobs.length === 0 ? (
            <div className="ssl-saved-empty">
              ยังไม่มีงานที่บันทึกไว้ — เมื่อสร้างงานครั้งแรก ระบบจะจำ Job นี้ให้อัตโนมัติ
            </div>
          ) : (
            <div className="ssl-saved-grid">
              {savedJobs.map((savedJob) => {
                const percent = savedJob.total_urls > 0
                  ? Math.min(100, Math.round((savedJob.checked_count / savedJob.total_urls) * 100))
                  : 0;
                const isCurrent = job?.id === savedJob.id;

                return (
                  <article
                    key={savedJob.id}
                    className={`ssl-saved-card ${isCurrent ? "is-current" : ""}`}
                  >
                    <div className="ssl-saved-card-top">
                      <div className="ssl-saved-file">
                        <span>Job #{savedJob.id}</span>
                        <strong title={savedJob.original_file_name}>
                          {savedJob.original_file_name}
                        </strong>
                      </div>

                      <span className={`ssl-saved-status ${savedJob.status.toLowerCase()}`}>
                        {jobStatusLabel(savedJob.status)}
                      </span>
                    </div>

                    <div className="ssl-saved-stats">
                      <div>
                        <span>ตรวจแล้ว</span>
                        <strong>{savedJob.checked_count.toLocaleString()} / {savedJob.total_urls.toLocaleString()}</strong>
                      </div>
                      <div>
                        <span>รอบ</span>
                        <strong>{savedJob.completed_rounds.toLocaleString()} / {savedJob.total_rounds.toLocaleString()}</strong>
                      </div>
                      <div>
                        <span>คืบหน้า</span>
                        <strong>{percent}%</strong>
                      </div>
                    </div>

                    <div className="ssl-saved-progress">
                      <div style={{ width: `${percent}%` }} />
                    </div>

                    <div className="ssl-saved-card-bottom">
                      <small>อัปเดต {formatDateTime(savedJob.updated_at)}</small>
                      <button
                        type="button"
                        disabled={busy || isCurrent}
                        onClick={() => openSavedJob(savedJob.id)}
                      >
                        {isCurrent
                          ? "กำลังเปิดอยู่"
                          : savedJob.status === "PAUSED"
                            ? "เปิดและตรวจต่อ"
                            : "เปิดดู"}
                      </button>
                    </div>
                  </article>
                );
              })}
            </div>
          )}
        </section>

        <section className="ssl-file-card">
          <div className="ssl-file-block">
            <span>
              ไฟล์ที่อัปโหลด
            </span>

            <strong>
              {fileName ||
                "ยังไม่ได้เลือกไฟล์"}
            </strong>

            <small>
              รองรับ TXT (.txt) และ Excel (.xlsx, .xls)
            </small>
          </div>

          <div className="ssl-file-stat">
            <span>
              URL ทั้งหมด
            </span>

            <strong>
              {totalUrlsDisplay.toLocaleString()}
            </strong>

            <small>
              URL
            </small>
          </div>

          <div className="ssl-file-stat">
            <span>
              สถานะ
            </span>

            <strong
              className={`ssl-job-state ${
                job?.status?.toLowerCase() ??
                "ready"
              }`}
            >
              {job
                ? jobStatusLabel(
                    job.status
                  )
                : "ยังไม่ได้สร้างงาน"}
            </strong>
          </div>

          <div className="ssl-file-stat">
            <span>
              ผู้ใช้งาน
            </span>

            <strong className="ssl-user-name">
              {currentUser?.name ??
                "-"}
            </strong>

            <small>
              {currentUser?.role ??
                "ไม่พบ role"}
            </small>
          </div>
        </section>

        <section className="ssl-control-card">
          <div className="ssl-control-main ssl-control-main-clean">
            <div className="ssl-run-setting">
              <span className="ssl-control-label">
                จำนวนรอบที่จะตรวจต่อ
              </span>

              <div className="ssl-round-run-control">
                <input
                  type="number"
                  min={1}
                  value={roundCount}
                  disabled={
                    isViewer ||
                    runAll ||
                    job?.status === "RUNNING" ||
                    busy
                  }
                  className={
                    runAll ? "is-disabled" : ""
                  }
                  onChange={(event) =>
                    setRoundCount(
                      Math.max(1, Number(event.target.value) || 1)
                    )
                  }
                />

                <span>รอบ</span>

                <label className={`ssl-all-toggle ${runAll ? "selected" : ""}`}>
                  <input
                    type="checkbox"
                    checked={runAll}
                    disabled={isViewer || job?.status === "RUNNING" || busy}
                    onChange={(event) => setRunAll(event.target.checked)}
                  />
                  ตรวจทั้งไฟล์
                </label>
              </div>

              <small>1 รอบ = 500 URL • บันทึกผลทุก 100 URL</small>
            </div>

            <div className="ssl-control-actions ssl-control-actions-clean">
              {job && rounds.length > 0 && (
                <button
                  type="button"
                  className="ssl-round-menu-button"
                  onClick={() => setRoundsOpen(true)}
                >
                  ☰ ดูรอบ
                </button>
              )}

              {canEdit && job && (
                <button
                  type="button"
                  className="ssl-edit-page-button ssl-edit-main-button"
                  disabled={job.status === "RUNNING" || job.checked_count === 0}
                  onClick={() => router.push(`/ssl-checker/edit?jobId=${job.id}`)}
                >
                  ✎ แก้ไขผล
                </button>
              )}

              {canEdit && job && (
                <button
                  type="button"
                  className="ssl-reset-button"
                  disabled={busy || job.status === "RUNNING" || job.checked_count === 0}
                  onClick={resetJob}
                >
                  ↺ ล้างผลตรวจ
                </button>
              )}

              {job?.status === "RUNNING" ? (
                <button
                  type="button"
                  className="ssl-pause-button"
                  disabled={busy || !canControl}
                  onClick={pauseJob}
                >
                  ⏸ หยุดพัก
                </button>
              ) : (
                <button
                  type="button"
                  className="ssl-check-button"
                  disabled={
                    busy ||
                    !canControl ||
                    (!job && urls.length === 0) ||
                    job?.status === "COMPLETED"
                  }
                  onClick={startOrResume}
                >
                  {job?.status === "PAUSED"
                    ? "▶ ตรวจต่อ"
                    : job
                      ? "▶ เริ่มตรวจ"
                      : "▶ สร้างงานและเริ่มตรวจ"}
                </button>
              )}
            </div>
          </div>

          {!canControl && (
            <div className="ssl-permission-note">
              บัญชี Viewer เป็นสิทธิ์ดูอย่างเดียว ไม่สามารถอัปโหลด เริ่ม/หยุด/ตรวจต่อ แก้ไข หรือล้างผลได้
            </div>
          )}
        </section>

        {job && (
          <section className="ssl-progress-card">
            <div className="ssl-progress-grid">
              <div>
                <span>
                  ความคืบหน้าโดยรวม
                </span>

                <strong className="ssl-progress-number">
                  {progressPercent}%
                </strong>

                <small>
                  ตรวจแล้ว{" "}
                  {job.checked_count.toLocaleString()}{" "}
                  /{" "}
                  {job.total_urls.toLocaleString()}{" "}
                  URL
                </small>
              </div>

              <div>
                <span>
                  รอบที่เสร็จแล้ว
                </span>

                <strong className="ssl-progress-number">
                  {job.completed_rounds.toLocaleString()}{" "}
                  /{" "}
                  {job.total_rounds.toLocaleString()}
                </strong>

                <small>
                  บันทึกลงฐานข้อมูลแล้ว
                </small>
              </div>

              <div>
                <span>
                  รอบปัจจุบัน
                </span>

                <strong className="ssl-progress-number">
                  {job.current_round > 0
                    ? `รอบที่ ${job.current_round}`
                    : "-"}
                </strong>

                <small>
                  สถานะ{" "}
                  {jobStatusLabel(
                    job.status
                  )}
                </small>
              </div>

              <div>
                <span>
                  อัปเดตล่าสุด
                </span>

                <strong className="ssl-progress-number ssl-progress-date">
                  {formatDateTime(
                    job.updated_at
                  )}
                </strong>

                <small>
                  Job ID #{job.id}
                </small>
              </div>
            </div>

            <div className="ssl-progress-track">
              <div
                className="ssl-progress-fill"
                style={{
                  width:
                    `${progressPercent}%`,
                }}
              />
            </div>
          </section>
        )}

        {errorMessage && (
          <div className="ssl-error-box">
            {errorMessage}
          </div>
        )}

        {job && (
          <section className="ssl-summary-grid">
            <article>
              <span>
                URL ที่ตรวจแล้ว
              </span>

              <strong>
                {job.checked_count.toLocaleString()}
              </strong>

              <small>
                Checked URLs
              </small>
            </article>

            <article>
              <span>
                ใช้งานได้
              </span>

              <strong>
                {Number(
                  job.valid_count ??
                    0
                ).toLocaleString()}
              </strong>

              <small className="valid">
                Valid
              </small>
            </article>

            <article>
              <span>
                ใกล้หมดอายุ
              </span>

              <strong>
                {Number(
                  job.expiring_count ??
                    0
                ).toLocaleString()}
              </strong>

              <small className="warning">
                ≤ 30 days
              </small>
            </article>

            <article>
              <span>
                หมดอายุ
              </span>

              <strong>
                {Number(
                  job.expired_count ??
                    0
                ).toLocaleString()}
              </strong>

              <small className="danger">
                Expired
              </small>
            </article>

            <article>
              <span>
                ตรวจสอบไม่ได้
              </span>

              <strong>
                {Number(
                  job.error_count ??
                    0
                ).toLocaleString()}
              </strong>

              <small className="danger">
                Error
              </small>
            </article>
          </section>
        )}

        <section className="ssl-results-card ssl-results-card-full">
            <div className="ssl-results-header">
              <div>
                <h2>
                  ผลการตรวจสอบ SSL / TLS{" "}
                  <span>
                    รอบที่{" "}
                    {selectedRound}
                  </span>
                </h2>

                <p>
                  เรียงตามลำดับในไฟล์ต้นฉบับจากบนลงล่าง
                </p>
              </div>

              <div className="ssl-result-header-actions">
                <span
                  style={{
                    color: "#617089",
                    fontSize: "12px",
                    fontWeight: 700,
                    whiteSpace: "nowrap",
                  }}
                >
                  {filteredResults.length.toLocaleString()} รายการ
                </span>

                <details
                  style={{
                    position: "relative",
                  }}
                >
                  <summary
                    style={{
                      minHeight: "40px",
                      display: "inline-flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "7px",
                      padding: "0 14px",
                      border: "1px solid #d6e0ec",
                      borderRadius: "9px",
                      background: "#ffffff",
                      color: "#18365d",
                      fontSize: "12px",
                      fontWeight: 750,
                      cursor: "pointer",
                      listStyle: "none",
                      userSelect: "none",
                    }}
                  >
                    ☷ เลือกคอลัมน์
                  </summary>

                  <div
                    style={{
                      width: "220px",
                      position: "absolute",
                      top: "calc(100% + 8px)",
                      right: 0,
                      zIndex: 50,
                      padding: "10px",
                      border: "1px solid #dce4ee",
                      borderRadius: "10px",
                      background: "#ffffff",
                      boxShadow:
                        "0 12px 28px rgba(21, 48, 84, 0.14)",
                    }}
                  >
                    <div
                      style={{
                        marginBottom: "7px",
                        color: "#7b8ba1",
                        fontSize: "10px",
                        fontWeight: 700,
                      }}
                    >
                      เลือกข้อมูลที่ต้องการแสดง
                    </div>

                    {RESULT_COLUMNS.map((column) => (
                      <label
                        key={column.key}
                        style={{
                          minHeight: "34px",
                          display: "flex",
                          alignItems: "center",
                          gap: "9px",
                          padding: "5px 6px",
                          borderRadius: "7px",
                          color: "#203a5f",
                          fontSize: "11px",
                          cursor: "pointer",
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={
                            selectedColumns.includes(
                              column.key
                            )
                          }
                          onChange={() =>
                            toggleResultColumn(
                              column.key
                            )
                          }
                        />
                        {column.label}
                      </label>
                    ))}

                    <button
                      type="button"
                      onClick={() =>
                        setSelectedColumns(
                          RESULT_COLUMNS.map(
                            (column) =>
                              column.key
                          )
                        )
                      }
                      style={{
                        width: "100%",
                        height: "32px",
                        marginTop: "7px",
                        border: "1px solid #cfe0f7",
                        borderRadius: "7px",
                        background: "#f3f7fd",
                        color: "#1768e6",
                        fontSize: "10px",
                        fontWeight: 750,
                        cursor: "pointer",
                      }}
                    >
                      แสดงทุกคอลัมน์
                    </button>
                  </div>
                </details>

                <button
                  type="button"
                  className="ssl-download-button"
                  disabled={
                    filteredResults.length ===
                    0
                  }
                  onClick={
                    downloadCurrentPage
                  }
                >
                  ↓ ดาวน์โหลด Excel
                </button>
              </div>
            </div>

            <div className="ssl-toolbar">
              <div className="ssl-search">
                <span>
                  ⌕
                </span>

                <input
                  type="text"
                  value={
                    search
                  }
                  placeholder="ค้นหา URL / Hostname / Issuer..."
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

              <div className="ssl-filter-tabs">
                {(
                  [
                    [
                      "ALL",
                      `ทั้งหมด (${roundPageSummary.total})`,
                    ],
                    [
                      "VALID",
                      `Valid (${roundPageSummary.valid})`,
                    ],
                    [
                      "EXPIRING_SOON",
                      `ใกล้หมดอายุ (${roundPageSummary.expiring})`,
                    ],
                    [
                      "EXPIRED",
                      `Expired (${roundPageSummary.expired})`,
                    ],
                    [
                      "ERROR",
                      `Error (${roundPageSummary.error})`,
                    ],
                  ] as const
                ).map(
                  ([
                    value,
                    label,
                  ]) => (
                    <button
                      key={
                        value
                      }
                      type="button"
                      className={
                        filter ===
                        value
                          ? "active"
                          : ""
                      }
                      onClick={() =>
                        setFilter(
                          value
                        )
                      }
                    >
                      {label}
                    </button>
                  )
                )}
              </div>
            </div>

            <div className="ssl-table-wrapper">
              <table className="ssl-table">
                <thead>
                  <tr>
                    {isColumnVisible("order") && (
                      <th>ลำดับในไฟล์</th>
                    )}

                    {isColumnVisible("url") && (
                      <th>URL / HOSTNAME</th>
                    )}

                    {isColumnVisible("validFrom") && (
                      <th>VALID FROM</th>
                    )}

                    {isColumnVisible("expirationDate") && (
                      <th>EXPIRATION DATE</th>
                    )}

                    {isColumnVisible("daysLeft") && (
                      <th>DAYS LEFT</th>
                    )}

                    {isColumnVisible("status") && (
                      <th>STATUS</th>
                    )}

                    {isColumnVisible("issuer") && (
                      <th>ISSUER</th>
                    )}

                    {isColumnVisible("checkedAt") && (
                      <th>CHECKED AT</th>
                    )}

                    {isColumnVisible("error") && (
                      <th>ERROR</th>
                    )}
                  </tr>
                </thead>

                <tbody>
                  {filteredResults.length ===
                  0 ? (
                    <tr>
                      <td
                        colSpan={
                          visibleColumnCount
                        }
                        className="ssl-empty-cell"
                      >
                        {selectedRoundData?.status ===
                        "RUNNING"
                          ? "กำลังตรวจสอบข้อมูลในรอบนี้..."
                          : "ยังไม่มีผลในรอบนี้"}
                      </td>
                    </tr>
                  ) : (
                    filteredResults.map(
                      (
                        item
                      ) => (
                        <tr
                          key={
                            item.id
                          }
                        >
                          {isColumnVisible("order") && (
                            <td className="ssl-order-cell">
                              {item.file_order.toLocaleString()}
                            </td>
                          )}

                          {isColumnVisible("url") && (
                            <td className="ssl-url-cell">
                              <strong>
                                {item.url}
                              </strong>

                              <span>
                                {item.hostname}
                              </span>
                            </td>
                          )}

                          {isColumnVisible("validFrom") && (
                            <td>
                              {formatDateTime(
                                item.valid_from
                              )}
                            </td>
                          )}

                          {isColumnVisible("expirationDate") && (
                            <td>
                              {formatDateTime(
                                item.expiration_date
                              )}
                            </td>
                          )}

                          {isColumnVisible("daysLeft") && (
                            <td className="ssl-days-cell">
                              {item.days_left ??
                                "-"}
                            </td>
                          )}

                          {isColumnVisible("status") && (
                            <td>
                              <span
                                className={`ssl-status-badge ${item.status.toLowerCase()}`}
                              >
                                {statusLabel(
                                  item.status
                                )}
                              </span>
                            </td>
                          )}

                          {isColumnVisible("issuer") && (
                            <td>
                              {item.issuer ||
                                "-"}
                            </td>
                          )}

                          {isColumnVisible("checkedAt") && (
                            <td>
                              {formatDateTime(
                                item.checked_at
                              )}
                            </td>
                          )}

                          {isColumnVisible("error") && (
                            <td className="ssl-error-cell">
                              {item.error ||
                                "-"}
                            </td>
                          )}
                        </tr>
                      )
                    )
                  )}
                </tbody>
              </table>
            </div>

            <div className="ssl-table-footer">
              <span>
                รอบนี้มีผลแล้ว{" "}
                {roundResultTotal.toLocaleString()}{" "}
                รายการ
              </span>

              {roundResultTotal >
                TABLE_PAGE_SIZE && (
                <div className="ssl-pagination">
                  <button
                    type="button"
                    disabled={
                      safePage <= 1
                    }
                    onClick={() =>
                      setTablePage(
                        Math.max(
                          1,
                          safePage -
                            1
                        )
                      )
                    }
                  >
                    ← ก่อนหน้า
                  </button>

                  <span>
                    หน้า{" "}
                    <strong>
                      {safePage}
                    </strong>{" "}
                    /{" "}
                    <strong>
                      {totalPages}
                    </strong>
                  </span>

                  <button
                    type="button"
                    disabled={
                      safePage >=
                      totalPages
                    }
                    onClick={() =>
                      setTablePage(
                        Math.min(
                          totalPages,
                          safePage +
                            1
                        )
                      )
                    }
                  >
                    ถัดไป →
                  </button>
                </div>
              )}
            </div>
          </section>
        {roundsOpen && (
          <div className="ssl-round-drawer-backdrop" onClick={() => setRoundsOpen(false)}>
            <aside className="ssl-round-drawer" onClick={(event) => event.stopPropagation()}>
              <div className="ssl-round-drawer-header">
                <div>
                  <h2>รอบการตรวจสอบ</h2>
                  <p>เลือกรอบเพื่อดูผล • รอบที่เสร็จแล้วดูได้ แต่ระบบจะไม่ตรวจซ้ำ</p>
                </div>
                <button type="button" className="ssl-round-drawer-close" onClick={() => setRoundsOpen(false)}>×</button>
              </div>

              <div className="ssl-round-drawer-summary">
                <span>เสร็จแล้ว <strong>{job?.completed_rounds.toLocaleString() ?? 0}</strong> รอบ</span>
                <span>ทั้งหมด <strong>{job?.total_rounds.toLocaleString() ?? 0}</strong> รอบ</span>
              </div>

              <div className="ssl-round-list ssl-round-list-drawer">
                {rounds.map((round) => {
                  const completed = round.status === "COMPLETED";
                  const active = round.status === "RUNNING";
                  return (
                    <button
                      key={round.id}
                      type="button"
                      className={`ssl-round-item ${selectedRound === round.round_number ? "selected" : ""} ${completed ? "completed" : active ? "active" : round.status === "PAUSED" ? "paused" : "waiting"}`}
                      onClick={() => chooseRound(round.round_number)}
                    >
                      <span className="ssl-round-icon">{completed ? "✓" : active ? "▶" : round.status === "PAUSED" ? "⏸" : "•"}</span>
                      <span className="ssl-round-copy">
                        <strong>รอบที่ {round.round_number} ({round.start_order.toLocaleString()}-{round.end_order.toLocaleString()})</strong>
                        <small>ตรวจแล้ว {round.checked_count.toLocaleString()} / {round.total_urls.toLocaleString()} URL</small>
                      </span>
                      <span className="ssl-round-status">{round.status}</span>
                    </button>
                  );
                })}
              </div>
            </aside>
          </div>
        )}

        <div className="ssl-note">
          ℹ เมื่อครบ 100 URL ระบบจะบันทึกผลลงฐานข้อมูล และเมื่อครบ 500 URL จะบันทึกสถานะรอบเป็น COMPLETED อัตโนมัติ
        </div>
      </div>
    </main>
  );
}

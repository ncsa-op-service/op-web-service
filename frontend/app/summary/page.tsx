"use client";

import {
  useCallback,
  useEffect,
  useMemo,
  useState,
} from "react";
import { useRouter } from "next/navigation";

import "./summary.css";

type ProviderKey = "ais" | "trueDtac" | "nt" | "cloudflare";
type ProviderView = "ALL" | ProviderKey;
type ProviderResultStatus = "success" | "nxdomain" | "failed" | null;
type FakeStatusFilter = "ALL" | "NXDOMAIN" | "FAILED" | "URL";
type FakeViewMode = "TABLE" | "SUMMARY";
type SSLStatus = "VALID" | "EXPIRING_SOON" | "EXPIRED" | "ERROR";
type JobStatus = "READY" | "RUNNING" | "PAUSED" | "COMPLETED" | "ERROR";
type RoundStatus = "WAITING" | "RUNNING" | "PAUSED" | "COMPLETED" | "ERROR";
type SSLStatusFilter = "ALL" | SSLStatus;

type UserInfo = {
  id: number;
  name: string;
  email: string;
  role: "super_admin" | "editor" | "viewer";
};

type SummaryRow = {
  id: number;
  batch_id: number;
  case_id: string;
  url_sms: string;
  ais_result: string | null;
  true_dtac_result: string | null;
  nt_result: string | null;
  cloudflare_result: string | null;
  ais_status?: ProviderResultStatus;
  true_dtac_status?: ProviderResultStatus;
  nt_status?: ProviderResultStatus;
  cloudflare_status?: ProviderResultStatus;
  created_at: string;
};

type GraphCount = {
  found: number;
  notFound: number;
  pending: number;
};

type SummaryData = {
  reportDate: string | null;
  updatedAt: string | null;
  urlFakeWeb: {
    batch: {
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
    } | null;
    results: SummaryRow[];
    graph: Record<ProviderKey, GraphCount>;
  };
};

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
  created_at: string;
  updated_at: string;
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

type ClabType =
  | "overview_va"
  | "report"
  | "tool_op02"
  | "tool_op03"
  | "waiting_organizations"
  | "email_service"
  | string;

type LatestSheet = {
  id: number;
  snapshot_id: number;
  original_sheet_name: string;
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

const PROVIDERS: Array<{ key: ProviderKey; label: string }> = [
  { key: "ais", label: "AIS" },
  { key: "trueDtac", label: "TRUE / DTAC" },
  { key: "nt", label: "NT" },
  { key: "cloudflare", label: "CloudFlare" },
];

const API_URL =
  process.env.NEXT_PUBLIC_API_URL &&
  !process.env.NEXT_PUBLIC_API_URL.includes("backend:")
    ? process.env.NEXT_PUBLIC_API_URL.replace(/\/$/, "")
    : "http://localhost:4000";

const URL_PAGE_SIZE = 10;
const SSL_PAGE_SIZE = 8;

function readCurrentUser(): UserInfo | null {
  if (typeof window === "undefined") return null;

  const keys = ["user", "currentUser", "authUser", "op_user"];
  for (const key of keys) {
    const raw = window.localStorage.getItem(key);
    if (!raw) continue;
    try {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed.id === "number" && typeof parsed.role === "string") {
        return parsed as UserInfo;
      }
    } catch {
      // ignore invalid local storage
    }
  }
  return null;
}

function getRoleLabel(role: UserInfo["role"]) {
  if (role === "super_admin") return "Super Admin";
  if (role === "editor") return "ผู้แก้ไขข้อมูล";
  if (role === "viewer") return "ผู้ดูข้อมูล";
  return "ผู้ใช้งาน";
}

function formatDate(value: string | null | undefined, includeTime = false) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString("th-TH", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    ...(includeTime ? { hour: "2-digit", minute: "2-digit" } : {}),
  });
}

function isFound(value: string | null | undefined) {
  const text = String(value ?? "").trim();
  return text.startsWith("http://") || text.startsWith("https://");
}

function getProviderResult(row: SummaryRow, provider: ProviderKey) {
  if (provider === "ais") return row.ais_result;
  if (provider === "trueDtac") return row.true_dtac_result;
  if (provider === "nt") return row.nt_result;
  return row.cloudflare_result;
}

function getProviderStatus(
  row: SummaryRow,
  provider: ProviderKey
): ProviderResultStatus {
  if (provider === "ais") return row.ais_status ?? null;
  if (provider === "trueDtac") return row.true_dtac_status ?? null;
  if (provider === "nt") return row.nt_status ?? null;
  return row.cloudflare_status ?? null;
}

function matchesRawResult(
  value: string | null | undefined,
  status: ProviderResultStatus,
  filter: FakeStatusFilter
) {
  const text = String(value ?? "").trim();
  const isNxDomain =
    status === "nxdomain" ||
    text.toUpperCase().includes("DNS_PROBE_FINISHED_NXDOMAIN");

  if (filter === "ALL") return true;
  if (filter === "NXDOMAIN") return isNxDomain;
  if (filter === "FAILED") return status === "failed";
  if (filter === "URL") {
    // ข้อมูลใหม่ใช้ status เป็นหลัก
    if (status != null) return status === "success";

    // fallback สำหรับข้อมูลเก่าที่ยังไม่มี *_status
    return isFound(text) && !isNxDomain;
  }

  return true;
}

function ResultCell({
  value,
  originalUrl,
  status,
}: {
  value: string | null | undefined;
  originalUrl: string;
  status?: ProviderResultStatus;
}) {
  const text = String(value ?? "").trim();

  if (!text) {
    return <span className="summary-result empty">-</span>;
  }

  const isNxDomain =
    status === "nxdomain" ||
    text.toUpperCase().includes("DNS_PROBE_FINISHED_NXDOMAIN");

  if (isNxDomain) {
    return (
      <a
        href={originalUrl}
        target="_blank"
        rel="noreferrer"
        className="summary-result not-found summary-result-link summary-result-raw"
        title={`เปิด URL ต้นฉบับ: ${originalUrl}`}
      >
        DNS_PROBE_FINISHED_NXDOMAIN
      </a>
    );
  }

  if (status === "failed") {
    const failedUrl = isFound(text) ? text : originalUrl;
    return (
      <a
        href={failedUrl}
        target="_blank"
        rel="noreferrer"
        className="summary-result not-found summary-result-link summary-result-raw"
        title={`URL ที่เปิดไม่ได้: ${failedUrl}`}
      >
        {failedUrl}
      </a>
    );
  }

  if (status === "success" || (status == null && isFound(text))) {
    const successUrl = isFound(text) ? text : originalUrl;
    return (
      <a
        href={successUrl}
        target="_blank"
        rel="noreferrer"
        className="summary-result found summary-result-link summary-result-raw"
        title={successUrl}
      >
        {successUrl}
      </a>
    );
  }

  return (
    <a
      href={originalUrl}
      target="_blank"
      rel="noreferrer"
      className="summary-result not-found summary-result-link summary-result-raw"
      title={`เปิด URL ต้นฉบับ: ${originalUrl}`}
    >
      {text}
    </a>
  );
}

function SSLStatusBadge({ status }: { status: SSLStatus }) {
  const labels: Record<SSLStatus, string> = {
    VALID: "ใช้งานได้",
    EXPIRING_SOON: "ใกล้หมดอายุ",
    EXPIRED: "หมดอายุ",
    ERROR: "ตรวจสอบไม่ได้",
  };
  return <span className={`summary-ssl-status ${status.toLowerCase()}`}>{labels[status]}</span>;
}

async function readJson<T>(response: Response): Promise<T> {
  const text = await response.text();
  const data = text ? JSON.parse(text) : {};
  if (!response.ok) {
    throw new Error(data?.message ?? data?.error ?? `HTTP ${response.status}`);
  }
  return data as T;
}

export default function SummaryPage() {
  const router = useRouter();

  const [currentUser, setCurrentUser] = useState<UserInfo | null>(null);
  const [authReady, setAuthReady] = useState(false);
  const [data, setData] = useState<SummaryData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  // URL Fake Web filters
  const [provider, setProvider] = useState<ProviderView>("ALL");
  const [fakeStatus, setFakeStatus] = useState<FakeStatusFilter>("ALL");
  const [search, setSearch] = useState("");
  const [urlPage, setUrlPage] = useState(1);
  const [fakeViewMode, setFakeViewMode] = useState<FakeViewMode>("TABLE");

  // SSL summary
  const [sslJobs, setSslJobs] = useState<SSLJob[]>([]);
  const [sslJobId, setSslJobId] = useState<number | null>(null);
  const [sslRounds, setSslRounds] = useState<SSLRound[]>([]);
  const [sslRoundNumber, setSslRoundNumber] = useState<number | null>(null);
  const [sslResults, setSslResults] = useState<SSLResultRow[]>([]);
  const [sslResultTotal, setSslResultTotal] = useState(0);
  const [sslPage, setSslPage] = useState(1);
  const [sslLoading, setSslLoading] = useState(false);
  const [sslError, setSslError] = useState("");
  const [sslStatusFilter, setSslStatusFilter] = useState<SSLStatusFilter>("ALL");
  const [sslSummaryOpen, setSslSummaryOpen] = useState(false);

  // CLAB saved data
  const [clabSheets, setClabSheets] = useState<LatestSheet[]>([]);
  const [clabLoading, setClabLoading] = useState(false);
  const [clabError, setClabError] = useState("");

  const canEdit = currentUser?.role === "super_admin" || currentUser?.role === "editor";
  const isViewer = currentUser?.role === "viewer";

  const loadSummary = useCallback(async () => {
    const response = await fetch(`${API_URL}/api/summary`, { cache: "no-store" });
    const result = await readJson<SummaryData>(response);
    setData(result);
  }, []);

  const loadSSLJobs = useCallback(async () => {
    try {
      setSslLoading(true);
      setSslError("");
      const response = await fetch(`${API_URL}/api/ssl-checker/jobs?limit=30`, {
        cache: "no-store",
      });
      const result = await readJson<{ jobs: SSLJob[] }>(response);
      const jobs = Array.isArray(result.jobs) ? result.jobs : [];
      setSslJobs(jobs);
      setSslJobId((current) => {
        if (current && jobs.some((job) => job.id === current)) return current;
        return jobs[0]?.id ?? null;
      });
    } catch (loadError) {
      setSslError(loadError instanceof Error ? loadError.message : "โหลดรายการ SSL/TLS ไม่สำเร็จ");
    } finally {
      setSslLoading(false);
    }
  }, []);

  const loadClab = useCallback(async () => {
    try {
      setClabLoading(true);
      setClabError("");
      const response = await fetch(`${API_URL}/api/clab/latest-by-type`, {
        cache: "no-store",
      });
      const result = await readJson<{ sheets: LatestSheet[] }>(response);
      setClabSheets(Array.isArray(result.sheets) ? result.sheets : []);
    } catch (loadError) {
      setClabError(loadError instanceof Error ? loadError.message : "โหลดข้อมูล CLAB ไม่สำเร็จ");
    } finally {
      setClabLoading(false);
    }
  }, []);

  const loadAll = useCallback(async () => {
    try {
      setLoading(true);
      setError("");
      await Promise.all([loadSummary(), loadSSLJobs(), loadClab()]);
    } catch (loadError) {
      console.error("Summary dashboard error:", loadError);
      setError(loadError instanceof Error ? loadError.message : "ไม่สามารถโหลด Main Summary ได้");
    } finally {
      setLoading(false);
    }
  }, [loadSummary, loadSSLJobs, loadClab]);

  useEffect(() => {
    const user = readCurrentUser();
    if (!user) {
      router.replace("/login");
      return;
    }

    if (!["super_admin", "editor", "viewer"].includes(user.role)) {
      window.localStorage.removeItem("user");
      router.replace("/login");
      return;
    }

    setCurrentUser(user);
    setAuthReady(true);
    void loadAll();
  }, [loadAll, router]);


  // อัปเดตข้อมูล URL Fake Web อัตโนมัติทุก 10 วินาที
  // โดยไม่ reload ทั้งหน้า
  useEffect(() => {
    if (!authReady) return;

    const intervalId = window.setInterval(() => {
      void loadSummary().catch((refreshError) => {
        console.error("Auto refresh summary error:", refreshError);
      });
    }, 10000);

    return () => {
      window.clearInterval(intervalId);
    };
  }, [authReady, loadSummary]);

  useEffect(() => {
    if (!sslJobId) {
      setSslRounds([]);
      setSslRoundNumber(null);
      return;
    }

    let cancelled = false;
    (async () => {
      try {
        setSslLoading(true);
        setSslError("");
        const response = await fetch(`${API_URL}/api/ssl-checker/jobs/${sslJobId}/rounds`, {
          cache: "no-store",
        });
        const result = await readJson<{ rounds: SSLRound[] }>(response);
        if (cancelled) return;

        const rounds = Array.isArray(result.rounds) ? result.rounds : [];
        setSslRounds(rounds);
        setSslRoundNumber((current) => {
          if (current && rounds.some((round) => round.round_number === current)) return current;
          const completed = [...rounds]
            .filter((round) => round.checked_count > 0 || round.status === "COMPLETED")
            .sort((a, b) => b.round_number - a.round_number)[0];
          return completed?.round_number ?? rounds[0]?.round_number ?? null;
        });
        setSslPage(1);
      } catch (loadError) {
        if (!cancelled) {
          setSslError(loadError instanceof Error ? loadError.message : "โหลดรอบ SSL/TLS ไม่สำเร็จ");
        }
      } finally {
        if (!cancelled) setSslLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [sslJobId]);

  useEffect(() => {
    if (!sslJobId || !sslRoundNumber) {
      setSslResults([]);
      setSslResultTotal(0);
      return;
    }

    let cancelled = false;
    (async () => {
      try {
        setSslLoading(true);
        setSslError("");

        const pageSize = 500;
        let page = 1;
        let total = 0;
        const allResults: SSLResultRow[] = [];

        while (true) {
          const response = await fetch(
            `${API_URL}/api/ssl-checker/jobs/${sslJobId}/rounds/${sslRoundNumber}/results?page=${page}&pageSize=${pageSize}`,
            { cache: "no-store" }
          );
          const result = await readJson<{ total: number; results: SSLResultRow[] }>(response);
          if (cancelled) return;

          const pageResults = Array.isArray(result.results) ? result.results : [];
          total = Number(result.total ?? 0);
          allResults.push(...pageResults);

          if (pageResults.length === 0 || allResults.length >= total) break;
          page += 1;
        }

        setSslResults(allResults);
        setSslResultTotal(total);
      } catch (loadError) {
        if (!cancelled) {
          setSslResults([]);
          setSslResultTotal(0);
          setSslError(loadError instanceof Error ? loadError.message : "โหลดผลตรวจ SSL/TLS ไม่สำเร็จ");
        }
      } finally {
        if (!cancelled) setSslLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [sslJobId, sslRoundNumber]);

  const rows = data?.urlFakeWeb?.results ?? [];
  const graph = data?.urlFakeWeb?.graph;
  const batch = data?.urlFakeWeb?.batch;

  const filteredRows = useMemo(() => {
    const keyword = search.trim().toLowerCase();

    return rows.filter((row) => {
      const matchesSearch =
        !keyword ||
        row.case_id.toLowerCase().includes(keyword) ||
        row.url_sms.toLowerCase().includes(keyword);

      const providerResults =
        provider === "ALL"
          ? PROVIDERS.map((item) => ({
              value: getProviderResult(row, item.key),
              status: getProviderStatus(row, item.key),
            }))
          : [
              {
                value: getProviderResult(row, provider),
                status: getProviderStatus(row, provider),
              },
            ];

      const matchesStatus = providerResults.some(({ value, status }) =>
        matchesRawResult(value, status, fakeStatus)
      );

      return matchesSearch && matchesStatus;
    });
  }, [rows, provider, fakeStatus, search]);

  const fakeSummary = useMemo(() => {
    let found = 0;
    let failed = 0;
    let nxdomain = 0;
    let pending = 0;

    // นับเป็น “ผลตรวจต่อเครือข่าย”
    // ALL = ทุก URL x ทุกเครือข่าย
    // เลือกเครือข่าย = นับเฉพาะเครือข่ายนั้น
    rows.forEach((row) => {
      const providerResults =
        provider === "ALL"
          ? PROVIDERS.map((item) => ({
              value: getProviderResult(row, item.key),
              status: getProviderStatus(row, item.key),
            }))
          : [
              {
                value: getProviderResult(row, provider),
                status: getProviderStatus(row, provider),
              },
            ];

      providerResults.forEach(({ value, status }) => {
        const text = String(value ?? "").trim();
        const isNxDomain =
          status === "nxdomain" ||
          text.toUpperCase().includes("DNS_PROBE_FINISHED_NXDOMAIN");

        if (isNxDomain) {
          nxdomain += 1;
          return;
        }

        if (status === "failed") {
          failed += 1;
          return;
        }

        if (status === "success") {
          found += 1;
          return;
        }

        // รองรับข้อมูลเก่าที่ยังไม่มี *_status
        if (!text) {
          pending += 1;
        } else if (isFound(text)) {
          found += 1;
        } else {
          failed += 1;
        }
      });
    });

    const total = found + failed + nxdomain + pending;
    return { total, found, failed, nxdomain, pending };
  }, [rows, provider]);

  const fakeSummaryMax = Math.max(
    1,
    fakeSummary.found,
    fakeSummary.failed,
    fakeSummary.nxdomain,
    fakeSummary.pending
  );

  const urlTotalPages = Math.max(1, Math.ceil(filteredRows.length / URL_PAGE_SIZE));
  const visibleUrlRows = filteredRows.slice((urlPage - 1) * URL_PAGE_SIZE, urlPage * URL_PAGE_SIZE);

  const filteredSSLResults = useMemo(() => {
    if (sslStatusFilter === "ALL") return sslResults;
    return sslResults.filter((row) => row.status === sslStatusFilter);
  }, [sslResults, sslStatusFilter]);

  const sslRoundSummary = useMemo(() => {
    return sslResults.reduce(
      (acc, row) => {
        acc[row.status] += 1;
        return acc;
      },
      { VALID: 0, EXPIRING_SOON: 0, EXPIRED: 0, ERROR: 0 } as Record<SSLStatus, number>
    );
  }, [sslResults]);

  const sslTotalPages = Math.max(1, Math.ceil(filteredSSLResults.length / SSL_PAGE_SIZE));
  const visibleSSLResults = filteredSSLResults.slice((sslPage - 1) * SSL_PAGE_SIZE, sslPage * SSL_PAGE_SIZE);

  useEffect(() => {
    setUrlPage(1);
  }, [provider, fakeStatus, search]);

  useEffect(() => {
    setSslPage(1);
  }, [sslStatusFilter, sslJobId, sslRoundNumber]);

  const selectedSSLJob = sslJobs.find((job) => job.id === sslJobId) ?? null;
  const selectedSSLRound = sslRounds.find((round) => round.round_number === sslRoundNumber) ?? null;

  if (!authReady) return null;

  if (loading && !data) {
    return (
      <main className="summary-page">
        <div className="summary-state">กำลังโหลดข้อมูล...</div>
      </main>
    );
  }

  if (error || !data) {
    return (
      <main className="summary-page">
        <div className="summary-state">
          <p>{error || "ไม่พบข้อมูล Summary"}</p>
          <button type="button" onClick={() => void loadAll()}>
            ลองใหม่
          </button>
        </div>
      </main>
    );
  }

  return (
    <main className="summary-page">
      <div className="summary-container">
        <section className="summary-heading">
          <div>
            <p className="summary-breadcrumb">Main Summary</p>
            <h1>Operations Overview</h1>
            <p className="summary-description">
              รวมผล URL Fake Web, SSL / TLS Checker และข้อมูล CLAB ที่บันทึกไว้ในหน้าเดียว
            </p>
          </div>

          <div className="summary-heading-actions">
            {currentUser && (
              <div className="summary-current-user">
                <div className="summary-user-avatar">
                  {(currentUser.name || currentUser.email).charAt(0).toUpperCase()}
                </div>
                <div className="summary-user-details">
                  <strong>{currentUser.name || currentUser.email}</strong>
                  <span>{getRoleLabel(currentUser.role)}</span>
                </div>
              </div>
            )}

            <div className="summary-batch-info">
              <span>อัปเดตล่าสุด</span>
              <strong>{formatDate(data.updatedAt, true)}</strong>
            </div>

            {isViewer && <span className="summary-readonly-badge">ดูอย่างเดียว</span>}

            {canEdit && (
              <button
                type="button"
                className="summary-edit-button"
                onClick={() => router.push("/summary/edit")}
              >
                ✎ แก้ไขข้อมูล
              </button>
            )}

            <button type="button" className="summary-refresh" onClick={() => void loadAll()}>
              ↻ Refresh
            </button>
          </div>
        </section>

        {/* ================= URL FAKE WEB ================= */}
        <section className="summary-dashboard-section">
          <div className="summary-section-heading">
            <div>
              <span className="summary-section-kicker">URL FAKE WEB</span>
              <h2>ผลตรวจ URL Fake Web</h2>
              <p>เลือกดูทุกเครือข่ายหรือเฉพาะเครือข่ายที่ต้องการ และกรองตามค่าผลตรวจจริง</p>
            </div>
            <div className="summary-file-chip" title={batch?.original_file_name ?? ""}>
              <span>ไฟล์ล่าสุด</span>
              <strong>{batch?.original_file_name ?? "-"}</strong>
            </div>
          </div>

          <div className="summary-fake-view-switch">
            <button
              type="button"
              className={fakeViewMode === "TABLE" ? "active" : ""}
              onClick={() => setFakeViewMode("TABLE")}
            >
              ตาราง
            </button>
            <button
              type="button"
              className={fakeViewMode === "SUMMARY" ? "active" : ""}
              onClick={() => setFakeViewMode("SUMMARY")}
            >
              สรุป
            </button>
          </div>

          <div className="summary-filter-row">
            <div className="summary-filter-group">
              <span>เครือข่าย</span>
              <div className="summary-segmented summary-provider-segmented">
                <button
                  type="button"
                  className={provider === "ALL" ? "active" : ""}
                  onClick={() => setProvider("ALL")}
                >
                  แสดงทั้งหมด
                </button>
                {PROVIDERS.map((item) => (
                  <button
                    key={item.key}
                    type="button"
                    className={provider === item.key ? "active" : ""}
                    onClick={() => setProvider(item.key)}
                  >
                    {item.label}
                  </button>
                ))}
              </div>
            </div>

            {fakeViewMode === "TABLE" && (
              <>
                <div className="summary-filter-group">
                  <span>ผลตรวจ</span>
                  <div className="summary-segmented compact">
                    {[
                      ["ALL", "ทั้งหมด"],
                      ["NXDOMAIN", "DNS_PROBE_FINISHED_NXDOMAIN"],
                      ["FAILED", "URL ที่เปิดไม่ได้"],
                      ["URL", "URL ที่ตรวจพบ"],
                    ].map(([key, label]) => (
                      <button
                        key={key}
                        type="button"
                        className={fakeStatus === key ? "active" : ""}
                        onClick={() => setFakeStatus(key as FakeStatusFilter)}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="summary-filter-search">
                  <span>ค้นหา</span>
                  <input
                    value={search}
                    onChange={(event) => setSearch(event.target.value)}
                    placeholder="CASE ID / URL"
                  />
                </div>
              </>
            )}
          </div>

          {fakeViewMode === "SUMMARY" ? (
            <div className="summary-fake-summary-card">
              <div className="summary-fake-summary-head">
                <div>
                  <strong>สรุปผลตรวจ URL Fake Web</strong>
                  <span>
                    {provider === "ALL"
                      ? "สรุปผลตรวจแยกทุกเครือข่าย (1 URL อาจมีได้หลายผลตรวจ)"
                      : `สรุปเฉพาะ ${PROVIDERS.find((item) => item.key === provider)?.label ?? ""}`}
                  </span>
                </div>
                <span className="summary-fake-summary-provider">
                  {provider === "ALL" ? "ทุกเครือข่าย" : PROVIDERS.find((item) => item.key === provider)?.label}
                </span>
              </div>

              <div className="summary-fake-stat-grid">
                <article>
                  <span>ผลตรวจทั้งหมด</span>
                  <strong>{fakeSummary.total.toLocaleString("th-TH")}</strong>
                  <small>ผลตรวจ</small>
                </article>
                <article className="found">
                  <span>ตรวจพบ URL</span>
                  <strong>{fakeSummary.found.toLocaleString("th-TH")}</strong>
                  <small>ผลตรวจ</small>
                </article>
                <article className="not-found">
                  <span>URL ที่เปิดไม่ได้</span>
                  <strong>{fakeSummary.failed.toLocaleString("th-TH")}</strong>
                  <small>ผลตรวจ</small>
                </article>
                <article className="not-found">
                  <span>DNS_PROBE_FINISHED_NXDOMAIN</span>
                  <strong>{fakeSummary.nxdomain.toLocaleString("th-TH")}</strong>
                  <small>ผลตรวจ</small>
                </article>
                <article className="pending">
                  <span>รอตรวจ / ไม่มีผล</span>
                  <strong>{fakeSummary.pending.toLocaleString("th-TH")}</strong>
                  <small>ผลตรวจ</small>
                </article>
              </div>

              <div className="summary-fake-chart">
                {[
                  ["ตรวจพบ URL", fakeSummary.found, "found"],
                  ["URL ที่เปิดไม่ได้", fakeSummary.failed, "not-found"],
                  ["DNS_PROBE_FINISHED_NXDOMAIN", fakeSummary.nxdomain, "not-found"],
                  ["รอตรวจ / ไม่มีผล", fakeSummary.pending, "pending"],
                ].map(([label, value, kind]) => {
                  const count = Number(value);
                  const width = `${Math.max(count > 0 ? 3 : 0, (count / fakeSummaryMax) * 100)}%`;
                  return (
                    <div className="summary-fake-chart-row" key={String(label)}>
                      <div className="summary-fake-chart-label">
                        <span>{String(label)}</span>
                        <strong>{count.toLocaleString("th-TH")}</strong>
                      </div>
                      <div className="summary-fake-chart-track">
                        <div className={`summary-fake-chart-bar ${String(kind)}`} style={{ width }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ) : (
            <div className="summary-compact-table-card">
              <div className="summary-compact-table-head">
                <div>
                  <strong>{provider === "ALL" ? "แสดงตารางทั้งหมด" : PROVIDERS.find((item) => item.key === provider)?.label}</strong>
                  <span>{filteredRows.length.toLocaleString("th-TH")} รายการตามตัวกรอง</span>
                </div>
                <span className="summary-page-counter">หน้า {urlPage} / {urlTotalPages}</span>
              </div>

              <div className="summary-table-wrapper">
                <table className={`summary-table summary-table-compact ${provider === "ALL" ? "summary-fake-all-table" : "summary-fake-single-table"}`}>
                  <thead>
                    <tr>
                      <th>CASE ID</th>
                      <th>URL SMS</th>
                      {provider === "ALL" ? (
                        <>
                          <th>AIS</th>
                          <th>TRUE / DTAC</th>
                          <th>NT</th>
                          <th>CLOUDFLARE</th>
                        </>
                      ) : (
                        <th>{PROVIDERS.find((item) => item.key === provider)?.label}</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {visibleUrlRows.length === 0 ? (
                      <tr>
                        <td colSpan={provider === "ALL" ? 6 : 3} className="summary-table-empty">ไม่พบข้อมูลตามตัวกรอง</td>
                      </tr>
                    ) : (
                      visibleUrlRows.map((row) => (
                        <tr key={row.id}>
                          <td><strong>{row.case_id}</strong></td>
                          <td>
                            <a href={row.url_sms} target="_blank" rel="noreferrer" className="summary-url summary-fixed-url" title={row.url_sms}>
                              {row.url_sms}
                            </a>
                          </td>
                          {provider === "ALL" ? (
                            <>
                              <td><ResultCell value={row.ais_result} originalUrl={row.url_sms} status={row.ais_status} /></td>
                              <td><ResultCell value={row.true_dtac_result} originalUrl={row.url_sms} status={row.true_dtac_status} /></td>
                              <td><ResultCell value={row.nt_result} originalUrl={row.url_sms} status={row.nt_status} /></td>
                              <td><ResultCell value={row.cloudflare_result} originalUrl={row.url_sms} status={row.cloudflare_status} /></td>
                            </>
                          ) : (
                            <td><ResultCell value={getProviderResult(row, provider)} originalUrl={row.url_sms} status={getProviderStatus(row, provider)} /></td>
                          )}
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              <div className="summary-pagination">
                <button type="button" disabled={urlPage <= 1} onClick={() => setUrlPage((page) => page - 1)}>← ก่อนหน้า</button>
                <button type="button" disabled={urlPage >= urlTotalPages} onClick={() => setUrlPage((page) => page + 1)}>ถัดไป →</button>
              </div>
            </div>
          )}
        </section>

        {/* ================= SSL / TLS ================= */}
        <section className="summary-dashboard-section">
          <div className="summary-section-heading">
            <div>
              <span className="summary-section-kicker">SSL / TLS CHECKER</span>
              <h2>ผลตรวจ SSL / TLS Checker</h2>
              <p>เลือกไฟล์งานและรอบที่ต้องการดู จะแสดงเฉพาะผลของรอบนั้น</p>
            </div>
            <button type="button" className="summary-open-module" onClick={() => router.push("/ssl-checker")}>เปิดหน้า SSL / TLS →</button>
          </div>

          <div className="summary-ssl-controls">
            <label>
              <span>ไฟล์ / Job</span>
              <select
                value={sslJobId ?? ""}
                onChange={(event) => {
                  setSslJobId(event.target.value ? Number(event.target.value) : null);
                  setSslRoundNumber(null);
                  setSslPage(1);
                }}
              >
                {sslJobs.length === 0 && <option value="">ยังไม่มีงาน</option>}
                {sslJobs.map((job) => (
                  <option key={job.id} value={job.id}>
                    #{job.id} • {job.original_file_name}
                  </option>
                ))}
              </select>
            </label>

            <label>
              <span>รอบที่ตรวจ</span>
              <select
                value={sslRoundNumber ?? ""}
                disabled={!sslJobId || sslRounds.length === 0}
                onChange={(event) => {
                  setSslRoundNumber(event.target.value ? Number(event.target.value) : null);
                  setSslPage(1);
                }}
              >
                {sslRounds.length === 0 && <option value="">ไม่มีรอบ</option>}
                {sslRounds.map((round) => (
                  <option key={round.id} value={round.round_number}>
                    รอบ {round.round_number} • {round.checked_count}/{round.total_urls} • {round.status}
                  </option>
                ))}
              </select>
            </label>

            <div className="summary-ssl-meta">
              <span>สถานะงาน</span>
              <strong>{selectedSSLJob?.status ?? "-"}</strong>
            </div>

            <div className="summary-ssl-meta">
              <span>ผลในรอบนี้</span>
              <strong>{sslResultTotal.toLocaleString("th-TH")}</strong>
            </div>
          </div>

          <div className="summary-ssl-action-row">
            <div className="summary-ssl-status-filters">
              {[
                ["ALL", "ทั้งหมด"],
                ["VALID", "ใช้งานได้"],
                ["EXPIRING_SOON", "ใกล้หมดอายุ"],
                ["EXPIRED", "หมดอายุ"],
                ["ERROR", "ตรวจไม่ได้"],
              ].map(([key, label]) => (
                <button
                  key={key}
                  type="button"
                  className={sslStatusFilter === key ? "active" : ""}
                  onClick={() => setSslStatusFilter(key as SSLStatusFilter)}
                >
                  {label}
                </button>
              ))}
            </div>

            <button
              type="button"
              className={`summary-ssl-summary-button ${sslSummaryOpen ? "active" : ""}`}
              onClick={() => setSslSummaryOpen((open) => !open)}
            >
              สรุปรวม
            </button>
          </div>

          {sslSummaryOpen && (
            <div className="summary-ssl-summary-panel">
              <div className="summary-ssl-summary-title">
                <div>
                  <strong>สรุปรวม SSL / TLS</strong>
                  <span>ตัวเลขรวมของไฟล์งานที่เลือก และผลของรอบที่กำลังดู</span>
                </div>
                <span className="summary-ssl-summary-round">
                  {selectedSSLRound ? `รอบ ${selectedSSLRound.round_number}` : "ยังไม่ได้เลือกรอบ"}
                </span>
              </div>

              <div className="summary-ssl-summary-grid">
                <article>
                  <span>URL ทั้งหมด</span>
                  <strong>{(selectedSSLJob?.total_urls ?? 0).toLocaleString("th-TH")}</strong>
                  <small>รายการ</small>
                </article>
                <article>
                  <span>ตรวจไปแล้ว</span>
                  <strong>{(selectedSSLJob?.checked_count ?? 0).toLocaleString("th-TH")}</strong>
                  <small>รายการ</small>
                </article>
                <article className="valid">
                  <span>ใช้งานได้</span>
                  <strong>{(selectedSSLJob?.valid_count ?? sslRoundSummary.VALID).toLocaleString("th-TH")}</strong>
                  <small>รายการ</small>
                </article>
                <article className="warning">
                  <span>ใกล้หมดอายุ</span>
                  <strong>{(selectedSSLJob?.expiring_count ?? sslRoundSummary.EXPIRING_SOON).toLocaleString("th-TH")}</strong>
                  <small>รายการ</small>
                </article>
                <article className="expired">
                  <span>หมดอายุ</span>
                  <strong>{(selectedSSLJob?.expired_count ?? sslRoundSummary.EXPIRED).toLocaleString("th-TH")}</strong>
                  <small>รายการ</small>
                </article>
                <article className="error">
                  <span>ตรวจไม่ได้</span>
                  <strong>{(selectedSSLJob?.error_count ?? sslRoundSummary.ERROR).toLocaleString("th-TH")}</strong>
                  <small>รายการ</small>
                </article>
              </div>
            </div>
          )}

          {sslError && <div className="summary-inline-error">{sslError}</div>}

          <div className="summary-compact-table-card">
            <div className="summary-compact-table-head">
              <div>
                <strong>{selectedSSLJob?.original_file_name ?? "ยังไม่มีไฟล์ SSL/TLS"}</strong>
                <span>
                  {selectedSSLRound ? `รอบ ${selectedSSLRound.round_number} • ${selectedSSLRound.status}` : "เลือกรอบที่ต้องการดู"}
                </span>
              </div>
              <span className="summary-page-counter">{filteredSSLResults.length.toLocaleString("th-TH")} รายการ • หน้า {sslPage} / {sslTotalPages}</span>
            </div>

            <div className="summary-table-wrapper">
              <table className="summary-table summary-table-compact summary-ssl-table">
                <thead>
                  <tr>
                    <th>ลำดับ</th>
                    <th>URL</th>
                    <th>วันเริ่ม</th>
                    <th>วันหมด</th>
                    <th>สถานะ</th>
                  </tr>
                </thead>
                <tbody>
                  {sslLoading && sslResults.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="summary-table-empty">
                        กำลังโหลดผล SSL/TLS...
                      </td>
                    </tr>
                  ) : filteredSSLResults.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="summary-table-empty">
                        ไม่พบข้อมูลตามสถานะที่เลือก
                      </td>
                    </tr>
                  ) : (
                    visibleSSLResults.map((row) => (
                      <tr key={row.id}>
                        <td><strong>{row.file_order}</strong></td>
                        <td>
                          <span className="summary-ssl-hostname" title={row.url}>
                            {row.hostname || row.url}
                          </span>
                        </td>
                        <td>{formatDate(row.valid_from)}</td>
                        <td>{formatDate(row.expiration_date)}</td>
                        <td>
                          <a
                            href={row.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="summary-ssl-status-link"
                            title={`เปิด ${row.url}`}
                          >
                            <SSLStatusBadge status={row.status} />
                          </a>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            <div className="summary-pagination">
              <button type="button" disabled={sslPage <= 1} onClick={() => setSslPage((page) => page - 1)}>
                ← ก่อนหน้า
              </button>
              <button type="button" disabled={sslPage >= sslTotalPages} onClick={() => setSslPage((page) => page + 1)}>
                ถัดไป →
              </button>
            </div>
          </div>
        </section>

        {/* ================= CLAB ================= */}
        <section className="summary-dashboard-section">
          <div className="summary-section-heading">
            <div>
              <span className="summary-section-kicker">CLAB</span>
              <h2>ข้อมูล CLAB ที่บันทึกไว้</h2>
              <p>แสดงข้อมูลล่าสุดที่บันทึกแยกตามประเภท โดยการ์ดมีความสูงเท่ากันทุกใบ</p>
            </div>
            <button type="button" className="summary-open-module" onClick={() => router.push("/clab")}>เปิดหน้า CLAB →</button>
          </div>

          {clabError && <div className="summary-inline-error">{clabError}</div>}

          {clabLoading && clabSheets.length === 0 ? (
            <div className="summary-section-empty">กำลังโหลดข้อมูล CLAB...</div>
          ) : clabSheets.length === 0 ? (
            <div className="summary-section-empty">ยังไม่มีข้อมูล CLAB ที่บันทึกไว้</div>
          ) : (
            <div className="summary-clab-grid">
              {clabSheets.map((sheet) => (
                <button
                  type="button"
                  className="summary-clab-card summary-clab-card-button"
                  key={sheet.id}
                  onClick={() => router.push("/clab")}
                  title="เปิดดูข้อมูลนี้ในหน้า CLAB"
                >
                  <div className="summary-clab-card-top">
                    <span className="summary-clab-type">{sheet.detected_label}</span>
                    <span className="summary-clab-confidence">{Math.round(Number(sheet.confidence || 0) * (Number(sheet.confidence || 0) <= 1 ? 100 : 1))}%</span>
                  </div>

                  <div className="summary-clab-file" title={sheet.original_file_name ?? ""}>
                    {sheet.original_file_name ?? "ไม่ระบุชื่อไฟล์"}
                  </div>

                  <div className="summary-clab-sheet" title={sheet.original_sheet_name}>
                    Sheet: <strong>{sheet.original_sheet_name}</strong>
                  </div>

                  <div className="summary-clab-card-bottom">
                    <span>{Array.isArray(sheet.rows) ? sheet.rows.length.toLocaleString("th-TH") : "0"} แถว</span>
                    <span>{formatDate(sheet.updated_at ?? sheet.created_at, true)}</span>
                  </div>
                </button>
              ))}
            </div>
          )}
        </section>
      </div>
    </main>
  );
}

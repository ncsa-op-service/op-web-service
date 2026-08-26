"use client";

import {
  ChangeEvent,
  useMemo,
  useState,
} from "react";

import * as XLSX from "xlsx";

import "./checker.css";

type ProviderKey =
  | "ais"
  | "trueDtac"
  | "nt"
  | "cloudflare";

type ViewMode =
  | "original"
  | "checked";

type ProviderData = {
  ais: string;
  trueDtac: string;
  nt: string;
  cloudflare: string;
};

type UrlItem = {
  id: number;
  caseId: string;
  urlSms: string;

  original: ProviderData;

  checked: ProviderData;
};

type ApiCheckResult = {
  url: string;
  reachable: boolean;
  statusCode:
    | number
    | null;
  finalUrl:
    | string
    | null;
  error:
    | string
    | null;
};

type ApiCheckResponse = {
  total: number;
  results:
    ApiCheckResult[];
};

type NetworkInfo = {
  ip:
    | string
    | null;

  asn:
    | string
    | null;

  org:
    | string
    | null;

  country:
    | string
    | null;

  provider:
    | ProviderKey
    | "unknown";
};

export default function CheckerPage() {
  const [
    fileName,
    setFileName,
  ] = useState("");

  const [
    items,
    setItems,
  ] =
    useState<
      UrlItem[]
    >([]);

  const [
    search,
    setSearch,
  ] = useState("");

  const [
    provider,
    setProvider,
  ] =
    useState<
      ProviderKey
    >("ais");

  const [
    viewMode,
    setViewMode,
  ] =
    useState<
      ViewMode
    >("original");

  const [
    isProcessing,
    setIsProcessing,
  ] =
    useState(false);

  const [
    isDetectingNetwork,
    setIsDetectingNetwork,
  ] =
    useState(false);

  const [
    networkInfo,
    setNetworkInfo,
  ] =
    useState<
      NetworkInfo | null
    >(null);

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
            .includes(
              keyword
            ) ||
          item.urlSms
            .toLowerCase()
            .includes(
              keyword
            )
      );
    }, [
      items,
      search,
    ]);

  /* =========================
     อ่านค่าจาก Excel
  ========================= */

  const getExcelValue = (
    row:
      Record<
        string,
        unknown
      >,
    possibleHeaders:
      string[]
  ) => {
    const keys =
      Object.keys(row);

    for (
      const header
      of possibleHeaders
    ) {
      const foundKey =
        keys.find(
          (key) =>
            key
              .trim()
              .toLowerCase() ===
            header
              .trim()
              .toLowerCase()
        );

      if (foundKey) {
        return String(
          row[
            foundKey
          ] ?? ""
        ).trim();
      }
    }

    return "";
  };

  /* =========================
     IMPORT EXCEL
  ========================= */

  const handleFileChange =
    async (
      event:
        ChangeEvent<HTMLInputElement>
    ) => {
      const file =
        event.target
          .files?.[0];

      if (!file) {
        setFileName(
          ""
        );

        setItems([]);

        return;
      }

      setFileName(
        file.name
      );

      try {
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

        if (
          workbook
            .SheetNames
            .length === 0
        ) {
          alert(
            "ไม่พบ Sheet ภายในไฟล์ Excel"
          );

          return;
        }

        const firstSheetName =
          workbook
            .SheetNames[0];

        const worksheet =
          workbook
            .Sheets[
              firstSheetName
            ];

        const rows =
          XLSX.utils
            .sheet_to_json<
              Record<
                string,
                unknown
              >
            >(
              worksheet,
              {
                defval:
                  "",
              }
            );

        const importedItems:
          UrlItem[] =
          rows
            .map(
              (
                row,
                index
              ) => {
                const urlSms =
                  getExcelValue(
                    row,
                    [
                      "URL SMS",
                      "URL",
                      "URLSMS",
                    ]
                  );

                const ais =
                  getExcelValue(
                    row,
                    [
                      "AIS",
                    ]
                  );

                const trueDtac =
                  getExcelValue(
                    row,
                    [
                      "TRUE/DTAC",
                      "TRUE / DTAC",
                      "TRUE-DTAC",
                      "TRUE DTAC",
                    ]
                  );

                const nt =
                  getExcelValue(
                    row,
                    [
                      "NT",
                    ]
                  );

                const cloudflare =
                  getExcelValue(
                    row,
                    [
                      "CloudFlare",
                      "Cloudflare",
                      "CLOUDFLARE",
                      "CLOUD FLARE",
                    ]
                  );

                return {
                  id:
                    index +
                    1,

                  caseId:
                    `CASE-${String(
                      index +
                        1
                    ).padStart(
                      3,
                      "0"
                    )}`,

                  urlSms,

                  original:
                    {
                      ais,
                      trueDtac,
                      nt,
                      cloudflare,
                    },

                  checked:
                    {
                      ais:
                        "",
                      trueDtac:
                        "",
                      nt:
                        "",
                      cloudflare:
                        "",
                    },
                };
              }
            )
            .filter(
              (
                item
              ) =>
                item.urlSms !==
                ""
            );

        setItems(
          importedItems
        );

        setViewMode(
          "original"
        );

        setNetworkInfo(
          null
        );

        alert(
          `อ่านไฟล์สำเร็จ ${importedItems.length} รายการ`
        );
      } catch (
        error
      ) {
        console.error(
          "Excel import error:",
          error
        );

        alert(
          "อ่านไฟล์ Excel ไม่สำเร็จ กรุณาตรวจสอบรูปแบบไฟล์และชื่อคอลัมน์"
        );
      }
    };

  /* =========================
     NETWORK NAME
  ========================= */

  const getProviderName =
    (
      value:
        ProviderKey =
        provider
    ) => {
      if (
        value === "ais"
      ) {
        return "AIS";
      }

      if (
        value ===
        "trueDtac"
      ) {
        return "TRUE / DTAC";
      }

      if (
        value === "nt"
      ) {
        return "NT";
      }

      return "CloudFlare";
    };

  /* =========================
     DETECT NETWORK
  ========================= */

  const handleDetectNetwork =
    async (): Promise<
      NetworkInfo | null
    > => {
      setIsDetectingNetwork(
        true
      );

      try {
        const response =
          await fetch(
            "http://localhost:4000/api/network-info"
          );

        if (
          !response.ok
        ) {
          throw new Error(
            "ตรวจ Network ไม่สำเร็จ"
          );
        }

        const data =
          (await response.json()) as NetworkInfo;

        setNetworkInfo(
          data
        );

        return data;
      } catch (
        error
      ) {
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
    };

  /* =========================
     RUN CHECK
  ========================= */

  const handleRunProcess =
    async () => {
      if (
        items.length ===
        0
      ) {
        alert(
          "ยังไม่มี URL สำหรับตรวจ"
        );

        return;
      }

      /*
        1. ตรวจ Network ก่อน
      */

      const detected =
        await handleDetectNetwork();

      if (!detected) {
        return;
      }

      /*
        2. ถ้าระบบรู้ Network
        แต่ไม่ตรงกับที่เลือก
        ให้หยุด
      */

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
            `Public IP: ${detected.ip ?? "-"}\n\n` +
            `กรุณาเปลี่ยนเครือข่าย หรือเลือก Network ให้ตรงก่อน`
        );

        return;
      }

      /*
        3. ถ้าตรวจ Provider ไม่ได้
        ให้คนยืนยันเอง
      */

      if (
        detected.provider ===
        "unknown"
      ) {
        const confirmed =
          window.confirm(
            `ระบบไม่สามารถระบุผู้ให้บริการอัตโนมัติได้\n\n` +
              `Public IP: ${detected.ip ?? "-"}\n` +
              `ISP/ORG: ${detected.org ?? "-"}\n\n` +
              `คุณยืนยันว่าเครื่องกำลังเชื่อมผ่าน ${getProviderName()} ใช่หรือไม่?`
          );

        if (
          !confirmed
        ) {
          return;
        }
      }

      /*
        4. Network ผ่านแล้ว
        ค่อยตรวจ URL
      */

      setIsProcessing(
        true
      );

      try {
        const response =
          await fetch(
            "http://localhost:4000/api/check-url",
            {
              method:
                "POST",

              headers:
                {
                  "Content-Type":
                    "application/json",
                },

              body:
                JSON.stringify(
                  {
                    urls:
                      items.map(
                        (
                          item
                        ) =>
                          item.urlSms
                      ),
                  }
                ),
            }
          );

        if (
          !response.ok
        ) {
          throw new Error(
            "ตรวจ URL ไม่สำเร็จ"
          );
        }

        const data =
          (await response.json()) as ApiCheckResponse;

        setItems(
          (
            current
          ) =>
            current.map(
              (
                item
              ) => {
                const result =
                  data.results.find(
                    (
                      r
                    ) =>
                      r.url ===
                      item.urlSms
                  );

                if (
                  !result
                ) {
                  return item;
                }

                const checkText =
                  result.reachable
                    ? `เปิดได้ (${result.statusCode})`
                    : result.error ||
                      "เปิดไม่ได้";

                return {
                  ...item,

                  checked:
                    {
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
          `ตรวจเสร็จ ${data.total} รายการ\nNetwork: ${getProviderName()}\nIP: ${detected.ip ?? "-"}`
        );
      } catch (
        error
      ) {
        console.error(
          "Run check error:",
          error
        );

        alert(
          "ตรวจ URL ไม่สำเร็จ กรุณาตรวจสอบว่า Backend ที่ port 4000 ยังทำงานอยู่"
        );
      } finally {
        setIsProcessing(
          false
        );
      }
    };

  /* =========================
     MANUAL QR RESULT
  ========================= */

  const handleFoundQr =
    (
      item:
        UrlItem
    ) => {
      setItems(
        (
          current
        ) =>
          current.map(
            (
              currentItem
            ) =>
              currentItem.id ===
              item.id
                ? {
                    ...currentItem,

                    checked:
                      {
                        ...currentItem.checked,

                        [provider]:
                          item.urlSms,
                      },
                  }
                : currentItem
          )
      );

      setViewMode(
        "checked"
      );
    };

  const handleNotFoundQr =
    (
      id:
        number
    ) => {
      setItems(
        (
          current
        ) =>
          current.map(
            (
              item
            ) =>
              item.id ===
              id
                ? {
                    ...item,

                    checked:
                      {
                        ...item.checked,

                        [provider]:
                          "ไม่พบ QR",
                      },
                  }
                : item
          )
      );

      setViewMode(
        "checked"
      );
    };

  const handleClearResult =
    (
      id:
        number
    ) => {
      setItems(
        (
          current
        ) =>
          current.map(
            (
              item
            ) =>
              item.id ===
              id
                ? {
                    ...item,

                    checked:
                      {
                        ...item.checked,

                        [provider]:
                          "",
                      },
                  }
                : item
          )
      );
    };

  /* =========================
     SAVE
  ========================= */

  const handleSaveAll =
    () => {
      console.log(
        "ข้อมูลต้นฉบับ + ผลตรวจ:",
        items
      );

      console.log(
        "Network:",
        networkInfo
      );

      alert(
        "ตอนนี้ข้อมูลยังเก็บในหน้าเว็บ ขั้นต่อไปค่อยบันทึกลง Database"
      );
    };

  /* =========================
     RESULT DISPLAY
  ========================= */

  const renderProviderValue =
    (
      value:
        string
    ) => {
      const cleanValue =
        String(
          value ??
            ""
        ).trim();

      if (
        !cleanValue
      ) {
        return (
          <span className="provider-empty">
            -
          </span>
        );
      }

      if (
        cleanValue.startsWith(
          "เปิดได้"
        )
      ) {
        return (
          <span className="provider-reachable">
            ✓{" "}
            {
              cleanValue
            }
          </span>
        );
      }

      if (
        cleanValue ===
          "เปิดไม่ได้" ||
        cleanValue ===
          "Timeout" ||
        cleanValue
          .toLowerCase()
          .includes(
            "fetch failed"
          ) ||
        cleanValue
          .toLowerCase()
          .includes(
            "failed"
          )
      ) {
        return (
          <span className="provider-unreachable">
            ✕{" "}
            {
              cleanValue
            }
          </span>
        );
      }

      if (
        cleanValue.toLowerCase() ===
          "ไม่พบ qr" ||
        cleanValue.toLowerCase() ===
          "not found"
      ) {
        return (
          <span className="provider-not-found">
            ✕ ไม่พบ QR
          </span>
        );
      }

      if (
        cleanValue.startsWith(
          "http://"
        ) ||
        cleanValue.startsWith(
          "https://"
        )
      ) {
        return (
          <a
            href={
              cleanValue
            }
            target="_blank"
            rel="noreferrer"
            className="provider-found"
            title={
              cleanValue
            }
          >
            ✓ พบลิงก์
          </a>
        );
      }

      return (
        <span className="provider-text">
          {
            cleanValue
          }
        </span>
      );
    };

  const getDisplayedData =
    (
      item:
        UrlItem
    ) => {
      if (
        viewMode ===
        "original"
      ) {
        return item.original;
      }

      return item.checked;
    };

  return (
    <main className="fake-page">
      <div className="fake-container">

        <p className="fake-breadcrumb">
          URL Fake Web
        </p>

        <div className="fake-divider" />

        <h1 className="fake-title">
          URL Fake Web
        </h1>

        {/* TOOLBAR */}

        <section className="fake-actions">

          <div className="action-group">
            <label className="action-button import-button">
              ↥{" "}
              {fileName ||
                "Import spreadsheet"}

              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={
                  handleFileChange
                }
                hidden
              />
            </label>

            <span className="action-description">
              รองรับไฟล์ .xlsx, .xls
            </span>
          </div>

          <div className="action-group">
            <select
              className="network-select"
              value={
                provider
              }
              onChange={(
                event
              ) =>
                setProvider(
                  event
                    .target
                    .value as ProviderKey
                )
              }
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

            <span className="action-description">
              เครือข่ายที่ต้องการทดสอบ
            </span>
          </div>

          <div className="action-group">
            <button
              type="button"
              className="action-button detect-network-button"
              onClick={
                handleDetectNetwork
              }
              disabled={
                isDetectingNetwork
              }
            >
              {isDetectingNetwork
                ? "กำลังตรวจ Network..."
                : "ตรวจ Network"}
            </button>

            <span className="action-description">
              ตรวจ Public IP / ISP
            </span>
          </div>

          <div className="action-group">
            <button
              type="button"
              className="action-button run-button"
              onClick={
                handleRunProcess
              }
              disabled={
                isProcessing ||
                isDetectingNetwork ||
                items.length ===
                  0
              }
            >
              ▶{" "}
              {isProcessing
                ? "กำลังตรวจ..."
                : "Run check"}
            </button>

            <span className="action-description">
              ตรวจ URL ผ่าน Network ปัจจุบัน
            </span>
          </div>

          <div className="action-group search-group">
            <div className="new-search-box">
              <span>
                ⌕
              </span>

              <input
                type="text"
                placeholder="Search CASE ID, URL..."
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
          </div>

          <div className="action-group">
            <button
              type="button"
              className="action-button save-button-new"
              onClick={
                handleSaveAll
              }
              disabled={
                items.length ===
                0
              }
            >
              บันทึกข้อมูล
            </button>

            <span className="action-description">
              บันทึกผลการตรวจทั้งหมด
            </span>
          </div>

        </section>

        {/* NETWORK INFO */}

        <div className="current-network">

          <div>
            เครือข่ายที่เลือก:
            <strong>
              {" "}
              {getProviderName()}
            </strong>
          </div>

          {networkInfo && (
            <div className="detected-network-info">

              <span>
                Public IP:
                <strong>
                  {" "}
                  {networkInfo.ip ??
                    "-"}
                </strong>
              </span>

              <span>
                ISP:
                <strong>
                  {" "}
                  {networkInfo.org ??
                    "-"}
                </strong>
              </span>

              <span>
                ตรวจพบ:
                <strong>
                  {" "}
                  {networkInfo.provider ===
                  "unknown"
                    ? "Unknown"
                    : getProviderName(
                        networkInfo.provider
                      )}
                </strong>
              </span>

            </div>
          )}

        </div>

        {/* TABLE */}

        <section className="verification-card">

          <div className="verification-header">

            <div>
              <h2>
                Verification queue
              </h2>

              <p>
                {filteredItems.length} cases loaded
              </p>
            </div>

            <div className="result-switch">

              <button
                type="button"
                className={
                  viewMode ===
                  "original"
                    ? "active"
                    : ""
                }
                onClick={() =>
                  setViewMode(
                    "original"
                  )
                }
              >
                ต้นฉบับจาก Excel
              </button>

              <button
                type="button"
                className={
                  viewMode ===
                  "checked"
                    ? "active"
                    : ""
                }
                onClick={() =>
                  setViewMode(
                    "checked"
                  )
                }
              >
                ผลตรวจของระบบ
              </button>

            </div>

          </div>

          <div className="table-wrapper">

            <table className="verification-table">

              <thead>
                <tr>
                  <th>
                    CASE ID
                  </th>

                  <th>
                    URL SMS
                  </th>

                  <th className="ais-column">
                    AIS
                  </th>

                  <th className="true-column">
                    TRUE / DTAC
                  </th>

                  <th className="nt-column">
                    NT
                  </th>

                  <th className="cloudflare-column">
                    CloudFlare
                  </th>

                  <th>
                    ACTION
                  </th>
                </tr>
              </thead>

              <tbody>

                {filteredItems.length ===
                0 ? (
                  <tr>
                    <td
                      colSpan={
                        7
                      }
                      className="empty-table"
                    >
                      ยังไม่มีข้อมูล กรุณา Import ไฟล์ Excel
                    </td>
                  </tr>
                ) : (
                  filteredItems.map(
                    (
                      item
                    ) => {
                      const displayed =
                        getDisplayedData(
                          item
                        );

                      return (
                        <tr
                          key={
                            item.id
                          }
                        >

                          <td className="case-id">
                            {item.caseId}
                          </td>

                          <td>
                            <a
                              href={
                                item.urlSms
                              }
                              target="_blank"
                              rel="noreferrer"
                              className="url-link"
                            >
                              {item.urlSms} ↗
                            </a>
                          </td>

                          <td>
                            {renderProviderValue(
                              displayed.ais
                            )}
                          </td>

                          <td>
                            {renderProviderValue(
                              displayed.trueDtac
                            )}
                          </td>

                          <td>
                            {renderProviderValue(
                              displayed.nt
                            )}
                          </td>

                          <td>
                            {renderProviderValue(
                              displayed.cloudflare
                            )}
                          </td>

                          <td>
                            <div className="action-cell">

                              <a
                                href={
                                  item.urlSms
                                }
                                target="_blank"
                                rel="noreferrer"
                                className="open-button"
                              >
                                Open
                              </a>

                              <button
                                type="button"
                                className="found-button"
                                onClick={() =>
                                  handleFoundQr(
                                    item
                                  )
                                }
                              >
                                ✓ พบ QR
                              </button>

                              <button
                                type="button"
                                className="not-found-button"
                                onClick={() =>
                                  handleNotFoundQr(
                                    item.id
                                  )
                                }
                              >
                                ✕ ไม่พบ QR
                              </button>

                              <button
                                type="button"
                                className="clear-button"
                                onClick={() =>
                                  handleClearResult(
                                    item.id
                                  )
                                }
                              >
                                ล้าง
                              </button>

                            </div>
                          </td>

                        </tr>
                      );
                    }
                  )
                )}

              </tbody>

            </table>

          </div>

        </section>

      </div>
    </main>
  );
}
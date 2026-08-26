"use client";

import { ChangeEvent, useMemo, useState } from "react";
import "./ssl-checker.css";

type SslStatus = "valid" | "expiring" | "expired";

type SslItem = {
  id: number;
  url: string;
  expirationDate: string;
  status: SslStatus;
};

const exampleItems: SslItem[] = [
  {
    id: 1,
    url: "https://example.com",
    expirationDate: "20/12/2026",
    status: "valid",
  },
  {
    id: 2,
    url: "https://example.org",
    expirationDate: "02/09/2026",
    status: "expiring",
  },
  {
    id: 3,
    url: "https://expired-example.com",
    expirationDate: "10/08/2026",
    status: "expired",
  },
];

export default function SslCheckerPage() {
  const [fileName, setFileName] = useState("");
  const [items, setItems] = useState<SslItem[]>(exampleItems);
  const [search, setSearch] = useState("");
  const [isChecking, setIsChecking] = useState(false);

  const filteredItems = useMemo(() => {
    const keyword = search.toLowerCase().trim();

    if (!keyword) {
      return items;
    }

    return items.filter(
      (item) =>
        item.url.toLowerCase().includes(keyword) ||
        item.expirationDate.toLowerCase().includes(keyword) ||
        item.status.toLowerCase().includes(keyword)
    );
  }, [items, search]);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];

    if (!file) {
      setFileName("");
      return;
    }

    setFileName(file.name);
  };

  const handleRunCheck = async () => {
    if (!fileName) {
      alert("กรุณาเลือกไฟล์ก่อน");
      return;
    }

    setIsChecking(true);

    // Mock เท่านั้น
    // ภายหลังเปลี่ยนเป็น Backend API ตรวจ Certificate จริง
    await new Promise((resolve) => setTimeout(resolve, 1000));

    setIsChecking(false);
  };

  const getStatusText = (status: SslStatus) => {
    if (status === "valid") {
      return "VALID";
    }

    if (status === "expiring") {
      return "EXPIRING SOON";
    }

    return "EXPIRED";
  };

  return (
    <main className="ssl-page">
      <div className="ssl-container">
        <p className="ssl-breadcrumb">SSL/TLS Checker</p>

        <div className="ssl-divider" />

        <h1 className="ssl-title">SSL/TLS Checker</h1>

        <section className="ssl-actions">
          <label className="ssl-import-button">
            <span className="ssl-button-dot" />

            {fileName || "Import spreadsheet"}

            <input
              type="file"
              accept=".xlsx,.xls,.txt"
              onChange={handleFileChange}
              hidden
            />
          </label>

          <button
            type="button"
            className="ssl-run-button"
            onClick={handleRunCheck}
            disabled={isChecking}
          >
            {isChecking ? "Checking..." : "Run check"}

            <span className="ssl-button-dot" />
          </button>

          <div className="ssl-search-box">
            <span>⌕</span>

            <input
              type="text"
              placeholder="Search..."
              value={search}
              onChange={(event) => setSearch(event.target.value)}
            />
          </div>
        </section>

        <section className="ssl-card">
          <div className="ssl-card-header">
            <h2>SSL Certificate</h2>
            <p>Live validation results</p>
          </div>

          <div className="ssl-table-wrapper">
            <table className="ssl-table">
              <thead>
                <tr>
                  <th>URL</th>
                  <th>EXPIRATION DATE</th>
                  <th>STATUS</th>
                </tr>
              </thead>

              <tbody>
                {filteredItems.map((item) => (
                  <tr key={item.id}>
                    <td>
                      <a
                        href={item.url}
                        target="_blank"
                        rel="noreferrer"
                        className="ssl-url"
                      >
                        {item.url}
                      </a>
                    </td>

                    <td>{item.expirationDate}</td>

                    <td>
                      <span
                        className={`ssl-status ssl-status-${item.status}`}
                      >
                        {getStatusText(item.status)}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}
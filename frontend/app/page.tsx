"use client";

import { ChangeEvent, useState } from "react";

type LineItem = {
  id: number;
  name: string;
  url: string;
  status: "waiting" | "found" | "not_found";
};

const exampleItems: LineItem[] = [
  {
    id: 1,
    name: "ร้านค้า A",
    url: "https://lin.ee/example1",
    status: "found",
  },
  {
    id: 2,
    name: "ร้านค้า B",
    url: "https://lin.ee/example2",
    status: "waiting",
  },
  {
    id: 3,
    name: "ร้านค้า C",
    url: "https://example.com",
    status: "not_found",
  },
];

export default function Home() {
  const [fileName, setFileName] = useState("");
  const [items, setItems] = useState<LineItem[]>(exampleItems);
  const [isChecking, setIsChecking] = useState(false);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];

    if (!file) {
      setFileName("");
      return;
    }

    setFileName(file.name);
  };

  const handleCheckAll = async () => {
    setIsChecking(true);

    // ตอนนี้เป็นข้อมูลจำลอง
    // ภายหลังค่อยเปลี่ยนเป็นเรียก Backend API
    await new Promise((resolve) => setTimeout(resolve, 1200));

    setItems((currentItems) =>
      currentItems.map((item) => ({
        ...item,
        status: item.url.includes("lin.ee") ? "found" : "not_found",
      }))
    );

    setIsChecking(false);
  };

  const getStatus = (status: LineItem["status"]) => {
    if (status === "found") {
      return <span className="status statusFound">พบลิงก์ LINE</span>;
    }

    if (status === "not_found") {
      return <span className="status statusNotFound">ไม่พบ</span>;
    }

    return <span className="status statusWaiting">รอตรวจสอบ</span>;
  };

  return (
    <main className="page">
      <header className="header">
        <div className="headerContent">
          <div>
            <p className="eyebrow">LINE QR CHECKER</p>
            <h1>ระบบตรวจสอบลิงก์ LINE</h1>
            <p className="subtitle">
              อัปโหลดไฟล์ Excel ตรวจสอบลิงก์ และสร้าง QR Code
            </p>
          </div>

          <div className="headerIcon">QR</div>
        </div>
      </header>

      <section className="container">
        <div className="summaryGrid">
          <div className="summaryCard">
            <span>รายการทั้งหมด</span>
            <strong>{items.length}</strong>
          </div>

          <div className="summaryCard">
            <span>พบลิงก์ LINE</span>
            <strong>
              {items.filter((item) => item.status === "found").length}
            </strong>
          </div>

          <div className="summaryCard">
            <span>ไม่พบ</span>
            <strong>
              {items.filter((item) => item.status === "not_found").length}
            </strong>
          </div>
        </div>

        <section className="card uploadCard">
          <div>
            <h2>อัปโหลดไฟล์ Excel</h2>
            <p>รองรับไฟล์ .xlsx และ .xls</p>
          </div>

          <div className="uploadArea">
            <label className="fileButton">
              เลือกไฟล์ Excel
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileChange}
                hidden
              />
            </label>

            <div className="fileName">
              {fileName || "ยังไม่ได้เลือกไฟล์"}
            </div>

            <button
              type="button"
              className="primaryButton"
              disabled={!fileName}
            >
              อัปโหลดไฟล์
            </button>
          </div>
        </section>

        <section className="card">
          <div className="tableHeader">
            <div>
              <h2>ผลการตรวจสอบ</h2>
              <p>รายการลิงก์ที่อ่านจากไฟล์ Excel</p>
            </div>

            <div className="actions">
              <button
                type="button"
                className="secondaryButton"
                onClick={handleCheckAll}
                disabled={isChecking}
              >
                {isChecking ? "กำลังตรวจสอบ..." : "ตรวจสอบทั้งหมด"}
              </button>

              <button type="button" className="primaryButton">
                Export Excel
              </button>
            </div>
          </div>

          <div className="tableWrapper">
            <table>
              <thead>
                <tr>
                  <th>ลำดับ</th>
                  <th>ชื่อ</th>
                  <th>ลิงก์ LINE</th>
                  <th>QR Code</th>
                  <th>สถานะ</th>
                  <th>การทำงาน</th>
                </tr>
              </thead>

              <tbody>
                {items.map((item) => (
                  <tr key={item.id}>
                    <td>{item.id}</td>
                    <td className="nameCell">{item.name}</td>

                    <td>
                      <div className="urlCell">{item.url}</div>
                    </td>

                    <td>
                      <div className="qrPlaceholder">
                        <span>QR</span>
                      </div>
                    </td>

                    <td>{getStatus(item.status)}</td>

                    <td>
                      <a
                        href={item.url}
                        target="_blank"
                        rel="noreferrer"
                        className="openLink"
                      >
                        เปิดลิงก์
                      </a>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </section>
    </main>
  );
}
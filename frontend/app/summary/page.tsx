"use client";

import { useState } from "react";
import "./summary.css";

import BarChart from "./charts/BarChart";
import PieChart from "./charts/PieChart";
import LineChart from "./charts/LineChart";

export default function SummaryPage() {
  const [chartType, setChartType] = useState<"bar" | "pie" | "line">("bar");

  return (
    <main className="summary-page">
      <div className="summary-container">
        <p className="summary-breadcrumb">main</p>

        <div className="summary-divider" />

        <h1 className="summary-title">Summary</h1>

        <section className="summary-section">
          <h2>CLAB</h2>

          <div className="clab-card">
            <div className="clab-card-header">
              <div>
                <h3>Signal distribution</h3>
                <p>Edits in this table update every chart instantly.</p>
              </div>

              <div className="chart-tabs">
                <button
                  className={chartType === "bar" ? "active" : ""}
                  onClick={() => setChartType("bar")}
                >
                  BAR
                </button>

                <button
                  className={chartType === "pie" ? "active" : ""}
                  onClick={() => setChartType("pie")}
                >
                  PIE
                </button>

                <button
                  className={chartType === "line" ? "active" : ""}
                  onClick={() => setChartType("line")}
                >
                  LINE
                </button>
              </div>
            </div>

            <div className="chart-display">
              {chartType === "bar" && <BarChart />}
              {chartType === "pie" && <PieChart />}
              {chartType === "line" && <LineChart />}
            </div>
          </div>
        </section>

        <section className="summary-cards">
          <div className="info-card">
            <h3>URL Fake Web</h3>
            <p className="info-number">120</p>
            <span>Total URLs checked</span>
          </div>

          <div className="info-card">
            <h3>SSL / TLS Checker</h3>
            <p className="info-number">84</p>
            <span>Certificates checked</span>
          </div>

          <div className="info-card">
            <h3>CLAB Requests</h3>
            <p className="info-number">32</p>
            <span>Total requests</span>
          </div>
        </section>
      </div>
    </main>
  );
}
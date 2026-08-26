"use client";

import { useState } from "react";
import "./clab.css";

import BarChart from "./charts/BarChart";
import PieChart from "./charts/PieChart";
import LineChart from "./charts/LineChart";


type ChartType = "bar" | "pie" | "line";

export default function ClabPage() {
  const [chartType, setChartType] = useState<ChartType>("bar");

  return (
    <main className="clab-page">
      <div className="clab-container">
        <p className="clab-breadcrumb">CLAB</p>

        <div className="clab-divider" />

        <h1 className="clab-title">CLAB</h1>

        <button className="import-btn" type="button">
          <span className="import-dot" />
          Import data
        </button>

        <section className="clab-chart-card">
          <div className="clab-card-header">
            <div>
              <h3>Signal distribution</h3>
              <p>Edits in the table update every chart instantly.</p>
            </div>

            <div className="clab-tabs">
              <button
                type="button"
                className={chartType === "bar" ? "active" : ""}
                onClick={() => setChartType("bar")}
              >
                BAR
              </button>

              <button
                type="button"
                className={chartType === "pie" ? "active" : ""}
                onClick={() => setChartType("pie")}
              >
                PIE
              </button>

              <button
                type="button"
                className={chartType === "line" ? "active" : ""}
                onClick={() => setChartType("line")}
              >
                LINE
              </button>
            </div>
          </div>

          <div className="chart-area">
            {chartType === "bar" && <BarChart />}
            {chartType === "pie" && <PieChart />}
            {chartType === "line" && <LineChart />}
          </div>
        </section>

       
      </div>
    </main>
  );
}
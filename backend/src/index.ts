import "dotenv/config";
import express from "express";
import cors from "cors";

import checkUrlRoutes from "./routes/checkUrl.js";
import networkInfoRoutes from "./routes/networkInfo.js";
import { initializeDatabase } from "./db.js";

const app = express();

const port = Number(process.env.PORT ?? 4000);

app.use(
  cors({
    origin: [
      "http://localhost:3000",
      "http://localhost:3001",
    ],
  })
);

app.use(express.json());

app.get("/", (_req, res) => {
  res.json({
    message: "OP Web Service API",
    status: "running",
  });
});

/* ตรวจ URL */
app.use("/api/check-url", checkUrlRoutes);

/* ตรวจ Public IP / ISP / Network */
app.use("/api/network-info", networkInfoRoutes);

async function startServer() {
  try {
    await initializeDatabase();

    console.log("Database initialized successfully");

    app.listen(port, "0.0.0.0", () => {
      console.log(
        `Backend running on http://localhost:${port}`
      );

      console.log(
        `Check URL API: http://localhost:${port}/api/check-url`
      );

      console.log(
        `Network Info API: http://localhost:${port}/api/network-info`
      );
    });
  } catch (error) {
    console.error("Failed to initialize database:", error);
    process.exit(1);
  }
}

startServer();
import "dotenv/config";

import express from "express";
import cors from "cors";

import checkUrlRoutes from "./routes/checkUrl.js";
import networkInfoRoutes from "./routes/networkInfo.js";
import authRoutes from "./routes/auth.js";
import usersRoutes from "./routes/users.js";
import logsRoutes from "./routes/logs.js";
import sessionsRoutes from "./routes/sessions.js";
import summaryRoutes from "./routes/summary.js";
import checkerResultsRoutes from "./routes/checkerResults.js";
import sslCheckerRouter from "./routes/sslChecker.js";
import clabRoutes from "./routes/clab.js";

import {
  initializeDatabase,
} from "./db.js";

import {
  initializeSessionDatabase,
} from "./sessionDb.js";

/* ==============================
   APP CONFIGURATION
================================ */

const app = express();

const port = Number(
  process.env.PORT ?? 4000
);

/*
 * ทำให้ req.ip อ่านค่า IP ผ่าน Docker
 * หรือ Reverse Proxy ได้ถูกต้อง
 */
app.set(
  "trust proxy",
  1
);

/* ==============================
   CORS
================================ */

app.use(
  cors({
    origin: [
      "http://localhost:3000",
      "http://localhost:3001",
    ],
  })
);

/* ==============================
   BODY PARSER
================================ */

app.use(
  express.json({
    limit: "100mb",
  })
);

app.use(
  express.urlencoded({
    extended: true,
    limit: "100mb",
  })
);

/* ==============================
   ROOT
================================ */

app.get(
  "/",
  (_req, res) => {
    return res.json({
      message:
        "OP Web Service API",

      status:
        "running",
    });
  }
);

/* ==============================
   ROUTES
================================ */

/* ตรวจ URL */
app.use(
  "/api/check-url",
  checkUrlRoutes
);

/* ตรวจ Public IP / ISP / Network */
app.use(
  "/api/network-info",
  networkInfoRoutes
);

/* Login */
app.use(
  "/api/auth",
  authRoutes
);

/* จัดการผู้ใช้งาน */
app.use(
  "/api/users",
  usersRoutes
);

/* System Log */
app.use(
  "/api/logs",
  logsRoutes
);

/* Session / Online Status */
app.use(
  "/api/sessions",
  sessionsRoutes
);

/* Main Summary */
app.use(
  "/api/summary",
  summaryRoutes
);

/* URL Fake Web Results */
app.use(
  "/api/checker-results",
  checkerResultsRoutes
);

/* SSL / TLS Checker */
app.use(
  "/api/ssl-checker",
  sslCheckerRouter
);

/* CLAB */
app.use(
  "/api/clab",
  clabRoutes
);

/* ==============================
   START SERVER
================================ */

async function startServer(): Promise<void> {
  try {
    /*
     * สร้างตารางเดิมของระบบก่อน
     */
    await initializeDatabase();

    /*
     * จากนั้นสร้างตาราง Session
     * เพราะต้องอ้างอิง users และ admin_logs
     */
    await initializeSessionDatabase();

    app.listen(
      port,
      "0.0.0.0",
      () => {
        console.log(
          `Backend running on http://localhost:${port}`
        );

        console.log(
          `Login API: http://localhost:${port}/api/auth/login`
        );

        console.log(
          `Session API: http://localhost:${port}/api/sessions`
        );

        console.log(
          `Session overview: http://localhost:${port}/api/sessions/overview`
        );
      }
    );
  } catch (error) {
    console.error(
      "Failed to initialize database:",
      error
    );

    process.exit(1);
  }
}

void startServer();
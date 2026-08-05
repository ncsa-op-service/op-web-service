import express from "express";
import cors from "cors";
import { initializeDatabase, pool } from "./db.js";
import uploadRouter from "./routes/upload.js";
import linksRouter from "./routes/links.js";

const app = express();
const port = Number(process.env.PORT ?? 4000);

app.use(
  cors({
    origin: "http://localhost:3000",
  })
);

app.use(express.json());

app.get("/", (_req, res) => {
  res.json({
    message: "LINE QR Checker API",
    status: "running",
  });
});

app.get("/db-test", async (_req, res) => {
  try {
    const result = await pool.query(
      "SELECT NOW() AS current_time"
    );

    res.json({
      connected: true,
      databaseTime: result.rows[0].current_time,
    });
  } catch (error) {
    console.error(error);

    res.status(500).json({
      connected: false,
      message: "เชื่อม PostgreSQL ไม่สำเร็จ",
    });
  }
});

app.use("/api/uploads", uploadRouter);
app.use("/api/links", linksRouter);

async function startServer(): Promise<void> {
  try {
    await initializeDatabase();

    app.listen(port, "0.0.0.0", () => {
      console.log(`Backend running on port ${port}`);
    });
  } catch (error) {
    console.error("เริ่ม Backend ไม่สำเร็จ:", error);
    process.exit(1);
  }
}

void startServer();
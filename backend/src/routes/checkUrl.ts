import { Router } from "express";

const router = Router();

type CheckResult = {
  url: string;
  reachable: boolean;
  statusCode: number | null;
  finalUrl: string | null;
  error: string | null;
};

router.post("/", async (req, res) => {
  try {
    const { urls } = req.body;

    if (!Array.isArray(urls)) {
      return res.status(400).json({
        message: "urls must be an array",
      });
    }

    const cleanUrls = urls
      .map((url) => String(url).trim())
      .filter((url) => url !== "");

    if (cleanUrls.length === 0) {
      return res.status(400).json({
        message: "ไม่พบ URL สำหรับตรวจสอบ",
      });
    }

    const results: CheckResult[] = await Promise.all(
      cleanUrls.map(async (url) => {
        try {
          let targetUrl = url;

          // ถ้า Excel มีแค่ domain เช่น example.com
          // ให้เติม https:// อัตโนมัติ
          if (
            !targetUrl.startsWith("http://") &&
            !targetUrl.startsWith("https://")
          ) {
            targetUrl = `https://${targetUrl}`;
          }

          // ตรวจรูปแบบ URL
          new URL(targetUrl);

          const controller = new AbortController();

          const timeout = setTimeout(() => {
            controller.abort();
          }, 10000);

          try {
            const response = await fetch(targetUrl, {
              method: "GET",
              redirect: "follow",
              signal: controller.signal,
              headers: {
                "User-Agent":
                  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OP-Web-Service-Link-Checker/1.0",
              },
            });

            return {
              url,
              reachable: response.ok,
              statusCode: response.status,
              finalUrl: response.url,
              error: null,
            };
          } finally {
            clearTimeout(timeout);
          }
        } catch (error) {
          let errorMessage = "ไม่สามารถเชื่อมต่อ URL ได้";

          if (error instanceof Error) {
            if (error.name === "AbortError") {
              errorMessage = "Timeout";
            } else {
              errorMessage = error.message;
            }
          }

          return {
            url,
            reachable: false,
            statusCode: null,
            finalUrl: null,
            error: errorMessage,
          };
        }
      })
    );

    return res.json({
      total: results.length,
      results,
    });
  } catch (error) {
    console.error("Check URL error:", error);

    return res.status(500).json({
      message: "Internal server error",
    });
  }
});

export default router;
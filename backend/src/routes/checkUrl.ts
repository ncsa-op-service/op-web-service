import { Router } from "express";

const router = Router();

type CheckErrorType =
  | "nxdomain"
  | "timeout"
  | "connection"
  | "tls"
  | "invalid_url"
  | "network"
  | null;

type CheckResult = {
  url: string;
  reachable: boolean;
  statusCode: number | null;
  finalUrl: string | null;
  error: string | null;
  errorType: CheckErrorType;
};

type ErrorLike = {
  name?: unknown;
  message?: unknown;
  code?: unknown;
  cause?: unknown;
};

function getErrorCode(
  error: unknown
): string | null {
  let current: unknown = error;

  for (let depth = 0; depth < 5; depth += 1) {
    if (
      !current ||
      typeof current !== "object"
    ) {
      break;
    }

    const value =
      current as ErrorLike;

    if (
      typeof value.code === "string" &&
      value.code.trim()
    ) {
      return value.code.trim();
    }

    current = value.cause;
  }

  return null;
}

function getErrorMessage(
  error: unknown
): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (
    error &&
    typeof error === "object" &&
    "message" in error &&
    typeof (error as ErrorLike).message ===
      "string"
  ) {
    return String(
      (error as ErrorLike).message
    );
  }

  return "ไม่สามารถเชื่อมต่อ URL ได้";
}

function classifyFetchError(
  error: unknown
): {
  errorType: Exclude<
    CheckErrorType,
    null
  >;
  errorMessage: string;
} {
  const code =
    getErrorCode(error);

  const message =
    getErrorMessage(error);

  const upperCode =
    (code ?? "").toUpperCase();

  const upperMessage =
    message.toUpperCase();

  const errorName =
    error instanceof Error
      ? error.name
      : "";

  /*
    ENOTFOUND = DNS resolver หา hostname ไม่เจอ
    กรณีนี้ให้ Frontend แสดง
    DNS_PROBE_FINISHED_NXDOMAIN
  */
  if (
    upperCode === "ENOTFOUND" ||
    upperMessage.includes(
      "GETADDRINFO ENOTFOUND"
    )
  ) {
    return {
      errorType: "nxdomain",
      errorMessage:
        code || "ENOTFOUND",
    };
  }

  if (
    errorName === "AbortError" ||
    upperCode ===
      "UND_ERR_CONNECT_TIMEOUT" ||
    upperCode === "ETIMEDOUT" ||
    upperMessage.includes("TIMEOUT")
  ) {
    return {
      errorType: "timeout",
      errorMessage:
        code || "Timeout",
    };
  }

  if (
    [
      "ECONNREFUSED",
      "ECONNRESET",
      "ENETUNREACH",
      "EHOSTUNREACH",
      "EPIPE",
    ].includes(upperCode)
  ) {
    return {
      errorType: "connection",
      errorMessage:
        code || message,
    };
  }

  if (
    upperCode.includes("CERT") ||
    upperCode.includes("TLS") ||
    upperCode.includes("SSL") ||
    upperMessage.includes(
      "CERTIFICATE"
    ) ||
    upperMessage.includes("TLS") ||
    upperMessage.includes("SSL")
  ) {
    return {
      errorType: "tls",
      errorMessage:
        code || message,
    };
  }

  return {
    errorType: "network",
    errorMessage:
      code || message,
  };
}

router.post("/", async (req, res) => {
  try {
    const { urls } = req.body;

    if (!Array.isArray(urls)) {
      return res.status(400).json({
        message: "urls must be an array",
      });
    }

    const cleanUrls = urls
      .map((url) =>
        String(url).trim()
      )
      .filter(
        (url) => url !== ""
      );

    if (cleanUrls.length === 0) {
      return res.status(400).json({
        message:
          "ไม่พบ URL สำหรับตรวจสอบ",
      });
    }

    const results: CheckResult[] =
      await Promise.all(
        cleanUrls.map(
          async (url): Promise<CheckResult> => {
            let targetUrl = url;

            try {
              // ถ้า Excel มีแค่ domain
              // เช่น example.com ให้เติม https://
              if (
                !targetUrl.startsWith(
                  "http://"
                ) &&
                !targetUrl.startsWith(
                  "https://"
                )
              ) {
                targetUrl =
                  `https://${targetUrl}`;
              }

              try {
                new URL(targetUrl);
              } catch {
                return {
                  url,
                  reachable: false,
                  statusCode: null,
                  finalUrl: null,
                  error: "Invalid URL",
                  errorType:
                    "invalid_url",
                };
              }

              const controller =
                new AbortController();

              const timeout =
                setTimeout(() => {
                  controller.abort();
                }, 10000);

              try {
                const response =
                  await fetch(
                    targetUrl,
                    {
                      method: "GET",
                      redirect: "follow",
                      signal:
                        controller.signal,
                      headers: {
                        "User-Agent":
                          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OP-Web-Service-Link-Checker/1.0",
                      },
                    }
                  );

                /*
                  ถ้า fetch ได้ HTTP response กลับมา
                  ถือว่า URL/Server "เปิดถึง"

                  ไม่ใช้ response.ok เพราะ 404 / 403 / 500
                  ถึงจะไม่ใช่ 2xx แต่ Browser ยังเปิดหน้าได้
                */
                return {
                  url,
                  reachable: true,
                  statusCode:
                    response.status,
                  finalUrl:
                    response.url ||
                    targetUrl,
                  error: null,
                  errorType: null,
                };
              } finally {
                clearTimeout(timeout);
              }
            } catch (error) {
              const classified =
                classifyFetchError(
                  error
                );

              return {
                url,
                reachable: false,
                statusCode: null,
                finalUrl: null,
                error:
                  classified.errorMessage,
                errorType:
                  classified.errorType,
              };
            }
          }
        )
      );

    return res.json({
      total: results.length,
      results,
    });
  } catch (error) {
    console.error(
      "Check URL error:",
      error
    );

    return res.status(500).json({
      message:
        "Internal server error",
    });
  }
});

export default router;

import { Router } from "express";

const router = Router();

type DetectedProvider =
  | "ais"
  | "trueDtac"
  | "nt"
  | "cloudflare"
  | "unknown";

function detectProvider(org: string): DetectedProvider {
  const value = org.toLowerCase();

  if (
    value.includes("advanced info service") ||
    value.includes("ais")
  ) {
    return "ais";
  }

  if (
    value.includes("true internet") ||
    value.includes("true corporation") ||
    value.includes("dtac") ||
    value.includes("total access")
  ) {
    return "trueDtac";
  }

  if (
    value.includes("national telecom") ||
    value.includes("tot public") ||
    value.includes("cat telecom")
  ) {
    return "nt";
  }

  if (value.includes("cloudflare")) {
    return "cloudflare";
  }

  return "unknown";
}

router.get("/", async (_req, res) => {
  try {
    const response = await fetch(
      "https://ipapi.co/json/",
      {
        headers: {
          "User-Agent":
            "OP-Web-Service-Network-Detector/1.0",
        },
      }
    );

    if (!response.ok) {
      return res.status(502).json({
        message: "ไม่สามารถตรวจสอบ Public IP ได้",
      });
    }

    const data = (await response.json()) as {
      ip?: string;
      org?: string;
      asn?: string;
      country_name?: string;
    };

    const org = data.org ?? "";

    return res.json({
      ip: data.ip ?? null,
      asn: data.asn ?? null,
      org: org || null,
      country: data.country_name ?? null,
      provider: detectProvider(org),
    });
  } catch (error) {
    console.error("Network detection error:", error);

    return res.status(500).json({
      message: "ตรวจสอบ Network ไม่สำเร็จ",
    });
  }
});

export default router;
"use client";

import { useEffect } from "react";

const HEARTBEAT_INTERVAL_MS = 30_000;

function getApiUrl(): string {
  const value = process.env.NEXT_PUBLIC_API_URL?.trim();
  if (!value || value.includes("backend:")) return "http://localhost:4000";
  return value.replace(/\/+$/, "");
}

export default function SessionHeartbeat() {
  useEffect(() => {
    let timer: number | undefined;
    let stopped = false;

    async function sendHeartbeat() {
      if (stopped || document.visibilityState !== "visible") return;

      const token = localStorage.getItem("token");
      if (!token) return;

      try {
        const response = await fetch(`${getApiUrl()}/api/sessions/heartbeat`, {
          method: "POST",
          headers: { Authorization: `Bearer ${token}` },
          cache: "no-store",
        });

        if (response.status === 401) {
          localStorage.removeItem("token");
          localStorage.removeItem("user");
          window.dispatchEvent(new Event("session-expired"));
        }
      } catch (error) {
        console.error("Heartbeat error:", error);
      }
    }

    function startHeartbeat() {
      if (timer !== undefined) window.clearInterval(timer);
      void sendHeartbeat();
      timer = window.setInterval(() => void sendHeartbeat(), HEARTBEAT_INTERVAL_MS);
    }

    function handleVisibilityChange() {
      if (document.visibilityState === "visible") void sendHeartbeat();
    }

    startHeartbeat();
    window.addEventListener("session-started", startHeartbeat);
    document.addEventListener("visibilitychange", handleVisibilityChange);

    return () => {
      stopped = true;
      if (timer !== undefined) window.clearInterval(timer);
      window.removeEventListener("session-started", startHeartbeat);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
    };
  }, []);

  return null;
}

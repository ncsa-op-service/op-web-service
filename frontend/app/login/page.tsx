"use client";

import { useState } from "react";
import type { FormEvent } from "react";
import { useRouter } from "next/navigation";
import "./login.css";

function getApiUrl(): string {
  const value = process.env.NEXT_PUBLIC_API_URL?.trim();

  if (!value || value.includes("backend:")) {
    return "http://localhost:4000";
  }

  return value.replace(/\/+$/, "");
}

const codeRows = [
  "01001010 11000101 SYSTEM_SECURE ACCESS_CONTROL SSL_TLS 00110101",
  "NETWORK_MONITORING // AUTH_SUCCESS // 101101001 // OP_SERVICE",
  "root@system:~$ sudo monitor --network --secure --access",
  "11010101 00101101 10110010 01011010 11100010 00110110",
  "const security = { status: 'ACTIVE', tls: true, access: true };",
  "NCSA // OP WEB SERVICE // SYSTEM READY // SSL ACTIVE // 010101",
  "<AUTH> VERIFY TOKEN ACCESS_GRANTED </AUTH> 1101001",
  "010101011010 NETWORK_PACKET 110100101 SECURITY_SCAN 010101",
  "TLS_1.3::CONNECTED::SECURE_CHANNEL::OP_SYSTEM::1011010",
  "sudo security-check --network --ssl --auth --monitor",
  "SYSTEM_EVENT 01001001 ACCESS_OK 11010101 CYBER_SECURITY",
  "{ user: 'admin', role: 'secure', status: 200 }",
  "PACKET_CAPTURE 001101101010101 NETWORK_ACTIVE 110010101",
  "OP_SECURITY // AUTH // TLS // MONITOR // ACCESS // SYSTEM",
  "101100101011010110101 SYSTEM_PROCESS ACTIVE CONNECTION_SECURE",
  '<NETWORK status="active" security="enabled" />',
  "01010110 00110101 11010010 NCSA_OPERATION_CENTER 01010",
  "AUTH_TOKEN::VERIFIED // REQUEST::200 // SYSTEM::ONLINE",
  "SECURITY_ENGINE 01101001 SCAN_RUNNING POLICY_OK 10101101",
  "GET /api/auth/login 200 OK // TOKEN_VALID // SESSION_ACTIVE",
  "01100110 FIREWALL_ACTIVE IDS_MONITORING WAF_ENABLED 10101010",
  "OP-CYBER://TRACE/NETWORK/AUTH/SSL/ACCESS/STATUS=SECURE",
];

export default function LoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (loading) return;

    setError("");
    setLoading(true);

    try {
      const response = await fetch(`${getApiUrl()}/api/auth/login`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json",
        },
        body: JSON.stringify({ email: email.trim(), password }),
      });

      const data = await response.json().catch(() => null);

      if (!response.ok) {
        setError(data?.message ?? "Email หรือ Password ไม่ถูกต้อง");
        return;
      }

      if (!data?.user || !data?.token) {
        setError("Backend ไม่ส่งข้อมูล Session กลับมา กรุณาตรวจสอบ auth.ts");
        return;
      }

      localStorage.setItem("user", JSON.stringify(data.user));
      localStorage.setItem("token", data.token);
      window.dispatchEvent(new Event("session-started"));

      router.replace(
        data.user.role === "super_admin" ? "/superadmin" : "/summary"
      );
      router.refresh();
    } catch (loginError) {
      console.error("Login error:", loginError);
      setError("ไม่สามารถเชื่อมต่อ Backend ได้");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="login-page">
      <div className="cyber-effects" aria-hidden="true">
        <div className="blue-code-layer">
          {codeRows.map((text, index) => (
            <div
              key={index}
              className={`blue-code-line code-line-${index + 1}`}
            >
              {text}
            </div>
          ))}
        </div>

        <div className="blue-cyber-glow glow-blue-1" />
        <div className="blue-cyber-glow glow-blue-2" />
      </div>

      <section className="login-left">
        <div className="login-top-brand">
          <div className="ncsa-logo-box">
            <img src="/ncsa-logo.png" alt="NCSA Logo" className="ncsa-logo" />
          </div>

          <div className="login-brand">
            <div className="login-logo">OP</div>

            <div className="login-brand-text">
              <h1>OP Web Service</h1>
              <p>Internal Operations Office Service System</p>
            </div>
          </div>
        </div>

        <p className="login-description">
          A platform for monitoring, managing, and supporting
          <br />
          cybersecurity services.
        </p>
      </section>

      <section className="login-right">
        <div className="login-card">
          <h2>Admin Login</h2>
          <div className="login-title-line" />
          <p className="login-subtitle">Sign in with your email and password.</p>

          <form className="login-form" onSubmit={handleSubmit}>
            <div className="login-input">
              <span className="input-icon">✉</span>
              <input
                type="email"
                name="email"
                placeholder="Email"
                autoComplete="email"
                value={email}
                onChange={(event) => setEmail(event.target.value)}
                disabled={loading}
                required
              />
            </div>

            <div className="login-input">
              <span className="input-icon">🔒</span>
              <input
                type="password"
                name="password"
                placeholder="Password"
                autoComplete="current-password"
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                disabled={loading}
                required
              />
            </div>

            {error && <p className="login-error-message">{error}</p>}

            <a
              href="#"
              className="forgot-password"
              onClick={(event) => event.preventDefault()}
            >
              Forgot Password?
            </a>

            <button type="submit" className="sign-in-button" disabled={loading}>
              {loading ? "Signing in..." : "Sign In"}
            </button>
          </form>

          <p className="contact-admin">
            Don&apos;t have an account? Contact the administrator.
          </p>

          <a href="/" className="back-home">
            ← กลับสู่หน้าหลัก
          </a>
        </div>
      </section>
    </main>
  );
}

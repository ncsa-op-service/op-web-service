export default function Home() {
  return (
    <main className="home-page">
      <div className="page-overlay">
        {/* โค้ดวิ่ง */}
        <div className="home-code-layer" aria-hidden="true">
          <span className="home-code-line home-code-line-1">
            01001010 11000101 SYSTEM_SECURE ACCESS_CONTROL SSL_TLS 00110101
          </span>
          <span className="home-code-line home-code-line-2">
            NETWORK_MONITORING // AUTH_SUCCESS // OP_SERVICE // 101101001
          </span>
          <span className="home-code-line home-code-line-3">
            NCSA • CYBER SECURITY • MONITORING • TLS • ACCESS CONTROL
          </span>
          <span className="home-code-line home-code-line-4">
            11010101 00101101 10110010 01011010 11100010 00110110
          </span>
          <span className="home-code-line home-code-line-5">
            SYSTEM READY // SECURE CHANNEL // NETWORK ACTIVE // 010101
          </span>
          <span className="home-code-line home-code-line-6">
            AUTH_TOKEN VERIFIED • REQUEST 200 • SESSION ACTIVE • OP WEB SERVICE
          </span>
        </div>

        {/* แสง / วงกลม / จุดฟ้า */}
        <div className="home-fx-layer" aria-hidden="true">
          <span className="home-orb orb-1" />
          <span className="home-orb orb-2" />
          <span className="home-orb orb-3" />
          <span className="home-orb orb-4" />
          <span className="home-orb orb-5" />
          <span className="home-orb orb-6" />

          <span className="home-dot dot-1" />
          <span className="home-dot dot-2" />
          <span className="home-dot dot-3" />
          <span className="home-dot dot-4" />
          <span className="home-dot dot-5" />
          <span className="home-dot dot-6" />
          <span className="home-dot dot-7" />
          <span className="home-dot dot-8" />

          <span className="home-ring ring-1" />
          <span className="home-ring ring-2" />
          <span className="home-ring ring-3" />
        </div>

        {/* Header */}
        <header className="header">
          <div className="brand">
            <div className="logo">OP</div>

            <div>
              <h1>OP Web Service</h1>
              <p>Internal Operations Office Service System</p>
            </div>
          </div>
        </header>

        {/* Hero */}
        <section className="hero">
          <div className="hero-content">
            <div className="welcome-small">
              <span className="shield-icon">♢</span>
              Welcome to OP Web Service
            </div>

            <h2>
              Welcome to
              <br />
              OP Web Service
            </h2>

            <div className="title-line" />

            <p className="description">
              A platform for monitoring, managing, and supporting
              <br />
              cybersecurity services.
            </p>

            <a href="/login" className="login-button">
              <span>♙</span>
              Login
            </a>
          </div>
        </section>

        {/* Service Cards */}
        <section className="services">
          <div className="service-card">
            <h3>Secure</h3>
            <p>
              Advanced security
              <br />
              measures to protect
              <br />
              critical data and
              <br />
              systems.
            </p>
          </div>

          <div className="service-card">
            <h3>Monitor</h3>
            <p>
              Real-time monitoring
              <br />
              and alerts to ensure
              <br />
              system health and
              <br />
              performance.
            </p>
          </div>

          <div className="service-card">
            <h3>Manage</h3>
            <p>
              Manage services and
              <br />
              resources efficiently
              <br />
              in one centralized
              <br />
              platform.
            </p>
          </div>

          <div className="service-card">
            <h3>Support</h3>
            <p>
              Dedicated support to
              <br />
              help you with any
              <br />
              issues or inquiries.
            </p>
          </div>
        </section>

        {/* Footer */}
        <footer className="footer">
          <p>
            <strong>OP Web Service</strong>
            <span> © 2026 All rights reserved.</span>
          </p>

          <div className="footer-links">
            <a href="#">Privacy Policy</a>
            <a href="#">Terms of Service</a>
          </div>
        </footer>
      </div>
    </main>
  );
}

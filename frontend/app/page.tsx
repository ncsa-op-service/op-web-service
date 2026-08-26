export default function Home() {
  return (
    <main className="home-page">
      <div className="page-overlay">
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
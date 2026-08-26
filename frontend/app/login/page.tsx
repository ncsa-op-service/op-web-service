import "./login.css";

export default function LoginPage() {
  return (
    <main className="login-page">
      {/* LEFT SIDE */}
      <section className="login-left">
        <div className="login-top-brand">
          <div className="ncsa-logo-box">
            <img
              src="/ncsa-logo.png"
              alt="NCSA Logo"
              className="ncsa-logo"
            />
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

      {/* RIGHT SIDE */}
      <section className="login-right">
        <div className="login-card">
          <h2>Admin Login</h2>

          <div className="login-title-line"></div>

          <p className="login-subtitle">
            Sign in with your email and password.
          </p>

          <form className="login-form">
            <div className="login-input">
              <span className="input-icon">✉</span>

              <input
                type="email"
                name="email"
                placeholder="Email"
                autoComplete="email"
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
                required
              />
            </div>

            <a href="#" className="forgot-password">
              Forgot Password?
            </a>

            <button type="submit" className="sign-in-button">
              Sign In
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
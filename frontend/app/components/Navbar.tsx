"use client";

import Link from "next/link";
import {
  useEffect,
  useRef,
  useState,
} from "react";
import {
  usePathname,
  useRouter,
} from "next/navigation";
import "./Navbar.css";

type LoginUser = {
  id: number;
  name: string;
  email: string;
  role: "super_admin" | "editor" | "viewer";
};

/* ==============================
   FONT SCALE
================================ */

const DEFAULT_FONT_SCALE = 1;
const MIN_FONT_SCALE = 0.9;
const MAX_FONT_SCALE = 1.3;
const FONT_STEP = 0.1;

export default function Navbar() {
  const pathname = usePathname();
  const router = useRouter();

  const [user, setUser] =
    useState<LoginUser | null>(null);

  const [menuOpen, setMenuOpen] =
    useState(false);

  const [fontScale, setFontScale] =
    useState(DEFAULT_FONT_SCALE);

  const menuRef =
    useRef<HTMLDivElement | null>(null);

  /* ==============================
     CURRENT USER
  ================================ */

  function loadCurrentUser() {
    const savedUser =
      localStorage.getItem("user");

    if (!savedUser) {
      setUser(null);
      return;
    }

    try {
      const parsedUser =
        JSON.parse(savedUser) as LoginUser;

      if (
        !parsedUser ||
        !parsedUser.id ||
        !parsedUser.role
      ) {
        setUser(null);
        return;
      }

      setUser(parsedUser);
    } catch {
      localStorage.removeItem("user");
      setUser(null);
    }
  }

  /* ==============================
     APPLY FONT SCALE
  ================================ */

  function applyFontScale(scale: number) {
    const safeScale = Math.min(
      MAX_FONT_SCALE,
      Math.max(MIN_FONT_SCALE, scale)
    );

    const roundedScale =
      Math.round(safeScale * 10) / 10;

    setFontScale(roundedScale);

    document.documentElement.style.setProperty(
      "--app-font-scale",
      String(roundedScale)
    );

    localStorage.setItem(
      "appFontScale",
      String(roundedScale)
    );
  }

  function increaseFontSize() {
    applyFontScale(
      fontScale + FONT_STEP
    );
  }

  function decreaseFontSize() {
    applyFontScale(
      fontScale - FONT_STEP
    );
  }

  function resetFontSize() {
    applyFontScale(
      DEFAULT_FONT_SCALE
    );
  }

  /* ==============================
     LOAD USER + FONT
  ================================ */

  useEffect(() => {
    loadCurrentUser();

    const savedScale =
      localStorage.getItem(
        "appFontScale"
      );

    const parsedScale =
      savedScale
        ? Number(savedScale)
        : DEFAULT_FONT_SCALE;

    if (
      Number.isFinite(parsedScale)
    ) {
      applyFontScale(parsedScale);
    } else {
      applyFontScale(
        DEFAULT_FONT_SCALE
      );
    }

    function handleSessionStarted() {
      loadCurrentUser();
    }

    function handleStorage(
      event: StorageEvent
    ) {
      if (event.key === "user") {
        loadCurrentUser();
      }

      if (
        event.key ===
        "appFontScale"
      ) {
        const nextScale =
          Number(event.newValue);

        if (
          Number.isFinite(nextScale)
        ) {
          applyFontScale(
            nextScale
          );
        }
      }
    }

    window.addEventListener(
      "session-started",
      handleSessionStarted
    );

    window.addEventListener(
      "storage",
      handleStorage
    );

    return () => {
      window.removeEventListener(
        "session-started",
        handleSessionStarted
      );

      window.removeEventListener(
        "storage",
        handleStorage
      );
    };
  }, []);

  /* ==============================
     PATH CHANGE
  ================================ */

  useEffect(() => {
    loadCurrentUser();
    setMenuOpen(false);
  }, [pathname]);

  /* ==============================
     CLICK OUTSIDE
  ================================ */

  useEffect(() => {
    function handleClickOutside(
      event: MouseEvent
    ) {
      if (
        menuRef.current &&
        !menuRef.current.contains(
          event.target as Node
        )
      ) {
        setMenuOpen(false);
      }
    }

    document.addEventListener(
      "mousedown",
      handleClickOutside
    );

    return () => {
      document.removeEventListener(
        "mousedown",
        handleClickOutside
      );
    };
  }, []);

  /* ==============================
     HIDE NAVBAR
  ================================ */

  if (
    pathname === "/" ||
    pathname === "/login" ||
    pathname === "/superadmin" ||
    pathname === "/viewer"
  ) {
    return null;
  }

  /* ==============================
     MENUS
  ================================ */

  const menus = [
    {
      name: "Main summary",
      href: "/summary",
    },
    {
      name: "URL Fake Web",
      href: "/checker",
    },
    {
      name: "SSL / TLS Checker",
      href: "/ssl-checker",
    },
    {
      name: "CLAB",
      href: "/clab",
    },
  ];

  /* ==============================
     ROLE NAME
  ================================ */

  function getRoleName() {
    if (
      user?.role ===
      "super_admin"
    ) {
      return "Super Admin";
    }

    if (
      user?.role === "editor"
    ) {
      return "ผู้แก้ไขข้อมูล";
    }

    if (
      user?.role === "viewer"
    ) {
      return "ผู้ดูข้อมูล";
    }

    return "ผู้ใช้งาน";
  }

  /* ==============================
     LOGOUT
  ================================ */

  function handleLogout() {
    localStorage.removeItem(
      "user"
    );

    localStorage.removeItem(
      "token"
    );

    setUser(null);
    setMenuOpen(false);

    window.dispatchEvent(
      new Event(
        "session-started"
      )
    );

    router.replace("/login");
    router.refresh();
  }

  /* ==============================
     SUPER ADMIN
  ================================ */

  function handleSuperAdminPage() {
    setMenuOpen(false);

    router.push(
      "/superadmin"
    );
  }

  /* ==============================
     DISPLAY USER
  ================================ */

  const displayName =
    user?.name?.trim() ||
    user?.email ||
    "ผู้ใช้งาน";

  const firstLetter =
    displayName
      .charAt(0)
      .toUpperCase();

  const fontPercent =
    Math.round(
      fontScale * 100
    );

  return (
    <header className="navbar">

      {/* ==========================
          TOP
      =========================== */}

      <div className="navbar-top">

        <Link
          href="/summary"
          className="brand"
        >
          <img
            src="/ncsa-logo.png"
            alt="NCSA Logo"
            className="navbar-logo"
          />

          <div>
            <div className="brand-title">
              OP Web Service
            </div>

            <div className="brand-subtitle">
              Internal Operations Office
              Service System
            </div>
          </div>
        </Link>

        {user && (
          <div className="navbar-user-area">

            <span className="navbar-login-label">
              Logged in as
            </span>

            <strong className="navbar-user-name">
              {displayName}
            </strong>

            <div className="navbar-user-avatar">
              {firstLetter}
            </div>

            <span
              className={
                `navbar-role-badge role-${user.role}`
              }
            >
              {getRoleName()}
            </span>

            <div className="navbar-user-divider" />

            <button
              type="button"
              className="logout-button"
              onClick={handleLogout}
            >
              Log out
            </button>

            {user.role ===
              "super_admin" && (
              <div
                className="navbar-more-wrapper"
                ref={menuRef}
              >
                <button
                  type="button"
                  className="navbar-more-button"
                  onClick={() =>
                    setMenuOpen(
                      (current) =>
                        !current
                    )
                  }
                  aria-label="Open Super Admin menu"
                >
                  ☰
                </button>

                {menuOpen && (
                  <div className="navbar-dropdown">
                    <button
                      type="button"
                      onClick={
                        handleSuperAdminPage
                      }
                    >
                      Super Admin Management
                    </button>
                  </div>
                )}
              </div>
            )}

          </div>
        )}
      </div>

      {/* ==========================
          MENU
      =========================== */}

      <nav className="navbar-menu">

        <div className="navbar-menu-links">
          {menus.map(
            (menu) => (
              <Link
                key={menu.href}
                href={menu.href}
                className={
                  `navbar-link ${
                    pathname ===
                    menu.href
                      ? "active"
                      : ""
                  }`
                }
              >
                {menu.name}
              </Link>
            )
          )}
        </div>

        {/* ========================
            FONT SIZE CONTROL
        ========================= */}

        <div className="navbar-font-control">

          <span className="font-control-label">
            ขนาดตัวอักษร
          </span>

          <div className="font-control-buttons">

            <button
              type="button"
              className="font-control-button font-small"
              onClick={
                decreaseFontSize
              }
              disabled={
                fontScale <=
                MIN_FONT_SCALE
              }
              title="ลดขนาดตัวอักษร"
              aria-label="ลดขนาดตัวอักษร"
            >
              A−
            </button>

            <button
              type="button"
              className="font-control-button font-reset"
              onClick={
                resetFontSize
              }
              title="กลับขนาดปกติ"
              aria-label="กลับขนาดตัวอักษรปกติ"
            >
              A
            </button>

            <button
              type="button"
              className="font-control-button font-large"
              onClick={
                increaseFontSize
              }
              disabled={
                fontScale >=
                MAX_FONT_SCALE
              }
              title="เพิ่มขนาดตัวอักษร"
              aria-label="เพิ่มขนาดตัวอักษร"
            >
              A+
            </button>

          </div>

          <span className="font-scale-value">
            {fontPercent}%
          </span>

        </div>

      </nav>

    </header>
  );
}
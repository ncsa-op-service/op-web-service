"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import "./Navbar.css";

export default function Navbar() {
  const pathname = usePathname();

  // ไม่แสดง Navbar ในหน้า Home และ Login
  if (pathname === "/" || pathname === "/login") {
    return null;
  }

  const menus = [
    { name: "Main summary", href: "/summary" },
    { name: "URL Fake Web", href: "/checker" },
    { name: "SSL / TLS Checker", href: "/ssl-checker" },
    { name: "CLAB", href: "/clab" },
  ];

  return (
    <header className="navbar">
      <div className="navbar-top">
        <Link href="/summary" className="brand">
          <img
            src="/ncsa-logo.png"
            alt="NCSA Logo"
            className="navbar-logo"
          />

          <div>
            <div className="brand-title">OP Web Service</div>
            <div className="brand-subtitle">
              Internal Operations Office Service System
            </div>
          </div>
        </Link>

        <button
          className="logout-button"
          onClick={() => {
            window.location.href = "/login";
          }}
        >
          Log out
        </button>
      </div>

      <nav className="navbar-menu">
        {menus.map((menu) => (
          <Link
            key={menu.href}
            href={menu.href}
            className={`navbar-link ${
              pathname === menu.href ? "active" : ""
            }`}
          >
            {menu.name}
          </Link>
        ))}
      </nav>
    </header>
  );
}
import type { Metadata } from "next";
import "./globals.css";
import Navbar from "./components/Navbar";
import SessionHeartbeat from "./superadmin/SessionHeartbeat";

export const metadata: Metadata = {
  title: "OP Web Service",
  description: "Internal Operations Office Service System",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="th">
      <body>
        <SessionHeartbeat />
        <Navbar />
        <main>{children}</main>
      </body>
    </html>
  );
}

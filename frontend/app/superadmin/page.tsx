"use client";

import { useEffect, useMemo, useState } from "react";
import type { FormEvent, ReactNode } from "react";
import "./superadmin.css";
import "./session-status.css";

type Role = "super_admin" | "editor" | "viewer";

type User = {
  id: number;
  name: string;
  email: string;
  role: Role;
  created_at: string;
  updated_at?: string;
};

type Visit = {
  visitNumber: number;
  startedAt: string;
  lastSeenAt: string;
  logoutAt: string | null;
  ipAddress: string | null;
  mergedLogins: number;
  isOnline: boolean;
  isIdle: boolean;
};

type PresenceAccount = {
  id: number;
  name: string;
  email: string;
  role: Role;
  visit_count: number;
  last_seen_at: string | null;
  is_online: boolean;
  is_idle: boolean;
  visits: Visit[];
};

type AddAdminForm = {
  name: string;
  email: string;
  password: string;
  confirmPassword: string;
  role: Role | "";
};

type IconName = "home" | "link" | "shield" | "grid" | "users" | "activity" | "plus" | "refresh" | "edit" | "trash" | "logout" | "close" | "chevron";

const iconPaths: Record<IconName, ReactNode> = {
  home: <><path d="m3 11 9-8 9 8" /><path d="M5 10v10h14V10" /><path d="M9 20v-6h6v6" /></>,
  link: <><path d="M10 13a5 5 0 0 0 7.1.1l2-2a5 5 0 0 0-7.1-7.1l-1.1 1.1" /><path d="M14 11a5 5 0 0 0-7.1-.1l-2 2A5 5 0 0 0 12 20l1.1-1.1" /></>,
  shield: <><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z" /><path d="m9 12 2 2 4-4" /></>,
  grid: <><rect x="3" y="3" width="7" height="7" rx="1" /><rect x="14" y="3" width="7" height="7" rx="1" /><rect x="3" y="14" width="7" height="7" rx="1" /><rect x="14" y="14" width="7" height="7" rx="1" /></>,
  users: <><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M22 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" /></>,
  activity: <path d="M3 12h4l2.5-7 5 14 2.5-7h4" />,
  plus: <path d="M12 5v14M5 12h14" />,
  refresh: <><path d="M20 11a8 8 0 1 0-2.3 5.7" /><path d="M20 4v7h-7" /></>,
  edit: <><path d="M12 20h9" /><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4Z" /></>,
  trash: <><path d="M3 6h18" /><path d="M8 6V4h8v2" /><path d="m19 6-1 15H6L5 6" /><path d="M10 11v5M14 11v5" /></>,
  logout: <><path d="M10 17l5-5-5-5" /><path d="M15 12H3" /><path d="M15 3h5a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1h-5" /></>,
  close: <path d="M6 6l12 12M18 6 6 18" />,
  chevron: <path d="m9 18 6-6-6-6" />,
};

function Icon({ name, size = 18 }: { name: IconName; size?: number }) {
  return <svg className="sa-icon" width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{iconPaths[name]}</svg>;
}

const emptyForm: AddAdminForm = { name: "", email: "", password: "", confirmPassword: "", role: "" };

function getApiUrl(): string {
  const value = process.env.NEXT_PUBLIC_API_URL?.trim();
  if (!value || value.includes("backend:")) return "http://localhost:4000";
  return value.replace(/\/+$/, "");
}

function authHeaders(): HeadersInit {
  const token = typeof window === "undefined" ? null : localStorage.getItem("token");
  return token ? { Authorization: `Bearer ${token}` } : {};
}

export default function SuperAdminPage() {
  const [users, setUsers] = useState<User[]>([]);
  const [accounts, setAccounts] = useState<PresenceAccount[]>([]);
  const [loading, setLoading] = useState(true);
  const [overviewLoading, setOverviewLoading] = useState(false);
  const [activeTab, setActiveTab] = useState<"manage" | "activity">("manage");
  const [expandedAccounts, setExpandedAccounts] = useState<Record<number, boolean>>({});
  const [showAddModal, setShowAddModal] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [formError, setFormError] = useState("");
  const [form, setForm] = useState<AddAdminForm>(emptyForm);
  const apiUrl = getApiUrl();

  async function loadUsers() {
    try {
      setLoading(true);
      const response = await fetch(`${apiUrl}/api/users`, { cache: "no-store" });
      if (!response.ok) throw new Error("โหลดข้อมูลผู้ใช้ไม่สำเร็จ");
      const data = await response.json();
      setUsers(data.users ?? []);
    } catch (error) {
      console.error("Load users error:", error);
    } finally {
      setLoading(false);
    }
  }

  async function loadOverview(showLoading = true) {
    try {
      if (showLoading) setOverviewLoading(true);
      const response = await fetch(`${apiUrl}/api/sessions/overview`, {
        headers: authHeaders(),
        cache: "no-store",
      });
      if (response.status === 401) {
        window.location.href = "/login";
        return;
      }
      if (!response.ok) throw new Error("โหลดสถานะผู้ใช้ไม่สำเร็จ");
      const data = await response.json();
      setAccounts(data.accounts ?? []);
    } catch (error) {
      console.error("Load presence error:", error);
    } finally {
      if (showLoading) setOverviewLoading(false);
    }
  }

  useEffect(() => { void loadUsers(); }, []);

  useEffect(() => {
    if (activeTab !== "activity") return;
    void loadOverview();
    const timer = window.setInterval(() => void loadOverview(false), 30_000);
    return () => window.clearInterval(timer);
  }, [activeTab]);

  const stats = useMemo(() => ({
    total: users.length,
    editors: users.filter((user) => user.role === "editor").length,
    viewers: users.filter((user) => user.role === "viewer").length,
  }), [users]);

  function formatRole(role: Role) {
    if (role === "super_admin") return "Super Admin";
    if (role === "editor") return "Editor";
    return "View-only";
  }

  function formatDate(date: string | null) {
    if (!date) return "—";
    return new Date(date).toLocaleString("th-TH", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
  }

  function presence(account: PresenceAccount) {
    if (account.is_online) return { label: "Online", className: "online" };
    if (account.is_idle) return { label: "Idle", className: "idle" };
    return { label: "Offline", className: "offline" };
  }

  async function handleCreateAdmin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setFormError("");
    if (!form.name || !form.email || !form.password || !form.confirmPassword || !form.role) return setFormError("กรุณากรอกข้อมูลให้ครบ");
    if (form.password !== form.confirmPassword) return setFormError("Password และ Confirm Password ไม่ตรงกัน");

    try {
      setSubmitting(true);
      const response = await fetch(`${apiUrl}/api/users`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: form.name, email: form.email, password: form.password, role: form.role }),
      });
      const data = await response.json();
      if (!response.ok) return setFormError(data.message ?? "เพิ่มผู้ใช้ไม่สำเร็จ");
      setUsers((current) => [...current, data.user]);
      setShowAddModal(false);
      setForm(emptyForm);
    } catch (error) {
      console.error(error);
      setFormError("ไม่สามารถเชื่อมต่อ Backend ได้");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleDelete(user: User) {
    if (!window.confirm(`ต้องการลบผู้ใช้ “${user.name}” หรือไม่?`)) return;
    const response = await fetch(`${apiUrl}/api/users/${user.id}`, { method: "DELETE" });
    const data = await response.json().catch(() => null);
    if (!response.ok) return alert(data?.message ?? "ลบผู้ใช้ไม่สำเร็จ");
    setUsers((current) => current.filter((item) => item.id !== user.id));
  }

  async function handleLogout() {
    const token = localStorage.getItem("token");
    if (token) {
      await fetch(`${apiUrl}/api/sessions/logout`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}` },
        keepalive: true,
      }).catch(() => undefined);
    }
    localStorage.removeItem("token");
    localStorage.removeItem("user");
    window.location.href = "/login";
  }

  return (
    <main className="sa-layout">
      <aside className="sa-sidebar">
        <div>
          <a className="sa-brand" href="/summary"><span className="sa-brand-mark">OP</span><span className="sa-brand-copy"><strong>OP Web Service</strong></span></a>
          <nav className="sa-nav">
            <a href="/summary"><Icon name="home" /><span>Main Summary</span></a>
            <a href="/checker"><Icon name="link" /><span>URL Fake Web</span></a>
            <a href="/ssl-checker"><Icon name="shield" /><span>SSL / TLS Checker</span></a>
            <a href="/clab"><Icon name="grid" /><span>CLAB</span></a>
            <a href="/superadmin" className="active"><Icon name="users" /><span>Super Admin</span></a>
          </nav>
        </div>
      </aside>

      <section className="sa-main">
        <header className="sa-topbar"><span /><div className="sa-topbar-actions"><div className="sa-top-user"><div className="sa-top-avatar">SA</div><span><strong>Super Admin</strong><small>สกมช. (NCSA)</small></span></div><button type="button" className="sa-logout" onClick={handleLogout}><Icon name="logout" />Log out</button></div></header>

        <div className="sa-content">
          <div className="sa-heading-row"><div className="sa-heading"><h1>User Management</h1><p>Manage system administrators and user access</p></div>{activeTab === "manage" && <button type="button" className="sa-primary-button" onClick={() => setShowAddModal(true)}><Icon name="plus" />Add administrator</button>}</div>

          <section className="sa-stats">
            <article className="sa-stat-total"><div><small>Total accounts</small><strong>{stats.total}</strong></div><span className="sa-stat-icon"><Icon name="users" size={30} /></span></article>
            <article className="sa-stat-editors"><div><small>Editors</small><strong>{stats.editors}</strong></div><span className="sa-stat-icon"><Icon name="edit" size={30} /></span></article>
            <article className="sa-stat-viewers"><div><small>View-only users</small><strong>{stats.viewers}</strong></div><span className="sa-stat-icon"><Icon name="shield" size={30} /></span></article>
          </section>

          <div className="sa-tabs">
            <button type="button" className={activeTab === "manage" ? "active" : ""} onClick={() => setActiveTab("manage")}>Administrators</button>
            <button type="button" className={activeTab === "activity" ? "active" : ""} onClick={() => setActiveTab("activity")}>Activity log</button>
          </div>

          {activeTab === "manage" ? (
            <section className="sa-card">
              {loading ? <div className="sa-state"><span className="sa-spinner" />Loading users...</div> : (
                <><div className="sa-table-wrapper"><table className="sa-table"><thead><tr><th>User</th><th>Email</th><th>Role</th><th>Created date</th><th className="sa-align-right">Actions</th></tr></thead><tbody>{users.map((user) => <tr key={user.id}><td><div className="sa-user-cell"><span className={`sa-user-avatar sa-user-avatar-${user.role}`}>{user.name.slice(0, 2).toUpperCase()}</span><strong>{user.name}</strong></div></td><td>{user.email}</td><td><span className={`sa-role sa-role-${user.role}`}>{formatRole(user.role)}</span></td><td className="sa-date">{formatDate(user.created_at)}</td><td><div className="sa-manage"><button type="button" className="sa-icon-button" onClick={() => alert(`Edit user: ${user.name}`)}><Icon name="edit" /></button><button type="button" className="sa-icon-button danger" onClick={() => void handleDelete(user)}><Icon name="trash" /></button></div></td></tr>)}</tbody></table></div><div className="sa-table-footer">Showing {users.length} results</div></>
              )}
            </section>
          ) : (
            <section className="sa-card sa-presence-card">
              <div className="sa-presence-header"><div><h2>User sessions</h2><p>เรียงตามจำนวนรอบเข้าใช้งานจากมากไปน้อย พร้อมสถานะปัจจุบัน</p></div><button type="button" className="sa-secondary-button" onClick={() => void loadOverview()}><Icon name="refresh" />Refresh</button></div>
              {overviewLoading ? <div className="sa-state"><span className="sa-spinner" />Loading sessions...</div> : accounts.length === 0 ? <div className="sa-state">ยังไม่มีข้อมูล Session</div> : <div className="sa-session-list">{accounts.map((account) => {
                const state = presence(account);
                const expanded = Boolean(expandedAccounts[account.id]);
                return <article className="sa-account" key={account.id}>
                  <button type="button" className="sa-account-summary" onClick={() => setExpandedAccounts((current) => ({ ...current, [account.id]: !current[account.id] }))}>
                    <span className={`sa-account-avatar sa-account-avatar-${account.role}`}>{account.name.slice(0, 2).toUpperCase()}</span>
                    <span className="sa-account-identity"><strong>{account.name}</strong><small>{account.email}</small></span>
                    <span className={`sa-presence ${state.className}`}><i />{state.label}</span>
                    <span className="sa-visit-count"><strong>{account.visit_count}</strong><small>รอบเข้าใช้งาน</small></span>
                    <span className="sa-last-seen"><small>ใช้งานล่าสุด</small><strong>{formatDate(account.last_seen_at)}</strong></span>
                    <span className={`sa-expand-icon ${expanded ? "expanded" : ""}`}><Icon name="chevron" /></span>
                  </button>
                  {expanded && <div className="sa-visits"><div className="sa-visit-head"><span>รอบ</span><span>เวลาเริ่ม</span><span>ล่าสุด/สิ้นสุด</span><span>IP address</span><span>สถานะ</span></div>{account.visits.map((visit, index) => {
                    const visitState = visit.isOnline ? "online" : visit.isIdle ? "idle" : "offline";
                    const visitLabel = visit.isOnline ? "Online" : visit.isIdle ? "Idle" : "Ended";
                    return <div className="sa-visit-row" key={`${account.id}-${visit.visitNumber}`}><span>#{account.visits.length - index}{visit.mergedLogins > 1 && <small>รวม {visit.mergedLogins} login</small>}</span><span>{formatDate(visit.startedAt)}</span><span>{formatDate(visit.logoutAt ?? visit.lastSeenAt)}</span><span className="sa-mono">{visit.ipAddress ?? "—"}</span><span className={`sa-presence compact ${visitState}`}><i />{visitLabel}</span></div>;
                  })}</div>}
                </article>;
              })}</div>}
            </section>
          )}
        </div>
      </section>

      {showAddModal && <div className="sa-modal-overlay" onMouseDown={() => !submitting && setShowAddModal(false)}><div className="sa-modal" onMouseDown={(event) => event.stopPropagation()}><div className="sa-modal-header"><div><span className="sa-modal-icon"><Icon name="users" /></span><div><h2>Add administrator</h2><p>Create a new account and assign access.</p></div></div><button type="button" onClick={() => setShowAddModal(false)}><Icon name="close" /></button></div><form className="sa-modal-form" onSubmit={handleCreateAdmin}><div className="sa-form-grid"><label className="sa-field"><span>Name</span><input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></label><label className="sa-field"><span>Email</span><input type="email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></label><label className="sa-field"><span>Password</span><input type="password" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} /></label><label className="sa-field"><span>Confirm password</span><input type="password" value={form.confirmPassword} onChange={(e) => setForm({ ...form, confirmPassword: e.target.value })} /></label><label className="sa-field sa-field-full"><span>Role</span><select value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value as Role | "" })}><option value="">Select access level</option><option value="super_admin">Super Admin</option><option value="editor">Editor</option><option value="viewer">View-only</option></select></label></div>{formError && <p className="sa-form-error">{formError}</p>}<div className="sa-modal-actions"><button type="button" className="sa-cancel" onClick={() => setShowAddModal(false)} disabled={submitting}>Cancel</button><button type="submit" className="sa-create" disabled={submitting}>{submitting ? "Creating..." : "Create account"}</button></div></form></div></div>}
    </main>
  );
}

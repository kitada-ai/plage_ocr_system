"use client";

import { useState } from "react";

export default function LoginPage() {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError(null);

    try {
      const res = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });

      if (!res.ok) {
        setError("ログインに失敗しました");
        return;
      }

      window.location.href = "/";
    } catch {
      setError("ログインに失敗しました");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f5f5f5" }}>
      <form onSubmit={onSubmit} style={{ width: 320, padding: 24, background: "#fff", borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,0.08)" }}>
        <h1 style={{ margin: "0 0 16px", fontSize: 20 }}>ログイン</h1>
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12 }}>
            ユーザー名
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              style={{ padding: "8px 10px", border: "1px solid #ddd", borderRadius: 6 }}
              required
            />
          </label>
          <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12 }}>
            パスワード
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              style={{ padding: "8px 10px", border: "1px solid #ddd", borderRadius: 6 }}
              required
            />
          </label>
          {error && <div style={{ color: "#c00", fontSize: 12 }}>{error}</div>}
          <button
            type="submit"
            disabled={loading}
            className="login-button"
            style={{
              marginTop: 4,
              padding: "10px 12px",
              background: "#111",
              color: "#fff",
              border: "1px solid #111",
              borderRadius: 6,
              cursor: "pointer",
              opacity: loading ? 0.7 : 1,
              boxShadow: "0 2px 0 rgba(0,0,0,0.6), 0 8px 18px rgba(0,0,0,0.12)",
              transform: "translateY(0)",
              transition: "transform 120ms ease, box-shadow 120ms ease, background 120ms ease",
            }}
          >
            {loading ? "送信中..." : "ログイン"}
          </button>
        </div>
      </form>
      <style jsx>{`
        .login-button:hover:not(:disabled) {
          background: #0d0d0d;
          box-shadow: 0 2px 0 rgba(0, 0, 0, 0.6), 0 10px 22px rgba(0, 0, 0, 0.14);
        }
        .login-button:active:not(:disabled) {
          transform: translateY(2px);
          box-shadow: 0 0 0 rgba(0, 0, 0, 0.6), 0 5px 12px rgba(0, 0, 0, 0.16);
        }
        .login-button:focus-visible {
          outline: 2px solid rgba(17, 17, 17, 0.35);
          outline-offset: 2px;
        }
      `}</style>
    </div>
  );
}

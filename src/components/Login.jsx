import React, { useState } from "react";
import { supabase } from "../lib/supabaseClient";

const NAVY = "#1F3864";
const GREEN = "#5FBF8F";

export default function Login() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    const { error: signInError } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (signInError) {
      setError(signInError.message);
    }
    // On success, the onAuthStateChange listener in App.jsx picks up the
    // new session automatically — nothing else to do here.
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#F7F6F1", fontFamily: "Arial, sans-serif" }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", padding: 36, borderRadius: 12, width: 340, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
        <p style={{ color: GREEN, fontWeight: 700, fontSize: 18, margin: "0 0 4px" }}>Fleet Tracker</p>
        <p style={{ color: "#5F5E5A", fontSize: 13, margin: "0 0 24px" }}>Sign in to continue</p>

        <label style={{ display: "block", fontSize: 12.5, fontWeight: 600, color: "#2C2C2A", marginBottom: 4 }}>Email</label>
        <input
          type="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          required
          style={{ width: "100%", padding: "9px 11px", fontSize: 14, border: "1px solid #D3D1C7", borderRadius: 8, marginBottom: 14, boxSizing: "border-box" }}
        />

        <label style={{ display: "block", fontSize: 12.5, fontWeight: 600, color: "#2C2C2A", marginBottom: 4 }}>Password</label>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          required
          style={{ width: "100%", padding: "9px 11px", fontSize: 14, border: "1px solid #D3D1C7", borderRadius: 8, marginBottom: 18, boxSizing: "border-box" }}
        />

        {error && (
          <p style={{ color: "#C0392B", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>
        )}

        <button
          type="submit"
          disabled={loading}
          style={{ width: "100%", padding: "10px 0", background: NAVY, color: "#fff", border: "none", borderRadius: 8, fontSize: 14, fontWeight: 700, cursor: loading ? "default" : "pointer", opacity: loading ? 0.7 : 1 }}
        >
          {loading ? "Signing in…" : "Sign in"}
        </button>

        <p style={{ fontSize: 11.5, color: "#898781", marginTop: 18, marginBottom: 0 }}>
          Accounts are created by your administrator — there's no self-signup here.
        </p>
      </form>
    </div>
  );
}

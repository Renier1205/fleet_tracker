import React from "react";
import { supabase } from "../lib/supabaseClient";

const NAVY = "#1F3864";

export default function AccessSuspended() {
  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#F7F6F1", fontFamily: "Arial, sans-serif", textAlign: "center", padding: 24 }}>
      <div style={{ background: "#fff", padding: 36, borderRadius: 12, maxWidth: 380, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
        <p style={{ color: NAVY, fontWeight: 700, fontSize: 17, margin: "0 0 10px" }}>Access suspended</p>
        <p style={{ color: "#5F5E5A", fontSize: 13.5, margin: "0 0 20px", lineHeight: 1.5 }}>
          This system's licence is currently inactive. Please contact Datavera Analytics to resolve this.
        </p>
        <button
          onClick={() => supabase.auth.signOut()}
          style={{ background: "none", border: `1px solid ${NAVY}`, color: NAVY, padding: "8px 16px", borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
        >
          Sign out
        </button>
      </div>
    </div>
  );
}

import React, { useEffect, useState } from "react";
import { supabase } from "./lib/supabaseClient";
import Login from "./components/Login";
import AccessSuspended from "./components/AccessSuspended";
import EngineeringApp from "./EngineeringApp";

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = still checking
  const [licenseActive, setLicenseActive] = useState(undefined);
  const [isAdmin, setIsAdmin] = useState(false);
  const [myRole, setMyRole] = useState("manager");
  const [mySites, setMySites] = useState(undefined);
  const [myPageAccess, setMyPageAccess] = useState([]); // empty = no restriction beyond role
  const [myFullName, setMyFullName] = useState(undefined); // undefined = still checking, null = needs to be set

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));

    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });

    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) {
      setLicenseActive(undefined);
      return;
    }
    supabase
      .rpc("is_license_active")
      .then(({ data, error }) => {
        if (error) {
          console.error("License check failed:", error.message);
          setLicenseActive(false);
        } else {
          setLicenseActive(!!data);
        }
      });
  }, [session]);

  useEffect(() => {
    if (!session) {
      setIsAdmin(false);
      return;
    }
    supabase
      .rpc("is_admin")
      .then(({ data, error }) => {
        if (error) {
          console.error("Admin check failed:", error.message);
          setIsAdmin(false);
        } else {
          setIsAdmin(!!data);
        }
      });
  }, [session]);

  useEffect(() => {
    if (!session) {
      setMyRole("manager");
      return;
    }
    supabase
      .rpc("get_my_role")
      .then(({ data, error }) => {
        if (error) {
          console.error("Role check failed:", error.message);
          setMyRole("manager"); // fail permissive, not locked-out
        } else {
          setMyRole(data || "manager");
        }
      });
  }, [session]);

  useEffect(() => {
    if (!session) {
      setMySites(undefined);
      return;
    }
    supabase
      .rpc("get_my_sites")
      .then(({ data, error }) => {
        if (error) {
          console.error("Site access check failed:", error.message);
          setMySites([]);
        } else {
          setMySites(data || []);
        }
      });
  }, [session]);

  useEffect(() => {
    if (!session) {
      setMyPageAccess([]);
      return;
    }
    supabase
      .from("user_page_access")
      .select("page_key")
      .eq("user_id", session.user.id)
      .then(({ data, error }) => {
        if (error) {
          console.error("Page access check failed:", error.message);
          setMyPageAccess([]); // fail open - role-based visibility still applies
        } else {
          setMyPageAccess((data || []).map((r) => r.page_key));
        }
      });
  }, [session]);

  const loadMyProfile = React.useCallback(() => {
    if (!session) {
      setMyFullName(undefined);
      return;
    }
    supabase
      .from("profiles")
      .select("full_name")
      .eq("user_id", session.user.id)
      .maybeSingle()
      .then(({ data, error }) => {
        if (error) {
          console.error("Profile check failed:", error.message);
          setMyFullName(null);
        } else {
          setMyFullName(data?.full_name || null);
        }
      });
  }, [session]);

  useEffect(() => { loadMyProfile(); }, [loadMyProfile]);

  // Still resolving the session on first load
  if (session === undefined) {
    return null;
  }

  if (!session) {
    return <Login />;
  }

  // Session exists but we haven't confirmed the licence yet
  if (licenseActive === undefined) {
    return null;
  }

  if (!licenseActive) {
    return <AccessSuspended />;
  }

  if (myFullName === undefined) {
    return null;
  }

  if (myFullName === null) {
    return <NamePrompt userEmail={session.user.email} onSaved={loadMyProfile} />;
  }

  if (mySites === undefined) {
    return null;
  }

  if (mySites.length === 0) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh", fontFamily: "-apple-system, sans-serif", padding: 24 }}>
        <div style={{ textAlign: "center", maxWidth: 380 }}>
          <p style={{ fontSize: 16, fontWeight: 700, color: "#1F6668", marginBottom: 8 }}>No site access yet</p>
          <p style={{ fontSize: 14, color: "#4B5659", marginBottom: 4 }}>Your account isn't linked to any site yet. Ask your administrator to grant you access.</p>
          <p style={{ fontSize: 12.5, color: "#859195", marginBottom: 20 }}>Signed in as {session.user.email}</p>
          <button
            onClick={() => supabase.auth.signOut()}
            style={{ background: "#fff", border: "1px solid #E2E6E3", color: "#183642", padding: "9px 18px", borderRadius: 8, fontSize: 13.5, fontWeight: 600, cursor: "pointer" }}
          >
            Sign out
          </button>
        </div>
      </div>
    );
  }

  return <EngineeringApp userEmail={session.user.email} isAdmin={isAdmin} myRole={myRole} mySites={mySites} myPageAccess={myPageAccess} myFullName={myFullName} onNameSaved={loadMyProfile} />;
}

function NamePrompt({ userEmail, onSaved }) {
  const [name, setName] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!name.trim()) return;
    setSaving(true);
    setError("");
    const { data: { user } } = await supabase.auth.getUser();
    const { error: err } = await supabase.from("profiles").upsert({ user_id: user.id, full_name: name.trim(), updated_at: new Date().toISOString() });
    setSaving(false);
    if (err) {
      setError(err.message);
    } else {
      onSaved();
    }
  };

  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh", fontFamily: "-apple-system, sans-serif", padding: 24, background: "#F7F8F6" }}>
      <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 12, padding: 32, width: 380, maxWidth: "100%", boxShadow: "0 1px 3px rgba(0,0,0,0.08)", boxSizing: "border-box" }}>
        <p style={{ fontSize: 16, fontWeight: 700, color: "#1F6668", margin: "0 0 6px" }}>What's your name?</p>
        <p style={{ fontSize: 13, color: "#4B5659", margin: "0 0 18px" }}>
          Used for the Audit Trail and anywhere the system shows who did something - your name, not your email ({userEmail}).
        </p>
        <input
          type="text"
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Full name and surname"
          autoFocus
          required
          style={{ width: "100%", padding: "10px 12px", fontSize: 16, border: "1px solid #E2E6E3", borderRadius: 8, boxSizing: "border-box", marginBottom: 16 }}
        />
        {error && <p style={{ color: "#B85450", fontSize: 12.5, margin: "0 0 14px" }}>{error}</p>}
        <button type="submit" disabled={saving || !name.trim()} style={{ width: "100%", padding: "11px 0", background: "#1F6668", color: "#fff", border: "none", borderRadius: 8, fontSize: 14, fontWeight: 700, cursor: saving ? "default" : "pointer", opacity: saving ? 0.7 : 1 }}>
          {saving ? "Saving…" : "Continue"}
        </button>
      </form>
    </div>
  );
}

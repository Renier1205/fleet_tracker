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

  // Banning someone (or suspending the license) in Supabase only blocks
  // future sign-ins - it doesn't revoke a session that's already open.
  // Without this, someone already working when they get banned would
  // keep working until their token happened to expire on its own. This
  // re-checks the real, current status directly from the database every
  // 20 seconds and force-signs-out the instant either one fails, rather
  // than trusting whatever the token they're already holding says.
  useEffect(() => {
    if (!session) return;
    const check = async () => {
      const [{ data: banned }, { data: licenseOk }] = await Promise.all([
        supabase.rpc("is_banned"),
        supabase.rpc("is_license_active"),
      ]);
      if (banned || licenseOk === false) {
        await supabase.auth.signOut();
      }
    };
    const interval = setInterval(check, 20000);
    return () => clearInterval(interval);
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

  // Only admins fill in names now - a non-admin with no name yet just
  // has to wait. An admin with no name yet still gets in, since they're
  // the one who has to set it (their own included) from the User Access
  // page - blocking them here would be a lock with no key.
  if (myFullName === null && !isAdmin) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh", fontFamily: "-apple-system, sans-serif", padding: 24, background: "#F7F8F6" }}>
        <div style={{ background: "#fff", borderRadius: 12, padding: 32, width: 380, maxWidth: "100%", boxShadow: "0 1px 3px rgba(0,0,0,0.08)", boxSizing: "border-box", textAlign: "center" }}>
          <p style={{ fontSize: 16, fontWeight: 700, color: "#1F6668", margin: "0 0 8px" }}>Almost there</p>
          <p style={{ fontSize: 13.5, color: "#4B5659", margin: "0 0 18px" }}>
            Your administrator still needs to set your name up before you can continue. Check back shortly, or let them know you're waiting.
          </p>
          <p style={{ fontSize: 12, color: "#859195", margin: "0 0 18px" }}>Signed in as {session.user.email}</p>
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

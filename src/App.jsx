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

  return <EngineeringApp userEmail={session.user.email} isAdmin={isAdmin} myRole={myRole} mySites={mySites} />;
}

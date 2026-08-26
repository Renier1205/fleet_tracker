import React, { useEffect, useState, useCallback } from "react";
import { supabase } from "./lib/supabaseClient";
import Login from "./components/Login";
import AccessSuspended from "./components/AccessSuspended";
import EngineeringApp from "./EngineeringApp";

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = still checking
  const [licenseActive, setLicenseActive] = useState(undefined);
  const [profileData, setProfileData] = useState(undefined); // undefined = still loading

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

  // Everything EngineeringApp needs about who's logged in - their role,
  // which sites they can see, whether they're an admin, which pages
  // they're restricted to, and their display name. Re-run via
  // onNameSaved whenever the user edits their own name, so the header
  // updates immediately without a full reload.
  const loadProfileData = useCallback(async () => {
    if (!session) return;
    const [roleRes, sitesRes, adminRes, pageAccessRes, profileRes] = await Promise.all([
      supabase.rpc("get_my_role"),
      supabase.rpc("get_my_sites"),
      supabase.rpc("is_admin"),
      supabase.from("user_page_access").select("page_key").eq("user_id", session.user.id),
      supabase.from("profiles").select("full_name").eq("user_id", session.user.id).maybeSingle(),
    ]);
    setProfileData({
      myRole: roleRes.data || "manager",
      mySites: sitesRes.data || [],
      isAdmin: !!adminRes.data,
      myPageAccess: (pageAccessRes.data || []).map((r) => r.page_key),
      myFullName: profileRes.data?.full_name || "",
    });
  }, [session]);

  useEffect(() => {
    if (session && licenseActive) loadProfileData();
  }, [session, licenseActive, loadProfileData]);

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

  // Licence confirmed but role/site/profile data still loading
  if (profileData === undefined) {
    return null;
  }

  return (
    <EngineeringApp
      userEmail={session.user.email}
      isAdmin={profileData.isAdmin}
      myRole={profileData.myRole}
      mySites={profileData.mySites}
      myPageAccess={profileData.myPageAccess}
      myFullName={profileData.myFullName}
      onNameSaved={loadProfileData}
    />
  );
}

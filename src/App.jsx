import React, { useEffect, useState } from "react";
import { supabase } from "./lib/supabaseClient";
import Login from "./components/Login";
import AccessSuspended from "./components/AccessSuspended";
import EngineeringApp from "./EngineeringApp";

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = still checking
  const [licenseActive, setLicenseActive] = useState(undefined);

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

  return <EngineeringApp />;
}

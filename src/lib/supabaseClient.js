import { createClient } from "@supabase/supabase-js";

// These come from your .env file (see .env.example) - never hardcode
// real values directly in this file, especially if this project ever
// ends up in a public repository.
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  console.error(
    "Missing Supabase environment variables. Copy .env.example to .env and fill in your Project URL and anon key from Supabase → Project Settings → API."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

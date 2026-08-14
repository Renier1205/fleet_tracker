# Fleet Equipment Tracker — Setup

## What this is

This is the real, deployable version of the app — as opposed to the single-file
preview you've been clicking around in chat. Same code, same design, but this
one actually connects to Supabase, requires login, and can be deployed to a
real URL.

## Run it locally

1. Install Node.js if you don't have it (nodejs.org — the LTS version).
2. In this folder, run:
   ```
   npm install
   ```
3. Copy `.env.example` to a new file called `.env`, and fill in your real
   Supabase Project URL and anon key (Supabase dashboard → Project Settings → API).
4. Run:
   ```
   npm run dev
   ```
5. Open the URL it gives you (usually http://localhost:5173). You should see
   the login screen.

## Before you can log in

You need at least one real user account. In Supabase: Authentication → Users →
Add user. Use that email/password to log in.

You also need the `licenses` row set to `active` (see `006_licensing.sql` —
this should already be done from setup, but if login works and you land on
an "Access suspended" screen, that's why).

## Deploy it for real

**Vercel (recommended, free at this scale):**
1. Push this folder to a GitHub repository (private is fine).
2. Go to vercel.com → New Project → import that repository.
3. In the project's Environment Variables settings, add `VITE_SUPABASE_URL`
   and `VITE_SUPABASE_ANON_KEY` with your real values (same as your `.env`).
4. Deploy. Vercel gives you a URL immediately; you can point a real domain
   at it later from the same settings page.

**Netlify** works the same way — import the repo, set the same two
environment variables, deploy.

Either way: never put your real Supabase values directly in code that goes
into a public GitHub repo. Environment variables (set in Vercel/Netlify's
dashboard, not in a committed file) are what keep them out of your repo
history while still being available to the running app.

## What's still mock data

Every screen still shows the same sample data as the chat preview — none of
it is wired to your real Supabase tables yet. That's the next phase: swapping
each tab's data source over to a real query, one at a time, verified against
real data before moving to the next.

## Project structure

```
src/
  main.jsx              - React entry point
  App.jsx                - checks login + licence status, shows the right screen
  EngineeringApp.jsx      - the main app (sidebar, all tabs) — same file as the chat preview
  lib/supabaseClient.js   - Supabase connection, reads from .env
  components/
    Login.jsx             - email/password sign-in screen
    AccessSuspended.jsx    - shown if the licence is inactive
```

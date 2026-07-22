# MoM Task Tracker

Replaces the old Sheet + reply-parsing MoM tracker with a web checklist.
You still send MoM emails and `@Name`-tag owners; Apps Script parses them into a
Google Sheet, then each owner gets **one permanent link** to a checklist they
can update themselves. You get a password-gated admin dashboard.

## How it works

1. `scanMoMEmails` scans your sent MoM emails for an `[Actions]` section and
   `@Name` tags, and writes new tasks to the `Tracker` sheet.
2. `notifyOwners` emails each owner their personal checklist link once their
   task is at least `NOTIFY_DELAY_DAYS` (default 3) days past the MoM date —
   a "welcome" email the first time, then a lighter "new items added" nudge
   afterwards. The link never changes. The delay is checked per-task inside
   the function itself, so it doesn't matter how often the trigger runs —
   e.g. running it daily just means newly-eligible tasks get picked up within
   a day of crossing the 3-day mark.
3. Owners open their link, check tasks off or mark them In Progress / Blocked /
   Revised Timeline — this calls the Apps Script Web App directly and updates
   the Sheet live.
4. `sendReminders` runs every 7 days and nags anyone not Done/Blocked (honoring
   revised timelines), linking to the same permanent checklist.
5. You open `admin.html`, enter your password, and see every owner's status in
   one table.

## One-time setup

### 1. Apps Script project
This can be either a standalone Apps Script project (script.google.com → New
Project) or one bound to a Sheet (Sheet → Extensions → Apps Script) — either
works, since the script is told which Sheet to use by ID rather than relying
on a container binding.

1. Create (or reuse) a Google Sheet to act as the database. Copy its ID out of
   the URL: `https://docs.google.com/spreadsheets/d/`**`THIS_PART`**`/edit`.
2. In the Apps Script editor, delete the default `Code.gs` content and paste in
   [`apps-script/Code.gs`](apps-script/Code.gs). Also update the manifest
   (Project Settings → "Show appsscript.json") with
   [`apps-script/appsscript.json`](apps-script/appsscript.json).
3. Apps Script's Run button always calls a function with zero arguments, so
   you can't run `setSpreadsheetId('...')` directly. Instead, find the `setup`
   function near the top of the file, replace its two placeholder strings
   with your Sheet ID and a password of your choosing, then select `setup` in
   the function dropdown and click Run once. This also runs `initializeSheets`
   (creates the `Tracker`, `Owners`, `Unmatched` sheets with headers) and
   `setAdminPassword` (stores it in Script Properties — never written to the
   Sheet or the public repo). Authorize the requested Gmail/Sheets scopes when
   prompted.
4. Open the **Owners** sheet and pre-fill `Name` + `Email` for everyone you
   might tag with `@Name` in a MoM. If `scanMoMEmails` sees a tag it can't
   match to a name here, it logs it to the `Unmatched` sheet instead of
   guessing — check that sheet periodically and add missing people.
5. Deploy → New deployment → type **Web app**. Execute as **Me**, who has
   access **Anyone**. Copy the deployment URL (ends in `/exec`).
6. Once you know your GitHub Pages URL (step 3 below), find the `setupSiteUrl`
   function, replace its placeholder with that URL, select `setupSiteUrl` in
   the dropdown, and click Run once — this is what gets embedded in the
   emailed links.
7. Triggers (clock icon in the left sidebar) → add three time-driven triggers:
   - `scanMoMEmails` — e.g. every few hours, or daily.
   - `notifyOwners` — daily, shortly after `scanMoMEmails`.
   - `sendReminders` — every 7 days.

### 2. Static site
1. Edit [`site/config.js`](site/config.js) and set `API_URL` to the Web App
   URL from step 1.5.

### 3. GitHub Pages
1. Create a new GitHub repo (public, unless you have GitHub Pro/Enterprise for
   private Pages) and push the contents of the `site/` folder to its root
   (or push the whole project and point Pages at `/site`).
2. Repo Settings → Pages → deploy from the branch/folder containing
   `index.html`.
3. Your owner checklist lives at `https://<you>.github.io/<repo>/` and the
   admin dashboard at `https://<you>.github.io/<repo>/admin.html`.
4. Go back to step 1.7 and set `setSiteBaseUrl` to this exact URL if you
   haven't already.

## Notes on security
- Per-owner tokens and the admin password are the only real secrets. Neither
  lives in the repo — tokens live in the `Owners` sheet, the admin password in
  Script Properties. `config.js`'s Web App URL is meant to be public; it's the
  API endpoint the site calls.
- The Web App runs "execute as Me," so it has your Gmail/Sheets permissions —
  don't share the deployment URL as if it were secret in itself, but there's
  no reason to publicize it either.
- Reads use GET, writes use POST with a `text/plain` body — this avoids CORS
  preflight requests, which Apps Script Web Apps can't handle (no `doOptions`).
- Concurrent checklist updates are serialized with `LockService` to avoid two
  people's edits clobbering each other in the Sheet.

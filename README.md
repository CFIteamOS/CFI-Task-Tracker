# MoM Task Tracker

Replaces the old Sheet + reply-parsing MoM tracker with a web checklist.
You still send MoM emails and `@Name`-tag owners; Apps Script parses them into a
Google Sheet, then each owner gets **one permanent link** to a checklist they
can update themselves. You get a password-gated admin dashboard.

## How it works

1. `scanMoMEmails` scans your sent MoM emails for an `[Actions]` section and
   `@Name` tags, and writes new tasks to the `Tracker` sheet.
2. `notifyOwners` emails each owner their personal checklist link — a
   "welcome" email the first time, then a lighter "new items added" nudge
   afterwards. The link never changes.
3. Owners open their link, check tasks off or mark them In Progress / Blocked /
   Revised Timeline — this calls the Apps Script Web App directly and updates
   the Sheet live.
4. `sendReminders` runs every 7 days and nags anyone not Done/Blocked (honoring
   revised timelines), linking to the same permanent checklist.
5. You open `admin.html`, enter your password, and see every owner's status in
   one table.

## One-time setup

### 1. Apps Script project
1. Go to a Google Sheet you want to use as the database (or create a new one).
2. Extensions → Apps Script. Delete the default `Code.gs` content and paste in
   [`apps-script/Code.gs`](apps-script/Code.gs). Also update the manifest
   (Project Settings → "Show appsscript.json") with
   [`apps-script/appsscript.json`](apps-script/appsscript.json).
3. In the Apps Script editor, run `initializeSheets` once (creates the
   `Tracker`, `Owners`, `Unmatched` sheets with headers). Authorize the
   requested Gmail/Sheets scopes when prompted.
4. Run `setAdminPassword('choose-a-strong-password')` once from the editor
   (select it in the function dropdown, then Run). This stores the password in
   Script Properties — it's never written to the Sheet or the public repo.
5. Open the **Owners** sheet and pre-fill `Name` + `Email` for everyone you
   might tag with `@Name` in a MoM. If `scanMoMEmails` sees a tag it can't
   match to a name here, it logs it to the `Unmatched` sheet instead of
   guessing — check that sheet periodically and add missing people.
6. Deploy → New deployment → type **Web app**. Execute as **Me**, who has
   access **Anyone**. Copy the deployment URL (ends in `/exec`).
7. Back in the editor, run `setSiteBaseUrl('https://<you>.github.io/<repo>/')`
   once you know your GitHub Pages URL (step 3 below) — this is what gets
   embedded in the emailed links.
8. Triggers (clock icon in the left sidebar) → add three time-driven triggers:
   - `scanMoMEmails` — e.g. every few hours, or daily.
   - `notifyOwners` — daily, shortly after `scanMoMEmails`.
   - `sendReminders` — every 7 days.

### 2. Static site
1. Edit [`site/config.js`](site/config.js) and set `API_URL` to the Web App
   URL from step 1.6.

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

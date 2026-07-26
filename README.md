# MoM Task Tracker

Replaces the old Sheet + reply-parsing MoM tracker with a web checklist.
MoM emails still get `@Name`-tagged action items; Apps Script parses them into
a Google Sheet, then each owner gets **one permanent link** to a checklist they
can update themselves. You get a password-gated admin dashboard.

## How it works

1. `scanMoMEmails` scans for MoM emails and parses their `[Actions]` section,
   writing new tasks to the `Tracker` sheet. Each bullet line looks like:
   ```
   [Actions]

   - [Category] Task text @Full Name @Another Person
   ```
   Any leading `[Category]` tag is left as plain text at the start of the
   task (nothing is parsed out of it). Tagging multiple people on one line
   creates a separate copy of that task on each of their checklists. Works
   with both a plainly typed `@Name` and Gmail's auto-inserted contact chip
   (`@Full Name <email@x.com>`).
   By default it searches your own Sent Mail (`in:sent subject:MoM`). If MoM
   emails instead arrive in your mailbox from a separate address (e.g. an
   automated `updates@yourcompany.com`), set that with `setMomSender` (see
   setup below) and it'll search `from:that-address subject:MoM` instead.
2. `notifyOwners` emails each owner their personal checklist link, once ever
   per owner, the first time they have a task at least `NOTIFY_DELAY_DAYS`
   (default 3) days past its MoM date. The link never changes. There's no
   separate "new items added" email after that — `sendReminders` already
   sweeps up anything still open regardless of whether it was ever announced,
   so a second notification would just be redundant. The delay is checked
   per-task inside the function itself, so it doesn't matter how often the
   trigger runs — e.g. running it daily just means newly-eligible tasks get
   picked up within a day of crossing the 3-day mark.
3. Owners open their link, check tasks off or mark them In Progress / Blocked /
   Revised Timeline — this calls the Apps Script Web App directly and updates
   the Sheet live.
4. `sendReminders` runs every 7 days and nags anyone not Done/Blocked (honoring
   revised timelines), linking to the same permanent checklist.
5. You open `admin.html`, enter your password, and see every owner's status in
   one table, filterable by all / not-done / blocked / revised timeline, plus
   an overall completion bar. From here you can:
   - Assign a new task to one or several existing owners at once (pick from
     the multi-select owner list) — creates one copy per person, each
     appearing on their checklist immediately and flowing through the same
     notification/reminder logic as MoM-parsed tasks.
   - Edit or delete any existing task (click Edit on a row to load it into
     the same form, restricted to a single owner while editing; changing the
     owner re-flags it as unnotified so they get
     a heads-up about the reassignment).
   - View (but not add) comments on any task.
6. Owners can also add their own tasks from their checklist page (no owner
   picker needed — it's always added under them), and can leave comments on
   any of their own tasks — a running, timestamped log, not a single
   overwritable note. Useful for context like "why is this blocked." Only the
   task's owner can add a comment; the admin dashboard can view but not add
   one. The full log lives in the `Comments` sheet; the `Tracker` sheet's own
   `Comments` column is just a synced summary for at-a-glance reading — feel
   free to drag that column next to `Task` (or anywhere else) in the Sheet UI,
   every function looks columns up by name, not position.

## Branding

Both pages look for `site/curefoods-logo.png` and show it in the header; if
that file doesn't exist yet, the header just shows the "Task Tracker" text on
its own (no broken image icon). Drop a PNG or SVG named exactly
`curefoods-logo.png` into the `site/` folder (GitHub web UI: Add file →
Upload files) and it'll pick it up automatically — no code change needed. If
you use an SVG instead, update the `src="curefoods-logo.png"` references in
`index.html`/`admin.html` to match the filename.

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
   function near the top of the file and fill in its placeholder strings:
   your Sheet ID, a password of your choosing, and (only if MoM emails arrive
   from a separate address rather than your own Sent Mail) that sender's
   email in `setMomSender`. Leave `setMomSender`'s placeholder untouched to
   keep searching your own Sent Mail instead. Then select `setup` in the
   function dropdown and click Run once. This also runs `initializeSheets`
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
1. Push this whole project to a GitHub repo (public, unless you have GitHub
   Pro/Enterprise for private Pages) — GitHub Pages' branch-deploy source can
   only be `/(root)` or `/docs`, not an arbitrary folder like `/site`, so we
   serve the whole repo from root rather than moving `site/` around.
2. Repo Settings → Pages → Source: **Deploy from a branch** → Branch **main**,
   folder **`/ (root)`** → Save.
3. Your owner checklist lives at `https://<you>.github.io/<repo>/site/index.html`
   and the admin dashboard at `https://<you>.github.io/<repo>/site/admin.html`
   (note the `/site/` segment, since Pages is serving the repo root).
4. Go back to step 1.6 and run `setupSiteUrl` with
   `https://<you>.github.io/<repo>/site/` (the base, so the emailed link
   becomes `.../site/index.html?token=...`).

## Migrating from an older tracker

[`apps-script/migrate-legacy-tracker.gs`](apps-script/migrate-legacy-tracker.gs)
is a one-time utility for bringing tasks over from a prior Sheet-based
tracker. It's a separate file, not part of the regular scan/notify/remind
flow — paste it into the same Apps Script project, run
`migrateLegacyTracker()` once, then `sendWelcomeToMigratedOwners()` once, and
you can delete the file afterward. See the comment at the top of that file
for exactly what it does and how it avoids duplicating anything the live
system has already picked up on its own.

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

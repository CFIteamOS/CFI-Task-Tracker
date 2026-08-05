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
   with a plainly typed `@Name`, Gmail's auto-inserted contact chip
   (`@Full Name <email@x.com>`), or a raw email address typed as the tag
   (`@person@domain.com` is read as one tag, not split at the embedded `@`) —
   `resolveOwner_` matches an email-shaped tag against the Owners sheet's
   Email column too, not just Name. A tag that starts with a digit (e.g.
   `@99` used as informal price shorthand, not a real mention) is ignored
   rather than logged to Unmatched as a fake person.
   By default it searches your own Sent Mail (`in:sent subject:MoM`). If MoM
   emails instead arrive in your mailbox from a separate address (e.g. an
   automated `updates@yourcompany.com`), set that with `setMomSender` (see
   setup below) and it'll search `from:that-address subject:MoM` instead.
2. `notifyOwners` emails each owner the same day they get a task (controlled
   by `NOTIFY_DELAY_DAYS`, default 0 — raise it if you'd rather batch up
   same-day MoM edits before emailing). The **first** task an owner ever gets
   triggers a welcome email with their permanent checklist link. Any task
   after that — from a new MoM, an admin assignment, or a reassignment —
   triggers a shorter "New tasks have been added!" nudge instead, listing
   just the new items, since they already have their link. It also backfills:
   anyone who already has tasks on file but was never actually welcomed (e.g.
   rows brought in by the legacy migration) gets caught and emailed their
   link on the next run, using their full current task list. Since Apps
   Script triggers aren't real-time, run this trigger frequently (e.g. every
   hour) if same-day/near-instant delivery matters.

   You can also trigger this manually, for specific people, straight from the
   Google Sheet: a **Task Tracker** menu appears in the Sheet's menu bar
   (added automatically whenever you open it) with a **Send welcome
   email...** item. It opens a dialog listing everyone in the `Owners`
   sheet — tick whoever you want (or "Select all") and it emails each of
   them their permanent link plus their current task list immediately,
   regardless of whether they were auto-welcomed already. Handy right after
   a bulk migration, or if someone lost their original email.
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
   - View (but not add or delete) comments on any task. The Comments column
     shows just the latest comment plus how long ago it landed (no author
     name), and any task with a comment added since your last visit to the
     dashboard gets a "New" badge, a highlighted row, and is counted in a
     banner + a "New comments" filter pill — tracked per browser (via
     localStorage), not per person, since there's no separate login per
     admin.
6. Owners can also add their own tasks from their checklist page (no owner
   picker needed — it's always added under them), and their checklist page
   shows their own completion progress bar. They can leave comments on any of
   their own tasks — a running, timestamped log, not a single overwritable
   note — using **bold** text (a Bold button wraps the selection, or type
   `**like this**` yourself) and real line breaks (the comment box is a
   textarea, so Enter just makes a new line). They can also delete their own
   comments. Only the task's owner can add or delete a comment; the admin
   dashboard can view but not add or delete one. The full log lives in the
   `Comments` sheet (each row has its own `CommentID`); the `Tracker` sheet's
   own `Comments` column is just a synced plain-text summary for at-a-glance
   reading — feel free to drag that column next to `Task` (or anywhere else)
   in the Sheet UI, every function looks columns up by name, not position.
   On both the owner checklist and the admin table, tasks are sorted with
   anything still in play first, `Done` below that, and `Blocked` at the very
   bottom (since a blocked task needs the least day-to-day attention right
   now).

## Branding

Both pages look for `site/curefoods-logo.png` and show it in the header; if
that file doesn't exist yet, the header just shows the "Task Tracker" text on
its own (no broken image icon). Drop a PNG or SVG named exactly
`curefoods-logo.png` into the `site/` folder (GitHub web UI: Add file →
Upload files) and it'll pick it up automatically — no code change needed. If
you use an SVG instead, update the `src="curefoods-logo.png"` references in
`index.html`/`admin.html` to match the filename.

## Running a pilot with a subset of people

The Owners sheet has a **Pilot** checkbox column. While anyone's box is
unchecked, `notifyOwners` and `sendReminders` silently skip them — no welcome
email, no nudge, no reminder — even though `scanMoMEmails` keeps filing their
tasks as normal in the background. Nothing is lost or needs re-doing later:
check their box whenever you're ready to include them, and the next
automatic run picks up wherever they'd have been.

To start a pilot: check the box for your chosen 5-10 people, then use the
Sheet's **Task Tracker → Send welcome email...** menu to welcome exactly
them right away (that manual path always ignores the Pilot flag — it's an
explicit pick, not the automatic path). Everyone else's data stays exactly as
it is, untouched, ready to switch on whenever you widen the rollout.

If you're adding the Pilot column to a sheet that already existed before this
feature, run `setupPilotCheckboxes()` once from the Apps Script editor so the
column renders as real clickable checkboxes instead of blank/TRUE-FALSE text.

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
   [`apps-script/appsscript.json`](apps-script/appsscript.json). Then add a
   new HTML file (File → New → HTML) named exactly `WelcomeDialog`, and paste
   in [`apps-script/WelcomeDialog.html`](apps-script/WelcomeDialog.html) — this
   is the dialog behind the Sheet's "Send welcome email" menu button (see
   step 2 above).
   **Note:** the "Task Tracker" menu only appears automatically if this is a
   Sheet-bound script (Sheet → Extensions → Apps Script), since `onOpen` is a
   simple trigger and those only auto-fire for bound scripts. If you're using
   a standalone project instead, you can still open the dialog by running
   `showWelcomeDialog` from the Apps Script editor directly.
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
   - `scanMoMEmails` — e.g. every hour, or every few hours.
   - `notifyOwners` — every hour (or as often as `scanMoMEmails`), shortly
     after it, so welcome/new-task emails go out same-day.
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

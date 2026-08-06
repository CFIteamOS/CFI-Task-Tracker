/**
 * MoM Task Tracker backend.
 *
 * Sheets used (created automatically by initializeSheets if missing):
 *   Tracker:  TaskID | Owner | OwnerEmail | Comments | Task | Meeting | MoM Date
 *             | Status | Revised Timeline Date | Reminder Count | Last Updated
 *             | Notified | SourceKey
 *   Owners:   Name | Email | Token | WelcomeSent
 *   Unmatched: raw @name tags scanMoMEmails couldn't resolve to an email, for manual fixup
 *   Comments: TaskID | Author | Text | Timestamp — the full running log per task
 *             (the Tracker sheet's "Comments" column is just a synced summary
 *             for at-a-glance reading; this sheet is the source of truth)
 *
 * Column order in the Tracker sheet can be freely rearranged by hand (e.g. to
 * move Comments before Task) — every function below looks up columns by name
 * from the sheet's actual header row, never by a fixed position.
 *
 * One-time setup (run once from the Apps Script editor):
 *   1. Create (or reuse) a Google Sheet to act as the database, and set its ID here:
 *      setSpreadsheetId('paste-the-sheet-id-from-its-url')
 *   2. initializeSheets()
 *   3. setAdminPassword('choose-a-password')
 *   4. Pre-populate the Owners sheet with Name + Email for everyone you tag in MoMs.
 *   5. Deploy > New deployment > Web app (execute as Me, access Anyone), copy the URL into site/config.js.
 *   6. Set time-driven triggers for scanMoMEmails, notifyOwners, sendReminders (Triggers panel).
 */

const TRACKER_SHEET = 'Tracker';
const OWNERS_SHEET = 'Owners';
const UNMATCHED_SHEET = 'Unmatched';
const COMMENTS_SHEET = 'Comments';

const NOTIFY_DELAY_DAYS = 0; // notify same-day; set higher if you want to batch up same-day MoM edits first

const STATUS = {
  PENDING: 'Pending',
  IN_PROGRESS: 'In Progress',
  BLOCKED: 'Blocked',
  REVISED: 'Revised Timeline',
  DONE: 'Done'
};

// Order here only matters for a brand-new sheet (this is the order columns
// get created in). On an existing sheet, new headers are appended wherever
// there's room, and you're free to drag columns around afterward.
const TRACKER_HEADERS = [
  'TaskID', 'Owner', 'OwnerEmail', 'Comments', 'Task', 'Meeting', 'MoM Date', 'Status',
  'Revised Timeline Date', 'Reminder Count', 'Last Updated', 'Notified', 'SourceKey'
];
// Pilot: checkbox — while a pilot is running, notifyOwners/sendReminders
// only ever email owners with this checked. Unchecked owners' tasks still
// get filed by scanMoMEmails as normal, just silently, until they're
// switched on (nothing to re-migrate or re-send when that happens).
const OWNERS_HEADERS = ['Name', 'Email', 'Token', 'WelcomeSent', 'Pilot'];
const UNMATCHED_HEADERS = ['Name Tag', 'Task', 'Meeting', 'MoM Date', 'Seen At'];
const COMMENTS_HEADERS = ['TaskID', 'Author', 'Text', 'Timestamp', 'CommentID'];

// ---------- setup ----------
//
// Apps Script's Run button always calls the selected function with zero
// arguments, so setSpreadsheetId('...')/setAdminPassword('...') can't be run
// directly from the dropdown. Fill in the placeholders below, select `setup`
// in the function dropdown, and click Run once.

function setup() {
  setSpreadsheetId('1ICqfy3hvX4yVeppcpTare6zz2QmaBE9pZ-1jZBICzb8');
  setAdminPassword('Curefoods11');
  setMomSender('PASTE_THE_MOM_SENDER_EMAIL_HERE');
  initializeSheets();
}

function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
}

// If MoM emails are sent by a separate address (e.g. an automated inbox) into
// your mailbox, set that address here and scanMoMEmails will search for mail
// *received* from it instead of mail you personally sent.
function setMomSender(email) {
  if (!email || email.indexOf('PASTE_') === 0) return; // ignore the untouched placeholder
  PropertiesService.getScriptProperties().setProperty('MOM_SENDER', email);
}

function getMomSearchQuery_() {
  const sender = PropertiesService.getScriptProperties().getProperty('MOM_SENDER');
  return sender ? `from:(${sender}) subject:MoM newer_than:2d` : 'in:sent subject:MoM newer_than:2d';
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error(
    "No spreadsheet configured. Run setSpreadsheetId('your-sheet-id') once from the editor " +
    '(the ID is the long string in the Sheet\'s URL between /d/ and /edit).'
  );
}

function initializeSheets() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  ensureSheet_(ss, UNMATCHED_SHEET, UNMATCHED_HEADERS);
  ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
}

// One-time setup — run this once from the editor to put scanMoMEmails,
// notifyOwners, and sendReminders on a real recurring schedule, instead of
// only running when you click Run yourself. Safe to run more than once:
// skips creating a trigger for a function that already has one, so it won't
// pile up duplicates.
function setupAutomationTriggers() {
  const existingHandlers = new Set(ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()));
  const created = [];
  const skipped = [];

  if (existingHandlers.has('scanMoMEmails')) {
    skipped.push('scanMoMEmails');
  } else {
    ScriptApp.newTrigger('scanMoMEmails').timeBased().everyHours(1).create();
    created.push('scanMoMEmails (every hour)');
  }

  if (existingHandlers.has('notifyOwners')) {
    skipped.push('notifyOwners');
  } else {
    ScriptApp.newTrigger('notifyOwners').timeBased().everyHours(1).create();
    created.push('notifyOwners (every hour)');
  }

  if (existingHandlers.has('sendReminders')) {
    skipped.push('sendReminders');
  } else {
    ScriptApp.newTrigger('sendReminders').timeBased().everyDays(7).atHour(9).create();
    created.push('sendReminders (every 7 days, ~9am)');
  }

  Logger.log(`Created: ${created.join(', ') || '(none)'}`);
  Logger.log(`Already had a trigger, left alone: ${skipped.join(', ') || '(none)'}`);
}

// One-time helper for an already-live sheet that just picked up the Pilot
// column: turns it into real clickable checkboxes instead of blank/TRUE-FALSE
// text cells. Safe to run more than once. Run it once from the editor after
// pulling in the Pilot column for the first time.
function setupPilotCheckboxes() {
  const ss = getSpreadsheet_();
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const pilotColIndex = OWNERS_HEADERS.indexOf('Pilot') + 1;
  const rowCount = Math.max(owners.getLastRow() - 1, 200); // cover existing rows, plus headroom for new owners
  owners.getRange(2, pilotColIndex, rowCount, 1).insertCheckboxes();
}

// Creates the sheet with the given headers if it doesn't exist. If it does
// exist but is missing headers this code now expects (e.g. after adding a new
// column), those are appended to the end of the existing header row so older
// live sheets pick up schema changes automatically. You can then freely drag
// that column wherever you want — nothing below assumes a fixed position.
function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = headers.filter(h => existingHeaders.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, existingHeaders.length + 1, 1, missing.length).setValues([missing]);
  }
  return sheet;
}

// Reads a sheet's actual current header row (row 1), so column lookups always
// reflect reality even if you've manually reordered or the sheet has extra
// columns beyond what this script created.
function getHeaderRow_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

// Builds a row array matching a sheet's actual header order from a
// name -> value map, so appendRow() always lands values in the right column
// regardless of how the sheet's columns happen to be arranged.
function buildRowByHeaders_(headerRow, fieldsByName) {
  return headerRow.map(header =>
    Object.prototype.hasOwnProperty.call(fieldsByName, header) ? fieldsByName[header] : ''
  );
}

// getRange() throws if asked for 0 rows, which happens whenever a sheet has
// only its header row (no data yet) — this guards that case. Column is
// looked up by name against the sheet's actual header row.
function getColumnValues_(sheet, columnName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const colIndex = getHeaderRow_(sheet).indexOf(columnName) + 1;
  if (colIndex === 0) return [];
  return sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat();
}

function setAdminPassword(password) {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', password);
}

// Temporary diagnostic — run this from the editor, then check View > Logs
// (Ctrl+Enter) for exactly what parseActionsSection_ saw and extracted from
// each matching MoM email, so a "missing action items" report can be
// root-caused against the real data instead of guesswork. Safe to delete
// once the issue's found.
function debugParseActions() {
  const threads = GmailApp.search(getMomSearchQuery_());
  Logger.log(`Search query: ${getMomSearchQuery_()}`);
  Logger.log(`Found ${threads.length} thread(s).`);

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const actions = parseActionsSection_(body);

      Logger.log('==================================================');
      Logger.log(`Subject: ${subject}`);
      Logger.log(`Date: ${message.getDate()}`);
      Logger.log(`--- Raw body (between [Actions] and end, or first 3000 chars) ---`);
      const actionsIdx = body.search(/\[actions\]/i);
      Logger.log(actionsIdx === -1 ? '(no [Actions] heading found at all)' : body.slice(actionsIdx, actionsIdx + 3000));
      Logger.log(`--- Parsed ${actions.length} action(s) ---`);
      actions.forEach((a, i) => Logger.log(`${i + 1}. @${a.nameTag} -> "${a.task}"`));
    });
  });
}

// One-time manual fixup for the Olio/Arambam batch that got misparsed before
// the @tag regex was fixed (split email addresses, digit-led price shorthand
// read as fake mentions). Run this once from the editor, then delete it —
// it's not part of the regular scan/notify/remind flow.
function fixupOlioArambamOwners_Aug2026() {
  const ss = getSpreadsheet_();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const headerRow = getHeaderRow_(tracker);

  addOwnerIfMissing_(ss, 'Sakshi Agrawal', 'sakshi.ag@curefoods.in');
  addOwnerIfMissing_(ss, 'Venkanna Godavarthi', 'venkanna@milletexpress.in');

  const olioMeeting = 'Olio App Launch | MoM | 31 Jul';
  const olioDate = new Date(2026, 6, 31);
  const arambamOwnlyMeeting = '99 Arambam stack selection on Ownly | MoM | 31 Jul';
  const arambamOwnlyDate = new Date(2026, 6, 31);
  const arambamRepeatMeeting = 'Scaling Arambam brand repeat | MoM | 30 Jul';
  const arambamRepeatDate = new Date(2026, 6, 30);

  const sakshi = { owner: 'Sakshi Agrawal', email: 'sakshi.ag@curefoods.in' };
  const venkanna = { owner: 'Venkanna Godavarthi', email: 'venkanna@milletexpress.in' };

  const rowsToAdd = [
    { ...sakshi, task: '[Olio] Add lock screen updates for order status on consumer phone', meeting: olioMeeting, momDate: olioDate },
    { ...sakshi, task: '[Olio] Remove "All" tab from menu', meeting: olioMeeting, momDate: olioDate },
    { ...sakshi, task: '[Olio] Simplify menu to show only a "Sides" tab, with item flow: Diet Coke → Desserts → Sides', meeting: olioMeeting, momDate: olioDate },

    { ...venkanna, task: '[Arambam] Ownly masala dosa', meeting: arambamOwnlyMeeting, momDate: arambamOwnlyDate },
    { ...venkanna, task: '[Arambam] Check whether competitors are selling veg biryani; Try veg biryani in ME', meeting: arambamOwnlyMeeting, momDate: arambamOwnlyDate },
    { ...venkanna, task: '[Arambam] Introduce Paneer 65 in Hyd', meeting: arambamOwnlyMeeting, momDate: arambamOwnlyDate },
    { ...venkanna, task: '[Arambam] Evaluate whether specific hyderabad biryani SKUs are required; Finalize exclusive Hyderabad menu and pricing for OWnly launch', meeting: arambamOwnlyMeeting, momDate: arambamOwnlyDate },
    { ...venkanna, task: '[General] Introduce a ₹99 price point across categories (rice bowls, noodles) on the Ownly platform wherever possible', meeting: arambamOwnlyMeeting, momDate: arambamOwnlyDate },

    { ...venkanna, task: '[Arambam] Region cut on whether the repeat items in Arambam are easily available in North & West', meeting: arambamRepeatMeeting, momDate: arambamRepeatDate },
    { ...venkanna, task: '[Arambam] Price increase for Arambam: ₹10 now + add ghee', meeting: arambamRepeatMeeting, momDate: arambamRepeatDate },
    { ...venkanna, task: '[All 3] New rice packaging across all 3 brands to be rolled out: Increase by ₹10', meeting: arambamRepeatMeeting, momDate: arambamRepeatDate },
    { ...venkanna, task: "[Arambam] Arambam SKUs to be renamed 'with ghee'", meeting: arambamRepeatMeeting, momDate: arambamRepeatDate },
    { ...venkanna, task: '[Arambam] Offer add-on ghee sachets (2 sachets for Rs 20)', meeting: arambamRepeatMeeting, momDate: arambamRepeatDate },
    { ...venkanna, task: '[Arambam] 4CP exclusivity to be checked', meeting: arambamRepeatMeeting, momDate: arambamRepeatDate }
  ];

  const newRows = rowsToAdd.map(r => buildRowByHeaders_(headerRow, {
    TaskID: generateTaskId_(tracker),
    Owner: r.owner,
    OwnerEmail: r.email,
    Task: r.task,
    Meeting: r.meeting,
    'MoM Date': r.momDate,
    Status: STATUS.PENDING,
    'Reminder Count': 0,
    'Last Updated': new Date(),
    Notified: false,
    SourceKey: `manual-fixup:${Utilities.getUuid()}`
  }));

  tracker.getRange(tracker.getLastRow() + 1, 1, newRows.length, headerRow.length).setValues(newRows);

  // Remove the now-resolved rows from Unmatched. Deliberately leaves the
  // "199" row alone (the Nomad/SuperYou task) — that one was only ever
  // tagged with a price reference, not a real person, so there's no correct
  // owner to assign it to automatically.
  const unmatched = ensureSheet_(ss, UNMATCHED_SHEET, UNMATCHED_HEADERS);
  const uData = unmatched.getDataRange().getValues();
  const nameTagCol = UNMATCHED_HEADERS.indexOf('Name Tag');
  const namesToRemove = new Set(['Sakshi Agrawal', '89 (Discount from 99 to 89)', 'venkanna', 'milletexpress.in']);
  for (let i = uData.length - 1; i >= 1; i--) {
    if (namesToRemove.has(String(uData[i][nameTagCol]))) {
      unmatched.deleteRow(i + 1);
    }
  }

  Logger.log(`Added ${newRows.length} tasks for Sakshi Agrawal and Venkanna Godavarthi. Cleaned up resolved Unmatched rows.`);
}

function addOwnerIfMissing_(ss, name, email) {
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const data = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol] || '').toLowerCase() === name.toLowerCase()) return;
  }
  const headerRow = getHeaderRow_(owners);
  owners.appendRow(buildRowByHeaders_(headerRow, { Name: name, Email: email }));
}

// ---------- step 1: scan sent MoM emails ----------

function scanMoMEmails() {
  const ss = getSpreadsheet_();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const headerRow = getHeaderRow_(tracker);
  const existingKeys = new Set(getColumnValues_(tracker, 'SourceKey').filter(String));

  const threads = GmailApp.search(getMomSearchQuery_());
  const newRows = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const subject = message.getSubject();
      const meetingTitle = subject.replace(/^\s*MoM\s*[:\-]?\s*/i, '').trim() || subject;
      const momDate = message.getDate();
      const body = message.getPlainBody();
      const actions = parseActionsSection_(body);

      actions.forEach((action, idx) => {
        const sourceKey = `${message.getId()}:${idx}`;
        if (existingKeys.has(sourceKey)) return;

        const owner = resolveOwner_(ss, action.nameTag);
        if (!owner) {
          logUnmatched_(ss, action.nameTag, action.task, meetingTitle, momDate);
          return;
        }

        newRows.push(buildRowByHeaders_(headerRow, {
          TaskID: generateTaskId_(tracker),
          Owner: owner.name,
          OwnerEmail: owner.email,
          Task: action.task,
          Meeting: meetingTitle,
          'MoM Date': momDate,
          Status: STATUS.PENDING,
          'Reminder Count': 0,
          'Last Updated': new Date(),
          Notified: false,
          SourceKey: sourceKey
        }));
        existingKeys.add(sourceKey);
      });
    });
  });

  if (newRows.length) {
    tracker.getRange(tracker.getLastRow() + 1, 1, newRows.length, headerRow.length).setValues(newRows);
  }
}

// Format (one bullet per line, under a "[Actions]" heading):
//   - [Category] Task text @Full Name @Another Person
// Any leading [Category] tag is left in place as part of the task text
// verbatim — it's just plain text here, not parsed out. Multiple @tags on
// one line means the same task goes to each of those people separately.
// Also handles Gmail's auto-inserted contact chip, which expands a typed
// "@Name" into "@Full Name <email@x.com>".
function parseActionsSection_(body) {
  const lines = body.split('\n');
  // Matches a line that's just the "[Actions]" heading, tolerating bold
  // markup (*[Actions]*, **[Actions]**) that Gmail's plain-text export adds
  // when the heading was bold in the original email, plus any stray
  // whitespace/punctuation around it.
  const startIdx = lines.findIndex(l => /^[\s*_]*\[actions\][\s*_:]*$/i.test(l));
  if (startIdx === -1) return [];

  // Long bullets get word-wrapped across multiple physical lines by Gmail's
  // plain-text export — a wrapped continuation line has no bullet marker of
  // its own, so it's collected here and merged onto the previous bullet
  // below, rather than being mistaken for the end of the Actions section.
  // The distinguishing signal: a real continuation is always immediately
  // adjacent (no blank line) to the bullet it wraps from. A non-bullet line
  // that follows a blank line is trailing prose (e.g. a "Thanks!" sign-off)
  // and means the list has ended — likewise the next "[Something]" heading.
  const rawLines = [];
  let sawBlank = false;
  for (let i = startIdx + 1; i < lines.length; i++) {
    const line = lines[i];
    if (/^\s*$/.test(line)) {
      sawBlank = true;
      continue;
    }
    if (/^[\s*_]*\[[^\]]+\][\s*_:]*$/.test(line)) break; // next section heading
    const isBullet = /^\s*[-*]\s+/.test(line);
    if (!isBullet && sawBlank) break; // non-bullet line after a blank = end of the list
    rawLines.push(line);
    sawBlank = false;
  }

  const actionLines = [];
  rawLines.forEach(line => {
    if (/^\s*[-*]\s+/.test(line)) {
      actionLines.push(line.replace(/^\s*[-*]\s+/, '').trim());
    } else if (actionLines.length) {
      actionLines[actionLines.length - 1] += ' ' + line.trim();
    }
  });

  const actions = [];
  actionLines.forEach(rest => {
    // A new tag only starts at an "@" preceded by whitespace (or the very
    // start) — this lets an embedded "@" survive as part of the SAME tag
    // when someone mentions a raw email address with no space before its
    // domain (e.g. "@venkanna@milletexpress.in" is one tag, not two).
    const tagPattern = /@([^<]+?)(?:\s*<[^>]*>)?(?=\s+@|\s*$)/g;
    const nameTags = [];
    let firstTagIndex = -1;
    let m;
    while ((m = tagPattern.exec(rest)) !== null) {
      if (firstTagIndex === -1) firstTagIndex = m.index;
      const tag = m[1].trim();
      // A real name never starts with a digit — this filters out informal
      // price/number shorthand like "@99" or "@199" that isn't a mention at
      // all, so it doesn't get logged to Unmatched as a fake person.
      if (/^\d/.test(tag)) continue;
      nameTags.push(tag);
    }
    if (!nameTags.length) return;

    const task = rest.slice(0, firstTagIndex).trim();
    if (!task) return;

    nameTags.forEach(nameTag => actions.push({ nameTag, task }));
  });

  return actions;
}

function resolveOwner_(ss, nameTag) {
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const data = owners.getDataRange().getValues();
  // Some MoMs tag people by raw email address instead of name (e.g. an
  // external contact like "@venkanna@milletexpress.in") — match against the
  // Owners' Email column too in that case, not just the Name column.
  const looksLikeEmail = nameTag.indexOf('@') !== -1;

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '');
    const email = String(data[i][1] || '');
    const matchesByName = name.toLowerCase() === nameTag.toLowerCase() ||
      name.toLowerCase().startsWith(nameTag.toLowerCase());
    const matchesByEmail = looksLikeEmail && email.toLowerCase() === nameTag.toLowerCase();
    if (!matchesByName && !matchesByEmail) continue;

    if (!email) return null;
    let token = data[i][2];
    if (!token) {
      token = Utilities.getUuid();
      owners.getRange(i + 1, 3).setValue(token);
    }
    return { name, email, token };
  }
  return null;
}

function logUnmatched_(ss, nameTag, task, meetingTitle, momDate) {
  const sheet = ensureSheet_(ss, UNMATCHED_SHEET, UNMATCHED_HEADERS);
  sheet.appendRow([nameTag, task, meetingTitle, momDate, new Date()]);
}

function generateTaskId_(tracker) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const existing = new Set(getColumnValues_(tracker, 'TaskID'));
  let id;
  do {
    id = Array.from({ length: 4 }, () => chars[Math.floor(Math.random() * chars.length)]).join('');
  } while (existing.has(id));
  return id;
}

// ---------- shared HTML email helpers ----------
//
// Plain-text emails have no hanging indent, so a long bullet that wraps onto
// a second line looks like a new, unindented line rather than a continuation
// of the same item. Sending an htmlBody alongside the plain-text body fixes
// this (a real <ul><li> wraps with proper indent) while keeping the plain
// text as a fallback for clients that don't render HTML.
function escapeHtml_(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function buildItemsListHtml_(items, formatItem) {
  const format = formatItem || (it => escapeHtml_(it.task));
  return '<ul style="margin:8px 0 16px; padding-left:20px;">' +
    items.map(it => `<li style="margin-bottom:6px;">${format(it)}</li>`).join('') +
    '</ul>';
}

// ---------- step 2: notify owners (welcome, or a "new tasks" nudge) ----------

function notifyOwners() {
  const ss = getSpreadsheet_();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const baseUrl = getSiteBaseUrl_();

  const trackerData = tracker.getDataRange().getValues();
  const col = name => trackerData[0].indexOf(name);
  const notifyCutoff = new Date(Date.now() - NOTIFY_DELAY_DAYS * 24 * 60 * 60 * 1000);

  const pendingByOwner = {};
  for (let i = 1; i < trackerData.length; i++) {
    if (trackerData[i][col('Notified')] === true) continue;
    if (new Date(trackerData[i][col('MoM Date')]) > notifyCutoff) continue; // not old enough yet, catch it on a later run
    const owner = trackerData[i][col('Owner')];
    if (!pendingByOwner[owner]) pendingByOwner[owner] = [];
    pendingByOwner[owner].push({ row: i + 1, task: trackerData[i][col('Task')] });
  }

  const ownersData = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  const emailCol = OWNERS_HEADERS.indexOf('Email');
  const tokenCol = OWNERS_HEADERS.indexOf('Token');
  const welcomeCol = OWNERS_HEADERS.indexOf('WelcomeSent');
  const pilotCol = OWNERS_HEADERS.indexOf('Pilot');
  const notifiedColIndex = col('Notified') + 1;

  // Backfill: anyone who has tasks on file but was never actually welcomed
  // (e.g. rows brought in by the legacy migration, or any that got marked
  // Notified without an email ever going out) gets swept up here too, using
  // their full current task list — not just the ones still flagged
  // unnotified — so nobody with real tasks is left without their link.
  for (let i = 1; i < ownersData.length; i++) {
    const name = ownersData[i][nameCol];
    if (!name || ownersData[i][welcomeCol] || pendingByOwner[name]) continue;
    const theirTasks = [];
    for (let r = 1; r < trackerData.length; r++) {
      if (trackerData[r][col('Owner')] === name) {
        theirTasks.push({ row: r + 1, task: trackerData[r][col('Task')] });
      }
    }
    if (theirTasks.length) pendingByOwner[name] = theirTasks;
  }

  Object.keys(pendingByOwner).forEach(ownerName => {
    for (let i = 1; i < ownersData.length; i++) {
      if (ownersData[i][nameCol] !== ownerName) continue;
      if (ownersData[i][pilotCol] !== true) break; // not in the pilot yet — leave their tasks queued for later

      const email = ownersData[i][emailCol];
      const token = ownersData[i][tokenCol];
      const welcomeSent = ownersData[i][welcomeCol];
      const link = `${baseUrl}?token=${token}`;
      const items = pendingByOwner[ownerName];
      if (!items.length) break;

      if (!welcomeSent) {
        MailApp.sendEmail({
          to: email,
          subject: 'Your action items checklist',
          body: `Hi ${ownerName},\n\nYou've been tagged with action items from a recent meeting. ` +
            `Bookmark this link — it always shows your current, live checklist:\n\n${link}\n\n` +
            `New items right now:\n${items.map(it => `- ${it.task}`).join('\n')}\n\n` +
            `Just tick things off (or mark them In Progress / Blocked / Revised Timeline) as you go.`,
          htmlBody: `<p>Hi ${escapeHtml_(ownerName)},</p>` +
            `<p>You've been tagged with action items from a recent meeting. ` +
            `Bookmark this link — it always shows your current, live checklist:</p>` +
            `<p><a href="${link}">${escapeHtml_(link)}</a></p>` +
            `<p>New items right now:</p>${buildItemsListHtml_(items)}` +
            `<p>Just tick things off (or mark them In Progress / Blocked / Revised Timeline) as you go.</p>`
        });
        owners.getRange(i + 1, welcomeCol + 1).setValue(true);
      } else {
        // Already welcomed before — a lighter nudge instead of the full
        // welcome copy, whenever fresh tasks land for them (from a new MoM,
        // an admin assignment, or a reassignment).
        MailApp.sendEmail({
          to: email,
          subject: 'New tasks have been added!',
          body: `Hi ${ownerName},\n\nNew tasks have been added!\n\n` +
            `${items.map(it => `- ${it.task}`).join('\n')}\n\n` +
            `View your full checklist here:\n${link}`,
          htmlBody: `<p>Hi ${escapeHtml_(ownerName)},</p>` +
            `<p><strong>New tasks have been added!</strong></p>` +
            buildItemsListHtml_(items) +
            `<p>View your full checklist here: <a href="${link}">${escapeHtml_(link)}</a></p>`
        });
      }

      items.forEach(it => tracker.getRange(it.row, notifiedColIndex).setValue(true));
      break;
    }
  });
}

// ---------- sheet menu: manually (re)send a welcome email to chosen owners ----------
//
// The automatic path (notifyOwners, above) covers everyone on its own — this
// is the manual override, reached from a menu button on the Google Sheet
// itself (Extensions menu isn't involved; this adds its own top-level
// "Task Tracker" menu). Useful right after a bulk migration, or if someone
// lost their original email and needs it resent on demand.
//
// onOpen() is a simple trigger — Sheets runs it automatically every time the
// spreadsheet is opened, no manual trigger setup needed.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Task Tracker')
    .addItem('Send welcome email...', 'showWelcomeDialog')
    .addToUi();
}

function showWelcomeDialog() {
  const ss = getSpreadsheet_();
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const ownersData = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  const names = ownersData.slice(1).map(r => r[nameCol]).filter(Boolean);

  const template = HtmlService.createTemplateFromFile('WelcomeDialog');
  template.names = names;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(360).setHeight(440),
    'Send welcome email'
  );
}

// Called from WelcomeDialog.html via google.script.run. Runs as whoever has
// the Sheet open (already an authorized editor), so no separate password
// check is needed here — that's only for the public web app.
function sendWelcomeFromSheet(ownerNames) {
  return sendWelcomeCore_(Array.isArray(ownerNames) ? ownerNames : []);
}

// Bypasses the WelcomeSent flag entirely — (re)sends the welcome link to
// each named owner right now, listing whatever tasks they currently have.
// Marks WelcomeSent true and their current rows Notified true, same
// bookkeeping as the automatic path in notifyOwners.
function sendWelcomeCore_(ownerNames) {
  if (!ownerNames.length) return { error: 'Pick at least one owner' };

  const ss = getSpreadsheet_();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const baseUrl = getSiteBaseUrl_();

  const trackerData = tracker.getDataRange().getValues();
  const tcol = name => trackerData[0].indexOf(name);
  const notifiedColIndex = tcol('Notified') + 1;

  const ownersData = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  const emailCol = OWNERS_HEADERS.indexOf('Email');
  const tokenCol = OWNERS_HEADERS.indexOf('Token');
  const welcomeCol = OWNERS_HEADERS.indexOf('WelcomeSent');

  const sent = [];
  ownerNames.forEach(ownerName => {
    for (let i = 1; i < ownersData.length; i++) {
      if (ownersData[i][nameCol] !== ownerName) continue;

      const email = ownersData[i][emailCol];
      const token = ownersData[i][tokenCol];
      const link = `${baseUrl}?token=${token}`;
      const items = [];
      for (let r = 1; r < trackerData.length; r++) {
        if (trackerData[r][tcol('Owner')] === ownerName) {
          items.push({ row: r + 1, task: trackerData[r][tcol('Task')] });
        }
      }

      MailApp.sendEmail({
        to: email,
        subject: 'Your action items checklist',
        body: `Hi ${ownerName},\n\nHere's your permanent link to your action items checklist:\n\n${link}\n\n` +
          (items.length ? `Current items:\n${items.map(it => `- ${it.task}`).join('\n')}\n\n` : '') +
          `Just tick things off (or mark them In Progress / Blocked / Revised Timeline) as you go.`,
        htmlBody: `<p>Hi ${escapeHtml_(ownerName)},</p>` +
          `<p>Here's your permanent link to your action items checklist:</p>` +
          `<p><a href="${link}">${escapeHtml_(link)}</a></p>` +
          (items.length ? `<p>Current items:</p>${buildItemsListHtml_(items)}` : '') +
          `<p>Just tick things off (or mark them In Progress / Blocked / Revised Timeline) as you go.</p>`
      });

      owners.getRange(i + 1, welcomeCol + 1).setValue(true);
      items.forEach(it => tracker.getRange(it.row, notifiedColIndex).setValue(true));
      sent.push(ownerName);
      break;
    }
  });

  return { ok: true, sent };
}

function getSiteBaseUrl_() {
  return PropertiesService.getScriptProperties().getProperty('SITE_BASE_URL') || 'https://REPLACE-WITH-YOUR-GITHUB-PAGES-URL/';
}

function setSiteBaseUrl(url) {
  PropertiesService.getScriptProperties().setProperty('SITE_BASE_URL', url);
}

// Once you know your GitHub Pages URL: fill it in below, select `setupSiteUrl`
// in the function dropdown, and click Run once (see comment on `setup` above
// for why this can't just be run directly on setSiteBaseUrl).
function setupSiteUrl() {
  setSiteBaseUrl('https://PASTE_YOUR_GITHUB_PAGES_URL_HERE/');
}

// ---------- step 3: escalating reminders ----------

function sendReminders() {
  const ss = getSpreadsheet_();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const baseUrl = getSiteBaseUrl_();

  const data = tracker.getDataRange().getValues();
  const col = name => data[0].indexOf(name);
  const today = new Date();

  const dueByOwner = {};
  for (let i = 1; i < data.length; i++) {
    const status = data[i][col('Status')];
    if (status === STATUS.DONE || status === STATUS.BLOCKED) continue;

    const revisedDate = data[i][col('Revised Timeline Date')];
    if (revisedDate && new Date(revisedDate) > today) continue;

    const owner = data[i][col('Owner')];
    if (!dueByOwner[owner]) dueByOwner[owner] = [];
    dueByOwner[owner].push({ row: i + 1, task: data[i][col('Task')], status });
  }

  const ownersData = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  const emailCol = OWNERS_HEADERS.indexOf('Email');
  const tokenCol = OWNERS_HEADERS.indexOf('Token');
  const pilotCol = OWNERS_HEADERS.indexOf('Pilot');
  const reminderColIndex = col('Reminder Count') + 1;

  Object.keys(dueByOwner).forEach(ownerName => {
    for (let i = 1; i < ownersData.length; i++) {
      if (ownersData[i][nameCol] !== ownerName) continue;
      if (ownersData[i][pilotCol] !== true) break; // not in the pilot yet — no reminder, no count bump
      const email = ownersData[i][emailCol];
      const token = ownersData[i][tokenCol];
      const link = `${baseUrl}?token=${token}`;
      const items = dueByOwner[ownerName];

      MailApp.sendEmail({
        to: email,
        subject: `Reminder: ${items.length} pending action item(s)`,
        body: `Hi ${ownerName},\n\nStill open on your checklist:\n\n` +
          `${items.map(it => `- ${it.task} (${it.status})`).join('\n')}\n\n` +
          `Update your status here:\n${link}`,
        htmlBody: `<p>Hi ${escapeHtml_(ownerName)},</p>` +
          `<p>Still open on your checklist:</p>` +
          buildItemsListHtml_(items, it => `${escapeHtml_(it.task)} (${escapeHtml_(it.status)})`) +
          `<p>Update your status here: <a href="${link}">${escapeHtml_(link)}</a></p>`
      });

      items.forEach(it => {
        const reminderCountCell = tracker.getRange(it.row, reminderColIndex);
        reminderCountCell.setValue((reminderCountCell.getValue() || 0) + 1);
      });
      break;
    }
  });
}

// ---------- Web App API ----------

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getTasks') return jsonOut_(getTasksForToken_(e.parameter.token));
  return jsonOut_({ error: 'Unknown action' });
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ error: 'Invalid request body' });
  }

  if (payload.action === 'updateStatus') return jsonOut_(updateStatus_(payload));
  if (payload.action === 'adminList') return jsonOut_(getAdminList_(payload.password));
  if (payload.action === 'createTask') return jsonOut_(createTask_(payload));
  if (payload.action === 'updateTask') return jsonOut_(updateTask_(payload));
  if (payload.action === 'deleteTask') return jsonOut_(deleteTask_(payload));
  if (payload.action === 'createOwnTask') return jsonOut_(createOwnTask_(payload));
  if (payload.action === 'getComments') return jsonOut_(getComments_(payload));
  if (payload.action === 'addComment') return jsonOut_(addComment_(payload));
  if (payload.action === 'deleteComment') return jsonOut_(deleteComment_(payload));
  return jsonOut_({ error: 'Unknown action' });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Looks up the owner name a token belongs to. Returns null if the token
// doesn't match anyone in the Owners sheet.
function resolveOwnerNameByToken_(ss, token) {
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const ownersData = owners.getDataRange().getValues();
  const tokenCol = OWNERS_HEADERS.indexOf('Token');
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  for (let i = 1; i < ownersData.length; i++) {
    if (ownersData[i][tokenCol] === token) return ownersData[i][nameCol];
  }
  return null;
}

function getTasksForToken_(token) {
  if (!token) return { error: 'Missing token' };
  const ss = getSpreadsheet_();
  const ownerName = resolveOwnerNameByToken_(ss, token);
  if (!ownerName) return { error: 'Invalid token' };

  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const data = tracker.getDataRange().getValues();
  const col = name => data[0].indexOf(name);

  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][col('Owner')] !== ownerName) continue;
    tasks.push({
      id: data[i][col('TaskID')],
      task: data[i][col('Task')],
      meeting: data[i][col('Meeting')],
      momDate: data[i][col('MoM Date')],
      status: data[i][col('Status')],
      revisedTimelineDate: data[i][col('Revised Timeline Date')]
    });
  }
  return { owner: ownerName, tasks };
}

function updateStatus_(payload) {
  const { token, id, status, revisedTimelineDate } = payload;
  if (!token || !id || !status) return { error: 'Missing fields' };
  if (!Object.values(STATUS).includes(status)) return { error: 'Invalid status' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet_();
    const ownerName = resolveOwnerNameByToken_(ss, token);
    if (!ownerName) return { error: 'Invalid token' };

    const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
    const data = tracker.getDataRange().getValues();
    const col = name => data[0].indexOf(name);

    for (let i = 1; i < data.length; i++) {
      if (data[i][col('TaskID')] !== id) continue;
      if (data[i][col('Owner')] !== ownerName) return { error: 'Task does not belong to this owner' };

      tracker.getRange(i + 1, col('Status') + 1).setValue(status);
      tracker.getRange(i + 1, col('Last Updated') + 1).setValue(new Date());
      if (status === STATUS.REVISED && revisedTimelineDate) {
        tracker.getRange(i + 1, col('Revised Timeline Date') + 1).setValue(revisedTimelineDate);
      }
      return { ok: true };
    }
    return { error: 'Task not found' };
  } finally {
    lock.releaseLock();
  }
}

function getAdminList_(password) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected || password !== expected) return { error: 'Unauthorized' };

  const ss = getSpreadsheet_();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const data = tracker.getDataRange().getValues();
  const col = name => data[0].indexOf(name);

  // Latest comment per task, computed fresh from the Comments sheet (the
  // source of truth) rather than the Tracker's joined summary column — the
  // admin table only ever needs to show the single most recent comment plus
  // when it landed, not the full thread.
  const commentsSheet = ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
  const commentsData = commentsSheet.getDataRange().getValues();
  const ccol = name => commentsData[0].indexOf(name);
  const latestCommentByTask = {};
  for (let i = 1; i < commentsData.length; i++) {
    const taskId = commentsData[i][ccol('TaskID')];
    const timestamp = commentsData[i][ccol('Timestamp')];
    const existing = latestCommentByTask[taskId];
    if (!existing || new Date(timestamp) > new Date(existing.timestamp)) {
      latestCommentByTask[taskId] = { text: commentsData[i][ccol('Text')], timestamp };
    }
  }

  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const taskId = data[i][col('TaskID')];
    const latestComment = latestCommentByTask[taskId];
    tasks.push({
      id: taskId,
      owner: data[i][col('Owner')],
      task: data[i][col('Task')],
      meeting: data[i][col('Meeting')],
      momDate: data[i][col('MoM Date')],
      status: data[i][col('Status')],
      revisedTimelineDate: data[i][col('Revised Timeline Date')],
      reminderCount: data[i][col('Reminder Count')],
      lastUpdated: data[i][col('Last Updated')],
      lastCommentText: latestComment ? latestComment.text : '',
      lastCommentAt: latestComment ? latestComment.timestamp : ''
    });
  }

  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const ownersData = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  const ownerNames = ownersData.slice(1).map(r => r[nameCol]).filter(Boolean);

  return { tasks, owners: ownerNames };
}

// ---------- shared helpers for creating/editing tasks ----------

// Looks up an owner row by exact name match. Returns null if not found, or an
// object with the row's 0-indexed position and email/token if found —
// generating a token on the spot if that owner never had one yet.
function findOwnerByName_(ss, ownerName) {
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const ownersData = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  const emailCol = OWNERS_HEADERS.indexOf('Email');
  const tokenCol = OWNERS_HEADERS.indexOf('Token');

  for (let i = 1; i < ownersData.length; i++) {
    if (ownersData[i][nameCol] !== ownerName) continue;
    const email = ownersData[i][emailCol];
    let token = ownersData[i][tokenCol];
    if (!token) {
      token = Utilities.getUuid();
      owners.getRange(i + 1, tokenCol + 1).setValue(token);
    }
    return { row: i, email, token };
  }
  return null;
}

// Builds a Tracker row from named fields, matching whatever the sheet's
// actual column order currently is.
function buildTaskRow_(tracker, fields) {
  const headerRow = getHeaderRow_(tracker);
  return buildRowByHeaders_(headerRow, {
    TaskID: generateTaskId_(tracker),
    Owner: fields.owner,
    OwnerEmail: fields.ownerEmail,
    Task: fields.task,
    Meeting: fields.meeting,
    'MoM Date': new Date(),
    Status: STATUS.PENDING,
    'Reminder Count': 0,
    'Last Updated': new Date(),
    Notified: fields.notified,
    SourceKey: fields.sourceKey
  });
}

// ---------- admin: create/assign a task directly ----------

// Accepts either payload.ownerNames (an array — assign the same task to
// several people at once, each gets their own row) or the older singular
// payload.ownerName (still supported for a single assignment).
function createTask_(payload) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected || payload.password !== expected) return { error: 'Unauthorized' };

  const ownerNames = Array.isArray(payload.ownerNames)
    ? payload.ownerNames
    : (payload.ownerName ? [payload.ownerName] : []);
  const task = payload.task;
  if (!ownerNames.length || !task) return { error: 'Missing fields' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet_();
    const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
    const ids = [];

    for (const ownerName of ownerNames) {
      const owner = findOwnerByName_(ss, ownerName);
      if (!owner) return { error: `Unknown owner: ${ownerName}` };
      if (!owner.email) return { error: `${ownerName} has no email on file` };

      const row = buildTaskRow_(tracker, {
        owner: ownerName,
        ownerEmail: owner.email,
        task,
        meeting: 'Assigned by admin',
        notified: false,
        sourceKey: `manual:${Utilities.getUuid()}`
      });
      tracker.appendRow(row);
      ids.push(row[getHeaderRow_(tracker).indexOf('TaskID')]);
    }

    return { ok: true, ids };
  } finally {
    lock.releaseLock();
  }
}

// ---------- admin: edit or delete an existing task ----------

function updateTask_(payload) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected || payload.password !== expected) return { error: 'Unauthorized' };
  if (!payload.id) return { error: 'Missing task id' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet_();
    const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
    const data = tracker.getDataRange().getValues();
    const col = name => data[0].indexOf(name);

    for (let i = 1; i < data.length; i++) {
      if (data[i][col('TaskID')] !== payload.id) continue;

      if (payload.task) {
        tracker.getRange(i + 1, col('Task') + 1).setValue(payload.task);
      }
      if (payload.ownerName && payload.ownerName !== data[i][col('Owner')]) {
        const owner = findOwnerByName_(ss, payload.ownerName);
        if (!owner) return { error: 'Unknown owner' };
        if (!owner.email) return { error: 'That owner has no email on file' };
        tracker.getRange(i + 1, col('Owner') + 1).setValue(payload.ownerName);
        tracker.getRange(i + 1, col('OwnerEmail') + 1).setValue(owner.email);
        tracker.getRange(i + 1, col('Notified') + 1).setValue(false); // reassigned — let them know
      }
      tracker.getRange(i + 1, col('Last Updated') + 1).setValue(new Date());
      return { ok: true };
    }
    return { error: 'Task not found' };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask_(payload) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected || payload.password !== expected) return { error: 'Unauthorized' };
  if (!payload.id) return { error: 'Missing task id' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet_();
    const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
    const data = tracker.getDataRange().getValues();
    const col = name => data[0].indexOf(name);

    for (let i = 1; i < data.length; i++) {
      if (data[i][col('TaskID')] !== payload.id) continue;
      tracker.deleteRow(i + 1);
      deleteCommentsForTask_(ss, payload.id);
      return { ok: true };
    }
    return { error: 'Task not found' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- owner: add their own task ----------

function createOwnTask_(payload) {
  const { token, task } = payload;
  if (!token || !task) return { error: 'Missing fields' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet_();
    const ownerName = resolveOwnerNameByToken_(ss, token);
    if (!ownerName) return { error: 'Invalid token' };
    const owner = findOwnerByName_(ss, ownerName);

    const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
    const row = buildTaskRow_(tracker, {
      owner: ownerName,
      ownerEmail: owner.email,
      task,
      meeting: 'Added by owner',
      notified: true, // they just added it themselves, no need to email them about it
      sourceKey: `self:${Utilities.getUuid()}`
    });

    tracker.appendRow(row);
    return { ok: true, id: row[getHeaderRow_(tracker).indexOf('TaskID')] };
  } finally {
    lock.releaseLock();
  }
}

// ---------- comments (running log per task) ----------

// Authorizes a request to VIEW comments on a specific task: an admin password
// always works, or an owner token if that task actually belongs to them.
// Returns { author } (the name to attribute a new comment to) or { error }.
function authorizeForTask_(ss, payload, taskId) {
  if (payload.password) {
    const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!expected || payload.password !== expected) return { error: 'Unauthorized' };
    return { author: 'Admin' };
  }
  if (payload.token) {
    const ownerName = resolveOwnerNameByToken_(ss, payload.token);
    if (!ownerName) return { error: 'Invalid token' };
    const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
    const data = tracker.getDataRange().getValues();
    const col = name => data[0].indexOf(name);
    const owns = data.some((row, i) => i > 0 && row[col('TaskID')] === taskId && row[col('Owner')] === ownerName);
    if (!owns) return { error: 'Task does not belong to this owner' };
    return { author: ownerName };
  }
  return { error: 'Missing credentials' };
}

function getComments_(payload) {
  const { id } = payload;
  if (!id) return { error: 'Missing task id' };
  const ss = getSpreadsheet_();
  const auth = authorizeForTask_(ss, payload, id);
  if (auth.error) return auth;

  const sheet = ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const col = name => data[0].indexOf(name);
  const commentIdCol = col('CommentID') + 1;

  const comments = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][col('TaskID')] !== id) continue;

    // Comments added before the CommentID column existed have a blank one —
    // backfill it here on first read so they become deletable too.
    let commentId = data[i][col('CommentID')];
    if (!commentId) {
      commentId = Utilities.getUuid();
      sheet.getRange(i + 1, commentIdCol).setValue(commentId);
    }

    comments.push({
      id: commentId,
      author: data[i][col('Author')],
      text: data[i][col('Text')],
      timestamp: data[i][col('Timestamp')]
    });
  }
  return { comments };
}

// Comments can only be added by the task's owner (via their token) — not by
// the admin. Admin can still view them (see authorizeForTask_/getComments_).
function addComment_(payload) {
  const { id, text } = payload;
  if (!id || !text) return { error: 'Missing fields' };
  if (!payload.token) return { error: 'Only the task owner can add a comment' };

  const ss = getSpreadsheet_();
  const auth = authorizeForTask_(ss, { token: payload.token }, id);
  if (auth.error) return auth;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const commentsSheet = ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
    const commentId = Utilities.getUuid();
    const headerRow = getHeaderRow_(commentsSheet);
    commentsSheet.appendRow(buildRowByHeaders_(headerRow, {
      TaskID: id,
      Author: auth.author,
      Text: text,
      Timestamp: new Date(),
      CommentID: commentId
    }));
    refreshTrackerCommentSummary_(ss, id);
    return { ok: true, id: commentId };
  } finally {
    lock.releaseLock();
  }
}

// Same ownership rule as adding: only the task's owner (via their token) can
// delete one of the comments on it.
function deleteComment_(payload) {
  const { id, commentId } = payload;
  if (!id || !commentId) return { error: 'Missing fields' };
  if (!payload.token) return { error: 'Only the task owner can delete a comment' };

  const ss = getSpreadsheet_();
  const auth = authorizeForTask_(ss, { token: payload.token }, id);
  if (auth.error) return auth;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const col = name => data[0].indexOf(name);
    for (let i = 1; i < data.length; i++) {
      if (data[i][col('TaskID')] === id && data[i][col('CommentID')] === commentId) {
        sheet.deleteRow(i + 1);
        refreshTrackerCommentSummary_(ss, id);
        return { ok: true };
      }
    }
    return { error: 'Comment not found' };
  } finally {
    lock.releaseLock();
  }
}

function deleteCommentsForTask_(ss, taskId) {
  const sheet = ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const col = data[0] ? data[0].indexOf('TaskID') : -1;
  if (col === -1) return;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][col] === taskId) sheet.deleteRow(i + 1);
  }
}

// Keeps the Tracker sheet's own "Comments" column in sync with the full log
// in the Comments sheet, so you can see recent comments at a glance without
// opening the web app. The Comments sheet remains the source of truth.
function refreshTrackerCommentSummary_(ss, taskId) {
  const commentsSheet = ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
  const commentsData = commentsSheet.getDataRange().getValues();
  const ccol = name => commentsData[0].indexOf(name);

  const summary = commentsData
    .slice(1)
    .filter(row => row[ccol('TaskID')] === taskId)
    .map(row => `${row[ccol('Author')]}: ${row[ccol('Text')]}`)
    .join(' | ');

  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const trackerData = tracker.getDataRange().getValues();
  const tcol = name => trackerData[0].indexOf(name);
  const idCol = tcol('TaskID');
  const commentsCol = tcol('Comments');
  if (commentsCol === -1) return;

  for (let i = 1; i < trackerData.length; i++) {
    if (trackerData[i][idCol] === taskId) {
      tracker.getRange(i + 1, commentsCol + 1).setValue(summary);
      return;
    }
  }
}

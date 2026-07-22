/**
 * MoM Task Tracker backend.
 *
 * Sheets used (created automatically by initializeSheets if missing):
 *   Tracker: TaskID | Owner | OwnerEmail | Task | Meeting | MoM Date | Status
 *            | Revised Timeline Date | Reminder Count | Last Updated | Notified | SourceKey
 *   Owners:  Name | Email | Token | WelcomeSent
 *   Unmatched: raw @name tags scanMoMEmails couldn't resolve to an email, for manual fixup
 *
 * One-time setup (run once from the Apps Script editor):
 *   1. initializeSheets()
 *   2. setAdminPassword('choose-a-password')
 *   3. Pre-populate the Owners sheet with Name + Email for everyone you tag in MoMs.
 *   4. Deploy > New deployment > Web app (execute as Me, access Anyone), copy the URL into site/config.js.
 *   5. Set time-driven triggers for scanMoMEmails, notifyOwners, sendReminders (Triggers panel).
 */

const TRACKER_SHEET = 'Tracker';
const OWNERS_SHEET = 'Owners';
const UNMATCHED_SHEET = 'Unmatched';

const STATUS = {
  PENDING: 'Pending',
  IN_PROGRESS: 'In Progress',
  BLOCKED: 'Blocked',
  REVISED: 'Revised Timeline',
  DONE: 'Done'
};

const TRACKER_HEADERS = [
  'TaskID', 'Owner', 'OwnerEmail', 'Task', 'Meeting', 'MoM Date', 'Status',
  'Revised Timeline Date', 'Reminder Count', 'Last Updated', 'Notified', 'SourceKey'
];
const OWNERS_HEADERS = ['Name', 'Email', 'Token', 'WelcomeSent'];
const UNMATCHED_HEADERS = ['Name Tag', 'Task', 'Meeting', 'MoM Date', 'Seen At'];

// ---------- setup ----------

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  ensureSheet_(ss, UNMATCHED_SHEET, UNMATCHED_HEADERS);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setAdminPassword(password) {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', password);
}

// ---------- step 1: scan sent MoM emails ----------

function scanMoMEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const existingKeys = new Set(
    tracker.getRange(2, TRACKER_HEADERS.indexOf('SourceKey') + 1, Math.max(tracker.getLastRow() - 1, 0), 1)
      .getValues().flat().filter(String)
  );

  const threads = GmailApp.search('in:sent subject:MoM newer_than:2d');
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

        newRows.push([
          generateTaskId_(tracker),
          owner.name,
          owner.email,
          action.task,
          meetingTitle,
          momDate,
          STATUS.PENDING,
          '',
          0,
          new Date(),
          false,
          sourceKey
        ]);
        existingKeys.add(sourceKey);
      });
    });
  });

  if (newRows.length) {
    tracker.getRange(tracker.getLastRow() + 1, 1, newRows.length, TRACKER_HEADERS.length).setValues(newRows);
  }
}

function parseActionsSection_(body) {
  const match = body.match(/\[Actions\]([\s\S]*?)(?:\n\s*\[[A-Za-z]|$)/i);
  if (!match) return [];
  const lines = match[1].split('\n');
  const actions = [];
  lines.forEach(line => {
    const lineMatch = line.match(/^\s*[-*]?\s*@(\S+)\s*[:\-]\s*(.+?)\s*$/);
    if (lineMatch) {
      actions.push({ nameTag: lineMatch[1], task: lineMatch[2] });
    }
  });
  return actions;
}

function resolveOwner_(ss, nameTag) {
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const data = owners.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '');
    if (name.toLowerCase() === nameTag.toLowerCase() ||
        name.toLowerCase().startsWith(nameTag.toLowerCase())) {
      const email = data[i][1];
      if (!email) return null;
      let token = data[i][2];
      if (!token) {
        token = Utilities.getUuid();
        owners.getRange(i + 1, 3).setValue(token);
      }
      return { name, email, token };
    }
  }
  return null;
}

function logUnmatched_(ss, nameTag, task, meetingTitle, momDate) {
  const sheet = ensureSheet_(ss, UNMATCHED_SHEET, UNMATCHED_HEADERS);
  sheet.appendRow([nameTag, task, meetingTitle, momDate, new Date()]);
}

function generateTaskId_(tracker) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const existing = new Set(
    tracker.getRange(2, 1, Math.max(tracker.getLastRow() - 1, 0), 1).getValues().flat()
  );
  let id;
  do {
    id = Array.from({ length: 4 }, () => chars[Math.floor(Math.random() * chars.length)]).join('');
  } while (existing.has(id));
  return id;
}

// ---------- step 2: notify owners (welcome or new-items nudge) ----------

function notifyOwners() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const baseUrl = getSiteBaseUrl_();

  const trackerData = tracker.getDataRange().getValues();
  const notifiedCol = TRACKER_HEADERS.indexOf('Notified');
  const ownerCol = TRACKER_HEADERS.indexOf('Owner');
  const taskCol = TRACKER_HEADERS.indexOf('Task');

  const pendingByOwner = {};
  for (let i = 1; i < trackerData.length; i++) {
    if (trackerData[i][notifiedCol] === true) continue;
    const owner = trackerData[i][ownerCol];
    if (!pendingByOwner[owner]) pendingByOwner[owner] = [];
    pendingByOwner[owner].push({ row: i + 1, task: trackerData[i][taskCol] });
  }

  const ownersData = owners.getDataRange().getValues();
  const nameCol = OWNERS_HEADERS.indexOf('Name');
  const emailCol = OWNERS_HEADERS.indexOf('Email');
  const tokenCol = OWNERS_HEADERS.indexOf('Token');
  const welcomeCol = OWNERS_HEADERS.indexOf('WelcomeSent');

  Object.keys(pendingByOwner).forEach(ownerName => {
    for (let i = 1; i < ownersData.length; i++) {
      if (ownersData[i][nameCol] !== ownerName) continue;

      const email = ownersData[i][emailCol];
      const token = ownersData[i][tokenCol];
      const welcomeSent = ownersData[i][welcomeCol];
      const link = `${baseUrl}?token=${token}`;
      const items = pendingByOwner[ownerName];

      if (!welcomeSent) {
        MailApp.sendEmail({
          to: email,
          subject: 'Your action items checklist',
          body: `Hi ${ownerName},\n\nYou've been tagged with action items from a recent meeting. ` +
            `Bookmark this link — it always shows your current, live checklist:\n\n${link}\n\n` +
            `New items right now:\n${items.map(it => `- ${it.task}`).join('\n')}\n\n` +
            `Just tick things off (or mark them In Progress / Blocked / Revised Timeline) as you go.`
        });
        owners.getRange(i + 1, welcomeCol + 1).setValue(true);
      } else {
        MailApp.sendEmail({
          to: email,
          subject: 'New action items added to your checklist',
          body: `Hi ${ownerName},\n\nNew action items were just added to your checklist:\n\n` +
            `${items.map(it => `- ${it.task}`).join('\n')}\n\nView/update your full list here:\n${link}`
        });
      }

      items.forEach(it => tracker.getRange(it.row, notifiedCol + 1).setValue(true));
      break;
    }
  });
}

function getSiteBaseUrl_() {
  return PropertiesService.getScriptProperties().getProperty('SITE_BASE_URL') || 'https://REPLACE-WITH-YOUR-GITHUB-PAGES-URL/';
}

function setSiteBaseUrl(url) {
  PropertiesService.getScriptProperties().setProperty('SITE_BASE_URL', url);
}

// ---------- step 3: escalating reminders ----------

function sendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const baseUrl = getSiteBaseUrl_();

  const data = tracker.getDataRange().getValues();
  const col = name => TRACKER_HEADERS.indexOf(name);
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

  Object.keys(dueByOwner).forEach(ownerName => {
    for (let i = 1; i < ownersData.length; i++) {
      if (ownersData[i][nameCol] !== ownerName) continue;
      const email = ownersData[i][emailCol];
      const token = ownersData[i][tokenCol];
      const link = `${baseUrl}?token=${token}`;
      const items = dueByOwner[ownerName];

      MailApp.sendEmail({
        to: email,
        subject: `Reminder: ${items.length} pending action item(s)`,
        body: `Hi ${ownerName},\n\nStill open on your checklist:\n\n` +
          `${items.map(it => `- ${it.task} (${it.status})`).join('\n')}\n\n` +
          `Update your status here:\n${link}`
      });

      items.forEach(it => {
        const reminderCountCell = tracker.getRange(it.row, col('Reminder Count') + 1);
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
  return jsonOut_({ error: 'Unknown action' });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getTasksForToken_(token) {
  if (!token) return { error: 'Missing token' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const ownersData = owners.getDataRange().getValues();
  const tokenCol = OWNERS_HEADERS.indexOf('Token');
  const nameCol = OWNERS_HEADERS.indexOf('Name');

  let ownerName = null;
  for (let i = 1; i < ownersData.length; i++) {
    if (ownersData[i][tokenCol] === token) {
      ownerName = ownersData[i][nameCol];
      break;
    }
  }
  if (!ownerName) return { error: 'Invalid token' };

  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const data = tracker.getDataRange().getValues();
  const col = name => TRACKER_HEADERS.indexOf(name);

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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
    const ownersData = owners.getDataRange().getValues();
    const tokenCol = OWNERS_HEADERS.indexOf('Token');
    const nameCol = OWNERS_HEADERS.indexOf('Name');

    let ownerName = null;
    for (let i = 1; i < ownersData.length; i++) {
      if (ownersData[i][tokenCol] === token) {
        ownerName = ownersData[i][nameCol];
        break;
      }
    }
    if (!ownerName) return { error: 'Invalid token' };

    const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
    const data = tracker.getDataRange().getValues();
    const col = name => TRACKER_HEADERS.indexOf(name);

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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const data = tracker.getDataRange().getValues();
  const col = name => TRACKER_HEADERS.indexOf(name);

  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    tasks.push({
      id: data[i][col('TaskID')],
      owner: data[i][col('Owner')],
      task: data[i][col('Task')],
      meeting: data[i][col('Meeting')],
      momDate: data[i][col('MoM Date')],
      status: data[i][col('Status')],
      revisedTimelineDate: data[i][col('Revised Timeline Date')],
      reminderCount: data[i][col('Reminder Count')],
      lastUpdated: data[i][col('Last Updated')]
    });
  }
  return { tasks };
}

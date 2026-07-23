/**
 * One-time migration from the old Sheet-based task tracker into this system.
 *
 * This is a standalone utility, separate from Code.gs's ongoing production
 * logic. Paste this into the SAME Apps Script project as an additional file
 * (File > + next to Files > Script), run migrateLegacyTracker() once, then
 * run sendWelcomeToMigratedOwners() once. You can delete this file afterward
 * — it's not part of the regular scan/notify/remind flow.
 *
 * What it does:
 *   1. Finds the old tracker's data table (by header signature) and its
 *      contacts table (by content shape — a sheet whose 2nd column is mostly
 *      email addresses), wherever they are in the old spreadsheet's tabs.
 *   2. Skips obvious test rows (MoM subject containing "test").
 *   3. Resolves each Owner Email to a name via the contacts table, falling
 *      back to the email's local-part if not found.
 *   4. Reuses the old 4-character IDs as TaskIDs. Safe to re-run — any ID
 *      already present in your Tracker sheet is skipped, not duplicated.
 *   5. Normalizes status casing/spelling to this system's four statuses.
 *   6. Carries over "Latest Reply" as a comment on that task, and the old
 *      Reminder Count.
 *   7. Marks every migrated task as already-notified, so nobody gets a
 *      "look at these new items" flood email. sendWelcomeToMigratedOwners()
 *      handles introducing people to their checklist link separately,
 *      without listing every task in the email body.
 *   8. Skips anything from a MoM sent in the last 2 days — that's the same
 *      window scanMoMEmails itself looks back over, so the live system will
 *      pick those up naturally; migrating them here too would duplicate them.
 */

function migrateLegacyTracker() {
  const OLD_SPREADSHEET_ID = '1bb84bb82qa8oDFjlj0rO1I4P3W483GhnPq2-eb8xnDU';
  const oldSs = SpreadsheetApp.openById(OLD_SPREADSHEET_ID);

  const oldTrackerSheet = findSheetByHeaders_(oldSs, ['ID', 'Owner Email', 'Action Item', 'Status']);
  if (!oldTrackerSheet) throw new Error('Could not find the old tracker table (looked for a sheet with headers ID / Owner Email / Action Item / Status).');

  const peopleSheet = findPeopleSheet_(oldSs, oldTrackerSheet);
  const emailToName = peopleSheet ? buildEmailNameMap_(peopleSheet) : {};

  const oldData = oldTrackerSheet.getDataRange().getValues();
  const oldHeaders = oldData[0];
  const oldCol = name => oldHeaders.indexOf(name);

  const ss = getSpreadsheet_();
  const tracker = ensureSheet_(ss, TRACKER_SHEET, TRACKER_HEADERS);
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const trackerHeaderRow = getHeaderRow_(tracker);
  const existingIds = new Set(getColumnValues_(tracker, 'TaskID'));

  const statusMap = {
    'pending': STATUS.PENDING,
    'in progress': STATUS.IN_PROGRESS,
    'in-progress': STATUS.IN_PROGRESS,
    'blocked': STATUS.BLOCKED,
    'done': STATUS.DONE,
    'revised timeline': STATUS.REVISED
  };

  // Anything this recent might get picked up naturally by the live
  // scanMoMEmails (which itself looks back 2 days) — skip it here so the two
  // systems can't both create a task for the same real MoM.
  const recentCutoff = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);

  const newRows = [];
  const commentsToAdd = [];
  let skippedTest = 0, skippedDuplicate = 0, skippedNoTask = 0, skippedRecent = 0;

  for (let i = 1; i < oldData.length; i++) {
    const row = oldData[i];
    const oldId = String(row[oldCol('ID')] || '').trim();
    if (!oldId) continue;

    if (existingIds.has(oldId)) { skippedDuplicate++; continue; }

    const momSubject = String(row[oldCol('MoM Subject')] || '').trim();
    if (/test/i.test(momSubject)) { skippedTest++; continue; }

    const task = String(row[oldCol('Action Item')] || '').trim();
    if (!task) { skippedNoTask++; continue; }

    const rawMomDate = row[oldCol('MoM Sent Date')];
    const momDate = rawMomDate instanceof Date ? rawMomDate : (rawMomDate ? new Date(rawMomDate) : new Date());
    if (momDate > recentCutoff) { skippedRecent++; continue; }

    const ownerEmail = String(row[oldCol('Owner Email')] || '').trim();
    if (!ownerEmail) continue;
    const ownerName = emailToName[ownerEmail.toLowerCase()] || ownerEmail.split('@')[0];
    ensureOwnerExists_(owners, ownerName, ownerEmail);

    const rawStatus = String(row[oldCol('Status')] || 'Pending').trim().toLowerCase();
    const status = statusMap[rawStatus] || STATUS.PENDING;

    const reminderCount = Number(row[oldCol('Reminder Count')]) || 0;

    newRows.push(buildRowByHeaders_(trackerHeaderRow, {
      TaskID: oldId,
      Owner: ownerName,
      OwnerEmail: ownerEmail,
      Task: task,
      Meeting: momSubject,
      'MoM Date': momDate,
      Status: status,
      'Reminder Count': reminderCount,
      'Last Updated': new Date(),
      Notified: true,
      SourceKey: `legacy:${oldId}`
    }));
    existingIds.add(oldId);

    const latestReply = String(row[oldCol('Latest Reply')] || '').trim();
    if (latestReply) {
      commentsToAdd.push({ taskId: oldId, author: ownerName, text: latestReply, timestamp: momDate });
    }
  }

  if (newRows.length) {
    tracker.getRange(tracker.getLastRow() + 1, 1, newRows.length, trackerHeaderRow.length).setValues(newRows);
  }

  if (commentsToAdd.length) {
    const commentsSheet = ensureSheet_(ss, COMMENTS_SHEET, COMMENTS_HEADERS);
    const seenTaskIds = new Set();
    commentsToAdd.forEach(c => {
      commentsSheet.appendRow([c.taskId, c.author, c.text, c.timestamp]);
      seenTaskIds.add(c.taskId);
    });
    seenTaskIds.forEach(taskId => refreshTrackerCommentSummary_(ss, taskId));
  }

  Logger.log(
    `Migrated ${newRows.length} tasks. Skipped: ${skippedTest} test rows, ` +
    `${skippedDuplicate} already-migrated, ${skippedNoTask} with no task text, ` +
    `${skippedRecent} too recent (left for scanMoMEmails to pick up naturally).`
  );
}

// Sends the "here's your permanent checklist link" welcome email to any
// owner who hasn't received one yet (WelcomeSent isn't set) — without
// listing individual tasks, since a bulk migration can mean a lot of them.
// Safe to run any time; only touches owners not already welcomed.
function sendWelcomeToMigratedOwners() {
  const ss = getSpreadsheet_();
  const owners = ensureSheet_(ss, OWNERS_SHEET, OWNERS_HEADERS);
  const data = owners.getDataRange().getValues();
  const col = name => data[0].indexOf(name);
  const baseUrl = getSiteBaseUrl_();

  let sent = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][col('WelcomeSent')] === true) continue;
    const name = data[i][col('Name')];
    const email = data[i][col('Email')];
    const token = data[i][col('Token')];
    if (!email || !token) continue;

    MailApp.sendEmail({
      to: email,
      subject: 'Your action items checklist',
      body: `Hi ${name},\n\nYou have action items assigned to you in the CFI Task Tracker. ` +
        `Bookmark this link — it always shows your current, live checklist:\n\n${baseUrl}?token=${token}\n\n` +
        `Just tick things off (or mark them In Progress / Blocked / Revised Timeline) as you go.`
    });
    owners.getRange(i + 1, col('WelcomeSent') + 1).setValue(true);
    sent++;
  }

  Logger.log(`Sent welcome email to ${sent} owner(s).`);
}

// Finds a sheet/tab whose header row (row 1) contains all of the given
// column names, regardless of which tab it's actually on.
function findSheetByHeaders_(ss, requiredHeaders) {
  const sheets = ss.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    if (sheet.getLastRow() < 1) continue;
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    if (requiredHeaders.every(h => header.indexOf(h) !== -1)) return sheet;
  }
  return null;
}

// Finds the contacts-like sheet by content shape (column B mostly looks like
// email addresses) rather than by header text, since header labels varied.
function findPeopleSheet_(ss, excludeSheet) {
  const sheets = ss.getSheets();
  let best = null, bestScore = 0;
  sheets.forEach(sheet => {
    if (sheet.getSheetId() === excludeSheet.getSheetId()) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const sampleSize = Math.min(lastRow - 1, 50);
    const sample = sheet.getRange(2, 1, sampleSize, 2).getValues();
    let emailCount = 0, total = 0;
    sample.forEach(r => {
      if (!r[0]) return;
      total++;
      if (r[1] && /@/.test(String(r[1]))) emailCount++;
    });
    const score = total ? emailCount / total : 0;
    if (score > bestScore && score > 0.5) { bestScore = score; best = sheet; }
  });
  return best;
}

// Builds a lowercased-email -> name lookup from a contacts sheet shaped like
// Name (col A) | Email (col B) | ...anything else.
function buildEmailNameMap_(peopleSheet) {
  const data = peopleSheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    const email = String(data[i][1] || '').trim().toLowerCase();
    if (name && email && email.indexOf('@') !== -1) map[email] = name;
  }
  return map;
}

// Adds an owner row (with a generated token) if one doesn't already exist
// for this email.
function ensureOwnerExists_(ownersSheet, name, email) {
  const data = ownersSheet.getDataRange().getValues();
  const emailLower = email.toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').toLowerCase() === emailLower) return;
  }
  ownersSheet.appendRow([name, email, Utilities.getUuid(), false]);
}

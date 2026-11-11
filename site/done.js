// One-click "Mark done" landing page, opened from a link in the emails.
//
// The update happens here, in the page's own JavaScript, rather than on the
// link itself. That's deliberate: some mail systems (Outlook Safe Links and
// the like) open every link in an email to scan it, and a link that changed
// something just by being fetched would silently mark tasks done. A scanner
// that only downloads this HTML never runs this script, so nothing changes.
//
// It reuses the existing getTasks / updateStatus calls, so there's no extra
// backend, and it uses the same per-owner token as the checklist link.

if (window.initThemeToggle) window.initThemeToggle('themeToggle');

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str == null ? '' : str;
  return div.innerHTML;
}

async function getTasks(token) {
  const url = new URL(API_URL);
  url.searchParams.set('action', 'getTasks');
  url.searchParams.set('token', token);
  const res = await fetch(url.toString());
  return res.json();
}

async function updateStatus(token, id, status) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: 'updateStatus', token, id, status })
  });
  return res.json();
}

function render(html) {
  document.getElementById('result').innerHTML = html;
}

function checklistLink(token) {
  return `index.html?token=${encodeURIComponent(token || '')}`;
}

function failure(message, token) {
  render(`
    <div class="done-icon done-icon-bad">!</div>
    <h1 class="done-title">Couldn't update that</h1>
    <p class="done-detail">${escapeHtml(message)}</p>
    ${token ? `<div class="done-actions"><a class="done-link" href="${checklistLink(token)}">Open my checklist</a></div>` : ''}
  `);
}

function openCount(tasks, exceptId) {
  return tasks.filter(t => t.id !== exceptId && t.status !== 'Done').length;
}

function summaryLine(n) {
  if (n === 0) return 'Nothing else is open. You\'re all caught up.';
  return n === 1 ? 'You have 1 other open item.' : `You have ${n} other open items.`;
}

async function main() {
  const params = new URLSearchParams(window.location.search);
  const token = params.get('token');
  const id = params.get('id');

  if (!token || !id) {
    failure('This link looks incomplete. Please use the link from your email.', null);
    return;
  }

  let data;
  try {
    data = await getTasks(token);
  } catch (err) {
    failure('Could not reach the server. Check your connection and try again.', token);
    return;
  }
  if (data.error) {
    failure(data.error === 'Invalid token' ? 'This link isn\'t valid.' : data.error, null);
    return;
  }

  const task = data.tasks.find(t => t.id === id);
  if (!task) {
    failure('That task is no longer on your list. It may have been removed.', token);
    return;
  }

  const others = openCount(data.tasks, id);

  // Already done: say so, change nothing (so reopening the email later is harmless).
  if (task.status === 'Done') {
    render(`
      <div class="done-icon">&#10003;</div>
      <h1 class="done-title">Already marked done</h1>
      <p class="done-task">${escapeHtml(task.task)}</p>
      <p class="done-detail">${summaryLine(others)}</p>
      <div class="done-actions"><a class="done-link" href="${checklistLink(token)}">Open my checklist</a></div>
    `);
    return;
  }

  const previousStatus = task.status;
  let result;
  try {
    result = await updateStatus(token, id, 'Done');
  } catch (err) {
    failure('Could not reach the server. Nothing was changed.', token);
    return;
  }
  if (result.error) {
    failure(result.error, token);
    return;
  }

  render(`
    <div class="done-icon">&#10003;</div>
    <h1 class="done-title">Marked done</h1>
    <p class="done-task">${escapeHtml(task.task)}</p>
    <p class="done-detail">${summaryLine(others)}</p>
    <div class="done-actions">
      <a class="done-link" href="${checklistLink(token)}">Open my checklist</a>
      <button type="button" class="small-btn secondary" id="undoBtn">Undo</button>
    </div>
  `);

  // Marked it by mistake? Put it back to exactly what it was.
  document.getElementById('undoBtn').addEventListener('click', async (e) => {
    const btn = e.currentTarget;
    btn.disabled = true;
    let undone;
    try {
      undone = await updateStatus(token, id, previousStatus);
    } catch (err) {
      btn.disabled = false;
      alert('Could not reach the server. It is still marked done.');
      return;
    }
    if (undone.error) {
      btn.disabled = false;
      alert(undone.error);
      return;
    }
    render(`
      <div class="done-icon done-icon-neutral">&#8617;</div>
      <h1 class="done-title">Undone</h1>
      <p class="done-task">${escapeHtml(task.task)}</p>
      <p class="done-detail">Back to ${escapeHtml(previousStatus)}.</p>
      <div class="done-actions"><a class="done-link" href="${checklistLink(token)}">Open my checklist</a></div>
    `);
  });
}

main();

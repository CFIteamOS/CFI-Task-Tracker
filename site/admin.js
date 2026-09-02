function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}

function badgeClass(status) {
  return 'badge ' + status.replace(/\s+/g, '-');
}

// Done sits below everything still in play; Blocked sits below even Done,
// since a blocked task needs the least day-to-day attention right now.
function statusRank(status) {
  if (status === 'Blocked') return 2;
  if (status === 'Done') return 1;
  return 0;
}

// Different MoMs tag the same brand inconsistently (SB vs Sharief Bhai, FB
// vs Frozen Bottle, etc.) - this folds known aliases into one canonical
// bucket for the brand summary/filter. Anything not listed here keeps its
// own bracket tag as-is.
const BRAND_ALIASES = {
  'sb': 'Sharief Bhai',
  'sharief bhai': 'Sharief Bhai',
  'fb': 'Frozen Bottle',
  'frozen bottle': 'Frozen Bottle',
  'b2b+cpg': 'B2B+CPG',
  'b2b': 'B2B+CPG',
  'arambam': 'Arambam',
  'me': 'Arambam',
  'mcrc': 'Arambam',
  'all 3': 'Arambam',
  'general': 'General',
  'other': 'General',
  'inventory': 'General'
};

function normalizeBrand(raw) {
  const key = raw.trim().toLowerCase();
  return BRAND_ALIASES[key] || raw.trim();
}

// Tasks are written as "[FB] Do the thing" - the leading bracket tag is the
// brand, kept inline in the task text rather than a separate column (see
// README). Anything without one is grouped under "General" (see aliases).
function extractBrand(taskText) {
  const m = /^\[([^\]]+)\]/.exec(taskText || '');
  return normalizeBrand(m ? m[1] : 'Other');
}

// The admin can assign one task to several owners at once, which creates a
// separate Tracker row (and TaskID) per owner. This groups those rows back
// together for display, keyed on the fields that were identical at
// creation time, so they show as one row with owners joined together.
function groupTasks(tasks) {
  const map = new Map();
  tasks.forEach(t => {
    const key = `${t.task}||${t.meeting}||${t.momDate}`;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(t);
  });
  return Array.from(map.values());
}

function formatRelativeTime(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d)) return '';
  const minutes = Math.floor((Date.now() - d.getTime()) / 60000);
  if (minutes < 1) return 'just now';
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  const days = Math.floor(hours / 24);
  return `${days}d ago`;
}

// "New" is tracked per browser, not per person — there's no per-admin login,
// just the one shared password. lastVisit is read once when the page loads
// (before this session's own visit gets stamped in), so anything commented
// since your last time here stays flagged as new for the rest of THIS
// session, even as the dashboard silently reloads after edits.
const LAST_VISIT_KEY = 'taskTrackerLastVisit';
let sessionLastVisit = null;
let hasStampedThisVisit = false;

function isNewComment(task) {
  if (!task.lastCommentAt || !sessionLastVisit) return false;
  return new Date(task.lastCommentAt) > sessionLastVisit;
}

let allTasks = [];
let ownerNames = [];
let currentBrand = 'all';
let adminPassword = null;
let editingTaskId = null;
let onlyNewComments = false;

async function postAction(body) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  return res.json();
}

async function loadDashboard(password) {
  let data;
  try {
    data = await postAction({ action: 'adminList', password });
  } catch (err) {
    throw new Error('unreachable');
  }
  if (data.error) throw new Error(data.error);
  allTasks = data.tasks;
  ownerNames = data.owners || [];
  adminPassword = password;

  if (!hasStampedThisVisit) {
    const stored = localStorage.getItem(LAST_VISIT_KEY);
    sessionLastVisit = stored ? new Date(stored) : null;
    localStorage.setItem(LAST_VISIT_KEY, new Date().toISOString());
    hasStampedThisVisit = true;
  }

  render();
  populateOwnerPicker();
}

function populateOwnerPicker() {
  const picker = document.getElementById('ownerPicker');
  if (!ownerNames.length) {
    picker.innerHTML = '<summary>No owners in the Owners sheet yet</summary>';
    return;
  }
  picker.innerHTML = `
    <summary id="ownerPickerSummary">Select owner(s)</summary>
    <div class="owner-picker-list">
      <input type="text" id="ownerPickerSearch" class="owner-picker-search" placeholder="Type a name...">
      ${ownerNames.map(n => `
        <label class="owner-picker-item">
          <input type="checkbox" value="${escapeHtml(n)}"> ${escapeHtml(n)}
        </label>
      `).join('')}
    </div>
  `;
  updateOwnerPickerSummary();

  picker.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      // While editing an existing task, only one owner can be selected —
      // checking a new one unchecks any other.
      if (editingTaskId && cb.checked) {
        picker.querySelectorAll('input[type="checkbox"]').forEach(other => {
          if (other !== cb) other.checked = false;
        });
      }
      updateOwnerPickerSummary();
    });
  });

  const search = document.getElementById('ownerPickerSearch');
  search.addEventListener('click', e => e.stopPropagation());
  search.addEventListener('input', () => {
    const q = search.value.trim().toLowerCase();
    picker.querySelectorAll('.owner-picker-item').forEach(item => {
      const name = item.textContent.trim().toLowerCase();
      item.style.display = !q || name.includes(q) ? '' : 'none';
    });
  });
}

function getSelectedOwners() {
  return Array.from(document.querySelectorAll('#ownerPicker input[type="checkbox"]:checked')).map(cb => cb.value);
}

function setSelectedOwners(names) {
  document.querySelectorAll('#ownerPicker input[type="checkbox"]').forEach(cb => {
    cb.checked = names.includes(cb.value);
  });
  updateOwnerPickerSummary();
}

function updateOwnerPickerSummary() {
  const summary = document.getElementById('ownerPickerSummary');
  if (!summary) return;
  const selected = getSelectedOwners();
  summary.textContent = selected.length
    ? (selected.length <= 2 ? selected.join(', ') : `${selected.length} owners selected`)
    : 'Select owner(s)';
}

function render() {
  const total = allTasks.length;
  const doneCount = allTasks.filter(t => t.status === 'Done').length;

  const pct = total ? Math.round((doneCount / total) * 100) : 0;
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressLabel').textContent = `${pct}% complete (${doneCount}/${total})`;

  const newCount = allTasks.filter(isNewComment).length;
  const banner = document.getElementById('newCommentsBanner');
  if (newCount) {
    banner.style.display = 'flex';
    banner.style.cursor = 'pointer';
    banner.textContent = onlyNewComments
      ? `Showing only tasks with new comments — click to show all`
      : `${newCount} task${newCount > 1 ? 's have' : ' has'} new comments since your last visit — click to filter`;
    banner.onclick = () => {
      onlyNewComments = !onlyNewComments;
      render();
    };
  } else {
    banner.style.display = 'none';
    onlyNewComments = false;
  }

  // Pending-by-brand summary: counts tasks not yet Done, grouped by the
  // leading [Bracket] tag in the task text. Each pill also sets the brand
  // filter when clicked.
  const pendingByBrand = {};
  allTasks.forEach(t => {
    if (t.status === 'Done') return;
    const brand = extractBrand(t.task);
    pendingByBrand[brand] = (pendingByBrand[brand] || 0) + 1;
  });
  const brandBar = document.getElementById('brandBar');
  const brands = Object.keys(pendingByBrand).sort((a, b) => pendingByBrand[b] - pendingByBrand[a]);
  brandBar.innerHTML = `
    <button class="brand-pill ${currentBrand === 'all' ? 'active' : ''}" data-brand="all">All brands</button>
    ${brands.map(b => `
      <button class="brand-pill ${currentBrand === b ? 'active' : ''}" data-brand="${escapeHtml(b)}">
        ${escapeHtml(b)} (${pendingByBrand[b]})
      </button>
    `).join('')}
  `;

  let filtered = allTasks;
  if (currentBrand !== 'all') filtered = filtered.filter(t => extractBrand(t.task) === currentBrand);
  if (onlyNewComments) filtered = filtered.filter(isNewComment);

  const table = document.getElementById('table');
  if (!filtered.length) {
    table.innerHTML = '<div class="empty">No matching tasks.</div>';
    return;
  }

  const groups = groupTasks(filtered);

  const rows = groups
    .sort((a, b) => statusRank(a[0].status) - statusRank(b[0].status))
    .map(group => {
      const ids = group.map(t => t.id);
      const owners = group.map(t => t.owner).join(', ');
      const sameStatus = group.every(t => t.status === group[0].status);
      const statusCell = sameStatus
        ? `<span class="${badgeClass(group[0].status)}">${group[0].status}</span>`
        : group.map(t => `<div><span class="${badgeClass(t.status)}">${t.status}</span> <span style="color:var(--muted); font-size:0.78rem;">${escapeHtml(t.owner)}</span></div>`).join('');

      const anyNew = group.some(isNewComment);
      const latest = group.reduce((a, b) =>
        (!a.lastCommentAt || (b.lastCommentAt && new Date(b.lastCommentAt) > new Date(a.lastCommentAt))) ? b : a, group[0]);
      const commentCell = latest.lastCommentAt
        ? `
          <div class="comment-cell-meta">
            ${isNewComment(latest) ? '<span class="badge-new">New</span>' : ''}
            <span class="comment-cell-time">${formatRelativeTime(latest.lastCommentAt)}</span>
          </div>
          <div class="comment-cell-text">${escapeHtml(latest.lastCommentText)}</div>
        `
        : '';

      return `
      <tr data-ids="${escapeHtml(ids.join(','))}" class="${anyNew ? 'row-new' : ''}">
        <td data-label="Task">${escapeHtml(group[0].task)}</td>
        <td data-label="Owner">${escapeHtml(owners)}</td>
        <td data-label="Status">${statusCell}</td>
        <td data-label="Comments">${commentCell}</td>
        <td data-label="Actions">
          <div class="row-actions">
            <button class="small-btn secondary" data-action="edit">Edit</button>
            <button class="small-btn danger" data-action="delete">Delete</button>
          </div>
        </td>
      </tr>
    `;
    }).join('');

  table.innerHTML = `
    <table>
      <thead>
        <tr><th>Task</th><th>Owner</th><th>Status</th><th>Comments</th><th>Actions</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function startEdit(task) {
  editingTaskId = task.id;
  document.getElementById('formTitle').textContent = `Editing task ${task.id}`;
  setSelectedOwners([task.owner]);
  document.getElementById('assignTask').value = task.task;
  document.getElementById('assignBtn').textContent = 'Save changes';
  document.getElementById('cancelEditBtn').style.display = 'inline-block';
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function stopEdit() {
  editingTaskId = null;
  document.getElementById('formTitle').textContent = 'Assign a new task';
  document.getElementById('assignTask').value = '';
  setSelectedOwners([]);
  document.getElementById('assignBtn').textContent = 'Assign task';
  document.getElementById('cancelEditBtn').style.display = 'none';
}

document.getElementById('loginBtn').addEventListener('click', async () => {
  const password = document.getElementById('password').value;
  const errorEl = document.getElementById('loginError');
  errorEl.textContent = '';
  try {
    await loadDashboard(password);
    document.getElementById('login').style.display = 'none';
    document.getElementById('dashboard').style.display = 'block';
  } catch (err) {
    errorEl.textContent = err.message === 'unreachable'
      ? 'Could not reach the server. Check your connection and try again.'
      : 'Incorrect password.';
  }
});

document.getElementById('password').addEventListener('keydown', e => {
  if (e.key === 'Enter') document.getElementById('loginBtn').click();
});

// Brand pills are rebuilt on every render(), so this listens on the
// container rather than on buttons that get thrown away each time.
document.getElementById('brandBar').addEventListener('click', (e) => {
  const btn = e.target.closest('button[data-brand]');
  if (!btn) return;
  currentBrand = btn.dataset.brand;
  render();
});

document.getElementById('cancelEditBtn').addEventListener('click', stopEdit);

document.getElementById('assignBtn').addEventListener('click', async () => {
  const ownerNames = getSelectedOwners();
  const task = document.getElementById('assignTask').value.trim();
  const statusEl = document.getElementById('assignStatus');
  const btn = document.getElementById('assignBtn');

  statusEl.textContent = '';
  statusEl.className = 'assign-form-status';

  if (!ownerNames.length || !task) {
    statusEl.textContent = 'Pick at least one owner and enter a task.';
    statusEl.classList.add('error');
    return;
  }

  btn.disabled = true;
  try {
    const body = editingTaskId
      ? { action: 'updateTask', password: adminPassword, id: editingTaskId, ownerName: ownerNames[0], task }
      : { action: 'createTask', password: adminPassword, ownerNames, task };
    const data = await postAction(body);
    if (data.error) {
      statusEl.textContent = data.error;
      statusEl.classList.add('error');
      return;
    }
    statusEl.textContent = editingTaskId
      ? 'Saved.'
      : (ownerNames.length > 1 ? `Assigned to ${ownerNames.length} people.` : `Assigned to ${ownerNames[0]}.`);
    statusEl.classList.add('success');
    stopEdit();
    await loadDashboard(adminPassword);
  } catch (err) {
    statusEl.textContent = 'Could not reach the server.';
    statusEl.classList.add('error');
  } finally {
    btn.disabled = false;
  }
});

document.getElementById('table').addEventListener('click', async (e) => {
  const btn = e.target.closest('button[data-action]');
  if (!btn) return;
  const row = btn.closest('tr[data-ids]');
  const ids = row.dataset.ids.split(',');
  const group = ids.map(id => allTasks.find(t => t.id === id)).filter(Boolean);
  if (!group.length) return;

  if (btn.dataset.action === 'edit') {
    // A multi-owner group edits as the first owner's copy - editing is
    // restricted to a single owner anyway (see startEdit/populateOwnerPicker).
    startEdit(group[0]);
  } else if (btn.dataset.action === 'delete') {
    const label = group.length > 1
      ? `"${group[0].task}" for ${group.map(t => t.owner).join(', ')}`
      : `"${group[0].task}"`;
    if (!confirm(`Delete ${label}? This can't be undone.`)) return;
    for (const id of ids) {
      const data = await postAction({ action: 'deleteTask', password: adminPassword, id });
      if (data.error) { alert(data.error); return; }
    }
    await loadDashboard(adminPassword);
  }
});

initThemeToggle('themeToggle');

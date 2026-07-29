function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}

function formatDate(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d)) return '';
  return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
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
let currentFilter = 'all';
let adminPassword = null;
let editingTaskId = null;

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
  const newCommentsBtn = document.getElementById('newCommentsFilterBtn');
  if (newCount) {
    banner.style.display = 'flex';
    banner.textContent = `${newCount} task${newCount > 1 ? 's have' : ' has'} new comments since your last visit`;
    newCommentsBtn.style.display = 'inline-block';
    newCommentsBtn.textContent = `New comments (${newCount})`;
  } else {
    banner.style.display = 'none';
    newCommentsBtn.style.display = 'none';
    if (currentFilter === 'newComments') {
      currentFilter = 'all';
      document.querySelectorAll('.filters button').forEach(b => b.classList.remove('active'));
      document.querySelector('.filters button[data-filter="all"]').classList.add('active');
    }
  }

  let filtered = allTasks;
  if (currentFilter === 'open') filtered = allTasks.filter(t => t.status !== 'Done');
  else if (currentFilter === 'newComments') filtered = allTasks.filter(isNewComment);
  else if (currentFilter !== 'all') filtered = allTasks.filter(t => t.status === currentFilter);

  const table = document.getElementById('table');
  if (!filtered.length) {
    table.innerHTML = '<div class="empty">No matching tasks.</div>';
    return;
  }

  const rows = filtered
    .sort((a, b) => statusRank(a.status) - statusRank(b.status))
    .map(t => {
      const isNew = isNewComment(t);
      const commentCell = t.lastCommentAt
        ? `
          <div class="comment-cell-meta">
            ${isNew ? '<span class="badge-new">New</span>' : ''}
            <span class="comment-cell-time">${formatRelativeTime(t.lastCommentAt)}</span>
          </div>
          <div class="comment-cell-text">${escapeHtml(t.lastCommentText)}</div>
        `
        : '';
      return `
      <tr data-id="${escapeHtml(t.id)}" class="${isNew ? 'row-new' : ''}">
        <td data-label="Owner">${escapeHtml(t.owner)}</td>
        <td data-label="Task">${escapeHtml(t.task)}</td>
        <td data-label="Status"><span class="${badgeClass(t.status)}">${t.status}</span></td>
        <td data-label="Revised to">${formatDate(t.revisedTimelineDate)}</td>
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
        <tr><th>Owner</th><th>Task</th><th>Status</th><th>Revised to</th><th>Comments</th><th>Actions</th></tr>
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

document.querySelectorAll('.filters button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.filters button').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentFilter = btn.dataset.filter;
    render();
  });
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
  const row = btn.closest('tr[data-id]');
  const taskId = row.dataset.id;
  const task = allTasks.find(t => t.id === taskId);

  if (btn.dataset.action === 'edit') {
    startEdit(task);
  } else if (btn.dataset.action === 'delete') {
    if (!confirm(`Delete "${task.task}"? This can't be undone.`)) return;
    const data = await postAction({ action: 'deleteTask', password: adminPassword, id: taskId });
    if (data.error) { alert(data.error); return; }
    await loadDashboard(adminPassword);
  }
});

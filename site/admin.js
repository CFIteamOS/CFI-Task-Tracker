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

function formatDateTime(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d)) return '';
  return d.toLocaleString(undefined, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
}

function badgeClass(status) {
  return 'badge ' + status.replace(/\s+/g, '-');
}

let allTasks = [];
let ownerNames = [];
let currentFilter = 'all';
let adminPassword = null;
let editingTaskId = null;
let expandedCommentsId = null;

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
  render();
  populateOwnerDropdown();
}

function populateOwnerDropdown() {
  const select = document.getElementById('assignOwner');
  select.innerHTML = ownerNames.length
    ? ownerNames.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('')
    : '<option value="">No owners in the Owners sheet yet</option>';
}

function render() {
  const summary = document.getElementById('summary');
  const total = allTasks.length;
  const doneCount = allTasks.filter(t => t.status === 'Done').length;
  const notDone = total - doneCount;
  summary.textContent = `${total} total - ${notDone} not done`;

  const pct = total ? Math.round((doneCount / total) * 100) : 0;
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressLabel').textContent = `${pct}% complete (${doneCount}/${total})`;

  let filtered = allTasks;
  if (currentFilter === 'open') filtered = allTasks.filter(t => t.status !== 'Done');
  else if (currentFilter !== 'all') filtered = allTasks.filter(t => t.status === currentFilter);

  const table = document.getElementById('table');
  if (!filtered.length) {
    table.innerHTML = '<div class="empty">No matching tasks.</div>';
    return;
  }

  const rows = filtered
    .sort((a, b) => (a.status === 'Done') - (b.status === 'Done'))
    .map(t => {
      const mainRow = `
        <tr data-id="${escapeHtml(t.id)}">
          <td>${escapeHtml(t.owner)}</td>
          <td>${escapeHtml(t.task)}</td>
          <td><span class="${badgeClass(t.status)}">${t.status}</span></td>
          <td>${formatDate(t.revisedTimelineDate)}</td>
          <td>
            <div class="row-actions">
              <button class="small-btn secondary" data-action="edit">Edit</button>
              <button class="small-btn danger" data-action="delete">Delete</button>
              <button class="comments-toggle" data-action="comments">Comments</button>
            </div>
          </td>
        </tr>
      `;
      const commentsRow = expandedCommentsId === t.id
        ? `<tr class="comments-row" data-comments-for="${escapeHtml(t.id)}"><td colspan="5"><div class="comment-list" id="commentList-${escapeHtml(t.id)}">Loading...</div></td></tr>`
        : '';
      return mainRow + commentsRow;
    }).join('');

  table.innerHTML = `
    <table>
      <thead>
        <tr><th>Owner</th><th>Task</th><th>Status</th><th>Revised to</th><th>Actions</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  if (expandedCommentsId) loadComments(expandedCommentsId);
}

// Admin can view comments (read-only) but not add them — only the task's
// owner can, from their own checklist.
async function loadComments(taskId) {
  const listEl = document.getElementById(`commentList-${taskId}`);
  if (!listEl) return;
  const data = await postAction({ action: 'getComments', password: adminPassword, id: taskId });
  if (data.error) {
    listEl.innerHTML = `<div class="error">${escapeHtml(data.error)}</div>`;
    return;
  }
  listEl.innerHTML = data.comments.length
    ? data.comments.map(c => `
        <div class="comment-item">
          <div class="comment-meta">${escapeHtml(c.author)} - ${formatDateTime(c.timestamp)}</div>
          <div>${escapeHtml(c.text)}</div>
        </div>
      `).join('')
    : '<div class="empty" style="padding:8px 0">No comments yet.</div>';
}

function startEdit(task) {
  editingTaskId = task.id;
  document.getElementById('formTitle').textContent = `Editing task ${task.id}`;
  document.getElementById('assignOwner').value = task.owner;
  document.getElementById('assignTask').value = task.task;
  document.getElementById('assignBtn').textContent = 'Save changes';
  document.getElementById('cancelEditBtn').style.display = 'inline-block';
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function stopEdit() {
  editingTaskId = null;
  document.getElementById('formTitle').textContent = 'Assign a new task';
  document.getElementById('assignTask').value = '';
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
  const ownerName = document.getElementById('assignOwner').value;
  const task = document.getElementById('assignTask').value.trim();
  const statusEl = document.getElementById('assignStatus');
  const btn = document.getElementById('assignBtn');

  statusEl.textContent = '';
  statusEl.className = 'assign-form-status';

  if (!ownerName || !task) {
    statusEl.textContent = 'Pick an owner and enter a task.';
    statusEl.classList.add('error');
    return;
  }

  btn.disabled = true;
  try {
    const body = editingTaskId
      ? { action: 'updateTask', password: adminPassword, id: editingTaskId, ownerName, task }
      : { action: 'createTask', password: adminPassword, ownerName, task };
    const data = await postAction(body);
    if (data.error) {
      statusEl.textContent = data.error;
      statusEl.classList.add('error');
      return;
    }
    statusEl.textContent = editingTaskId ? 'Saved.' : `Assigned to ${ownerName}.`;
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
  } else if (btn.dataset.action === 'comments') {
    expandedCommentsId = expandedCommentsId === taskId ? null : taskId;
    render();
  }
});

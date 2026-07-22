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

let allTasks = [];
let ownerNames = [];
let currentFilter = 'all';
let adminPassword = null;

async function loadDashboard(password) {
  let res;
  try {
    res = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'adminList', password })
    });
  } catch (err) {
    throw new Error('unreachable');
  }
  const data = await res.json();
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
  const notDone = allTasks.filter(t => t.status !== 'Done').length;
  summary.textContent = `${allTasks.length} total - ${notDone} not done`;

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
    .map(t => `
      <tr>
        <td>${escapeHtml(t.owner)}</td>
        <td>${escapeHtml(t.task)}</td>
        <td>${escapeHtml(t.meeting)}</td>
        <td>${formatDate(t.momDate)}</td>
        <td><span class="${badgeClass(t.status)}">${t.status}</span></td>
        <td>${formatDate(t.revisedTimelineDate)}</td>
        <td>${formatDate(t.dueDate)}</td>
        <td>${t.reminderCount || 0}</td>
      </tr>
    `).join('');

  table.innerHTML = `
    <table>
      <thead>
        <tr><th>Owner</th><th>Task</th><th>Meeting</th><th>MoM date</th><th>Status</th><th>Revised to</th><th>Due</th><th>Reminders</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
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

document.getElementById('assignBtn').addEventListener('click', async () => {
  const ownerName = document.getElementById('assignOwner').value;
  const task = document.getElementById('assignTask').value.trim();
  const dueDate = document.getElementById('assignDueDate').value;
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
    const res = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'createTask', password: adminPassword, ownerName, task, dueDate })
    });
    const data = await res.json();
    if (data.error) {
      statusEl.textContent = data.error;
      statusEl.classList.add('error');
      return;
    }
    statusEl.textContent = `Assigned to ${ownerName}.`;
    statusEl.classList.add('success');
    document.getElementById('assignTask').value = '';
    document.getElementById('assignDueDate').value = '';
    await loadDashboard(adminPassword);
  } catch (err) {
    statusEl.textContent = 'Could not reach the server.';
    statusEl.classList.add('error');
  } finally {
    btn.disabled = false;
  }
});

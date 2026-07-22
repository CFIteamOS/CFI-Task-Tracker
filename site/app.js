const STATUSES = ['Pending', 'In Progress', 'Blocked', 'Revised Timeline', 'Done'];

function badgeClass(status) {
  return 'badge ' + status.replace(/\s+/g, '-');
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

async function callApi(params) {
  const url = new URL(API_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  return res.json();
}

async function postApi(body) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  return res.json();
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

let expandedCommentsId = null;

function renderTask(task, token, onChange) {
  const row = document.createElement('div');
  row.className = 'card';

  const isDone = task.status === 'Done';
  const isOverdue = !isDone && task.dueDate && new Date(task.dueDate) < new Date(new Date().toDateString());
  const dueHtml = task.dueDate
    ? `<span class="sep">-</span><span class="${isOverdue ? 'due-overdue' : ''}">Due ${formatDate(task.dueDate)}</span>`
    : '';
  row.innerHTML = `
    <div class="task-row">
      <input type="checkbox" ${isDone ? 'checked' : ''}>
      <div class="task-main">
        <div class="task-text ${isDone ? 'done' : ''}">${escapeHtml(task.task)}</div>
        <div class="task-meta">
          <span>${escapeHtml(task.meeting || '')}</span>
          <span class="sep">-</span>
          <span>${formatDate(task.momDate)}</span>
          ${dueHtml}
          <span class="sep">-</span>
          <span class="${badgeClass(task.status)}">${task.status}</span>
        </div>
        <div class="controls"></div>
        <button class="comments-toggle" style="margin-top:8px">Comments</button>
        <div class="comments-area" style="display:none; margin-top:10px"></div>
      </div>
    </div>
  `;

  const checkbox = row.querySelector('input[type="checkbox"]');
  const controls = row.querySelector('.controls');
  const commentsToggle = row.querySelector('.comments-toggle');
  const commentsArea = row.querySelector('.comments-area');

  function renderControls() {
    controls.innerHTML = '';
    if (checkbox.checked) return;

    const select = document.createElement('select');
    ['Pending', 'In Progress', 'Blocked', 'Revised Timeline'].forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      if (s === task.status) opt.selected = true;
      select.appendChild(opt);
    });

    const dateInput = document.createElement('input');
    dateInput.type = 'date';
    dateInput.style.marginLeft = '8px';
    dateInput.style.display = select.value === 'Revised Timeline' ? 'inline-block' : 'none';
    if (task.revisedTimelineDate) {
      dateInput.value = new Date(task.revisedTimelineDate).toISOString().slice(0, 10);
    }

    select.addEventListener('change', async () => {
      dateInput.style.display = select.value === 'Revised Timeline' ? 'inline-block' : 'none';
      if (select.value !== 'Revised Timeline') {
        await update(select.value, null);
      }
    });
    dateInput.addEventListener('change', async () => {
      await update('Revised Timeline', dateInput.value);
    });

    controls.appendChild(select);
    controls.appendChild(dateInput);
  }

  async function update(status, revisedTimelineDate) {
    checkbox.disabled = true;
    const result = await postApi({ action: 'updateStatus', token, id: task.id, status, revisedTimelineDate });
    checkbox.disabled = false;
    if (result.error) {
      alert('Could not update task: ' + result.error);
      return;
    }
    task.status = status;
    if (revisedTimelineDate) task.revisedTimelineDate = revisedTimelineDate;
    onChange();
  }

  checkbox.addEventListener('change', async () => {
    await update(checkbox.checked ? 'Done' : 'Pending', null);
  });

  async function loadComments() {
    commentsArea.innerHTML = '<div class="comment-list">Loading...</div><div class="comment-add"><input type="text" placeholder="Add a comment"><button class="small-btn secondary">Add</button></div>';
    const data = await postApi({ action: 'getComments', token, id: task.id });
    const list = commentsArea.querySelector('.comment-list');
    if (data.error) {
      list.innerHTML = `<div class="error">${escapeHtml(data.error)}</div>`;
      return;
    }
    list.innerHTML = data.comments.length
      ? data.comments.map(c => `
          <div class="comment-item">
            <div class="comment-meta">${escapeHtml(c.author)} - ${formatDateTime(c.timestamp)}</div>
            <div>${escapeHtml(c.text)}</div>
          </div>
        `).join('')
      : '<div class="empty" style="padding:8px 0">No comments yet.</div>';

    const input = commentsArea.querySelector('input');
    const addBtn = commentsArea.querySelector('button');
    addBtn.addEventListener('click', async () => {
      const text = input.value.trim();
      if (!text) return;
      addBtn.disabled = true;
      const result = await postApi({ action: 'addComment', token, id: task.id, text });
      addBtn.disabled = false;
      if (result.error) { alert(result.error); return; }
      input.value = '';
      loadComments();
    });
  }

  commentsToggle.addEventListener('click', () => {
    const isOpen = commentsArea.style.display !== 'none';
    commentsArea.style.display = isOpen ? 'none' : 'block';
    if (!isOpen) loadComments();
  });

  renderControls();
  return row;
}

async function init() {
  const params = new URLSearchParams(window.location.search);
  const token = params.get('token');
  const content = document.getElementById('content');
  const subtitle = document.getElementById('subtitle');
  const ownTaskForm = document.getElementById('ownTaskForm');

  if (!token) {
    subtitle.textContent = '';
    content.innerHTML = '<div class="error">No access token in the URL. Use the personal link from your email.</div>';
    return;
  }

  content.innerHTML = '<div class="loading">Loading your tasks...</div>';

  let data;
  try {
    data = await callApi({ action: 'getTasks', token });
  } catch (err) {
    subtitle.textContent = '';
    content.innerHTML = '<div class="error">Could not reach the server. Check your connection and try again.</div>';
    return;
  }

  if (data.error) {
    subtitle.textContent = '';
    content.innerHTML = `<div class="error">${escapeHtml(data.error)}</div>`;
    return;
  }

  ownTaskForm.style.display = 'flex';

  function rerender() {
    subtitle.textContent = `${data.owner} - ${data.tasks.filter(t => t.status !== 'Done').length} open`;
    content.innerHTML = '';
    if (!data.tasks.length) {
      content.innerHTML = '<div class="empty">No action items yet.</div>';
      return;
    }
    const sorted = [...data.tasks].sort((a, b) => (a.status === 'Done') - (b.status === 'Done'));
    sorted.forEach(task => content.appendChild(renderTask(task, token, rerender)));
  }

  rerender();

  document.getElementById('ownTaskBtn').addEventListener('click', async () => {
    const textInput = document.getElementById('ownTaskText');
    const dueInput = document.getElementById('ownTaskDueDate');
    const text = textInput.value.trim();
    if (!text) return;
    const btn = document.getElementById('ownTaskBtn');
    btn.disabled = true;
    const result = await postApi({ action: 'createOwnTask', token, task: text, dueDate: dueInput.value });
    btn.disabled = false;
    if (result.error) {
      alert('Could not add task: ' + result.error);
      return;
    }
    data.tasks.push({
      id: result.id, task: text, meeting: 'Added by owner', momDate: new Date().toISOString(),
      status: 'Pending', dueDate: dueInput.value || null
    });
    textInput.value = '';
    dueInput.value = '';
    rerender();
  });
}

init();

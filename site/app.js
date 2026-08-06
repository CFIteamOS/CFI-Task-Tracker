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

// Comments are entered as plain text with **bold** markers and real line
// breaks (via the textarea) — this renders those two things as HTML without
// allowing any other markup through, since the raw text is escaped first.
function formatCommentText(text) {
  return escapeHtml(text)
    .replace(/\n/g, '<br>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
}

// Done sits below everything still in play; Blocked sits below even Done,
// since a blocked task needs the least day-to-day attention right now.
function statusRank(status) {
  if (status === 'Blocked') return 2;
  if (status === 'Done') return 1;
  return 0;
}

let expandedCommentsId = null;

function renderTask(task, token, onChange) {
  const row = document.createElement('div');
  row.className = 'card';

  const isDone = task.status === 'Done';
  row.innerHTML = `
    <div class="task-row">
      <input type="checkbox" ${isDone ? 'checked' : ''}>
      <div class="task-main">
        <div class="task-text ${isDone ? 'done' : ''}">${escapeHtml(task.task)}</div>
        <div class="task-meta">
          <span>${escapeHtml(task.meeting || '')}</span>
          <span class="sep">-</span>
          <span>${formatDate(task.momDate)}</span>
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
    commentsArea.innerHTML = `
      <div class="comment-list">Loading...</div>
      <div class="comment-add">
        <div class="comment-add-toolbar">
          <button type="button" class="small-btn secondary comment-bold-btn" title="Bold the selected text"><strong>B</strong></button>
        </div>
        <textarea rows="2" placeholder="Add a comment - use **bold**, press Enter for a new line"></textarea>
        <button class="small-btn secondary comment-add-btn">Add</button>
      </div>
    `;
    const data = await postApi({ action: 'getComments', token, id: task.id });
    const list = commentsArea.querySelector('.comment-list');
    if (data.error) {
      list.innerHTML = `<div class="error">${escapeHtml(data.error)}</div>`;
      return;
    }
    list.innerHTML = data.comments.length
      ? data.comments.map(c => `
          <div class="comment-item" data-comment-id="${escapeHtml(c.id)}">
            <div class="comment-meta">
              <span>${escapeHtml(c.author)} - ${formatDateTime(c.timestamp)}</span>
              <button type="button" class="comment-delete-btn">Delete</button>
            </div>
            <div class="comment-text">${formatCommentText(c.text)}</div>
          </div>
        `).join('')
      : '<div class="empty" style="padding:8px 0">No comments yet.</div>';

    list.querySelectorAll('.comment-delete-btn').forEach(btn => {
      btn.addEventListener('click', async () => {
        if (!confirm('Delete this comment?')) return;
        const commentId = btn.closest('.comment-item').dataset.commentId;
        btn.disabled = true;
        const result = await postApi({ action: 'deleteComment', token, id: task.id, commentId });
        if (result.error) { alert(result.error); return; }
        loadComments();
      });
    });

    const textarea = commentsArea.querySelector('textarea');
    const boldBtn = commentsArea.querySelector('.comment-bold-btn');
    const addBtn = commentsArea.querySelector('.comment-add-btn');

    boldBtn.addEventListener('click', () => {
      const start = textarea.selectionStart;
      const end = textarea.selectionEnd;
      const value = textarea.value;
      const selected = value.slice(start, end) || 'bold text';
      textarea.value = value.slice(0, start) + '**' + selected + '**' + value.slice(end);
      textarea.focus();
      textarea.selectionStart = start + 2;
      textarea.selectionEnd = start + 2 + selected.length;
    });

    addBtn.addEventListener('click', async () => {
      const text = textarea.value.trim();
      if (!text) return;
      addBtn.disabled = true;
      const result = await postApi({ action: 'addComment', token, id: task.id, text });
      addBtn.disabled = false;
      if (result.error) { alert(result.error); return; }
      textarea.value = '';
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
  const content = document.getElementById('content');
  const subtitle = document.getElementById('subtitle');
  const ownTaskForm = document.getElementById('ownTaskForm');

  // A Home Screen icon (Android/iOS "Add to Home Screen") always relaunches
  // the app's static start URL, which has no token in it — so remember the
  // last-used token here and fall back to it whenever the URL is missing one.
  const TOKEN_KEY = 'taskTrackerToken';
  let token = params.get('token');
  if (token) {
    localStorage.setItem(TOKEN_KEY, token);
  } else {
    token = localStorage.getItem(TOKEN_KEY);
  }

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
    const total = data.tasks.length;
    const doneCount = data.tasks.filter(t => t.status === 'Done').length;
    const pct = total ? Math.round((doneCount / total) * 100) : 0;
    document.getElementById('progressFill').style.width = pct + '%';
    document.getElementById('progressLabel').textContent = total ? `${pct}% complete (${doneCount}/${total})` : '';

    subtitle.textContent = `${data.owner} - ${data.tasks.filter(t => t.status !== 'Done').length} open`;
    content.innerHTML = '';
    if (!data.tasks.length) {
      content.innerHTML = '<div class="empty">No action items yet.</div>';
      return;
    }
    const sorted = [...data.tasks].sort((a, b) => statusRank(a.status) - statusRank(b.status));
    sorted.forEach(task => content.appendChild(renderTask(task, token, rerender)));
  }

  rerender();

  document.getElementById('ownTaskBtn').addEventListener('click', async () => {
    const textInput = document.getElementById('ownTaskText');
    const text = textInput.value.trim();
    if (!text) return;
    const btn = document.getElementById('ownTaskBtn');
    btn.disabled = true;
    const result = await postApi({ action: 'createOwnTask', token, task: text });
    btn.disabled = false;
    if (result.error) {
      alert('Could not add task: ' + result.error);
      return;
    }
    data.tasks.push({
      id: result.id, task: text, meeting: 'Added by owner', momDate: new Date().toISOString(),
      status: 'Pending'
    });
    textInput.value = '';
    rerender();
  });
}

initThemeToggle('themeToggle');
init();

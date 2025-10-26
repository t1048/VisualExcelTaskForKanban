let api;
let RUN_MODE = 'mock';
let TASKS = [];
let STATUSES = [];

window.addEventListener('pywebviewready', async () => {
  try {
    api = window.pywebview.api;
    RUN_MODE = 'pywebview';
    await init(true);
  } catch (err) {
    console.error('pywebviewready failed', err);
  }
});

document.addEventListener('DOMContentLoaded', async () => {
  api = window.pywebview?.api || createMockApi();
  if (window.pywebview?.api) RUN_MODE = 'pywebview';
  await init(true);
  wireControls();
});

function createMockApi() {
  const today = new Date();
  const format = d => d.toISOString().slice(0, 10);
  return {
    async get_tasks() {
      const sample = [];
      for (let i = 0; i < 6; i++) {
        const due = new Date(today);
        due.setDate(today.getDate() + i * 2);
        sample.push({
          No: i + 1,
          ステータス: ['未着手', '進行中', '完了'][i % 3],
          タスク: `サンプルタスク ${i + 1}`,
          担当者: ['田中', '佐藤', '鈴木'][i % 3],
          期限: format(due),
          備考: 'モックデータ'
        });
      }
      return sample;
    },
    async get_statuses() {
      return ['未着手', '進行中', '完了', '保留'];
    }
  };
}

async function init(force = false) {
  if (!api) return;
  if (force) {
    try {
      if (RUN_MODE === 'pywebview' && typeof api.reload_from_excel === 'function') {
        const result = await api.reload_from_excel();
        TASKS = result?.tasks || await api.get_tasks();
        STATUSES = result?.statuses || await api.get_statuses?.() || [];
      } else {
        STATUSES = typeof api.get_statuses === 'function' ? await api.get_statuses() : [];
        TASKS = await api.get_tasks();
      }
    } catch (err) {
      console.error('init failed', err);
      TASKS = [];
    }
  }

  ensureRangeDefaults();
  renderSummary();
  renderLegend();
  renderAssigneeFilter();
  renderTimeline();
}

function ensureRangeDefaults() {
  const inputFrom = document.getElementById('date-from');
  const inputTo = document.getElementById('date-to');
  if (inputFrom.value && inputTo.value) return;

  const validDueDates = TASKS.map(t => parseISO(t.期限)).filter(Boolean).sort((a, b) => a - b);
  let start;
  let end;
  if (validDueDates.length >= 1) {
    start = new Date(validDueDates[0]);
    end = new Date(validDueDates[validDueDates.length - 1]);
  } else {
    const today = new Date();
    start = startOfWeek(today);
    end = new Date(start);
    end.setDate(end.getDate() + 13);
  }

  // guard against overly wide ranges
  if ((end - start) / (1000 * 60 * 60 * 24) > 60) {
    end = new Date(start);
    end.setDate(start.getDate() + 30);
  }

  inputFrom.value = toISODate(start);
  inputTo.value = toISODate(end);
}

function wireControls() {
  document.getElementById('btn-apply').addEventListener('click', () => {
    renderSummary();
    renderAssigneeFilter();
    renderTimeline();
  });

  document.querySelectorAll('.quick button[data-shift]').forEach(btn => {
    btn.addEventListener('click', () => {
      const delta = Number(btn.dataset.shift || '0');
      shiftRange(delta);
      renderSummary();
      renderAssigneeFilter();
      renderTimeline();
    });
  });

  const btnThisWeek = document.querySelector('.quick button[data-range="this-week"]');
  btnThisWeek.addEventListener('click', () => {
    const today = new Date();
    const start = startOfWeek(today);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    setRange(start, end);
    renderSummary();
    renderAssigneeFilter();
    renderTimeline();
  });

  const assigneeSelect = document.getElementById('assignee-filter');
  assigneeSelect.addEventListener('change', () => {
    renderTimeline();
  });
}

function shiftRange(deltaDays) {
  const fromInput = document.getElementById('date-from');
  const toInput = document.getElementById('date-to');
  const from = parseISO(fromInput.value);
  const to = parseISO(toInput.value);
  if (!from || !to) return;
  from.setDate(from.getDate() + deltaDays);
  to.setDate(to.getDate() + deltaDays);
  setRange(from, to);
}

function setRange(from, to) {
  const inputFrom = document.getElementById('date-from');
  const inputTo = document.getElementById('date-to');
  inputFrom.value = toISODate(from);
  inputTo.value = toISODate(to);
}

function renderSummary() {
  const from = parseISO(document.getElementById('date-from').value);
  const to = parseISO(document.getElementById('date-to').value);
  const summary = document.getElementById('summary');
  if (!from || !to) {
    summary.textContent = '開始日と終了日を指定してください。';
    return;
  }
  const diff = Math.round((to - from) / (1000 * 60 * 60 * 24)) + 1;
  const inRange = TASKS.filter(t => {
    const due = parseISO(t.期限);
    return due && due >= from && due <= to;
  }).length;
  const withoutDue = TASKS.filter(t => !parseISO(t.期限)).length;
  summary.textContent = `表示期間: ${toLocale(from)} 〜 ${toLocale(to)} （${diff}日間） / 対象タスク ${inRange} 件、期限未設定 ${withoutDue} 件`;
}

function renderAssigneeFilter() {
  const select = document.getElementById('assignee-filter');
  if (!select) return;

  const previous = select.value;
  const assignees = collectAllAssignees();
  select.innerHTML = '';

  const optionAll = document.createElement('option');
  optionAll.value = '';
  optionAll.textContent = 'すべて';
  select.appendChild(optionAll);

  assignees.forEach(name => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });

  if (previous && assignees.includes(previous)) {
    select.value = previous;
  } else {
    select.value = '';
  }
}

function renderLegend() {
  const legend = document.getElementById('legend');
  legend.innerHTML = '';
  const baseStatuses = Array.isArray(STATUSES) && STATUSES.length
    ? STATUSES
    : Array.from(new Set(TASKS.map(t => t.ステータス).filter(Boolean)));
  const seen = new Set();
  baseStatuses.forEach(name => {
    const dot = document.createElement('span');
    dot.className = 'legend-dot';
    dot.style.background = statusColor(name);
    const item = document.createElement('span');
    item.className = 'legend-item';
    item.appendChild(dot);
    const label = document.createElement('span');
    label.textContent = name;
    item.appendChild(label);
    legend.appendChild(item);
    seen.add(name);
  });

  const hasOther = TASKS.some(t => t.ステータス && !seen.has(t.ステータス));
  if (hasOther) {
    const dot = document.createElement('span');
    dot.className = 'legend-dot';
    dot.style.background = statusColor('other');
    const item = document.createElement('span');
    item.className = 'legend-item';
    item.appendChild(dot);
    const label = document.createElement('span');
    label.textContent = 'その他';
    item.appendChild(label);
    legend.appendChild(item);
  }
}

function renderTimeline() {
  const wrapper = document.getElementById('timeline-wrapper');
  const from = parseISO(document.getElementById('date-from').value);
  const to = parseISO(document.getElementById('date-to').value);
  const assigneeFilter = document.getElementById('assignee-filter')?.value || '';
  if (!from || !to || from > to) {
    wrapper.innerHTML = '<div class="message">期間の指定が正しくありません。</div>';
    return;
  }

  const days = enumerateDays(from, to);
  if (!days.length) {
    wrapper.innerHTML = '<div class="message">表示できる日付がありません。</div>';
    return;
  }

  const assignees = collectAssignees(from, to);
  const filteredAssignees = assigneeFilter ? assignees.filter(name => name === assigneeFilter) : assignees;
  if (!filteredAssignees.length) {
    if (assigneeFilter) {
      wrapper.innerHTML = '<div class="message">選択した担当者のタスクが表示期間内にありません。</div>';
    } else {
      wrapper.innerHTML = '<div class="message">表示できるタスクがありません。</div>';
    }
    return;
  }

  const table = document.createElement('table');
  table.className = 'timeline';

  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');
  const thCorner = document.createElement('th');
  thCorner.textContent = '担当者';
  trHead.appendChild(thCorner);
  days.forEach(day => {
    const th = document.createElement('th');
    const label = document.createElement('div');
    label.className = 'day-label';
    const dateSpan = document.createElement('span');
    dateSpan.className = 'date';
    dateSpan.textContent = `${day.month}/${day.date}`;
    const weekdaySpan = document.createElement('span');
    weekdaySpan.className = 'weekday';
    weekdaySpan.textContent = day.weekday;
    label.appendChild(dateSpan);
    label.appendChild(weekdaySpan);
    th.appendChild(label);
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const taskMap = buildTaskLookup(from, to, assigneeFilter);

  filteredAssignees.forEach(name => {
    const tr = document.createElement('tr');
    const th = document.createElement('th');
    th.textContent = name;
    tr.appendChild(th);

    days.forEach(day => {
      const td = document.createElement('td');
      const key = day.iso;
      const items = (taskMap.get(name)?.get(key)) || [];
      if (!items.length) {
        const empty = document.createElement('div');
        empty.className = 'empty-cell';
        empty.textContent = '-';
        td.appendChild(empty);
      } else {
        items.forEach(task => {
          td.appendChild(renderTaskChip(task));
        });
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);

  wrapper.innerHTML = '';
  wrapper.appendChild(table);
}

function renderTaskChip(task) {
  const div = document.createElement('div');
  div.className = 'task-chip';
  if (task.ステータス) {
    const rawClass = `status-${task.ステータス}`;
    div.classList.add(rawClass);
    div.classList.add(sanitizeClass(rawClass));
  } else {
    div.classList.add('status-other');
  }

  const title = document.createElement('div');
  title.className = 'task-title';
  title.textContent = task.タスク || '(タイトルなし)';
  div.appendChild(title);

  const meta = document.createElement('div');
  meta.className = 'task-meta';
  if (task.ステータス) {
    const status = document.createElement('span');
    status.textContent = `ステータス: ${task.ステータス}`;
    meta.appendChild(status);
  }
  if (task.No) {
    const no = document.createElement('span');
    no.textContent = `No.${task.No}`;
    meta.appendChild(no);
  }
  if (task.備考) {
    const notes = document.createElement('span');
    notes.textContent = task.備考;
    meta.appendChild(notes);
  }
  div.appendChild(meta);
  return div;
}

function buildTaskLookup(from, to, assigneeFilter = '') {
  const map = new Map();
  TASKS.forEach(task => {
    const due = parseISO(task.期限);
    if (!due || due < from || due > to) return;
    const assignee = task.担当者?.trim() || '（担当者未設定）';
    if (assigneeFilter && assignee !== assigneeFilter) return;
    if (!map.has(assignee)) map.set(assignee, new Map());
    const byDate = map.get(assignee);
    const key = toISODate(due);
    if (!byDate.has(key)) byDate.set(key, []);
    byDate.get(key).push(task);
  });

  // sort each bucket by No -> title
  map.forEach(byDate => {
    byDate.forEach(list => {
      list.sort((a, b) => {
        if (a.No && b.No && a.No !== b.No) return a.No - b.No;
        return (a.タスク || '').localeCompare(b.タスク || '');
      });
    });
  });
  return map;
}

function collectAssignees(rangeFrom, rangeTo) {
  const assignees = new Set();
  TASKS.forEach(task => {
    const due = parseISO(task.期限);
    if (!due) return;
    if (!rangeFrom || !rangeTo || due < rangeFrom || due > rangeTo) return;
    assignees.add(task.担当者?.trim() || '（担当者未設定）');
  });
  return Array.from(assignees).sort((a, b) => a.localeCompare(b, 'ja'));
}

function collectAllAssignees() {
  const assignees = new Set();
  TASKS.forEach(task => {
    assignees.add(task.担当者?.trim() || '（担当者未設定）');
  });
  return Array.from(assignees).sort((a, b) => a.localeCompare(b, 'ja'));
}

function enumerateDays(from, to) {
  const days = [];
  const names = ['日', '月', '火', '水', '木', '金', '土'];
  const cursor = new Date(from);
  while (cursor <= to) {
    days.push({
      iso: toISODate(cursor),
      month: cursor.getMonth() + 1,
      date: cursor.getDate(),
      weekday: `${names[cursor.getDay()]}曜`
    });
    cursor.setDate(cursor.getDate() + 1);
  }
  return days;
}

function parseISO(value) {
  if (!value) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!m) return null;
  const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(dt.getTime()) ? null : dt;
}

function toISODate(date) {
  const tzOffset = date.getTimezoneOffset() * 60000;
  return new Date(date.getTime() - tzOffset).toISOString().slice(0, 10);
}

function toLocale(date) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}

function startOfWeek(date) {
  const start = new Date(date);
  start.setDate(date.getDate() - date.getDay());
  start.setHours(0, 0, 0, 0);
  return start;
}

function statusColor(name) {
  switch (name) {
    case '未着手':
      return 'rgba(248, 113, 113, 0.8)';
    case '進行中':
      return 'rgba(56, 189, 248, 0.8)';
    case '完了':
      return 'rgba(34, 197, 94, 0.8)';
    case '保留':
      return 'rgba(250, 204, 21, 0.8)';
    default:
      return 'rgba(129, 140, 248, 0.8)';
  }
}

function sanitizeClass(value) {
  return value.replace(/[^\w-]/g, '-');
}

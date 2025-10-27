let api;
let RUN_MODE = 'mock';
let TASKS = [];
let STATUSES = [];
const PRIORITY_DEFAULT_OPTIONS = ['高', '中', '低'];
const UNSET_STATUS_LABEL = 'ステータス未設定';
let VALIDATIONS = {};
let CURRENT_EDIT = null;
const ASSIGNEE_FILTER_ALL = '';
const ASSIGNEE_FILTER_UNASSIGNED = '__UNASSIGNED__';
const ASSIGNEE_UNASSIGNED_LABEL = '（未割り当て）';

function sanitizeTaskRecord(task, fallbackIndex = 0) {
  if (!task || typeof task !== 'object') return null;
  const title = String(task.タスク ?? '').trim();
  if (!title) return null;
  const sanitized = { ...task, タスク: title };
  const noValue = sanitized.No;
  const noText = noValue === null || noValue === undefined ? '' : String(noValue).trim();
  if (!noText) {
    sanitized.No = fallbackIndex + 1;
  }
  return sanitized;
}

function sanitizeTaskList(rawList) {
  if (!Array.isArray(rawList)) return [];
  const result = [];
  rawList.forEach(item => {
    const sanitized = sanitizeTaskRecord(item, result.length);
    if (sanitized) {
      result.push(sanitized);
    }
  });
  return result;
}

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
  const baseStatuses = ['未着手', '進行中', '完了', '保留'];
  const statusSet = new Set(baseStatuses);
  const pad = n => String(n).padStart(2, '0');
  const toISO = date => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  const today = new Date();
  const majorCategories = ['プロジェクトA', 'プロジェクトB', 'プロジェクトC'];
  const minorCategories = ['企画', '設計', '実装', '検証'];
  const sampleTasks = Array.from({ length: 8 }).map((_, idx) => {
    const due = new Date(today);
    due.setDate(today.getDate() + idx - 2);
    const status = baseStatuses[idx % baseStatuses.length];
    statusSet.add(status);
    const major = majorCategories[idx % majorCategories.length];
    const minor = minorCategories[idx % minorCategories.length];
    return {
      ステータス: status,
      大分類: major,
      中分類: minor,
      タスク: `サンプルタスク ${idx + 1}`,
      担当者: ['田中', '佐藤', '鈴木', '高橋'][idx % 4],
      優先度: ['高', '中', '低'][idx % 3],
      期限: toISO(due),
      備考: idx % 2 === 0 ? 'モックデータ' : ''
    };
  });
  const tasks = [...sampleTasks];
  let validations = {
    'ステータス': Array.from(statusSet),
    '大分類': Array.from(new Set(majorCategories)),
    '中分類': Array.from(new Set(minorCategories)),
    '優先度': [...PRIORITY_DEFAULT_OPTIONS]
  };

  const cloneTask = task => ({ ...task });

  const sanitizeStatus = status => {
    const text = String(status ?? '').trim();
    if (text) return text;
    return baseStatuses[0];
  };

  const normalizeDue = value => {
    if (!value) return '';
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return '';
    return toISO(parsed);
  };

  const normalizePriority = value => {
    if (value === null || value === undefined) return '';
    const text = String(value).trim();
    return text;
  };

  const normalizeTask = (payload) => {
    const status = sanitizeStatus(payload?.ステータス);
    statusSet.add(status);

    const major = String(payload?.大分類 ?? '').trim();
    const minor = String(payload?.中分類 ?? '').trim();

    const title = String(payload?.タスク ?? '').trim();
    if (!title) {
      throw new Error('タスクは必須です');
    }

    return {
      ステータス: status,
      大分類: major,
      中分類: minor,
      タスク: title,
      担当者: String(payload?.担当者 ?? '').trim(),
      優先度: normalizePriority(payload?.優先度),
      期限: normalizeDue(payload?.期限),
      備考: String(payload?.備考 ?? '')
    };
  };

  const withSequentialNo = () => tasks.map((task, idx) => ({ ...cloneTask(task), No: idx + 1 }));

  const locateTask = (no) => {
    const idx = Number(no) - 1;
    if (!Number.isInteger(idx) || idx < 0 || idx >= tasks.length) {
      return null;
    }
    return { index: idx, record: tasks[idx] };
  };

  const updateValidations = (payload) => {
    const withFallbacks = (source) => {
      const merged = { ...source };
      if (!Array.isArray(merged['ステータス']) || merged['ステータス'].length === 0) {
        merged['ステータス'] = Array.from(statusSet);
      }
      if (!Array.isArray(merged['優先度']) || merged['優先度'].length === 0) {
        merged['優先度'] = [...PRIORITY_DEFAULT_OPTIONS];
      }
      return merged;
    };

    if (!payload || typeof payload !== 'object') {
      validations = withFallbacks({});
      return validations;
    }
    const cleaned = {};
    Object.keys(payload).forEach(key => {
      const raw = Array.isArray(payload[key]) ? payload[key] : [];
      const seen = new Set();
      const values = [];
      raw.forEach(v => {
        const text = String(v ?? '').trim();
        if (!text || seen.has(text)) return;
        seen.add(text);
        values.push(text);
      });
      if (values.length > 0) cleaned[key] = values;
    });
    validations = withFallbacks(cleaned);
    if (Array.isArray(validations['ステータス'])) {
      validations['ステータス'].forEach(v => statusSet.add(v));
    }
    return validations;
  };

  return {
    async get_tasks() {
      return withSequentialNo();
    },
    async get_statuses() {
      return Array.from(statusSet);
    },
    async get_validations() {
      return { ...validations };
    },
    async update_validations(payload) {
      const updated = updateValidations(payload);
      return { ok: true, validations: { ...updated }, statuses: Array.from(statusSet) };
    },
    async add_task(payload) {
      const record = normalizeTask(payload);
      tasks.push(record);
      return { ...cloneTask(record), No: tasks.length };
    },
    async update_task(no, payload) {
      const located = locateTask(no);
      if (!located) throw new Error('指定したタスクが見つかりません');
      const updated = normalizeTask({ ...located.record, ...payload });
      tasks[located.index] = updated;
      return { ...cloneTask(updated), No: located.index + 1 };
    },
    async delete_task(no) {
      const located = locateTask(no);
      if (!located) return false;
      tasks.splice(located.index, 1);
      return true;
    },
    async reload_from_excel() {
      return {
        tasks: withSequentialNo(),
        statuses: Array.from(statusSet),
        validations: { ...validations }
      };
    },
    async save_excel() {
      return 'mock-data.xlsx';
    }
  };
}

async function init(force = false) {
  if (!api) return;
  if (force) {
    try {
      if (RUN_MODE === 'pywebview' && typeof api.reload_from_excel === 'function') {
        const result = await api.reload_from_excel();
        const rawTasks = Array.isArray(result?.tasks)
          ? result.tasks
          : await api.get_tasks();
        TASKS = sanitizeTaskList(rawTasks);
        STATUSES = Array.isArray(result?.statuses) ? result.statuses : await api.get_statuses?.() || [];
        applyValidationState(result?.validations);
        if (!result?.validations && typeof api.get_validations === 'function') {
          try {
            applyValidationState(await api.get_validations());
          } catch (err) {
            console.warn('get_validations failed:', err);
          }
        }
      } else {
        STATUSES = typeof api.get_statuses === 'function' ? await api.get_statuses() : [];
        TASKS = sanitizeTaskList(await api.get_tasks());
        if (typeof api.get_validations === 'function') {
          try {
            applyValidationState(await api.get_validations());
          } catch (err) {
            console.warn('get_validations failed:', err);
            applyValidationState(null);
          }
        } else {
          applyValidationState(null);
        }
      }
    } catch (err) {
      console.error('init failed', err);
      TASKS = [];
      STATUSES = [];
      applyValidationState(null);
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

  const saveBtn = document.getElementById('btn-save');
  if (saveBtn) {
    saveBtn.addEventListener('click', async () => {
      try {
        if (typeof api?.save_excel !== 'function') {
          alert('保存機能が利用できません。');
          return;
        }
        const result = await api.save_excel();
        const message = result ? `Excelへ保存しました\n${result}` : 'Excelへ保存しました';
        alert(message);
      } catch (err) {
        alert('保存に失敗: ' + (err?.message || err));
      }
    });
  }
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
  const assignees = collectAllAssignees().filter(name => name !== ASSIGNEE_UNASSIGNED_LABEL);
  select.innerHTML = '';

  const optionAll = document.createElement('option');
  optionAll.value = ASSIGNEE_FILTER_ALL;
  optionAll.textContent = 'すべて';
  select.appendChild(optionAll);

  const optionUnassigned = document.createElement('option');
  optionUnassigned.value = ASSIGNEE_FILTER_UNASSIGNED;
  optionUnassigned.textContent = ASSIGNEE_UNASSIGNED_LABEL;
  select.appendChild(optionUnassigned);

  assignees.forEach(name => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });

  if (previous === ASSIGNEE_FILTER_UNASSIGNED) {
    select.value = ASSIGNEE_FILTER_UNASSIGNED;
  } else if (previous && assignees.includes(previous)) {
    select.value = previous;
  } else {
    select.value = ASSIGNEE_FILTER_ALL;
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
  const assigneeSelect = document.getElementById('assignee-filter');
  const assigneeFilter = assigneeSelect ? assigneeSelect.value : ASSIGNEE_FILTER_ALL;
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
  let filteredAssignees;
  if (assigneeFilter === ASSIGNEE_FILTER_UNASSIGNED) {
    filteredAssignees = assignees.includes(ASSIGNEE_UNASSIGNED_LABEL) ? [ASSIGNEE_UNASSIGNED_LABEL] : [];
  } else if (assigneeFilter) {
    filteredAssignees = assignees.filter(name => name === assigneeFilter);
  } else {
    filteredAssignees = assignees;
  }
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
  const statusText = String(task.ステータス ?? '').trim();
  if (statusText) {
    const rawClass = `status-${statusText}`;
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
  const status = document.createElement('span');
  status.textContent = `ステータス: ${statusText || '未設定'}`;
  meta.appendChild(status);
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
  div.addEventListener('dblclick', () => {
    if (task?.No) {
      openEdit(task.No);
    }
  });
  return div;
}

function buildTaskLookup(from, to, assigneeFilter = ASSIGNEE_FILTER_ALL) {
  const map = new Map();
  TASKS.forEach(task => {
    const due = parseISO(task.期限);
    if (!due || due < from || due > to) return;
    const name = task.担当者?.trim() || ASSIGNEE_UNASSIGNED_LABEL;
    if (assigneeFilter === ASSIGNEE_FILTER_UNASSIGNED) {
      if (name !== ASSIGNEE_UNASSIGNED_LABEL) return;
    } else if (assigneeFilter && name !== assigneeFilter) {
      return;
    }
    if (!map.has(name)) map.set(name, new Map());
    const byDate = map.get(name);
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
    assignees.add(task.担当者?.trim() || ASSIGNEE_UNASSIGNED_LABEL);
  });
  return Array.from(assignees).sort((a, b) => a.localeCompare(b, 'ja'));
}

function collectAllAssignees() {
  const assignees = new Set();
  TASKS.forEach(task => {
    assignees.add(task.担当者?.trim() || ASSIGNEE_UNASSIGNED_LABEL);
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

function normalizeStatusLabel(value) {
  const text = String(value ?? '').trim();
  return text || UNSET_STATUS_LABEL;
}

function denormalizeStatusLabel(value) {
  const text = String(value ?? '').trim();
  return text === UNSET_STATUS_LABEL ? '' : text;
}

function normalizeValidationValues(rawList) {
  if (!Array.isArray(rawList)) return [];
  const seen = new Set();
  const values = [];
  rawList.forEach(item => {
    const text = String(item ?? '').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    values.push(text);
  });
  return values;
}

function applyValidationState(raw) {
  const next = {};
  if (raw && typeof raw === 'object') {
    Object.keys(raw).forEach(key => {
      next[key] = normalizeValidationValues(raw[key]);
    });
  }
  if (!Array.isArray(next['優先度']) || next['優先度'].length === 0) {
    next['優先度'] = [...PRIORITY_DEFAULT_OPTIONS];
  }
  VALIDATIONS = next;
}

function getPriorityOptions() {
  const base = Array.isArray(VALIDATIONS['優先度']) && VALIDATIONS['優先度'].length > 0
    ? VALIDATIONS['優先度']
    : PRIORITY_DEFAULT_OPTIONS;
  const seen = new Set();
  const options = [];
  base.forEach(value => {
    const text = String(value ?? '').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    options.push(text);
  });
  if (options.length === 0) {
    PRIORITY_DEFAULT_OPTIONS.forEach(value => {
      if (!seen.has(value)) {
        seen.add(value);
        options.push(value);
      }
    });
  }
  return options;
}

function getDefaultPriorityValue() {
  const options = getPriorityOptions();
  if (options.includes('中')) return '中';
  return options[0] || '';
}

function applyPriorityOptions(selectEl, currentValue, preferDefault = false) {
  if (!selectEl) return;
  const normalized = currentValue === null || currentValue === undefined
    ? ''
    : String(currentValue).trim();
  const options = getPriorityOptions();
  const fragments = [];
  const addOption = (value, label = value) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = label;
    fragments.push(opt);
  };

  if (!normalized && !preferDefault) {
    addOption('', '（未設定）');
  }
  options.forEach(value => addOption(value));
  if (normalized && !options.includes(normalized)) {
    addOption(normalized);
  }

  selectEl.innerHTML = '';
  fragments.forEach(opt => selectEl.appendChild(opt));

  const values = Array.from(selectEl.options).map(opt => opt.value);
  let selection = normalized;
  if (!selection || !values.includes(selection)) {
    if (preferDefault) {
      selection = getDefaultPriorityValue();
    } else if (values.includes('')) {
      selection = '';
    } else {
      selection = getDefaultPriorityValue();
    }
  }
  if (!values.includes(selection)) {
    selection = values[0] || '';
  }
  selectEl.value = selection;
}

function openEdit(no) {
  const task = TASKS.find(t => t.No === no);
  if (!task) return;
  CURRENT_EDIT = no;
  openModal(task, { mode: 'edit' });
}

function openModal(task, { mode }) {
  const modal = document.getElementById('modal');
  if (!modal) return;
  const title = document.getElementById('modal-title');
  const fno = document.getElementById('f-no');
  const fstat = document.getElementById('f-status');
  const fmajor = document.getElementById('f-major');
  const fminor = document.getElementById('f-minor');
  const fttl = document.getElementById('f-title');
  const fwho = document.getElementById('f-assignee');
  const fprio = document.getElementById('f-priority');
  const fdue = document.getElementById('f-due');
  const fnote = document.getElementById('f-notes');
  const btnDelete = document.getElementById('btn-delete');

  title.textContent = mode === 'create' ? 'タスク追加' : 'タスク編集';

  const statuses = Array.isArray(STATUSES) ? STATUSES : [];
  const seen = new Set();
  fstat.innerHTML = '';
  statuses.forEach(status => {
    const normalized = normalizeStatusLabel(status);
    if (seen.has(normalized)) return;
    seen.add(normalized);
    const opt = document.createElement('option');
    opt.value = normalized;
    opt.textContent = normalized;
    fstat.appendChild(opt);
  });
  if (!seen.has(UNSET_STATUS_LABEL)) {
    const opt = document.createElement('option');
    opt.value = UNSET_STATUS_LABEL;
    opt.textContent = UNSET_STATUS_LABEL;
    fstat.appendChild(opt);
  }

  fno.value = task.No ?? '';
  fstat.value = normalizeStatusLabel(task.ステータス);
  if (!Array.from(fstat.options).some(opt => opt.value === fstat.value)) {
    const opt = document.createElement('option');
    opt.value = fstat.value;
    opt.textContent = fstat.value;
    fstat.appendChild(opt);
  }
  fstat.value = normalizeStatusLabel(task.ステータス);

  if (fmajor) fmajor.value = task.大分類 || '';
  if (fminor) fminor.value = task.中分類 || '';
  fttl.value = task.タスク || '';
  fwho.value = task.担当者 || '';
  applyPriorityOptions(fprio, task.優先度, mode === 'create');
  fdue.value = (task.期限 || '').slice(0, 10);
  fnote.value = task.備考 || '';

  btnDelete.style.display = mode === 'edit' ? 'inline-flex' : 'none';

  document.getElementById('btn-close').onclick = closeModal;
  document.getElementById('btn-cancel').onclick = closeModal;

  btnDelete.onclick = async () => {
    if (!CURRENT_EDIT) return;
    if (!confirm('削除しますか？')) return;
    try {
      const ok = await api.delete_task(CURRENT_EDIT);
      if (ok) {
        if (typeof api.get_tasks === 'function') {
          TASKS = sanitizeTaskList(await api.get_tasks());
        } else {
          const remaining = TASKS
            .filter(x => x.No !== CURRENT_EDIT)
            .map((record, idx) => ({ ...record, No: idx + 1 }));
          TASKS = sanitizeTaskList(remaining);
        }
        CURRENT_EDIT = null;
        ensureRangeDefaults();
        closeModal();
        renderSummary();
        renderLegend();
        renderAssigneeFilter();
        renderTimeline();
      } else {
        alert('削除できませんでした');
      }
    } catch (err) {
      alert('削除に失敗: ' + (err?.message || err));
    }
  };

  const form = document.getElementById('task-form');
  form.onsubmit = async (e) => {
    e.preventDefault();
    const payload = {
      ステータス: denormalizeStatusLabel(fstat.value),
      大分類: fmajor ? fmajor.value.trim() : '',
      中分類: fminor ? fminor.value.trim() : '',
      タスク: fttl.value.trim(),
      担当者: fwho.value.trim(),
      優先度: (fprio.value ?? '').trim(),
      期限: fdue.value ? fdue.value : '',
      備考: fnote.value
    };

    if (!payload.タスク) {
      alert('タスクを入力してください。');
      fttl.focus();
      return;
    }

    try {
      if (mode === 'create') {
        const created = await api.add_task(payload);
        const sanitized = sanitizeTaskRecord(created, TASKS.length);
        if (sanitized) {
          TASKS.push(sanitized);
        }
      } else {
        const no = CURRENT_EDIT;
        const updated = await api.update_task(no, payload);
        const idx = TASKS.findIndex(x => x.No === no);
        if (idx >= 0) {
          const sanitized = sanitizeTaskRecord(updated, idx);
          if (sanitized) {
            TASKS[idx] = sanitized;
          } else {
            TASKS.splice(idx, 1);
          }
        }
      }
      CURRENT_EDIT = null;
      closeModal();
      ensureRangeDefaults();
      renderSummary();
      renderLegend();
      renderAssigneeFilter();
      renderTimeline();
    } catch (err) {
      alert('保存に失敗: ' + (err?.message || err));
    }
  };

  modal.classList.add('open');
  modal.setAttribute('aria-hidden', 'false');
}

function closeModal() {
  const modal = document.getElementById('modal');
  if (!modal) return;
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden', 'true');
}

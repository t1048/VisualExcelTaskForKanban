/* ===================== ランタイム切替（mock / pywebview） ===================== */
let api;                  // 実際に使う API （後で差し替える）
let RUN_MODE = 'mock';    // 'mock' | 'pywebview'
let WIRED = false;        // ツールバー多重バインド防止
const PRIORITY_DEFAULT_OPTIONS = ['高', '中', '低'];

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
    async move_task(no, status) {
      return this.update_task(no, { ステータス: status });
    },
    async save_excel() {
      return 'mock://task.xlsx';
    },
    async reload_from_excel() {
      return {
        ok: true,
        tasks: withSequentialNo(),
        statuses: Array.from(statusSet),
        validations: { ...validations }
      };
    }
  };
}


// DOM 準備
function ready(fn) {
  if (document.readyState !== 'loading') return fn();
  document.addEventListener('DOMContentLoaded', fn);
}

// pywebview が利用可能になったタイミングで API を差し替えて強制再初期化
window.addEventListener('pywebviewready', async () => {
  try {
    if (window.pywebview?.api) {
      api = window.pywebview.api;
      RUN_MODE = 'pywebview';
      console.log('[kanban] switched to pywebview API');
      await init(true);  // Excelから取り直し
    }
  } catch (e) {
    console.error('pywebviewready error:', e);
  }
});

ready(async () => {
  if (window.pywebview?.api) {
    api = window.pywebview.api;
    RUN_MODE = 'pywebview';
  } else {
    api = createMockApi();
    RUN_MODE = 'mock';
  }
  console.log('[kanban] run mode:', RUN_MODE);
  await init(true);
});

/* ===================== 状態 ===================== */
const VALIDATION_COLUMNS = ["ステータス", "大分類", "中分類", "タスク", "担当者", "優先度", "期限", "備考"];
const DEFAULT_STATUSES = ['未着手', '進行中', '完了', '保留'];
const UNSET_STATUS_LABEL = 'ステータス未設定';
const ASSIGNEE_FILTER_ALL = '__ALL__';
const ASSIGNEE_FILTER_UNASSIGNED = '__UNASSIGNED__';
const ASSIGNEE_UNASSIGNED_LABEL = '（未割り当て）';
const CATEGORY_FILTER_ALL = '__CATEGORY_ALL__';
const CATEGORY_FILTER_MINOR_ALL = '__CATEGORY_MINOR_ALL__';
let STATUSES = [];
let FILTERS = {
  assignee: ASSIGNEE_FILTER_ALL,
  statuses: new Set(),              // 初期化時に全ONにする
  keyword: '',
  date: { mode: 'none', from: '', to: '' },
  category: { major: CATEGORY_FILTER_ALL, minor: CATEGORY_FILTER_MINOR_ALL }
};
let TASKS = [];
let CURRENT_EDIT = null;
let VALIDATIONS = {};

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

function normalizeStatusLabel(value) {
  const text = String(value ?? '').trim();
  return text || UNSET_STATUS_LABEL;
}

function denormalizeStatusLabel(value) {
  const text = String(value ?? '').trim();
  return text === UNSET_STATUS_LABEL ? '' : text;
}

const STATUS_SORT_SEQUENCE = ['', UNSET_STATUS_LABEL, '未着手', '進行中', '完了', '保留中'];

const TABLE_COLUMN_CONFIG = [
  { key: 'no', label: 'No', width: '70px', sortable: true },
  { key: 'major', label: '大分類', width: '160px', sortable: true },
  { key: 'minor', label: '中分類', width: '160px', sortable: true },
  { key: 'task', label: 'タスク', sortable: true },
  { key: 'status', label: 'ステータス', width: '160px', sortable: true },
  { key: 'assignee', label: '担当者', width: '160px', sortable: true },
  { key: 'priority', label: '優先度', width: '140px', sortable: false },
  { key: 'due', label: '期限', width: '160px', sortable: true },
  { key: 'notes', label: '備考', sortable: false }
];

const COLUMN_CONFIG_BY_KEY = new Map();
TABLE_COLUMN_CONFIG.forEach(col => {
  COLUMN_CONFIG_BY_KEY.set(col.key, col);
});

let SORT_STATE = [];

function normalizedText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function compareLocaleStrings(a, b) {
  if (!a && !b) return 0;
  if (!a) return 1;
  if (!b) return -1;
  return a.localeCompare(b, 'ja');
}

function statusSortWeight(name) {
  const normalized = normalizeStatusLabel(name);
  const idx = STATUS_SORT_SEQUENCE.indexOf(normalized);
  if (idx >= 0) return idx;
  const trimmed = normalizedText(name);
  const legacyIdx = STATUS_SORT_SEQUENCE.indexOf(trimmed);
  return legacyIdx >= 0 ? legacyIdx : 100;
}

function compareStatusValues(a, b) {
  const sa = normalizeStatusLabel(a?.ステータス);
  const sb = normalizeStatusLabel(b?.ステータス);
  const wa = statusSortWeight(sa);
  const wb = statusSortWeight(sb);
  if (wa !== wb) return wa - wb;
  return sa.localeCompare(sb, 'ja');
}

function compareMajorValues(a, b) {
  const ma = normalizedText(a?.大分類);
  const mb = normalizedText(b?.大分類);
  return compareLocaleStrings(ma, mb);
}

function compareMinorValues(a, b) {
  const majorCmp = compareMajorValues(a, b);
  if (majorCmp !== 0) return majorCmp;
  const mina = normalizedText(a?.中分類);
  const minb = normalizedText(b?.中分類);
  return compareLocaleStrings(mina, minb);
}

function compareTaskValues(a, b) {
  const ta = normalizedText(a?.タスク);
  const tb = normalizedText(b?.タスク);
  return compareLocaleStrings(ta, tb);
}

function compareAssigneeValues(a, b) {
  const aa = normalizedText(a?.担当者);
  const ab = normalizedText(b?.担当者);
  return compareLocaleStrings(aa, ab);
}

function compareDueValues(a, b) {
  const da = parseISO(a?.期限 || '');
  const db = parseISO(b?.期限 || '');
  if (da && db) {
    const diff = da.getTime() - db.getTime();
    if (diff !== 0) return diff;
    return 0;
  }
  if (da) return -1;
  if (db) return 1;
  return 0;
}

function compareNoValues(a, b) {
  const sa = normalizedText(a?.No);
  const sb = normalizedText(b?.No);
  const na = Number(sa);
  const nb = Number(sb);
  const validA = sa !== '' && Number.isFinite(na);
  const validB = sb !== '' && Number.isFinite(nb);
  if (validA && validB) {
    if (na !== nb) return na - nb;
    return 0;
  }
  if (validA) return -1;
  if (validB) return 1;
  return sa.localeCompare(sb, 'ja');
}

const SORT_COMPARATORS = {
  no: compareNoValues,
  status: compareStatusValues,
  major: compareMajorValues,
  minor: compareMinorValues,
  task: compareTaskValues,
  assignee: compareAssigneeValues,
  due: compareDueValues,
};

function defaultListComparator(a, b, statusOrder) {
  const normalizedA = normalizeStatusLabel(a.ステータス);
  const normalizedB = normalizeStatusLabel(b.ステータス);
  const sa = statusOrder.has(normalizedA) ? statusOrder.get(normalizedA) : 999;
  const sb = statusOrder.has(normalizedB) ? statusOrder.get(normalizedB) : 999;
  if (sa !== sb) return sa - sb;
  const pri = comparePriorityValues(a.優先度, b.優先度);
  if (pri !== 0) return pri;
  return compareNoValues(a, b);
}

function normalizeStatePayload(payload) {
  if (!payload) return {};
  if (typeof payload === 'string') {
    try {
      return JSON.parse(payload) || {};
    } catch (err) {
      console.warn('[kanban] failed to parse payload string', err);
      return {};
    }
  }
  if (typeof payload === 'object') return payload;
  return {};
}

async function applyStateFromPayload(payload, options = {}) {
  const { preserveFilters = true, fallbackToApi = true } = options;
  const data = normalizeStatePayload(payload);
  const prevSelection = preserveFilters ? new Set(FILTERS.statuses) : new Set();

  if (Array.isArray(data.tasks)) {
    TASKS = sanitizeTaskList(data.tasks);
  }
  if (Array.isArray(data.statuses)) {
    STATUSES = data.statuses;
  }

  let validationPayload = data.validations ?? VALIDATIONS;
  if (!data.validations && fallbackToApi && typeof api?.get_validations === 'function') {
    try {
      validationPayload = await api.get_validations();
    } catch (err) {
      console.warn('get_validations failed:', err);
    }
  }

  applyValidationState(validationPayload);
  syncFilterStatuses(prevSelection);
  renderList();
  buildFiltersUI();
}

window.__kanban_receive_update = (payload) => {
  Promise.resolve(
    applyStateFromPayload(payload, { preserveFilters: true, fallbackToApi: false })
  ).catch(err => {
    console.error('[kanban] failed to apply pushed payload', err);
  });
};

function normalizeValidationValues(rawList) {
  if (!Array.isArray(rawList)) return [];
  const seen = new Set();
  const values = [];
  rawList.forEach(v => {
    const text = String(v ?? '').trim();
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
      const values = normalizeValidationValues(raw[key]);
      if (values.length > 0) {
        next[key] = values;
      }
    });
  }
  const merged = { ...next };
  if (!Array.isArray(merged['ステータス']) || merged['ステータス'].length === 0) {
    merged['ステータス'] = [...DEFAULT_STATUSES];
  }
  if (!Array.isArray(merged['優先度']) || merged['優先度'].length === 0) {
    merged['優先度'] = [...PRIORITY_DEFAULT_OPTIONS];
  }
  VALIDATIONS = merged;

  const validatedStatuses = VALIDATIONS['ステータス'] || [];
  if (validatedStatuses.length > 0) {
    const extras = Array.isArray(STATUSES) ? STATUSES.filter(s => !validatedStatuses.includes(s)) : [];
    STATUSES = [...validatedStatuses, ...extras];
  }

  if (!Array.isArray(STATUSES) || STATUSES.length === 0) {
    STATUSES = [...DEFAULT_STATUSES];
  }

  const seen = new Set();
  let ordered = [];
  STATUSES.forEach(s => {
    const text = String(s ?? '').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    ordered.push(text);
  });

  let hasEmptyTaskStatus = false;
  if (Array.isArray(TASKS)) {
    TASKS.forEach(t => {
      const raw = String(t?.ステータス ?? '').trim();
      if (!raw) {
        hasEmptyTaskStatus = true;
        return;
      }
      if (seen.has(raw)) return;
      seen.add(raw);
      ordered.push(raw);
    });
  }

  if (ordered.length === 0) {
    DEFAULT_STATUSES.forEach(s => {
      if (!seen.has(s)) {
        ordered.push(s);
        seen.add(s);
      }
    });
  }

  if (hasEmptyTaskStatus) {
    ordered = [UNSET_STATUS_LABEL, ...ordered.filter(s => s !== UNSET_STATUS_LABEL)];
  } else {
    ordered = ordered.filter((s, idx) => s !== UNSET_STATUS_LABEL || ordered.indexOf(s) === idx);
  }

  STATUSES = ordered;
}

function getPriorityOptions() {
  const raw = Array.isArray(VALIDATIONS['優先度']) ? VALIDATIONS['優先度'] : [];
  const base = raw.length > 0 ? raw : PRIORITY_DEFAULT_OPTIONS;
  const seen = new Set();
  const list = [];
  base.forEach(value => {
    const text = String(value ?? '').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    list.push(text);
  });
  if (list.length === 0) {
    PRIORITY_DEFAULT_OPTIONS.forEach(value => {
      if (seen.has(value)) return;
      seen.add(value);
      list.push(value);
    });
  }
  return list;
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
  const optionElements = [];
  const addOption = (value, label = value) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = label;
    optionElements.push(opt);
  };

  if (!normalized && !preferDefault) {
    addOption('', '（未設定）');
  }
  options.forEach(value => addOption(value));
  if (normalized && !options.includes(normalized)) {
    addOption(normalized);
  }

  selectEl.innerHTML = '';
  optionElements.forEach(opt => selectEl.appendChild(opt));

  const optionValues = Array.from(selectEl.options).map(opt => opt.value);
  let selection = normalized;
  if (!selection || !optionValues.includes(selection)) {
    if (preferDefault) {
      selection = getDefaultPriorityValue();
    } else if (optionValues.includes('')) {
      selection = '';
    } else {
      selection = getDefaultPriorityValue();
    }
  }
  if (!optionValues.includes(selection)) {
    selection = optionValues[0] || '';
  }
  selectEl.value = selection;
}

function syncFilterStatuses(prevSelection) {
  const statuses = Array.isArray(STATUSES) ? STATUSES : [];
  const base = prevSelection instanceof Set ? prevSelection : new Set();
  const next = new Set();
  statuses.forEach(s => {
    if (base.has(s)) next.add(s);
  });
  const hasUnset = statuses.includes(UNSET_STATUS_LABEL);
  if (hasUnset) {
    const hadUnset = base.has(UNSET_STATUS_LABEL);
    const emptyExists = Array.isArray(TASKS) && TASKS.some(t => !String(t?.ステータス ?? '').trim());
    if (hadUnset || emptyExists || base.size === 0) {
      next.add(UNSET_STATUS_LABEL);
    }
  }
  if (next.size === 0) {
    statuses.forEach(s => next.add(s));
  }
  FILTERS.statuses = next;
}

/* ===================== 初期化 ===================== */
async function init(force = false) {
  let payload = {};
  if (force) {
    if (RUN_MODE === 'pywebview' && typeof api.reload_from_excel === 'function') {
      try {
        payload = normalizeStatePayload(await api.reload_from_excel());
      } catch (e) {
        console.warn('reload_from_excel failed, fallback to get_*', e);
        payload = {};
      }
    }

    if (!Array.isArray(payload.tasks) && typeof api.get_tasks === 'function') {
      try {
        payload.tasks = await api.get_tasks();
      } catch (err) {
        console.error('get_tasks failed:', err);
      }
    }

    if (!Array.isArray(payload.statuses) && typeof api.get_statuses === 'function') {
      try {
        payload.statuses = await api.get_statuses();
      } catch (err) {
        console.error('get_statuses failed:', err);
      }
    }
  }

  await applyStateFromPayload(payload, { preserveFilters: true, fallbackToApi: true });
  // 初回＆再読込時にフィルタUIを最新へ
  if (!WIRED) { wireToolbar(); WIRED = true; }
}

/* ===================== レンダリング ===================== */
function renderList() {
  const container = document.getElementById('list-container');
  const summary = document.getElementById('list-summary');
  if (!container) return;

  const filtered = getFilteredTasks();
  const total = filtered.length;
  const overall = Array.isArray(TASKS) ? TASKS.length : 0;

  if (summary) {
    const parts = [];
    parts.push(`表示件数: ${total} 件`);
    if (overall && overall !== total) {
      parts.push(`全体: ${overall} 件`);
    }
    summary.textContent = parts.join(' / ');
  }

  container.innerHTML = '';

  if (filtered.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'empty-message';
    empty.textContent = '該当するタスクはありません。';
    container.appendChild(empty);
    updateDueIndicators(filtered);
    return;
  }

  const statusOrder = new Map();
  STATUSES.forEach((s, idx) => { statusOrder.set(s, idx); });

  const table = document.createElement('table');
  table.className = 'task-list';

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  TABLE_COLUMN_CONFIG.forEach(col => {
    const th = document.createElement('th');
    if (col.width) th.style.width = col.width;
    th.textContent = col.label;
    if (col.sortable) {
      th.classList.add('sortable');
      th.dataset.columnKey = col.key;
      th.setAttribute('aria-sort', 'none');
      th.tabIndex = 0;
      th.addEventListener('click', () => handleSortToggle(col.key));
      th.addEventListener('keydown', (ev) => {
        if (ev.key === 'Enter' || ev.key === ' ') {
          ev.preventDefault();
          handleSortToggle(col.key);
        }
      });
    }
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);
  applySortHeaderState(thead);

  const tbody = document.createElement('tbody');

  const activeSorts = SORT_STATE.filter(entry => SORT_COMPARATORS[entry.key]);

  const sortedTasks = filtered
    .slice()
    .sort((a, b) => {
      if (activeSorts.length > 0) {
        for (const entry of activeSorts) {
          const comparator = SORT_COMPARATORS[entry.key];
          if (!comparator) continue;
          const result = comparator(a, b);
          if (result !== 0) {
            return entry.direction === 'asc' ? result : -result;
          }
        }
      }
      return defaultListComparator(a, b, statusOrder);
    });

  sortedTasks.forEach(task => {
      const tr = document.createElement('tr');
      tr.dataset.no = String(task.No || '');
      tr.addEventListener('dblclick', () => openEdit(task.No));

      const noTd = document.createElement('td');
      noTd.textContent = task.No ? `#${task.No}` : '';
      tr.appendChild(noTd);

      const majorTd = document.createElement('td');
      if ((task.大分類 || '').trim()) {
        const badge = document.createElement('span');
        badge.className = 'category-pill category-major';
        badge.textContent = task.大分類.trim();
        majorTd.appendChild(badge);
      }
      tr.appendChild(majorTd);

      const minorTd = document.createElement('td');
      if ((task.中分類 || '').trim()) {
        const badge = document.createElement('span');
        badge.className = 'category-pill category-minor';
        badge.textContent = task.中分類.trim();
        minorTd.appendChild(badge);
      }
      tr.appendChild(minorTd);

      const titleTd = document.createElement('td');
      titleTd.textContent = task.タスク || '(無題)';
      tr.appendChild(titleTd);

      const statusTd = document.createElement('td');
      const statusPill = document.createElement('span');
      statusPill.className = 'status-pill';
      statusPill.textContent = normalizeStatusLabel(task.ステータス);
      statusTd.appendChild(statusPill);
      tr.appendChild(statusTd);

      const assigneeTd = document.createElement('td');
      assigneeTd.textContent = (task.担当者 || '').trim();
      tr.appendChild(assigneeTd);

      const priorityTd = document.createElement('td');
      if (task.優先度 !== undefined && task.優先度 !== null && String(task.優先度).trim() !== '') {
        const pill = document.createElement('span');
        pill.className = 'priority-pill';
        pill.textContent = `優先度: ${task.優先度}`;
        priorityTd.appendChild(pill);
      }
      tr.appendChild(priorityTd);

      const dueTd = document.createElement('td');
      if (task.期限) {
        const due = document.createElement('span');
        due.className = 'due-badge';
        let label = task.期限;
        const state = getDueState(task);
        if (state) {
          if (state.level === 'overdue') due.classList.add('due-overdue');
          if (state.level === 'warning') due.classList.add('due-warning');
          label += `（${state.label}）`;
        }
        due.textContent = label;
        dueTd.appendChild(due);
      }
      tr.appendChild(dueTd);

      const notesTd = document.createElement('td');
      notesTd.className = 'notes-cell';
      notesTd.textContent = task.備考 || '';
      tr.appendChild(notesTd);

      tbody.appendChild(tr);
    });

  table.appendChild(tbody);
  container.appendChild(table);

  updateDueIndicators(filtered);
}


function handleSortToggle(columnKey) {
  const column = COLUMN_CONFIG_BY_KEY.get(columnKey);
  if (!column || !column.sortable) return;

  const idx = SORT_STATE.findIndex(entry => entry.key === columnKey);
  if (idx === -1) {
    SORT_STATE.push({ key: columnKey, direction: 'asc' });
  } else if (SORT_STATE[idx].direction === 'asc') {
    SORT_STATE[idx].direction = 'desc';
  } else {
    SORT_STATE.splice(idx, 1);
  }

  renderList();
}

function applySortHeaderState(thead) {
  if (!thead) return;
  const stateMap = new Map();
  SORT_STATE.forEach((entry, idx) => {
    stateMap.set(entry.key, { direction: entry.direction, order: idx + 1 });
  });

  thead.querySelectorAll('th[data-column-key]').forEach(th => {
    const key = th.dataset.columnKey;
    const info = stateMap.get(key);
    if (info) {
      th.dataset.sortDirection = info.direction;
      th.dataset.sortOrder = String(info.order);
      th.setAttribute('aria-sort', info.direction === 'asc' ? 'ascending' : 'descending');
    } else {
      th.removeAttribute('data-sort-direction');
      th.removeAttribute('data-sort-order');
      th.setAttribute('aria-sort', 'none');
    }
  });
}


function uniqAssignees() {
  const set = new Set();
  TASKS.forEach(t => { if ((t.担当者 || '').trim()) set.add(t.担当者.trim()); });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'ja'));
}

function collectCategoryOptions() {
  const majorMap = new Map();
  TASKS.forEach(task => {
    const major = String(task?.大分類 ?? '').trim();
    const minor = String(task?.中分類 ?? '').trim();
    if (!major) return;
    if (!majorMap.has(major)) {
      majorMap.set(major, new Set());
    }
    if (minor) {
      majorMap.get(major).add(minor);
    }
  });

  const majorList = Array.from(majorMap.keys()).sort((a, b) => a.localeCompare(b, 'ja'));
  const minorMap = new Map();
  majorList.forEach(major => {
    const minors = majorMap.get(major) ?? new Set();
    minorMap.set(major, Array.from(minors).sort((a, b) => a.localeCompare(b, 'ja')));
  });

  return { majorList, minorMap };
}

function buildFiltersUI() {
  // ステータス（チェックボックス）
  const wrap = document.getElementById('flt-statuses');
  wrap.innerHTML = '';
  if (FILTERS.statuses.size === 0) {
    // 初期は全ON
    STATUSES.forEach(s => FILTERS.statuses.add(s));
  }
  STATUSES.forEach(s => {
    const id = 'st-' + btoa(unescape(encodeURIComponent(s))).replace(/=/g, '');
    const lbl = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.id = id; cb.value = s; cb.checked = FILTERS.statuses.has(s);
    cb.addEventListener('change', () => {
      if (cb.checked) FILTERS.statuses.add(s); else FILTERS.statuses.delete(s);
      renderList();
    });
    const span = document.createElement('span'); span.textContent = s;
    lbl.appendChild(cb); lbl.appendChild(span);
    wrap.appendChild(lbl);
  });

  // 大分類・中分類
  const majorSel = document.getElementById('flt-major');
  const minorSel = document.getElementById('flt-minor');
  if (majorSel && minorSel) {
    const { majorList, minorMap } = collectCategoryOptions();
    let currentMajor = FILTERS.category?.major ?? CATEGORY_FILTER_ALL;
    let currentMinor = FILTERS.category?.minor ?? CATEGORY_FILTER_MINOR_ALL;

    if (!majorList.includes(currentMajor)) {
      currentMajor = CATEGORY_FILTER_ALL;
      FILTERS.category.major = CATEGORY_FILTER_ALL;
    }

    const majorOptions = [
      `<option value="${CATEGORY_FILTER_ALL}">（すべて）</option>`
    ].concat(majorList.map(name => `<option value="${name}">${name}</option>`));
    majorSel.innerHTML = majorOptions.join('');
    majorSel.value = currentMajor;

    const renderMinorOptions = () => {
      const majorsMinors = minorMap.get(currentMajor) || [];
      const minorOptions = [
        `<option value="${CATEGORY_FILTER_MINOR_ALL}">（すべて）</option>`
      ].concat(majorsMinors.map(name => `<option value="${name}">${name}</option>`));
      minorSel.innerHTML = minorOptions.join('');

      if (currentMajor === CATEGORY_FILTER_ALL || majorsMinors.length === 0) {
        minorSel.disabled = true;
        currentMinor = CATEGORY_FILTER_MINOR_ALL;
        minorSel.value = CATEGORY_FILTER_MINOR_ALL;
        FILTERS.category.minor = CATEGORY_FILTER_MINOR_ALL;
      } else {
        minorSel.disabled = false;
        if (majorsMinors.includes(currentMinor)) {
          minorSel.value = currentMinor;
        } else {
          currentMinor = CATEGORY_FILTER_MINOR_ALL;
          minorSel.value = CATEGORY_FILTER_MINOR_ALL;
        }
        FILTERS.category.minor = currentMinor;
      }
    };

    renderMinorOptions();

    majorSel.onchange = () => {
      currentMajor = majorSel.value;
      FILTERS.category.major = currentMajor;
      currentMinor = CATEGORY_FILTER_MINOR_ALL;
      FILTERS.category.minor = currentMinor;
      renderMinorOptions();
      renderList();
    };

    minorSel.onchange = () => {
      currentMinor = minorSel.value;
      FILTERS.category.minor = currentMinor;
      renderList();
    };
  }

  // 担当者（セレクト）
  const sel = document.getElementById('flt-assignee');
  const selected = FILTERS.assignee;
  const list = uniqAssignees();
  const options = [
    `<option value="${ASSIGNEE_FILTER_ALL}">（全員）</option>`,
    `<option value="${ASSIGNEE_FILTER_UNASSIGNED}">${ASSIGNEE_UNASSIGNED_LABEL}</option>`
  ].concat(list.map(a => `<option value="${a}">${a}</option>`));
  sel.innerHTML = options.join('');
  if (selected === ASSIGNEE_FILTER_UNASSIGNED) {
    sel.value = ASSIGNEE_FILTER_UNASSIGNED;
  } else if (list.includes(selected)) {
    sel.value = selected;
  } else {
    sel.value = ASSIGNEE_FILTER_ALL;
  }
  sel.onchange = () => { FILTERS.assignee = sel.value; renderList(); };

  // キーワード
  const keywordEl = document.getElementById('flt-keyword');
  keywordEl.value = FILTERS.keyword || '';
  keywordEl.oninput = () => {
    FILTERS.keyword = keywordEl.value;
    renderList();
  };

  // 期限（モード＆日付）
  const modeSel = document.getElementById('flt-date-mode');
  const fromEl = document.getElementById('flt-date-from');
  const toEl = document.getElementById('flt-date-to');
  const sepEl = document.getElementById('flt-date-sep');

  // 既存値の反映
  modeSel.value = FILTERS.date.mode || 'none';
  fromEl.value = FILTERS.date.from || '';
  toEl.value = FILTERS.date.to || '';

  const updateVisibility = () => {
    const m = modeSel.value;
    if (m === 'none') {
      fromEl.style.display = '';
      toEl.style.display = 'none';
      sepEl.style.display = 'none';
    } else if (m === 'before') {
      fromEl.style.display = '';
      toEl.style.display = 'none';
      sepEl.style.display = 'none';
    } else if (m === 'after') {
      fromEl.style.display = '';
      toEl.style.display = 'none';
      sepEl.style.display = 'none';
    } else { // range
      fromEl.style.display = '';
      toEl.style.display = '';
      sepEl.style.display = '';
    }
  };
  updateVisibility();

  modeSel.onchange = () => { FILTERS.date.mode = modeSel.value; updateVisibility(); renderList(); };
  fromEl.onchange = () => { FILTERS.date.from = fromEl.value; renderList(); };
  toEl.onchange = () => { FILTERS.date.to = toEl.value; renderList(); };

  // 解除ボタン
  document.getElementById('btn-clear-filters').onclick = () => {
    FILTERS = {
      assignee: ASSIGNEE_FILTER_ALL,
      statuses: new Set(STATUSES),
      keyword: '',
      date: { mode: 'none', from: '', to: '' },
      category: { major: CATEGORY_FILTER_ALL, minor: CATEGORY_FILTER_MINOR_ALL }
    };
    buildFiltersUI();
    renderList();
  };
}


const PRIORITY_LABEL_ORDER = new Map([
  ['最優先', 0],
  ['緊急', 0],
  ['高', 1],
  ['High', 1],
  ['中', 2],
  ['Medium', 2],
  ['低', 3],
  ['Low', 3],
  ['通常', 4],
  ['Normal', 4],
  ['後回し', 5],
  ['Low Priority', 5],
]);

function prioritySortKey(value) {
  if (value === null || value === undefined) return { weight: 1001, label: '' };
  if (typeof value === 'number' && !Number.isNaN(value)) return { weight: value, label: '' };
  const str = String(value).trim();
  if (!str) return { weight: 1001, label: '' };
  const num = Number(str);
  if (!Number.isNaN(num)) return { weight: num, label: '' };
  if (PRIORITY_LABEL_ORDER.has(str)) return { weight: PRIORITY_LABEL_ORDER.get(str), label: '' };
  return { weight: 1000, label: str };
}

function comparePriorityValues(a, b) {
  const ka = prioritySortKey(a);
  const kb = prioritySortKey(b);
  if (ka.weight !== kb.weight) return ka.weight - kb.weight;
  return ka.label.localeCompare(kb.label, 'ja');
}

function updateDueIndicators(tasks) {
  const container = document.getElementById('toolbar-due');
  const overdueEl = document.getElementById('due-overdue-count');
  const warningEl = document.getElementById('due-warning-count');
  const toastEl = document.getElementById('due-toast');
  if (!container || !overdueEl || !warningEl || !toastEl) return;

  let overdue = 0;
  let warning = 0;

  tasks.forEach(task => {
    const state = getDueState(task);
    if (!state) return;
    if (state.level === 'overdue') {
      overdue += 1;
    } else if (state.level === 'warning') {
      warning += 1;
    }
  });

  overdueEl.querySelector('.count').textContent = overdue;
  overdueEl.hidden = overdue === 0;

  warningEl.querySelector('.count').textContent = warning;
  warningEl.hidden = warning === 0;

  if (overdue > 0) {
    toastEl.hidden = false;
    toastEl.textContent = `⚠️ 期限を過ぎたカードが ${overdue} 件あります。`;
  } else if (warning > 0) {
    toastEl.hidden = false;
    toastEl.textContent = `⏰ 期限が近いカードが ${warning} 件あります。`;
  } else {
    toastEl.hidden = true;
    toastEl.textContent = '';
  }

  const hasAlerts = overdue > 0 || warning > 0;
  container.classList.toggle('active', hasAlerts);
}

/* ===================== ツールバー ===================== */
function wireToolbar() {
  document.getElementById('btn-add').addEventListener('click', () => openCreate());
  document.getElementById('btn-save').addEventListener('click', async () => {
    try {
      const p = await api.save_excel();
      alert('Excelへ保存しました\n' + p);
    } catch (e) {
      alert('保存に失敗: ' + (e?.message || e));
    }
  });
  document.getElementById('btn-validations').addEventListener('click', () => openValidationModal());
  document.getElementById('btn-reload').addEventListener('click', async () => {
    try {
      const payload = await api.reload_from_excel();
      await applyStateFromPayload(payload, { preserveFilters: true, fallbackToApi: true });
    } catch (e) {
      alert('再読込に失敗: ' + (e?.message || e));
    }
  });
}

/* ===================== 入力規則モーダル ===================== */
function openValidationModal() {
  const modal = document.getElementById('validation-modal');
  const editor = document.getElementById('validation-editor');
  if (!modal || !editor) return;

  editor.innerHTML = '';
  VALIDATION_COLUMNS.forEach(column => {
    const item = document.createElement('div');
    item.className = 'validation-item';

    const label = document.createElement('label');
    const id = 'val-' + btoa(unescape(encodeURIComponent(column))).replace(/=/g, '');
    label.setAttribute('for', id);
    label.textContent = column;

    const textarea = document.createElement('textarea');
    textarea.id = id;
    textarea.dataset.column = column;
    textarea.placeholder = '1 行に 1 候補を入力';
    textarea.value = (Array.isArray(VALIDATIONS[column]) ? VALIDATIONS[column] : []).join('\n');
    textarea.spellcheck = false;

    item.appendChild(label);
    item.appendChild(textarea);
    editor.appendChild(item);
  });

  const closeBtn = document.getElementById('btn-validation-close');
  const cancelBtn = document.getElementById('btn-validation-cancel');
  const saveBtn = document.getElementById('btn-validation-save');

  if (closeBtn) closeBtn.onclick = closeValidationModal;
  if (cancelBtn) cancelBtn.onclick = closeValidationModal;
  modal.onclick = (ev) => {
    if (ev.target === modal) closeValidationModal();
  };

  if (saveBtn) {
    saveBtn.onclick = async () => {
      const payload = {};
      editor.querySelectorAll('textarea[data-column]').forEach(area => {
        const col = area.dataset.column;
        const lines = area.value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
        payload[col] = lines;
      });

      try {
        const prevSelection = new Set(FILTERS.statuses);
        if (typeof api?.update_validations === 'function') {
          const res = await api.update_validations(payload);
          if (res?.statuses && Array.isArray(res.statuses)) {
            STATUSES = res.statuses;
          }
          const received = res?.validations ?? payload;
          applyValidationState(received);
        } else {
          applyValidationState(payload);
        }
        syncFilterStatuses(prevSelection);
        closeValidationModal();
        renderList();
        buildFiltersUI();
      } catch (err) {
        alert('入力規則の保存に失敗: ' + (err?.message || err));
      }
    };
  }

  modal.classList.add('open');
  modal.setAttribute('aria-hidden', 'false');
}

function closeValidationModal() {
  const modal = document.getElementById('validation-modal');
  if (!modal) return;
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden', 'true');
}

function parseISO(d) {
  // 'YYYY-MM-DD' を Date に（失敗は null）
  if (!d) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d);
  if (!m) return null;
  const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(dt.getTime()) ? null : dt;
}

function getDueState(task) {
  const dueDate = parseISO(task?.期限 || '');
  if (!dueDate) return null;

  const due = new Date(dueDate.getTime());
  due.setHours(0, 0, 0, 0);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((due.getTime() - today.getTime()) / 86400000);

  if (diffDays < 0) {
    const abs = Math.abs(diffDays);
    return {
      level: 'overdue',
      diff: abs,
      label: `${abs}日超過`
    };
  }

  if (diffDays === 0) {
    return {
      level: 'warning',
      diff: 0,
      label: '本日期限'
    };
  }

  const label = `あと${diffDays}日`;
  if (diffDays <= 3) {
    return {
      level: 'warning',
      diff: diffDays,
      label
    };
  }

  return {
    level: 'normal',
    diff: diffDays,
    label
  };
}

function getFilteredTasks() {
  const assignee = FILTERS.assignee;
  const statuses = FILTERS.statuses;
  const df = FILTERS.date;
  const keyword = (FILTERS.keyword || '').trim().toLowerCase();
  const majorFilter = FILTERS.category?.major ?? CATEGORY_FILTER_ALL;
  const minorFilter = FILTERS.category?.minor ?? CATEGORY_FILTER_MINOR_ALL;

  return TASKS.filter(t => {
    if (majorFilter !== CATEGORY_FILTER_ALL) {
      const major = String(t.大分類 ?? '').trim();
      if (major !== majorFilter) return false;
      if (minorFilter !== CATEGORY_FILTER_MINOR_ALL) {
        const minor = String(t.中分類 ?? '').trim();
        if (minor !== minorFilter) return false;
      }
    }

    // 担当者
    const who = String(t.担当者 ?? '').trim();
    if (assignee === ASSIGNEE_FILTER_UNASSIGNED) {
      if (who) return false;
    } else if (assignee !== ASSIGNEE_FILTER_ALL) {
      if (who !== assignee) return false;
    }
    // ステータス
    const normalizedStatus = normalizeStatusLabel(t.ステータス);
    if (!statuses.has(normalizedStatus)) return false;

    // キーワード（タスク・備考）
    if (keyword) {
      const title = String(t.タスク ?? '').toLowerCase();
      const note = String(t.備考 ?? '').toLowerCase();
      if (!title.includes(keyword) && !note.includes(keyword)) return false;
    }

    // 期限
    if (df.mode === 'none') return true;
    const due = parseISO(t.期限 || '');
    if (!due) return false; // 期限が無いカードは条件指定時は除外

    if (df.mode === 'before') {
      const d = parseISO(df.from);
      if (!d) return true; // 入力未指定なら全通し
      // 期限 <= 指定日
      return (due.getTime() <= d.getTime());
    }
    if (df.mode === 'after') {
      const d = parseISO(df.from);
      if (!d) return true;
      // 期限 >= 指定日
      return (due.getTime() >= d.getTime());
    }
    if (df.mode === 'range') {
      const f = parseISO(df.from);
      const t2 = parseISO(df.to);
      if (f && t2) return (f.getTime() <= due.getTime() && due.getTime() <= t2.getTime());
      if (f && !t2) return (f.getTime() <= due.getTime());
      if (!f && t2) return (due.getTime() <= t2.getTime());
      return true;
    }
    return true;
  });
}


/* ===================== モーダル: 追加/編集 ===================== */
function openCreate() {
  CURRENT_EDIT = null;
  openModal({
    No: '',
    ステータス: STATUSES[0] || '未着手',
    大分類: '',
    中分類: '',
    タスク: '',
    担当者: '',
    優先度: getDefaultPriorityValue(),
    期限: '',
    備考: ''
  }, { mode: 'create' });
}
function openEdit(no) {
  const t = TASKS.find(x => x.No === no);
  if (!t) return;
  CURRENT_EDIT = no;
  openModal(t, { mode: 'edit' });
}

function openModal(task, { mode }) {
  const modal = document.getElementById('modal');
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

  // ステータス選択肢
  fstat.innerHTML = '';
  STATUSES.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    fstat.appendChild(opt);
  });

  fno.value = task.No ?? '';
  fstat.value = task.ステータス || STATUSES[0] || '未着手';
  if (fmajor) fmajor.value = task.大分類 || '';
  if (fminor) fminor.value = task.中分類 || '';
  fttl.value = task.タスク || '';
  fwho.value = task.担当者 || '';
  applyPriorityOptions(fprio, task.優先度, mode === 'create');
  fdue.value = (task.期限 || '').slice(0, 10);
  fnote.value = task.備考 || '';

  btnDelete.style.display = (mode === 'edit') ? 'inline-flex' : 'none';

  // ハンドラ
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
            .map((task, idx) => ({ ...task, No: idx + 1 }));
          TASKS = sanitizeTaskList(remaining);
        }
        CURRENT_EDIT = null;
        closeModal(); renderList(); buildFiltersUI();
      } else {
        alert('削除できませんでした');
      }
    } catch (e) {
      alert('削除に失敗: ' + (e?.message || e));
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
        const i = TASKS.findIndex(x => x.No === no);
        if (i >= 0) {
          const sanitized = sanitizeTaskRecord(updated, i);
          if (sanitized) {
            TASKS[i] = sanitized;
          } else {
            TASKS.splice(i, 1);
          }
        }
      }
      closeModal(); renderList(); buildFiltersUI();
    } catch (err) {
      alert('保存に失敗: ' + (err?.message || err));
    }
  };

  modal.classList.add('open');
  modal.setAttribute('aria-hidden', 'false');
}
function closeModal() {
  const modal = document.getElementById('modal');
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden', 'true');
}

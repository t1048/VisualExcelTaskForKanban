/* ===================== ランタイム切替（mock / pywebview） ===================== */
const {
  createMockApi,
  ready,
  sanitizeTaskRecord,
  sanitizeTaskList,
  normalizeStatePayload,
  normalizeStatusLabel,
  denormalizeStatusLabel,
  normalizeValidationValues,
  createPriorityHelper,
  setupRuntime,
  parseISO,
  getDueState,
  createWorkloadSummary,
  PRIORITY_DEFAULT_OPTIONS,
  DEFAULT_STATUSES,
  UNSET_STATUS_LABEL,
  getPriorityLevel,
} = window.TaskAppCommon;

let api;                  // 実際に使う API （後で差し替える）
let RUN_MODE = 'mock';    // 'mock' | 'pywebview'
let WIRED = false;        // ツールバー多重バインド防止

/* ===================== 状態 ===================== */
const VALIDATION_COLUMNS = ["ステータス", "大分類", "中分類", "タスク", "担当者", "優先度", "期限", "備考"];
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
const priorityHelper = createPriorityHelper({
  getValidations: () => VALIDATIONS,
  defaultOptions: PRIORITY_DEFAULT_OPTIONS,
});
const getPriorityOptions = () => priorityHelper.getOptions();
const getDefaultPriorityValue = () => priorityHelper.getDefaultValue();
const applyPriorityOptions = (selectEl, currentValue, preferDefault = false) => (
  priorityHelper.applyOptions(selectEl, currentValue, preferDefault)
);

const WORKLOAD_IN_PROGRESS_KEYWORDS = ['進行', '作業中', 'inprogress', 'wip'];
const WORKLOAD_HEAVY_THRESHOLD = 5;

function workloadInProgressCount(entry) {
  if (!entry || typeof entry !== 'object' || !entry.statusCounts) return 0;
  return Object.entries(entry.statusCounts).reduce((acc, [status, count]) => {
    const numeric = Number(count) || 0;
    if (numeric <= 0) return acc;
    const normalized = String(status ?? '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '');
    if (!normalized) return acc;
    if (WORKLOAD_IN_PROGRESS_KEYWORDS.some(keyword => normalized.includes(keyword))) {
      return acc + numeric;
    }
    return acc;
  }, 0);
}

function workloadHighlightPredicate(entry) {
  return workloadInProgressCount(entry) > WORKLOAD_HEAVY_THRESHOLD;
}

const workloadSummary = createWorkloadSummary({
  container: document.getElementById('workload-summary'),
  getStatuses: () => STATUSES,
  normalizeStatusLabel,
  unassignedKey: ASSIGNEE_FILTER_UNASSIGNED,
  unassignedLabel: ASSIGNEE_UNASSIGNED_LABEL,
  allKey: ASSIGNEE_FILTER_ALL,
  getActiveAssignee: () => FILTERS.assignee,
  getDueState,
  highlightPredicate: workloadHighlightPredicate,
  onSelectAssignee: (value) => {
    const next = (() => {
      if (!value || value === ASSIGNEE_FILTER_ALL) return ASSIGNEE_FILTER_ALL;
      if (FILTERS.assignee === value) return ASSIGNEE_FILTER_ALL;
      return value;
    })();
    FILTERS.assignee = next;
    const select = document.getElementById('flt-assignee');
    if (select) {
      select.value = next;
    }
    renderBoard();
  },
});

setupRuntime({
  mockApiFactory: createMockApi,
  onApiChanged: ({ api: nextApi, runMode }) => {
    api = nextApi;
    RUN_MODE = runMode;
    console.log('[kanban] run mode:', RUN_MODE);
  },
  onInit: async () => {
    await init(true);
  },
  onRealtimeUpdate: (payload) => (
    applyStateFromPayload(payload, { preserveFilters: true, fallbackToApi: false })
  ),
});

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
  renderBoard();
  buildFiltersUI();
}

window.__kanban_receive_update = (payload) => {
  Promise.resolve(
    applyStateFromPayload(payload, { preserveFilters: true, fallbackToApi: false })
  ).catch(err => {
    console.error('[kanban] failed to apply pushed payload', err);
  });
};

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
function renderBoard() {
  const board = document.getElementById('board');
  board.innerHTML = '';

  const FILTERED = getFilteredTasks();  // ← 追加

  STATUSES.forEach(status => {
    const columnTasks = FILTERED.filter(t => normalizeStatusLabel(t.ステータス) === status);
    if (status === UNSET_STATUS_LABEL && columnTasks.length === 0) {
      return;
    }

    const col = document.createElement('section');
    col.className = 'column';
    col.dataset.status = status;

    const header = document.createElement('div');
    header.className = 'column-header';
    const title = document.createElement('div');
    title.className = 'column-title';
    title.textContent = status;
    const count = document.createElement('div');
    count.className = 'column-count';
    // ↓ 絞り込み済みから件数を出す
    count.textContent = `${columnTasks.length} 件`;

    header.appendChild(title);
    header.appendChild(count);

    const body = document.createElement('div');
    body.className = 'column-body';
    const drop = document.createElement('div');
    drop.className = 'dropzone';
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
    drop.addEventListener('drop', e => onDropCard(e, status, drop));

    // ↓ 絞り込み済みから、その列のカードを描画
    columnTasks
      .sort((a, b) => comparePriorityValues(a.優先度, b.優先度) || (a.No || 0) - (b.No || 0))
      .forEach(task => drop.appendChild(renderCard(task)));

    body.appendChild(drop);
    col.appendChild(header);
    col.appendChild(body);
    board.appendChild(col);
  });

  updateDueIndicators(FILTERED);
  buildAssigneeWorkload(FILTERED);
}


function buildAssigneeWorkload(tasks) {
  const list = Array.isArray(tasks) ? tasks : [];
  workloadSummary.update(list, {
    total: list.length,
    overall: Array.isArray(TASKS) ? TASKS.length : list.length,
  });
}


function uniqAssignees() {
  const set = new Set();
  TASKS.forEach(t => { if ((t.担当者 || '').trim()) set.add(t.担当者.trim()); });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'ja'));
}

function collectCategoryOptions() {
  const majorMap = new Map();
  const looseMinors = new Set();

  const ensureMajor = (name) => {
    const text = String(name ?? '').trim();
    if (!text) return null;
    if (!majorMap.has(text)) {
      majorMap.set(text, new Set());
    }
    return majorMap.get(text);
  };

  if (Array.isArray(TASKS)) {
    TASKS.forEach(task => {
      const major = String(task?.大分類 ?? '').trim();
      const minor = String(task?.中分類 ?? '').trim();
      if (major) {
        const set = ensureMajor(major);
        if (minor && set) {
          set.add(minor);
        }
      } else if (minor) {
        looseMinors.add(minor);
      }
    });
  }

  const validatedMajors = Array.isArray(VALIDATIONS['大分類']) ? VALIDATIONS['大分類'] : [];
  validatedMajors.forEach(name => {
    ensureMajor(name);
  });

  const validatedMinors = Array.isArray(VALIDATIONS['中分類']) ? VALIDATIONS['中分類'] : [];
  validatedMinors.forEach(name => {
    const text = String(name ?? '').trim();
    if (!text) return;
    let assigned = false;
    majorMap.forEach(set => {
      if (set.has(text)) assigned = true;
    });
    if (!assigned) {
      looseMinors.add(text);
    }
  });

  const majorList = Array.from(majorMap.keys()).sort((a, b) => a.localeCompare(b, 'ja'));
  const minorMap = new Map();
  majorList.forEach(major => {
    const minors = Array.from(majorMap.get(major) ?? new Set())
      .sort((a, b) => a.localeCompare(b, 'ja'));
    minorMap.set(major, minors);
  });

  const allMinorsSet = new Set();
  minorMap.forEach(list => {
    list.forEach(value => allMinorsSet.add(value));
  });
  looseMinors.forEach(value => allMinorsSet.add(value));

  const allMinors = Array.from(allMinorsSet).sort((a, b) => a.localeCompare(b, 'ja'));

  return { majorList, minorMap, allMinors };
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
      renderBoard();
    });
    const span = document.createElement('span'); span.textContent = s;
    lbl.appendChild(cb); lbl.appendChild(span);
    wrap.appendChild(lbl);
  });

  // 大分類・中分類
  const majorSel = document.getElementById('flt-major');
  const minorSel = document.getElementById('flt-minor');
  if (majorSel && minorSel) {
    const { majorList, minorMap, allMinors } = collectCategoryOptions();
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

    const renderMinorOptions = ({ preserve = false } = {}) => {
      const isAllMajor = currentMajor === CATEGORY_FILTER_ALL;
      const majorsMinors = isAllMajor ? (allMinors || []) : (minorMap.get(currentMajor) || []);
      const minorOptions = [
        `<option value="${CATEGORY_FILTER_MINOR_ALL}">（すべて）</option>`
      ].concat(majorsMinors.map(name => `<option value="${name}">${name}</option>`));
      minorSel.innerHTML = minorOptions.join('');

      if (!preserve) {
        currentMinor = CATEGORY_FILTER_MINOR_ALL;
      }

      if (currentMinor !== CATEGORY_FILTER_MINOR_ALL && !majorsMinors.includes(currentMinor)) {
        currentMinor = CATEGORY_FILTER_MINOR_ALL;
      }

      if (majorsMinors.length === 0) {
        currentMinor = CATEGORY_FILTER_MINOR_ALL;
        minorSel.disabled = true;
      } else {
        minorSel.disabled = false;
      }

      FILTERS.category.minor = currentMinor;
      minorSel.value = currentMinor;
    };

    renderMinorOptions({ preserve: true });

    majorSel.onchange = () => {
      currentMajor = majorSel.value;
      FILTERS.category.major = currentMajor;
      if (currentMajor !== CATEGORY_FILTER_ALL) {
        currentMinor = CATEGORY_FILTER_MINOR_ALL;
      }
      renderMinorOptions({ preserve: currentMajor === CATEGORY_FILTER_ALL });
      renderBoard();
    };

    minorSel.onchange = () => {
      currentMinor = minorSel.value;
      FILTERS.category.minor = currentMinor;
      renderBoard();
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
  sel.onchange = () => { FILTERS.assignee = sel.value; renderBoard(); };

  // キーワード
  const keywordEl = document.getElementById('flt-keyword');
  keywordEl.value = FILTERS.keyword || '';
  keywordEl.oninput = () => {
    FILTERS.keyword = keywordEl.value;
    renderBoard();
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

  modeSel.onchange = () => { FILTERS.date.mode = modeSel.value; updateVisibility(); renderBoard(); };
  fromEl.onchange = () => { FILTERS.date.from = fromEl.value; renderBoard(); };
  toEl.onchange = () => { FILTERS.date.to = toEl.value; renderBoard(); };

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
    renderBoard();
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

function renderCard(task) {
  const el = document.createElement('article');
  el.className = 'card';
  el.draggable = true;
  el.dataset.no = task.No;

  const dueState = getDueState(task);
  if (dueState?.level === 'overdue') {
    el.classList.add('due-overdue');
  } else if (dueState?.level === 'warning') {
    el.classList.add('due-warning');
  }

  el.addEventListener('dragstart', e => {
    e.dataTransfer.setData('text/plain', String(task.No));
    e.dataTransfer.dropEffect = 'move';
  });

  el.addEventListener('dblclick', () => openEdit(task.No));

  const category = document.createElement('div');
  category.className = 'card-category';
  let hasCategory = false;
  if (task.大分類) {
    const major = document.createElement('span');
    major.className = 'badge badge-major';
    major.textContent = task.大分類;
    category.appendChild(major);
    hasCategory = true;
  }
  if (task.中分類) {
    const minor = document.createElement('span');
    minor.className = 'badge badge-minor';
    minor.textContent = task.中分類;
    category.appendChild(minor);
    hasCategory = true;
  }

  const header = document.createElement('div');
  header.className = 'card-header';
  const title = document.createElement('div');
  title.className = 'card-title';
  title.textContent = task.タスク || '(無題)';
  const no = document.createElement('div');
  no.className = 'card-no';
  no.textContent = `#${task.No}`;
  header.appendChild(title);
  header.appendChild(no);

  if (hasCategory) {
    el.appendChild(category);
  }
  const meta = document.createElement('div');
  meta.className = 'card-meta';
  if (task.優先度 !== undefined && task.優先度 !== null && String(task.優先度).trim() !== '') {
    const bp = document.createElement('span');
    bp.className = 'badge badge-priority';
    bp.textContent = `優先度: ${task.優先度}`;
    const priorityLevel = getPriorityLevel(task.優先度);
    if (priorityLevel && priorityLevel !== 'unset' && priorityLevel !== 'custom') {
      bp.classList.add(`badge-priority-${priorityLevel}`);
    }
    meta.appendChild(bp);
  }
  if (task.担当者) {
    const b1 = document.createElement('span'); b1.className = 'badge badge-assignee'; b1.textContent = task.担当者; meta.appendChild(b1);
  }
  if (task.期限) {
    const b2 = document.createElement('span');
    b2.className = 'badge badge-date';
    let dueText = task.期限;
    if (dueState) {
      dueText += `（${dueState.label}）`;
      if (dueState.level === 'overdue') b2.classList.add('due-overdue');
      if (dueState.level === 'warning') b2.classList.add('due-warning');
    }
    b2.textContent = dueText;
    meta.appendChild(b2);
  }

  const notes = document.createElement('div');
  notes.className = 'card-notes';
  notes.textContent = task.備考 || '';

  el.appendChild(header);
  el.appendChild(meta);
  el.appendChild(notes);
  return el;
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

/* ===================== DnD ===================== */
async function onDropCard(e, newStatus, dropzone) {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  const no = parseInt(e.dataTransfer.getData('text/plain'), 10);
  if (!no) return;

  try {
    const nextStatus = denormalizeStatusLabel(newStatus);
    await api.move_task(no, nextStatus);
    const idx = TASKS.findIndex(t => t.No === no);
    if (idx >= 0) TASKS[idx].ステータス = nextStatus;
    renderBoard();
  } catch (err) {
    alert('移動に失敗しました: ' + (err?.message || err));
  }
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
  const timelineBtn = document.getElementById('btn-timeline');
  if (timelineBtn) {
    timelineBtn.addEventListener('click', () => {
      window.location.href = 'timeline.html';
    });
  }

  const listBtn = document.getElementById('btn-list');
  if (listBtn) {
    listBtn.addEventListener('click', () => {
      window.location.href = 'list.html';
    });
  }
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
        renderBoard();
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

function getFilteredTasks() {
  const assignee = FILTERS.assignee;
  const statuses = FILTERS.statuses;
  const df = FILTERS.date;
  const keyword = (FILTERS.keyword || '').trim().toLowerCase();
  const majorFilter = FILTERS.category?.major ?? CATEGORY_FILTER_ALL;
  const minorFilter = FILTERS.category?.minor ?? CATEGORY_FILTER_MINOR_ALL;

  const shouldFilterMajor = majorFilter !== CATEGORY_FILTER_ALL;
  const shouldFilterMinor = minorFilter !== CATEGORY_FILTER_MINOR_ALL;

  return TASKS.filter(t => {
    if (shouldFilterMajor || shouldFilterMinor) {
      const major = String(t.大分類 ?? '').trim();
      const minor = String(t.中分類 ?? '').trim();

      if (shouldFilterMajor && major !== majorFilter) return false;
      if (shouldFilterMinor && minor !== minorFilter) return false;
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

function setupCategoryInputSuggestions(fmajor, fminor) {
  const majorListEl = document.getElementById('modal-major-list');
  const minorListEl = document.getElementById('modal-minor-list');
  if (!majorListEl && !minorListEl) return;

  const { majorList, minorMap, allMinors } = collectCategoryOptions();

  const fillOptions = (datalist, values) => {
    if (!datalist) return;
    datalist.innerHTML = '';
    if (!Array.isArray(values)) return;
    const seen = new Set();
    values.forEach(value => {
      const text = String(value ?? '').trim();
      if (!text || seen.has(text)) return;
      seen.add(text);
      const opt = document.createElement('option');
      opt.value = text;
      datalist.appendChild(opt);
    });
  };

  fillOptions(majorListEl, majorList);

  const fallbackMinors = Array.isArray(allMinors) ? allMinors : [];
  const updateMinorOptions = () => {
    if (!minorListEl) return;
    const majorValue = String(fmajor?.value ?? '').trim();
    let candidates = fallbackMinors;
    if (majorValue && minorMap.has(majorValue)) {
      const list = minorMap.get(majorValue) || [];
      if (Array.isArray(list) && list.length > 0) {
        candidates = list;
      }
    }
    fillOptions(minorListEl, candidates);
  };

  updateMinorOptions();

  if (fmajor && minorListEl) {
    if (typeof fmajor.__minorListHandler === 'function') {
      fmajor.removeEventListener('input', fmajor.__minorListHandler);
      fmajor.removeEventListener('change', fmajor.__minorListHandler);
    }
    const handler = () => updateMinorOptions();
    fmajor.__minorListHandler = handler;
    fmajor.addEventListener('input', handler);
    fmajor.addEventListener('change', handler);
  }

  if (fminor && minorListEl) {
    // Ensure the list is refreshed when the field gains focus after manual edits.
    if (typeof fminor.__minorListRefresh === 'function') {
      fminor.removeEventListener('focus', fminor.__minorListRefresh);
    }
    const refresh = () => updateMinorOptions();
    fminor.__minorListRefresh = refresh;
    fminor.addEventListener('focus', refresh);
  }
}

function setupAssigneeInputSuggestions(fassignee) {
  const datalist = document.getElementById('modal-assignee-list');
  if (!fassignee || !datalist) return;

  const candidates = Array.isArray(TASKS) ? uniqAssignees() : [];
  const seen = new Set();
  datalist.innerHTML = '';

  candidates.forEach(name => {
    const text = String(name ?? '').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    const opt = document.createElement('option');
    opt.value = text;
    datalist.appendChild(opt);
  });

  const current = String(fassignee.value ?? '').trim();
  if (current && !seen.has(current)) {
    const opt = document.createElement('option');
    opt.value = current;
    datalist.appendChild(opt);
    seen.add(current);
  }

  fassignee.setAttribute('list', datalist.id);
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
  setupCategoryInputSuggestions(fmajor, fminor);
  fttl.value = task.タスク || '';
  fwho.value = task.担当者 || '';
  setupAssigneeInputSuggestions(fwho);
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
        closeModal(); renderBoard(); buildFiltersUI();
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
      closeModal(); renderBoard(); buildFiltersUI();
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

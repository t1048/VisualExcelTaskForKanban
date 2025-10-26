/* ===================== ランタイム切替（mock / pywebview） ===================== */
let api;                  // 実際に使う API （後で差し替える）
let RUN_MODE = 'mock';    // 'mock' | 'pywebview'
let WIRED = false;        // ツールバー多重バインド防止

function createMockApi() {
  const baseStatuses = ['未着手', '進行中', '完了', '保留'];
  const statusSet = new Set(baseStatuses);
  const pad = n => String(n).padStart(2, '0');
  const toISO = date => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  const today = new Date();
  const sampleTasks = Array.from({ length: 8 }).map((_, idx) => {
    const due = new Date(today);
    due.setDate(today.getDate() + idx - 2);
    const status = baseStatuses[idx % baseStatuses.length];
    statusSet.add(status);
    return {
      ステータス: status,
      タスク: `サンプルタスク ${idx + 1}`,
      担当者: ['田中', '佐藤', '鈴木', '高橋'][idx % 4],
      優先度: ['高', '中', '低'][idx % 3],
      期限: toISO(due),
      備考: idx % 2 === 0 ? 'モックデータ' : ''
    };
  });
  const tasks = [...sampleTasks];
  let validations = { 'ステータス': Array.from(statusSet) };

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

    return {
      ステータス: status,
      タスク: String(payload?.タスク ?? '').trim(),
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
    if (!payload || typeof payload !== 'object') {
      validations = { 'ステータス': Array.from(statusSet) };
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
    validations = cleaned;
    if (Array.isArray(cleaned['ステータス'])) {
      cleaned['ステータス'].forEach(v => statusSet.add(v));
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
const VALIDATION_COLUMNS = ["ステータス", "タスク", "担当者", "優先度", "期限", "備考"];
const DEFAULT_STATUSES = ['未着手', '進行中', '完了', '保留'];
let STATUSES = [];
let FILTERS = {
  assignee: '__ALL__',
  statuses: new Set(),              // 初期化時に全ONにする
  keyword: '',
  date: { mode: 'none', from: '', to: '' }
};
let TASKS = [];
let CURRENT_EDIT = null;
let VALIDATIONS = {};

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
  VALIDATIONS = next;

  const validatedStatuses = next['ステータス'] || [];
  if (validatedStatuses.length > 0) {
    const extras = Array.isArray(STATUSES) ? STATUSES.filter(s => !validatedStatuses.includes(s)) : [];
    STATUSES = [...validatedStatuses, ...extras];
  }

  if (!Array.isArray(STATUSES) || STATUSES.length === 0) {
    STATUSES = [...DEFAULT_STATUSES];
  }

  const seen = new Set();
  const ordered = [];
  STATUSES.forEach(s => {
    const text = String(s ?? '').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    ordered.push(text);
  });

  if (Array.isArray(TASKS)) {
    TASKS.forEach(t => {
      const name = String(t?.ステータス ?? '').trim();
      if (!name || seen.has(name)) return;
      seen.add(name);
      ordered.push(name);
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

  STATUSES = ordered;
}

function syncFilterStatuses(prevSelection) {
  const statuses = Array.isArray(STATUSES) ? STATUSES : [];
  const base = prevSelection instanceof Set ? prevSelection : new Set();
  const next = new Set();
  statuses.forEach(s => {
    if (base.has(s)) next.add(s);
  });
  if (next.size === 0) {
    statuses.forEach(s => next.add(s));
  }
  FILTERS.statuses = next;
}

/* ===================== 初期化 ===================== */
async function init(force = false) {
  const prevSelection = new Set(FILTERS.statuses);
  let validationPayload = VALIDATIONS;
  if (force) {
    if (RUN_MODE === 'pywebview' && typeof api.reload_from_excel === 'function') {
      try {
        const r = await api.reload_from_excel();
        if (r?.tasks) TASKS = r.tasks;
        if (r?.statuses) STATUSES = r.statuses;
        validationPayload = r?.validations ?? VALIDATIONS;
      } catch (e) {
        console.warn('reload_from_excel failed, fallback to get_*', e);
        STATUSES = await api.get_statuses();
        TASKS = await api.get_tasks();
        if (typeof api.get_validations === 'function') {
          validationPayload = await api.get_validations();
        } else {
          validationPayload = VALIDATIONS;
        }
      }
    } else {
      STATUSES = await api.get_statuses();
      TASKS = await api.get_tasks();
      if (typeof api.get_validations === 'function') {
        validationPayload = await api.get_validations();
      } else {
        validationPayload = VALIDATIONS;
      }
    }
  }
  applyValidationState(validationPayload);
  syncFilterStatuses(prevSelection);
  renderBoard();
  buildFiltersUI();   // 初回＆再読込時にフィルタUIを最新へ
  if (!WIRED) { wireToolbar(); WIRED = true; }
}

/* ===================== レンダリング ===================== */
function renderBoard() {
  const board = document.getElementById('board');
  board.innerHTML = '';

  const FILTERED = getFilteredTasks();  // ← 追加

  STATUSES.forEach(status => {
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
    count.textContent = `${FILTERED.filter(t => t.ステータス === status).length} 件`;

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
    FILTERED
      .filter(t => t.ステータス === status)
      .sort((a, b) => comparePriorityValues(a.優先度, b.優先度) || (a.No || 0) - (b.No || 0))
      .forEach(task => drop.appendChild(renderCard(task)));

    body.appendChild(drop);
    col.appendChild(header);
    col.appendChild(body);
    board.appendChild(col);
  });

  updateDueIndicators(FILTERED);
}


function uniqAssignees() {
  const set = new Set();
  TASKS.forEach(t => { if ((t.担当者 || '').trim()) set.add(t.担当者.trim()); });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'ja'));
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

  // 担当者（セレクト）
  const sel = document.getElementById('flt-assignee');
  const selected = FILTERS.assignee;
  const list = uniqAssignees();
  sel.innerHTML = '<option value="__ALL__">（全員）</option>' +
    list.map(a => `<option value="${a}">${a}</option>`).join('');
  sel.value = list.includes(selected) ? selected : '__ALL__';
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
    FILTERS = { assignee: '__ALL__', statuses: new Set(STATUSES), keyword: '', date: { mode: 'none', from: '', to: '' } };
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

  const meta = document.createElement('div');
  meta.className = 'card-meta';
  if (task.優先度 !== undefined && task.優先度 !== null && String(task.優先度).trim() !== '') {
    const bp = document.createElement('span');
    bp.className = 'badge badge-priority';
    bp.textContent = `優先度: ${task.優先度}`;
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

  if (overdue > 0) {
    overdueEl.hidden = false;
    overdueEl.querySelector('.count').textContent = overdue;
  } else {
    overdueEl.hidden = true;
  }

  if (warning > 0) {
    warningEl.hidden = false;
    warningEl.querySelector('.count').textContent = warning;
  } else {
    warningEl.hidden = true;
  }

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
    await api.move_task(no, newStatus);
    const idx = TASKS.findIndex(t => t.No === no);
    if (idx >= 0) TASKS[idx].ステータス = newStatus;
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
      const r = await api.reload_from_excel();
      if (r?.tasks) TASKS = r.tasks;
      if (r?.statuses) STATUSES = r.statuses;
      let validationPayload = VALIDATIONS;
      if (r?.validations) {
        validationPayload = r.validations;
      } else if (typeof api.get_validations === 'function') {
        validationPayload = await api.get_validations();
      }
      const prevSelection = new Set(FILTERS.statuses);
      applyValidationState(validationPayload);
      syncFilterStatuses(prevSelection);
      renderBoard();
      buildFiltersUI();
    } catch (e) {
      alert('再読込に失敗: ' + (e?.message || e));
    }
  });
  document.getElementById('btn-timeline').addEventListener('click', () => {
    window.location.href = 'timeline.html';
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

  return TASKS.filter(t => {
    // 担当者
    if (assignee !== '__ALL__') {
      if ((t.担当者 || '') !== assignee) return false;
    }
    // ステータス
    if (!statuses.has(t.ステータス)) return false;

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
    タスク: '',
    担当者: '',
    優先度: '',
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
  fttl.value = task.タスク || '';
  fwho.value = task.担当者 || '';
  fprio.value = task.優先度 !== undefined && task.優先度 !== null ? String(task.優先度) : '';
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
          TASKS = await api.get_tasks();
        } else {
          TASKS = TASKS.filter(x => x.No !== CURRENT_EDIT).map((task, idx) => ({ ...task, No: idx + 1 }));
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
      ステータス: fstat.value.trim(),
      タスク: fttl.value.trim(),
      担当者: fwho.value.trim(),
      優先度: fprio.value.trim(),
      期限: fdue.value ? fdue.value : '',
      備考: fnote.value
    };

    try {
      if (mode === 'create') {
        const created = await api.add_task(payload);
        TASKS.push(created);
      } else {
        const no = CURRENT_EDIT;
        const updated = await api.update_task(no, payload);
        const i = TASKS.findIndex(x => x.No === no);
        if (i >= 0) TASKS[i] = updated;
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

const {
  createMockApi,
  ready,
  sanitizeTaskRecord,
  sanitizeTaskList,
  normalizeStatePayload,
  setupDragViewportAutoScroll,
  parseISO: parseISODate,
} = window.TaskAppCommon;

const {
  setupRuntime,
  hasInitialExcelLoadFlag,
  markInitialExcelLoadFlag,
  resetInitialExcelLoadFlag,
  bindExcelActions,
} = window.TaskAppRuntime || {};

const {
  applyValidationState,
  createPriorityHelper,
  normalizeStatusLabel,
  denormalizeStatusLabel,
  PRIORITY_DEFAULT_OPTIONS,
  DEFAULT_STATUSES,
  UNSET_STATUS_LABEL,
} = window.TaskValidation || {};
const { createExcelSyncHandlers } = window.TaskExcelSync || {};

const { createFilterPresetManager } = window.TaskFilterPresets || {};

let api;
let RUN_MODE = 'mock';
let TASKS = [];
let STATUSES = [];
let VALIDATIONS = {};
let excelSyncHandlers = null;
const VALIDATION_COLUMNS = ["ステータス", "大分類", "中分類", "タスク", "担当者", "優先度", "期限", "備考"];
let CURRENT_EDIT = null;
let CURRENT_DRAG = null;
let cleanupAutoScroll = null;

let FILTERS = {
  month: '',
};

const filterPresetManager = typeof createFilterPresetManager === 'function'
  ? createFilterPresetManager({
      viewKey: 'calendar',
      selectors: {
        select: '#calendar-preset',
        apply: '#btn-calendar-preset-apply',
        save: '#btn-calendar-preset-save',
        delete: '#btn-calendar-preset-delete',
      },
      serialize: () => {
        syncFiltersFromUI();
        return {
          month: FILTERS.month || '',
        };
      },
      applyToUI: (raw) => {
        const data = raw && typeof raw === 'object' ? raw : {};
        const monthValue = String(data.month ?? '').trim();
        FILTERS.month = monthValue;
        applyFiltersToUI();
        syncFiltersFromUI();
        updateMonthLabel();
        renderCalendar();
        renderBacklog();
      },
    })
  : {
      maybeApplyInitialPreset: () => {},
      updateUI: () => {},
      getActivePresetName: () => '',
      reload: () => {},
    };

const headerController = window.TaskAppHeader?.initHeader({
  title: 'タスク・カレンダー',
  currentView: 'calendar',
  onAdd: () => openCreate(),
  onSave: () => handleSaveToExcel(),
  onValidations: () => openValidationModal(),
  onReload: () => handleReloadFromExcel(),
});

function updateHeaderDueSummary(tasks) {
  if (headerController && typeof headerController.updateDueSummary === 'function') {
    headerController.updateDueSummary(tasks);
  } else {
    window.TaskAppHeader?.updateDueSummary(tasks);
  }
}

if (typeof createExcelSyncHandlers === 'function') {
  excelSyncHandlers = createExcelSyncHandlers({
    apiAccessor: () => api,
    onAfterValidationSave: async ({ payload, response, closeModal }) => {
      const validationsInput = response?.validations ?? payload;
      const baseStatuses = Array.isArray(response?.statuses) ? response.statuses : STATUSES;
      if (Array.isArray(baseStatuses)) {
        STATUSES = baseStatuses;
      }

      const snapshot = applyValidationState({
        tasks: TASKS,
        statuses: STATUSES,
        validations: validationsInput,
      }) || {};

      TASKS = Array.isArray(snapshot.tasks) ? snapshot.tasks : TASKS;
      if (Array.isArray(snapshot.statuses) && snapshot.statuses.length > 0) {
        STATUSES = snapshot.statuses;
      }
      if (snapshot.validations && typeof snapshot.validations === 'object') {
        VALIDATIONS = snapshot.validations;
      }

      ensureMonthDefault();
      renderLegend();
      renderCalendar();
      renderBacklog();
      filterPresetManager.updateUI();
      updateHeaderDueSummary(TASKS);

      if (typeof closeModal === 'function') {
        closeModal();
      }
    },
  });
}

const priorityHelper = createPriorityHelper({
  getValidations: () => VALIDATIONS,
  defaultOptions: PRIORITY_DEFAULT_OPTIONS,
});
const applyPriorityOptions = (selectEl, currentValue, preferDefault = false) => (
  priorityHelper.applyOptions(selectEl, currentValue, preferDefault)
);
const getDefaultPriorityValue = () => priorityHelper.getDefaultValue();

function syncFiltersFromUI() {
  const monthInput = document.getElementById('month-picker');
  FILTERS.month = monthInput ? (monthInput.value || '') : '';
}

function applyFiltersToUI() {
  const monthInput = document.getElementById('month-picker');
  if (monthInput) {
    monthInput.value = FILTERS.month || '';
  }
}


if (typeof setupRuntime === 'function') {
  setupRuntime({
    mockApiFactory: createMockApi,
    onApiChanged: ({ api: nextApi, runMode }) => {
      api = nextApi;
      RUN_MODE = runMode;
      console.log('[calendar] run mode:', RUN_MODE);
    },
    onInit: async () => {
      try {
        await init(true);
        if (RUN_MODE === 'pywebview') {
          markInitialExcelLoadFlag?.();
        }
      } catch (err) {
        resetInitialExcelLoadFlag?.();
        throw err;
      }
    },
    onRealtimeUpdate: (payload) => applyStateFromPayload(payload, { fallbackToApi: false }),
  });
}

if (typeof bindExcelActions === 'function') {
  bindExcelActions({
    onSave: () => handleSaveToExcel(),
    onReload: () => handleReloadFromExcel(),
  });
}

ready(() => {
  wireControls();
  if (!cleanupAutoScroll && typeof setupDragViewportAutoScroll === 'function') {
    cleanupAutoScroll = setupDragViewportAutoScroll();
  }
});

async function init(force = false) {
  if (!api) return;
  if (force) {
    let payload = {};
    try {
      const isPywebview = RUN_MODE === 'pywebview';
      let loadedViaReload = false;
      if (isPywebview && !hasInitialExcelLoadFlag?.() && typeof api.reload_from_excel === 'function') {
        payload = normalizeStatePayload(await api.reload_from_excel());
        loadedViaReload = true;
      }
      if (isPywebview && !loadedViaReload && typeof api.get_state_snapshot === 'function') {
        payload = normalizeStatePayload(await api.get_state_snapshot());
      }
      if (!Array.isArray(payload.tasks) && typeof api.get_tasks === 'function') {
        payload.tasks = await api.get_tasks();
      }
      if (!Array.isArray(payload.statuses) && typeof api.get_statuses === 'function') {
        payload.statuses = await api.get_statuses();
      }
      if (payload.validations === undefined && typeof api.get_validations === 'function') {
        payload.validations = await api.get_validations();
      }
    } catch (err) {
      console.error('init failed', err);
      payload = {};
    }
    await applyStateFromPayload(payload, { fallbackToApi: true });
    return;
  }

  ensureMonthDefault();
  filterPresetManager.maybeApplyInitialPreset();
  applyFiltersToUI();
  syncFiltersFromUI();
  renderLegend();
  renderCalendar();
  renderBacklog();
  filterPresetManager.updateUI();
  updateHeaderDueSummary(TASKS);
}

async function applyStateFromPayload(payload, { fallbackToApi = false } = {}) {
  const data = normalizeStatePayload(payload);

  let tasksUpdated = false;
  if (Array.isArray(data.tasks)) {
    TASKS = sanitizeTaskList(data.tasks);
    tasksUpdated = true;
  }
  if (!tasksUpdated && fallbackToApi && typeof api?.get_tasks === 'function') {
    try {
      TASKS = sanitizeTaskList(await api.get_tasks());
      tasksUpdated = true;
    } catch (err) {
      console.error('get_tasks failed:', err);
      TASKS = [];
    }
  }

  let statusesUpdated = false;
  if (Array.isArray(data.statuses)) {
    STATUSES = data.statuses;
    statusesUpdated = true;
  }
  if (!statusesUpdated && fallbackToApi && typeof api?.get_statuses === 'function') {
    try {
      STATUSES = await api.get_statuses();
      statusesUpdated = true;
    } catch (err) {
      console.error('get_statuses failed:', err);
      STATUSES = [];
    }
  }

  const hasSnapshotValidations = Object.prototype.hasOwnProperty.call(data, 'validations');
  let validationPayload = hasSnapshotValidations ? data.validations : null;
  const shouldFetchValidations = (!hasSnapshotValidations || validationPayload == null)
    && fallbackToApi
    && typeof api?.get_validations === 'function';
  if (shouldFetchValidations) {
    try {
      validationPayload = await api.get_validations();
    } catch (err) {
      console.warn('get_validations failed:', err);
      validationPayload = null;
    }
  }
  if (!validationPayload || typeof validationPayload !== 'object') {
    validationPayload = VALIDATIONS;
  }
  const snapshot = applyValidationState({
    tasks: TASKS,
    statuses: STATUSES,
    validations: validationPayload,
  }) || {};
  TASKS = Array.isArray(snapshot.tasks) ? snapshot.tasks : TASKS;
  STATUSES = Array.isArray(snapshot.statuses) ? snapshot.statuses : STATUSES;
  VALIDATIONS = snapshot.validations && typeof snapshot.validations === 'object'
    ? snapshot.validations
    : VALIDATIONS;

  ensureMonthDefault();
  filterPresetManager.maybeApplyInitialPreset();
  applyFiltersToUI();
  syncFiltersFromUI();
  renderLegend();
  renderCalendar();
  renderBacklog();
  filterPresetManager.updateUI();
}

function wireControls() {
  const monthInput = document.getElementById('month-picker');
  const prevBtn = document.getElementById('btn-prev-month');
  const nextBtn = document.getElementById('btn-next-month');

  const shiftMonth = (delta) => {
    if (!monthInput) return;
    if (!monthInput.value) {
      ensureMonthDefault();
    }
    const target = parseMonthValue(monthInput.value);
    if (!target) return;
    target.setMonth(target.getMonth() + delta);
    monthInput.value = formatMonthValue(target);
    FILTERS.month = monthInput.value || '';
    updateMonthLabel();
    renderCalendar();
    renderBacklog();
    filterPresetManager.updateUI();
  };

  if (prevBtn) {
    prevBtn.addEventListener('click', () => shiftMonth(-1));
  }
  if (nextBtn) {
    nextBtn.addEventListener('click', () => shiftMonth(1));
  }
  if (monthInput) {
    monthInput.addEventListener('change', () => {
      FILTERS.month = monthInput.value || '';
      updateMonthLabel();
      renderCalendar();
      renderBacklog();
      filterPresetManager.updateUI();
    });
  }

  setupBacklogDropTarget();
  updateMonthLabel();
}

function handleSaveToExcel() {
  if (excelSyncHandlers?.handleSaveToExcel) {
    return excelSyncHandlers.handleSaveToExcel();
  }
  alert('保存機能が利用できません。');
}

async function handleReloadFromExcel() {
  if (!excelSyncHandlers?.handleReloadFromExcel) {
    alert('再読込機能が利用できません。');
    return;
  }
  await excelSyncHandlers.handleReloadFromExcel({
    onBeforeReload: () => resetInitialExcelLoadFlag?.(),
    onAfterReload: (payload) => applyStateFromPayload(payload, { fallbackToApi: true }),
  });
}

function setupBacklogDropTarget() {
  const dropZone = document.getElementById('calendar-backlog-drop');
  if (!dropZone) return;

  dropZone.addEventListener('dragover', (event) => {
    if (!getDraggedTaskNo(event)) return;
    event.preventDefault();
    dropZone.classList.add('drop-target');
    event.dataTransfer.dropEffect = 'move';
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drop-target');
  });

  dropZone.addEventListener('drop', async (event) => {
    event.preventDefault();
    dropZone.classList.remove('drop-target');
    const no = getDraggedTaskNo(event);
    if (!no) return;
    await updateTaskDue(no, '');
    CURRENT_DRAG = null;
    renderCalendar();
    renderBacklog();
  });
}

function ensureMonthDefault() {
  const monthInput = document.getElementById('month-picker');
  if (!monthInput || monthInput.value) {
    updateMonthLabel();
    return;
  }
  const validDueDates = TASKS
    .map(task => parseISO(task.期限))
    .filter(Boolean)
    .sort((a, b) => a - b);
  const base = validDueDates.length ? validDueDates[0] : new Date();
  monthInput.value = formatMonthValue(base);
  FILTERS.month = monthInput.value || '';
  updateMonthLabel();
}

const LEGEND_STATUS_ORDER = ['未着手', '進行中', '保留', '完了'];

function renderLegend() {
  const legend = document.getElementById('calendar-legend');
  if (!legend) return;
  legend.innerHTML = '';

  const baseStatuses = Array.isArray(STATUSES) && STATUSES.length ? STATUSES : [];
  const taskStatuses = TASKS.map(t => t.ステータス).filter(v => v !== undefined && v !== null);
  const combined = [
    ...LEGEND_STATUS_ORDER,
    ...DEFAULT_STATUSES,
    ...baseStatuses,
    ...taskStatuses,
  ];

  const seen = new Set();

  const appendLegendItem = (labelText, colorKey = labelText) => {
    const dot = document.createElement('span');
    dot.className = 'legend-dot';
    dot.style.background = statusColor(colorKey);
    const item = document.createElement('span');
    item.className = 'legend-item';
    item.appendChild(dot);
    const label = document.createElement('span');
    label.textContent = labelText;
    item.appendChild(label);
    legend.appendChild(item);
  };

  const registerStatus = (labelText) => {
    appendLegendItem(labelText);
    seen.add(labelText);
  };

  combined.forEach((rawName) => {
    const normalized = normalizeStatusLabel(rawName);
    if (normalized === UNSET_STATUS_LABEL) return;
    if (seen.has(normalized)) return;
    registerStatus(normalized);
  });

  const hasUnset = TASKS.some(task => normalizeStatusLabel(task.ステータス) === UNSET_STATUS_LABEL);
  if (hasUnset && !seen.has(UNSET_STATUS_LABEL)) {
    registerStatus(UNSET_STATUS_LABEL);
  }

  const hasOther = TASKS.some((task) => {
    const normalized = normalizeStatusLabel(task.ステータス);
    if (normalized === UNSET_STATUS_LABEL) return false;
    return !seen.has(normalized);
  });
  if (hasOther) {
    appendLegendItem('その他', 'other');
  }
}

function renderCalendar() {
  const container = document.getElementById('calendar-grid');
  if (!container) return;

  syncFiltersFromUI();
  const range = getCalendarRange();
  if (!range) {
    container.innerHTML = '<div class="message">表示する月を選択してください。</div>';
    return;
  }

  const { gridStart, gridEnd, monthStart, monthEnd } = range;
  const days = enumerateDays(gridStart, gridEnd);

  const tasksByDate = new Map();
  TASKS.forEach(task => {
    const iso = (task.期限 || '').slice(0, 10);
    if (!iso) return;
    if (!tasksByDate.has(iso)) tasksByDate.set(iso, []);
    tasksByDate.get(iso).push(task);
  });

  tasksByDate.forEach(list => {
    list.sort((a, b) => {
      const statusA = String(a.ステータス || '');
      const statusB = String(b.ステータス || '');
      if (statusA !== statusB) return statusA.localeCompare(statusB, 'ja');
      if (a.No && b.No && a.No !== b.No) return a.No - b.No;
      return (a.タスク || '').localeCompare(b.タスク || '', 'ja');
    });
  });

  const fragment = document.createDocumentFragment();
  days.forEach(day => {
    const cell = document.createElement('div');
    cell.className = 'calendar-cell';
    if (day.date.getMonth() !== monthStart.getMonth()) {
      cell.classList.add('outside-month');
    }
    cell.dataset.date = day.iso;

    cell.addEventListener('dragover', (event) => {
      if (!getDraggedTaskNo(event)) return;
      event.preventDefault();
      cell.classList.add('drop-target');
      event.dataTransfer.dropEffect = 'move';
    });
    cell.addEventListener('dragleave', () => {
      cell.classList.remove('drop-target');
    });
    cell.addEventListener('drop', async (event) => {
      event.preventDefault();
      cell.classList.remove('drop-target');
      const no = getDraggedTaskNo(event);
      if (!no) return;
      await updateTaskDue(no, day.iso);
      CURRENT_DRAG = null;
      renderCalendar();
      renderBacklog();
    });

    const header = document.createElement('div');
    header.className = 'calendar-cell-header';
    const dateLabel = document.createElement('span');
    dateLabel.className = 'calendar-date';
    dateLabel.textContent = `${day.date.getDate()}`;
    header.appendChild(dateLabel);
    const weekday = document.createElement('span');
    weekday.className = 'weekday';
    weekday.textContent = day.weekday;
    header.appendChild(weekday);
    cell.appendChild(header);

    const list = document.createElement('div');
    list.className = 'calendar-tasks';

    const items = tasksByDate.get(day.iso) || [];
    if (!items.length) {
      const empty = document.createElement('div');
      empty.className = 'empty-message';
      empty.textContent = 'タスクなし';
      list.appendChild(empty);
    } else {
      items.forEach(task => {
        const card = renderTaskCard(task);
        list.appendChild(card);
      });
    }

    cell.appendChild(list);
    fragment.appendChild(cell);
  });

  container.innerHTML = '';
  container.appendChild(fragment);

  const summary = document.getElementById('calendar-current');
  if (summary) {
    const inRange = TASKS.filter(task => {
      const due = parseISO(task.期限);
      return due && due >= monthStart && due <= monthEnd;
    }).length;
    summary.textContent = `${monthStart.getFullYear()}年${monthStart.getMonth() + 1}月のタスク: ${inRange} 件`;
  }
}

function renderBacklog() {
  const container = document.getElementById('calendar-backlog-content');
  if (!container) return;

  const tasks = TASKS.filter(task => !parseISO(task.期限));

  container.innerHTML = '';
  if (!tasks.length) {
    const empty = document.createElement('div');
    empty.className = 'message backlog-empty';
    empty.textContent = '期限未設定のタスクはありません。';
    container.appendChild(empty);
    return;
  }

  tasks.sort((a, b) => {
    const dueA = parseISO(a.期限);
    const dueB = parseISO(b.期限);
    if (dueA && dueB && dueA.getTime() !== dueB.getTime()) return dueA - dueB;
    if (dueA && !dueB) return 1;
    if (!dueA && dueB) return -1;
    if (a.No && b.No && a.No !== b.No) return a.No - b.No;
    return (a.タスク || '').localeCompare(b.タスク || '', 'ja');
  });

  const list = document.createElement('div');
  list.className = 'calendar-backlog-list';

  tasks.forEach(task => {
    const item = document.createElement('div');
    item.className = 'calendar-backlog-item';
    item.dataset.no = task.No ?? '';
    item.draggable = true;

    item.addEventListener('dragstart', (event) => {
      startTaskDrag(task, item, event);
    });
    item.addEventListener('dragend', () => {
      finishTaskDrag(item);
    });

    item.addEventListener('dblclick', () => {
      if (task?.No) {
        openEdit(task.No);
      }
    });

    const title = document.createElement('div');
    title.className = 'calendar-backlog-item-title';
    title.textContent = task.タスク || '(タイトルなし)';
    item.appendChild(title);

    const meta = document.createElement('div');
    meta.className = 'calendar-backlog-item-meta';
    meta.appendChild(createLabelValueSpan('ステータス', task.ステータス));
    if (task.No) {
      const no = document.createElement('span');
      no.textContent = `No.${task.No}`;
      meta.appendChild(no);
    }
    const due = parseISO(task.期限);
    const dueLabel = document.createElement('span');
    dueLabel.textContent = due ? `期限: ${toLocale(due)}` : '期限: 未設定';
    meta.appendChild(dueLabel);
    meta.appendChild(createLabelValueSpan('大分類', task.大分類));
    meta.appendChild(createLabelValueSpan('中分類', task.中分類));
    meta.appendChild(createLabelValueSpan('重要度', task.優先度));
    meta.appendChild(createLabelValueSpan('担当', task.担当者));
    item.appendChild(meta);

    list.appendChild(item);
  });

  container.appendChild(list);
}

function createLabelValueSpan(label, value) {
  const span = document.createElement('span');
  const text = String(value ?? '').trim();
  span.textContent = `${label}: ${text || '未設定'}`;
  return span;
}

function renderTaskCard(task) {
  const card = document.createElement('div');
  card.className = 'calendar-task';
  if (task.ステータス) {
    card.dataset.status = task.ステータス;
  }
  card.dataset.no = task.No ?? '';
  card.draggable = true;

  card.addEventListener('dragstart', (event) => {
    startTaskDrag(task, card, event);
  });
  card.addEventListener('dragend', () => {
    finishTaskDrag(card);
  });

  card.addEventListener('dblclick', () => {
    if (task?.No) {
      openEdit(task.No);
    }
  });

  const title = document.createElement('div');
  title.className = 'task-title';
  title.textContent = task.タスク || '(タイトルなし)';
  card.appendChild(title);

  const meta = document.createElement('div');
  meta.className = 'task-meta';

  meta.appendChild(createLabelValueSpan('ステータス', task.ステータス));

  const majorLabel = String(task.大分類 ?? '').trim();
  meta.appendChild(createLabelValueSpan('大分類', majorLabel));

  const minorLabel = String(task.中分類 ?? '').trim();
  meta.appendChild(createLabelValueSpan('中分類', minorLabel));

  const priorityText = String(task.優先度 ?? '').trim();
  meta.appendChild(createLabelValueSpan('重要度', priorityText));

  const assigneeLabel = String(task.担当者 ?? '').trim();
  meta.appendChild(createLabelValueSpan('担当', assigneeLabel));

  if (task.No) {
    const no = document.createElement('span');
    no.textContent = `No.${task.No}`;
    meta.appendChild(no);
  }

  card.appendChild(meta);

  return card;
}

function startTaskDrag(task, element, event) {
  CURRENT_DRAG = task?.No ?? null;
  element.classList.add('dragging');
  if (event?.dataTransfer) {
    event.dataTransfer.setData('text/plain', String(task?.No ?? ''));
    event.dataTransfer.effectAllowed = 'move';
  }
}

function finishTaskDrag(element) {
  element.classList.remove('dragging');
  CURRENT_DRAG = null;
}

function getDraggedTaskNo(event) {
  if (event?.dataTransfer) {
    const text = event.dataTransfer.getData('text/plain');
    const no = Number(text);
    if (Number.isInteger(no) && no > 0) return no;
  }
  if (CURRENT_DRAG && Number.isInteger(CURRENT_DRAG)) {
    return CURRENT_DRAG;
  }
  return null;
}

function parseISO(value) {
  return parseISODate(value);
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

function enumerateDays(from, to) {
  const days = [];
  const cursor = new Date(from);
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  while (cursor <= to) {
    days.push({
      iso: toISODate(cursor),
      date: new Date(cursor),
      weekday: `${weekdays[cursor.getDay()]}曜`,
    });
    cursor.setDate(cursor.getDate() + 1);
  }
  return days;
}

function parseMonthValue(value) {
  if (!value) return null;
  const m = /^(\d{4})-(\d{2})$/.exec(value);
  if (!m) return null;
  const dt = new Date(Number(m[1]), Number(m[2]) - 1, 1);
  if (Number.isNaN(dt.getTime())) return null;
  dt.setHours(0, 0, 0, 0);
  return dt;
}

function formatMonthValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

function getCalendarRange() {
  const monthInput = document.getElementById('month-picker');
  const value = FILTERS.month || (monthInput ? monthInput.value : '');
  if (!value) return null;
  const monthDate = parseMonthValue(value);
  if (!monthDate) return null;
  const monthStart = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  monthStart.setHours(0, 0, 0, 0);
  const monthEnd = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
  monthEnd.setHours(0, 0, 0, 0);
  const gridStart = startOfWeek(monthStart);
  const gridEnd = new Date(monthEnd);
  gridEnd.setDate(gridEnd.getDate() + (6 - gridEnd.getDay()));
  gridEnd.setHours(0, 0, 0, 0);
  return { monthStart, monthEnd, gridStart, gridEnd };
}

function updateMonthLabel() {
  const monthInput = document.getElementById('month-picker');
  const label = document.getElementById('calendar-current');
  const value = FILTERS.month || (monthInput ? monthInput.value : '');
  if (!label) return;
  if (!value) {
    label.textContent = '';
    return;
  }
  if (monthInput) {
    monthInput.value = value;
  }
  const monthDate = parseMonthValue(value);
  if (!monthDate) {
    label.textContent = '';
    return;
  }
  label.textContent = `${monthDate.getFullYear()}年${monthDate.getMonth() + 1}月のタスク`;
}

function openValidationModal(options = {}) {
  if (!excelSyncHandlers?.openValidationModal) return;
  excelSyncHandlers.openValidationModal({
    columns: VALIDATION_COLUMNS,
    getCurrentValues: () => VALIDATIONS,
    onAfterRender: options.onAfterRender,
  });
}

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
  setupAssigneeInputSuggestions(fwho);
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
        closeModal();
        ensureMonthDefault();
        renderLegend();
        renderCalendar();
        renderBacklog();
        updateHeaderDueSummary(TASKS);
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
      備考: fnote.value,
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
      ensureMonthDefault();
      renderLegend();
      renderCalendar();
      renderBacklog();
      updateHeaderDueSummary(TASKS);
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

function setupAssigneeInputSuggestions(fassignee) {
  const datalist = document.getElementById('modal-assignee-list');
  if (!fassignee || !datalist) return;

  const candidates = collectAllAssignees()
    .map(name => (name === '' ? '' : String(name ?? '').trim()))
    .filter(Boolean);

  const seen = new Set();
  datalist.innerHTML = '';

  candidates.forEach(name => {
    if (seen.has(name)) return;
    seen.add(name);
    const opt = document.createElement('option');
    opt.value = name;
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

function collectAllAssignees() {
  const assignees = new Set();
  TASKS.forEach(task => {
    const name = String(task.担当者 ?? '').trim();
    if (!name) return;
    assignees.add(name);
  });
  return Array.from(assignees).sort((a, b) => a.localeCompare(b, 'ja'));
}

function statusColor(name) {
  const normalized = normalizeStatusLabel(name);
  if (normalized === UNSET_STATUS_LABEL) {
    return 'rgba(148, 163, 184, 0.8)';
  }

  const key = normalized.replace(/\s+/g, '').toLowerCase();
  switch (key) {
    case '未着手':
    case 'notstarted':
      return 'rgba(248, 113, 113, 0.8)';
    case '進行中':
    case '対応中':
    case 'inprogress':
      return 'rgba(56, 189, 248, 0.8)';
    case '保留':
    case 'pending':
      return 'rgba(250, 204, 21, 0.8)';
    case '完了':
    case '完了済み':
    case '完了済':
    case 'done':
    case 'completed':
      return 'rgba(34, 197, 94, 0.8)';
    default:
      return 'rgba(129, 140, 248, 0.8)';
  }
}

async function updateTaskDue(no, dueIso) {
  const idx = TASKS.findIndex(task => task.No === no);
  if (idx < 0) return;
  const current = TASKS[idx];
  const previousDue = current.期限 || '';
  if ((previousDue || '') === (dueIso || '')) return;

  try {
    if (typeof api?.update_task === 'function') {
      const updated = await api.update_task(no, { 期限: dueIso });
      const sanitized = sanitizeTaskRecord(updated, idx);
      if (sanitized) {
        TASKS[idx] = sanitized;
      } else {
        TASKS[idx] = { ...TASKS[idx], 期限: dueIso };
      }
    } else {
      TASKS[idx] = { ...TASKS[idx], 期限: dueIso };
    }
  } catch (err) {
    alert('期限の更新に失敗: ' + (err?.message || err));
  }
}

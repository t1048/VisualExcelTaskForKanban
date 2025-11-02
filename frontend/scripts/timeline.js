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
  setupDragViewportAutoScroll,
  loadFilterPresets,
  saveFilterPreset,
  deleteFilterPreset,
  applyFilterPreset,
  PRIORITY_DEFAULT_OPTIONS,
  UNSET_STATUS_LABEL,
} = window.TaskAppCommon;

let api;
let RUN_MODE = 'mock';
let TASKS = [];
let STATUSES = [];
let VALIDATIONS = {};
const VALIDATION_COLUMNS = ["ステータス", "大分類", "中分類", "タスク", "担当者", "優先度", "期限", "備考"];
let CURRENT_EDIT = null;
const ASSIGNEE_FILTER_ALL = '';
const ASSIGNEE_FILTER_UNASSIGNED = '__UNASSIGNED__';
const ASSIGNEE_UNASSIGNED_LABEL = '（未割り当て）';
let cleanupAutoScroll = null;

let FILTERS = {
  range: { from: '', to: '' },
  assignee: ASSIGNEE_FILTER_ALL,
};

const FILTER_PRESET_VIEW_KEY = 'timeline';
let FILTER_PRESETS = [];
let ACTIVE_FILTER_PRESET = '';
let PRESET_INITIAL_APPLIED = false;

function initializeFilterPresetsState() {
  try {
    const { presets, lastApplied } = loadFilterPresets(FILTER_PRESET_VIEW_KEY) || {};
    FILTER_PRESETS = Array.isArray(presets) ? presets : [];
    ACTIVE_FILTER_PRESET = lastApplied?.name || '';
  } catch (err) {
    console.warn('[timeline] failed to load filter presets', err);
    FILTER_PRESETS = [];
    ACTIVE_FILTER_PRESET = '';
  }
}

initializeFilterPresetsState();

const INITIAL_LOAD_FLAG_KEY = 'kanban:excelLoaded';

const headerController = window.TaskAppHeader?.initHeader({
  title: '担当者タイムライン',
  currentView: 'timeline',
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

function hasInitialExcelLoadFlag() {
  try {
    return window.sessionStorage?.getItem(INITIAL_LOAD_FLAG_KEY) === '1';
  } catch (err) {
    return false;
  }
}

function markInitialExcelLoadFlag() {
  try {
    window.sessionStorage?.setItem(INITIAL_LOAD_FLAG_KEY, '1');
  } catch (err) {
    // ignore
  }
}

function resetInitialExcelLoadFlag() {
  try {
    window.sessionStorage?.removeItem(INITIAL_LOAD_FLAG_KEY);
  } catch (err) {
    // ignore
  }
}

const priorityHelper = createPriorityHelper({
  getValidations: () => VALIDATIONS,
  defaultOptions: PRIORITY_DEFAULT_OPTIONS,
});
const getPriorityOptions = () => priorityHelper.getOptions();
const getDefaultPriorityValue = () => priorityHelper.getDefaultValue();
const applyPriorityOptions = (selectEl, currentValue, preferDefault = false) => (
  priorityHelper.applyOptions(selectEl, currentValue, preferDefault)
);

function syncFiltersFromUI() {
  const fromInput = document.getElementById('date-from');
  const toInput = document.getElementById('date-to');
  FILTERS.range = {
    from: fromInput ? fromInput.value : '',
    to: toInput ? toInput.value : '',
  };
  const assigneeSelect = document.getElementById('assignee-filter');
  FILTERS.assignee = assigneeSelect ? assigneeSelect.value : ASSIGNEE_FILTER_ALL;
}

function applyFiltersToUI() {
  const fromInput = document.getElementById('date-from');
  const toInput = document.getElementById('date-to');
  if (fromInput) fromInput.value = FILTERS.range.from || '';
  if (toInput) toInput.value = FILTERS.range.to || '';
  const assigneeSelect = document.getElementById('assignee-filter');
  if (assigneeSelect) {
    const target = FILTERS.assignee || ASSIGNEE_FILTER_ALL;
    const hasOption = Array.from(assigneeSelect.options).some(opt => opt.value === target);
    assigneeSelect.value = hasOption ? target : ASSIGNEE_FILTER_ALL;
    FILTERS.assignee = assigneeSelect.value;
  }
}

function serializeFiltersForPreset() {
  return {
    range: {
      from: FILTERS.range.from || '',
      to: FILTERS.range.to || '',
    },
    assignee: FILTERS.assignee || ASSIGNEE_FILTER_ALL,
  };
}

function applyPresetFilters(raw) {
  const data = raw && typeof raw === 'object' ? raw : {};
  const range = data.range && typeof data.range === 'object' ? data.range : {};
  FILTERS.range = {
    from: String(range.from ?? ''),
    to: String(range.to ?? ''),
  };
  const assignee = String(data.assignee ?? '').trim();
  FILTERS.assignee = assignee || ASSIGNEE_FILTER_ALL;
  applyFiltersToUI();
}

function maybeApplyInitialPreset() {
  if (PRESET_INITIAL_APPLIED) return;
  PRESET_INITIAL_APPLIED = true;
  if (!ACTIVE_FILTER_PRESET) return;
  const preset = FILTER_PRESETS.find(item => item?.name === ACTIVE_FILTER_PRESET);
  if (!preset) {
    ACTIVE_FILTER_PRESET = '';
    return;
  }
  applyPresetFilters(preset.filters);
}

function ensureFilterPresetHandlers() {
  const select = document.getElementById('timeline-preset');
  const applyBtn = document.getElementById('btn-timeline-preset-apply');
  const saveBtn = document.getElementById('btn-timeline-preset-save');
  const deleteBtn = document.getElementById('btn-timeline-preset-delete');
  if (!select || !applyBtn || !saveBtn || !deleteBtn) return;
  if (select.dataset.bound === '1') return;
  select.dataset.bound = '1';

  const refreshButtonState = () => {
    const selected = select.value;
    const exists = Boolean(selected) && FILTER_PRESETS.some(preset => preset?.name === selected);
    applyBtn.disabled = !exists;
    deleteBtn.disabled = !exists;
  };

  select.addEventListener('change', () => {
    ACTIVE_FILTER_PRESET = select.value;
    refreshButtonState();
  });

  applyBtn.addEventListener('click', () => {
    const targetName = select.value;
    if (!targetName) {
      alert('プリセットを選択してください。');
      return;
    }
    const result = applyFilterPreset(FILTER_PRESET_VIEW_KEY, targetName, (filters) => {
      applyPresetFilters(filters);
      return true;
    });
    FILTER_PRESETS = result.presets;
    if (result.applied) {
      ACTIVE_FILTER_PRESET = result.applied.name;
      PRESET_INITIAL_APPLIED = true;
      syncFiltersFromUI();
      renderSummary();
      renderAssigneeFilter();
      renderTimeline();
    } else {
      alert('選択したプリセットが見つかりません。');
      updateFilterPresetUI();
    }
  });

  saveBtn.addEventListener('click', () => {
    const defaultName = select.value || '';
    const name = window.prompt('プリセット名を入力してください', defaultName);
    if (name === null) return;
    const trimmed = name.trim();
    if (!trimmed) {
      alert('プリセット名を入力してください。');
      return;
    }
    syncFiltersFromUI();
    const payload = serializeFiltersForPreset();
    const result = saveFilterPreset(FILTER_PRESET_VIEW_KEY, trimmed, payload);
    FILTER_PRESETS = result.presets;
    if (result.saved) {
      ACTIVE_FILTER_PRESET = result.saved.name;
      PRESET_INITIAL_APPLIED = true;
    }
    updateFilterPresetUI();
  });

  deleteBtn.addEventListener('click', () => {
    const targetName = select.value;
    if (!targetName) {
      alert('削除するプリセットを選択してください。');
      return;
    }
    if (!window.confirm(`プリセット「${targetName}」を削除しますか？`)) {
      return;
    }
    const result = deleteFilterPreset(FILTER_PRESET_VIEW_KEY, targetName);
    FILTER_PRESETS = result.presets;
    if (ACTIVE_FILTER_PRESET === targetName) {
      ACTIVE_FILTER_PRESET = '';
    }
    updateFilterPresetUI();
  });

  refreshButtonState();
}

function updateFilterPresetUI() {
  ensureFilterPresetHandlers();
  const select = document.getElementById('timeline-preset');
  const applyBtn = document.getElementById('btn-timeline-preset-apply');
  const deleteBtn = document.getElementById('btn-timeline-preset-delete');
  if (!select) return;

  const previousValue = select.value;
  select.innerHTML = '';

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '（プリセット未選択）';
  select.appendChild(placeholder);

  FILTER_PRESETS.forEach(preset => {
    if (!preset || typeof preset.name !== 'string') return;
    const opt = document.createElement('option');
    opt.value = preset.name;
    opt.textContent = preset.name;
    select.appendChild(opt);
  });

  let nextValue = '';
  if (ACTIVE_FILTER_PRESET && FILTER_PRESETS.some(p => p?.name === ACTIVE_FILTER_PRESET)) {
    nextValue = ACTIVE_FILTER_PRESET;
  } else if (FILTER_PRESETS.some(p => p?.name === previousValue)) {
    nextValue = previousValue;
    ACTIVE_FILTER_PRESET = previousValue;
  } else {
    ACTIVE_FILTER_PRESET = '';
  }

  select.value = nextValue;
  const hasSelection = Boolean(select.value);
  if (applyBtn) applyBtn.disabled = !hasSelection;
  if (deleteBtn) deleteBtn.disabled = !hasSelection;
}

setupRuntime({
  mockApiFactory: createMockApi,
  onApiChanged: ({ api: nextApi, runMode }) => {
    api = nextApi;
    RUN_MODE = runMode;
    console.log('[timeline] run mode:', RUN_MODE);
  },
  onInit: async () => {
    try {
      await init(true);
      if (RUN_MODE === 'pywebview') {
        markInitialExcelLoadFlag();
      }
    } catch (err) {
      resetInitialExcelLoadFlag();
      throw err;
    }
  },
  onRealtimeUpdate: (payload) => {
    resetInitialExcelLoadFlag();
    return applyStateFromPayload(payload, { fallbackToApi: false });
  },
});

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
      if (isPywebview && !hasInitialExcelLoadFlag() && typeof api.reload_from_excel === 'function') {
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

  ensureRangeDefaults();
  maybeApplyInitialPreset();
  syncFiltersFromUI();
  renderSummary();
  renderLegend();
  renderAssigneeFilter();
  renderTimeline();
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

  let validationPayload = data.validations;
  if (!validationPayload && fallbackToApi && typeof api?.get_validations === 'function') {
    try {
      validationPayload = await api.get_validations();
    } catch (err) {
      console.warn('get_validations failed:', err);
      validationPayload = null;
    }
  }
  if (validationPayload !== undefined) {
    applyValidationState(validationPayload);
  }

  ensureRangeDefaults();
  maybeApplyInitialPreset();
  syncFiltersFromUI();
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
  syncFiltersFromUI();
}

function wireControls() {
  document.getElementById('btn-apply').addEventListener('click', () => {
    syncFiltersFromUI();
    renderSummary();
    renderAssigneeFilter();
    renderTimeline();
  });

  document.querySelectorAll('.quick button[data-shift]').forEach(btn => {
    btn.addEventListener('click', () => {
      const delta = Number(btn.dataset.shift || '0');
      shiftRange(delta);
      syncFiltersFromUI();
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
    syncFiltersFromUI();
    renderSummary();
    renderAssigneeFilter();
    renderTimeline();
  });

  const assigneeSelect = document.getElementById('assignee-filter');
  assigneeSelect.addEventListener('change', () => {
    FILTERS.assignee = assigneeSelect.value || ASSIGNEE_FILTER_ALL;
    renderTimeline();
  });

}

async function handleSaveToExcel() {
  if (!api || typeof api.save_excel !== 'function') {
    alert('保存機能が利用できません。');
    return;
  }
  try {
    const result = await api.save_excel();
    const message = result ? `Excelへ保存しました\n${result}` : 'Excelへ保存しました';
    alert(message);
  } catch (err) {
    alert('保存に失敗: ' + (err?.message || err));
  }
}

async function handleReloadFromExcel() {
  if (!api || typeof api.reload_from_excel !== 'function') {
    alert('再読込機能が利用できません。');
    return;
  }
  resetInitialExcelLoadFlag();
  try {
    const payload = await api.reload_from_excel();
    await applyStateFromPayload(payload, { fallbackToApi: true });
  } catch (err) {
    alert('再読込に失敗: ' + (err?.message || err));
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
  syncFiltersFromUI();
}

function setRange(from, to) {
  const inputFrom = document.getElementById('date-from');
  const inputTo = document.getElementById('date-to');
  inputFrom.value = toISODate(from);
  inputTo.value = toISODate(to);
  FILTERS.range = {
    from: inputFrom.value,
    to: inputTo.value,
  };
}

function renderSummary() {
  syncFiltersFromUI();
  const from = parseISO(FILTERS.range.from);
  const to = parseISO(FILTERS.range.to);
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
  const backlogTasks = collectBacklogTasks(from, to);
  const backlogCount = backlogTasks.length;
  const withoutDue = backlogTasks.filter(t => !parseISO(t.期限)).length;
  summary.textContent = `表示期間: ${toLocale(from)} 〜 ${toLocale(to)} （${diff}日間） / 対象タスク ${inRange} 件、バックログ ${backlogCount} 件（期限未設定 ${withoutDue} 件）`;
}

function renderAssigneeFilter() {
  const select = document.getElementById('assignee-filter');
  if (!select) return;

  const targetAssignee = FILTERS.assignee || ASSIGNEE_FILTER_ALL;
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

  if (targetAssignee === ASSIGNEE_FILTER_UNASSIGNED) {
    select.value = ASSIGNEE_FILTER_UNASSIGNED;
  } else if (targetAssignee && assignees.includes(targetAssignee)) {
    select.value = targetAssignee;
  } else {
    select.value = ASSIGNEE_FILTER_ALL;
  }
  FILTERS.assignee = select.value;
  updateFilterPresetUI();
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

function collectBacklogTasks(rangeFrom, rangeTo) {
  return TASKS.filter(task => {
    const due = parseISO(task.期限);
    const assignee = String(task.担当者 ?? '').trim();
    return !due || !assignee;
  });
}

function renderTimeline() {
  const wrapper = document.getElementById('timeline-wrapper');
  syncFiltersFromUI();
  const from = parseISO(FILTERS.range.from);
  const to = parseISO(FILTERS.range.to);
  const assigneeFilter = FILTERS.assignee || ASSIGNEE_FILTER_ALL;
  const backlogTasks = collectBacklogTasks(from, to);
  renderBacklog(backlogTasks, { from, to });
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

function renderBacklog(tasks, { from, to } = {}) {
  const container = document.getElementById('backlog-content');
  if (!container) return;

  container.innerHTML = '';

  if (!tasks.length) {
    const empty = document.createElement('div');
    empty.className = 'message backlog-empty';
    empty.textContent = 'バックログはありません。';
    container.appendChild(empty);
    return;
  }

  const groups = new Map();
  tasks.forEach(task => {
    const name = task.担当者?.trim() || ASSIGNEE_UNASSIGNED_LABEL;
    if (!groups.has(name)) {
      groups.set(name, []);
    }
    groups.get(name).push(task);
  });

  const assignees = Array.from(groups.keys()).sort((a, b) => a.localeCompare(b, 'ja'));

  assignees.forEach(name => {
    const group = document.createElement('div');
    group.className = 'backlog-group';

    const heading = document.createElement('h3');
    heading.className = 'backlog-group-title';
    heading.textContent = name;
    group.appendChild(heading);

    const list = document.createElement('ul');
    list.className = 'backlog-list';

    const items = groups.get(name);
    items.sort((a, b) => {
      const dueA = parseISO(a.期限);
      const dueB = parseISO(b.期限);
      if (dueA && dueB && dueA.getTime() !== dueB.getTime()) return dueA - dueB;
      if (dueA && !dueB) return -1;
      if (!dueA && dueB) return 1;
      if (a.No && b.No && a.No !== b.No) return a.No - b.No;
      return (a.タスク || '').localeCompare(b.タスク || '', 'ja');
    });

    items.forEach(task => {
      const item = document.createElement('li');
      item.className = 'backlog-item';
      item.dataset.no = task.No ?? '';

      item.addEventListener('dblclick', () => {
        if (task?.No) {
          openEdit(task.No);
        }
      });

      const title = document.createElement('div');
      title.className = 'backlog-item-title';
      title.textContent = task.タスク || '(タイトルなし)';
      item.appendChild(title);

      const meta = document.createElement('div');
      meta.className = 'backlog-item-meta';

      const status = document.createElement('span');
      status.textContent = `ステータス: ${task.ステータス || '未設定'}`;
      meta.appendChild(status);

      if (task.No) {
        const no = document.createElement('span');
        no.textContent = `No.${task.No}`;
        meta.appendChild(no);
      }

      const due = parseISO(task.期限);
      const dueLabel = document.createElement('span');
      if (due) {
        let label = `期限: ${toLocale(due)}`;
        if (from && due < from) {
          label += '（開始前）';
        } else if (to && due > to) {
          label += '（終了後）';
        }
        dueLabel.textContent = label;
      } else {
        dueLabel.textContent = '期限: 未設定';
      }
      meta.appendChild(dueLabel);

      if (task.備考) {
        const notes = document.createElement('span');
        notes.textContent = task.備考;
        meta.appendChild(notes);
      }

      item.appendChild(meta);
      list.appendChild(item);
    });

    group.appendChild(list);
    container.appendChild(group);
  });
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
  const majorLabel = String(task.大分類 ?? '').trim();
  const major = document.createElement('span');
  major.textContent = `大分類: ${majorLabel || '未設定'}`;
  meta.appendChild(major);

  const minorLabel = String(task.中分類 ?? '').trim();
  const minor = document.createElement('span');
  minor.textContent = `中分類: ${minorLabel || '未設定'}`;
  meta.appendChild(minor);

  const priorityLabel = String(task.優先度 ?? '').trim();
  const priority = document.createElement('span');
  priority.textContent = `重要度: ${priorityLabel || '未設定'}`;
  meta.appendChild(priority);
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

function setupAssigneeInputSuggestions(fassignee) {
  const datalist = document.getElementById('modal-assignee-list');
  if (!fassignee || !datalist) return;

  const candidates = collectAllAssignees()
    .map(name => (name === ASSIGNEE_UNASSIGNED_LABEL ? '' : String(name ?? '').trim()))
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
  modal.onclick = (ev) => { if (ev.target === modal) closeValidationModal(); };

  if (saveBtn) {
    saveBtn.onclick = async () => {
      const payload = {};
      editor.querySelectorAll('textarea[data-column]').forEach(area => {
        const col = area.dataset.column;
        const lines = area.value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
        payload[col] = lines;
      });

      try {
        if (typeof api?.update_validations === 'function') {
          const res = await api.update_validations(payload);
          if (Array.isArray(res?.statuses)) {
            STATUSES = res.statuses;
          }
          const received = res?.validations ?? payload;
          applyValidationState(received);
        } else {
          applyValidationState(payload);
        }
        ensureRangeDefaults();
        renderSummary();
        renderLegend();
        renderAssigneeFilter();
        renderTimeline();
        updateHeaderDueSummary(TASKS);
        closeValidationModal();
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
        ensureRangeDefaults();
        closeModal();
        renderSummary();
        renderLegend();
        renderAssigneeFilter();
        renderTimeline();
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

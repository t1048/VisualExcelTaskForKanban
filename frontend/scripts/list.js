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
const MAJOR_EMPTY_LABEL = '（大分類なし）';
const MINOR_EMPTY_LABEL = '（中分類なし）';
const MAJOR_EMPTY_KEY = '__EMPTY_MAJOR__';
const MINOR_EMPTY_KEY = '__EMPTY_MINOR__';
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
const FILTER_COLLAPSED_STORAGE_KEY = 'taskList.filtersCollapsed';
const GROUP_CONTEXT_MENU_ID = 'group-context-menu';
let GROUP_CONTEXT_STATE = null;

function applyFilterCollapsedState() {
  const container = document.getElementById('filters-bar');
  if (!container) return;

  const toggle = container.querySelector('[data-filter-toggle]');
  const COLLAPSED_CLASS = 'is-collapsed';

  const setCollapsed = (collapsed) => {
    container.classList.toggle(COLLAPSED_CLASS, Boolean(collapsed));
    if (toggle) {
      toggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    }
  };

  let stored = null;
  try {
    stored = window.localStorage?.getItem(FILTER_COLLAPSED_STORAGE_KEY) ?? null;
  } catch (err) {
    console.warn('[list] failed to read filter collapsed state:', err);
  }

  const collapsed = stored === '1';
  setCollapsed(collapsed);

  if (!toggle || toggle.dataset.bound === '1') {
    return;
  }

  toggle.dataset.bound = '1';
  toggle.addEventListener('click', () => {
    const nextCollapsed = !container.classList.contains(COLLAPSED_CLASS);
    setCollapsed(nextCollapsed);
    try {
      window.localStorage?.setItem(FILTER_COLLAPSED_STORAGE_KEY, nextCollapsed ? '1' : '0');
    } catch (err) {
      console.warn('[list] failed to store filter collapsed state:', err);
    }
  });
}

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
    renderList();
  },
});

setupRuntime({
  mockApiFactory: createMockApi,
  onApiChanged: ({ api: nextApi, runMode }) => {
    api = nextApi;
    RUN_MODE = runMode;
    console.log('[list] run mode:', RUN_MODE);
  },
  onInit: async () => {
    await init(true);
  },
  onRealtimeUpdate: (payload) => (
    applyStateFromPayload(payload, { preserveFilters: true, fallbackToApi: false })
  ),
});

const STATUS_SORT_SEQUENCE = ['', UNSET_STATUS_LABEL, '未着手', '進行中', '完了', '保留中'];

const TABLE_COLUMN_CONFIG = [
  { key: 'no', label: 'No', width: 70, minWidth: 60, sortable: true },
  { key: 'major', label: '大分類', width: 160, minWidth: 120, sortable: true },
  { key: 'minor', label: '中分類', width: 160, minWidth: 120, sortable: true },
  { key: 'task', label: 'タスク', sortable: true, className: 'col-task', minWidth: 260 },
  { key: 'status', label: 'ステータス', width: 160, minWidth: 140, sortable: true },
  { key: 'assignee', label: '担当者', width: 160, minWidth: 140, sortable: true },
  { key: 'priority', label: '優先度', width: 140, minWidth: 120, sortable: false },
  { key: 'due', label: '期限', width: 160, minWidth: 140, sortable: true },
  { key: 'notes', label: '備考', sortable: false, className: 'col-notes', minWidth: 280 }
];

const COLUMN_CONFIG_BY_KEY = new Map();
TABLE_COLUMN_CONFIG.forEach((col, index) => {
  col.index = index;
  COLUMN_CONFIG_BY_KEY.set(col.key, col);
});

const COLUMN_WIDTH_STORAGE_KEY = 'taskList.columnWidths';
let COLUMN_WIDTHS = loadColumnWidths();

let horizontalScrollInitialized = false;
let horizontalScrollSyncing = false;
let horizontalScrollElements = null;
let pendingHorizontalScrollbarUpdate = false;

function getHorizontalScrollElements() {
  if (horizontalScrollElements) {
    const { scroller, bar, inner } = horizontalScrollElements;
    if (scroller?.isConnected && bar?.isConnected && inner?.isConnected) {
      return horizontalScrollElements;
    }
  }
  const scroller = document.getElementById('list-panel-scroll');
  const bar = document.getElementById('list-scrollbar');
  const inner = document.getElementById('list-scrollbar-inner');
  if (!scroller || !bar || !inner) {
    horizontalScrollElements = null;
    return null;
  }
  horizontalScrollElements = { scroller, bar, inner };
  return horizontalScrollElements;
}

function ensureHorizontalScrollBindings() {
  const elements = getHorizontalScrollElements();
  if (!elements) return null;
  if (!horizontalScrollInitialized) {
    const { scroller, bar } = elements;
    const syncFromScroller = () => {
      if (horizontalScrollSyncing) return;
      horizontalScrollSyncing = true;
      bar.scrollLeft = scroller.scrollLeft;
      window.requestAnimationFrame(() => {
        horizontalScrollSyncing = false;
      });
    };
    const syncFromBar = () => {
      if (horizontalScrollSyncing) return;
      horizontalScrollSyncing = true;
      scroller.scrollLeft = bar.scrollLeft;
      window.requestAnimationFrame(() => {
        horizontalScrollSyncing = false;
      });
    };
    scroller.addEventListener('scroll', syncFromScroller, { passive: true });
    bar.addEventListener('scroll', syncFromBar, { passive: true });
    horizontalScrollInitialized = true;
  }
  return elements;
}

function updateHorizontalScrollbar(table) {
  const elements = ensureHorizontalScrollBindings();
  if (!elements) return;
  const { scroller, bar, inner } = elements;
  const tableWidth = table?.scrollWidth ?? 0;
  const scrollerWidth = scroller?.scrollWidth ?? 0;
  const maxWidth = Math.max(tableWidth, scrollerWidth, bar.clientWidth);
  inner.style.width = `${Math.max(0, Math.round(maxWidth))}px`;

  const currentLeft = scroller.scrollLeft;
  horizontalScrollSyncing = true;
  if (Math.abs(bar.scrollLeft - currentLeft) > 1) {
    bar.scrollLeft = currentLeft;
  }
  window.requestAnimationFrame(() => {
    horizontalScrollSyncing = false;
  });

  if (maxWidth <= bar.clientWidth + 1) {
    bar.classList.add('is-scroll-disabled');
  } else {
    bar.classList.remove('is-scroll-disabled');
  }
}

function scheduleHorizontalScrollbarUpdate() {
  if (pendingHorizontalScrollbarUpdate) return;
  pendingHorizontalScrollbarUpdate = true;
  window.requestAnimationFrame(() => {
    pendingHorizontalScrollbarUpdate = false;
    const table = document.querySelector('#list-container table.task-list');
    updateHorizontalScrollbar(table || null);
  });
}

function numericWidth(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') {
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }
  return undefined;
}

function getColumnMinWidth(col) {
  if (!col) return 60;
  if (typeof col.minWidth === 'number' && Number.isFinite(col.minWidth)) {
    return col.minWidth;
  }
  const parsed = numericWidth(col.minWidth);
  return Number.isFinite(parsed) ? parsed : 60;
}

function getStoredColumnWidth(columnKey) {
  if (!COLUMN_WIDTHS || typeof COLUMN_WIDTHS !== 'object') return undefined;
  const value = COLUMN_WIDTHS[columnKey];
  const numeric = numericWidth(value);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : undefined;
}

function setStoredColumnWidth(columnKey, value) {
  const numeric = numericWidth(value);
  if (!Number.isFinite(numeric)) return;
  if (!COLUMN_WIDTHS || typeof COLUMN_WIDTHS !== 'object') {
    COLUMN_WIDTHS = {};
  }
  COLUMN_WIDTHS[columnKey] = Math.max(0, Math.round(numeric));
}

function loadColumnWidths() {
  try {
    const raw = window.localStorage?.getItem(COLUMN_WIDTH_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return {};
    const result = {};
    Object.entries(parsed).forEach(([key, value]) => {
      const numeric = numericWidth(value);
      if (Number.isFinite(numeric) && numeric > 0) {
        result[key] = Math.round(numeric);
      }
    });
    return result;
  } catch (err) {
    console.warn('[list] failed to load column widths', err);
    return {};
  }
}

function saveColumnWidths() {
  try {
    window.localStorage?.setItem(COLUMN_WIDTH_STORAGE_KEY, JSON.stringify(COLUMN_WIDTHS || {}));
  } catch (err) {
    console.warn('[list] failed to save column widths', err);
  }
}

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

function normalizeCategoryValue(value) {
  return String(value ?? '').trim();
}

function buildGroupedTaskList(tasks) {
  const majorGroups = [];
  const majorMap = new Map();

  tasks.forEach(task => {
    const majorRaw = normalizeCategoryValue(task?.大分類);
    const minorRaw = normalizeCategoryValue(task?.中分類);
    const majorKey = majorRaw || MAJOR_EMPTY_KEY;
    const minorKey = minorRaw || MINOR_EMPTY_KEY;
    const majorLabel = majorRaw || MAJOR_EMPTY_LABEL;
    const minorLabel = minorRaw || MINOR_EMPTY_LABEL;

    let majorGroup = majorMap.get(majorKey);
    if (!majorGroup) {
      majorGroup = {
        key: majorKey,
        label: majorLabel,
        value: majorRaw,
        minors: [],
        minorMap: new Map(),
      };
      majorMap.set(majorKey, majorGroup);
      majorGroups.push(majorGroup);
    }

    let minorGroup = majorGroup.minorMap.get(minorKey);
    if (!minorGroup) {
      minorGroup = {
        key: minorKey,
        label: minorLabel,
        value: minorRaw,
        tasks: [],
      };
      majorGroup.minorMap.set(minorKey, minorGroup);
      majorGroup.minors.push(minorGroup);
    }

    minorGroup.tasks.push(task);
  });

  majorGroups.forEach(group => {
    group.count = group.minors.reduce((sum, minor) => sum + minor.tasks.length, 0);
    delete group.minorMap;
  });

  return majorGroups;
}

function createGroupRow(type, label, count, meta = {}) {
  const row = document.createElement('tr');
  row.classList.add('group-row');
  row.classList.add(type === 'major' ? 'group-major' : 'group-minor');

  const cell = document.createElement('td');
  cell.colSpan = TABLE_COLUMN_CONFIG.length;
  cell.classList.add('group-cell');

  const title = document.createElement('span');
  title.className = 'group-label';
  title.textContent = label;
  cell.appendChild(title);

  if (typeof count === 'number' && Number.isFinite(count)) {
    const countSpan = document.createElement('span');
    countSpan.className = 'group-count';
    countSpan.textContent = `（${count}件）`;
    cell.appendChild(countSpan);
  }

  row.appendChild(cell);

  if (meta.majorValue !== undefined) {
    row.dataset.majorValue = meta.majorValue ?? '';
  }
  if (meta.minorValue !== undefined) {
    row.dataset.minorValue = meta.minorValue ?? '';
  }
  row.dataset.groupType = type;
  let suppressNextContextMenu = false;
  const triggerContextMenu = (event) => {
    showGroupContextMenu(event, {
      type,
      label,
      majorValue: meta.majorValue ?? '',
      minorValue: meta.minorValue,
    });
  };

  row.addEventListener('contextmenu', (event) => {
    if (suppressNextContextMenu) {
      suppressNextContextMenu = false;
      return;
    }
    triggerContextMenu(event);
  });

  const handleSecondaryPress = (event) => {
    if (event.button !== 2) return;
    suppressNextContextMenu = true;
    triggerContextMenu(event);
    window.setTimeout(() => {
      suppressNextContextMenu = false;
    }, 0);
  };

  if (window.PointerEvent) {
    row.addEventListener('pointerdown', handleSecondaryPress);
  } else {
    row.addEventListener('mousedown', handleSecondaryPress);
  }

  return row;
}

function ensureGroupContextMenu() {
  let menu = document.getElementById(GROUP_CONTEXT_MENU_ID);
  if (menu) return menu;

  menu = document.createElement('div');
  menu.id = GROUP_CONTEXT_MENU_ID;
  menu.className = 'group-context-menu';
  menu.setAttribute('role', 'menu');
  menu.setAttribute('aria-hidden', 'true');
  menu.style.display = 'none';

  const item = document.createElement('button');
  item.type = 'button';
  item.className = 'group-context-menu-item';
  item.textContent = 'ここの分類へタスク追加';
  item.addEventListener('click', () => {
    if (!GROUP_CONTEXT_STATE) return;
    const { majorValue, minorValue } = GROUP_CONTEXT_STATE;
    hideGroupContextMenu();
    openCreate({
      major: majorValue,
      minor: minorValue,
    });
  });

  menu.appendChild(item);
  document.body.appendChild(menu);

  return menu;
}

function hideGroupContextMenu() {
  const menu = document.getElementById(GROUP_CONTEXT_MENU_ID);
  if (menu) {
    menu.classList.remove('is-visible');
    menu.setAttribute('aria-hidden', 'true');
    menu.style.left = '';
    menu.style.top = '';
    menu.style.display = 'none';
  }
  GROUP_CONTEXT_STATE = null;
}

function positionContextMenu(menu, x, y) {
  const width = menu.offsetWidth || 0;
  const height = menu.offsetHeight || 0;
  let left = x;
  let top = y;

  const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;

  if (left + width > viewportWidth) {
    left = Math.max(0, viewportWidth - width - 8);
  }
  if (top + height > viewportHeight) {
    top = Math.max(0, viewportHeight - height - 8);
  }

  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

function showGroupContextMenu(event, context) {
  event.preventDefault();
  event.stopPropagation();
  hideGroupContextMenu();
  const menu = ensureGroupContextMenu();
  GROUP_CONTEXT_STATE = context;
  menu.classList.add('is-visible');
  menu.setAttribute('aria-hidden', 'false');

  // ensure menu has dimensions before positioning
  menu.style.left = '-9999px';
  menu.style.top = '-9999px';
  menu.style.display = 'block';
  const clientX = event.clientX ?? (event.pageX - (window.scrollX || window.pageXOffset || 0));
  const clientY = event.clientY ?? (event.pageY - (window.scrollY || window.pageYOffset || 0));
  positionContextMenu(menu, clientX, clientY);
}

document.addEventListener('click', (event) => {
  const menu = document.getElementById(GROUP_CONTEXT_MENU_ID);
  if (!menu) return;
  if (menu.contains(event.target)) return;
  hideGroupContextMenu();
});

document.addEventListener('scroll', hideGroupContextMenu, true);
window.addEventListener('blur', hideGroupContextMenu);
window.addEventListener('resize', hideGroupContextMenu);

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

function getEffectiveColumnWidth(columnKey) {
  const column = COLUMN_CONFIG_BY_KEY.get(columnKey);
  const stored = getStoredColumnWidth(columnKey);
  if (Number.isFinite(stored)) return stored;
  if (column?.width !== undefined) {
    const numeric = numericWidth(column.width);
    if (Number.isFinite(numeric)) return numeric;
  }
  return undefined;
}

function applyColumnBaseStyles(cell, columnKey) {
  const column = COLUMN_CONFIG_BY_KEY.get(columnKey);
  if (!column) return;
  cell.classList.add(`col-${columnKey}`);
  const minWidth = getColumnMinWidth(column);
  if (Number.isFinite(minWidth) && minWidth > 0) {
    cell.style.minWidth = `${minWidth}px`;
  }
  const width = getEffectiveColumnWidth(columnKey);
  if (Number.isFinite(width) && width > 0) {
    cell.style.width = `${width}px`;
  }
}

function applyColumnWidth(table, columnKey, width) {
  const column = COLUMN_CONFIG_BY_KEY.get(columnKey);
  if (!column) return;
  const minWidth = getColumnMinWidth(column);
  const normalized = Math.max(minWidth, Math.round(width));
  const px = `${normalized}px`;

  const headerCell = table.querySelector(`thead th[data-column-key="${columnKey}"]`);
  if (headerCell) {
    headerCell.style.width = px;
    headerCell.style.minWidth = `${minWidth}px`;
  }

  const colElement = table.querySelector(`col[data-column-key="${columnKey}"]`);
  if (colElement) {
    colElement.style.width = px;
  }

  table.querySelectorAll(`tbody tr`).forEach(row => {
    const cell = row.children[column.index];
    if (cell) {
      cell.style.width = px;
      cell.style.minWidth = `${minWidth}px`;
    }
  });
}

let columnResizeState = null;

function startColumnResize(ev, table, columnKey) {
  if (ev.button !== undefined && ev.button !== 0) return;
  const column = COLUMN_CONFIG_BY_KEY.get(columnKey);
  if (!column) return;
  if (typeof ev.preventDefault === 'function') ev.preventDefault();
  if (typeof ev.stopPropagation === 'function') ev.stopPropagation();

  const clientX = typeof ev.clientX === 'number' ? ev.clientX : (ev.touches?.[0]?.clientX ?? 0);
  const headerCell = table.querySelector(`thead th[data-column-key="${columnKey}"]`);
  const startWidth = headerCell ? headerCell.getBoundingClientRect().width : getEffectiveColumnWidth(columnKey) || getColumnMinWidth(column);

  columnResizeState = {
    table,
    columnKey,
    startX: clientX,
    startWidth,
    minWidth: getColumnMinWidth(column),
    lastWidth: startWidth,
  };

  const usePointerEvents = 'PointerEvent' in window;
  if (usePointerEvents) {
    document.addEventListener('pointermove', handleColumnResizeMove);
    document.addEventListener('pointerup', handleColumnResizeEnd);
    document.addEventListener('pointercancel', handleColumnResizeEnd);
  } else {
    document.addEventListener('mousemove', handleColumnResizeMove);
    document.addEventListener('mouseup', handleColumnResizeEnd);
    document.addEventListener('touchmove', handleColumnResizeMove, { passive: false });
    document.addEventListener('touchend', handleColumnResizeEnd);
    document.addEventListener('touchcancel', handleColumnResizeEnd);
  }
  document.body.classList.add('is-column-resizing');
}

function handleColumnResizeMove(ev) {
  if (!columnResizeState) return;
  if ('buttons' in ev && ev.buttons === 0) {
    handleColumnResizeEnd(ev);
    return;
  }
  if (typeof ev.preventDefault === 'function') ev.preventDefault();
  const clientX = typeof ev.clientX === 'number' ? ev.clientX : (ev.touches?.[0]?.clientX ?? columnResizeState.startX);
  const delta = clientX - columnResizeState.startX;
  const nextWidth = Math.max(columnResizeState.minWidth, columnResizeState.startWidth + delta);
  applyColumnWidth(columnResizeState.table, columnResizeState.columnKey, nextWidth);
  columnResizeState.lastWidth = nextWidth;
  scheduleHorizontalScrollbarUpdate();
}

function handleColumnResizeEnd() {
  if (!columnResizeState) return;
  if ('PointerEvent' in window) {
    document.removeEventListener('pointermove', handleColumnResizeMove);
    document.removeEventListener('pointerup', handleColumnResizeEnd);
    document.removeEventListener('pointercancel', handleColumnResizeEnd);
  } else {
    document.removeEventListener('mousemove', handleColumnResizeMove);
    document.removeEventListener('mouseup', handleColumnResizeEnd);
    document.removeEventListener('touchmove', handleColumnResizeMove);
    document.removeEventListener('touchend', handleColumnResizeEnd);
    document.removeEventListener('touchcancel', handleColumnResizeEnd);
  }
  document.body.classList.remove('is-column-resizing');

  if (Number.isFinite(columnResizeState.lastWidth)) {
    setStoredColumnWidth(columnResizeState.columnKey, columnResizeState.lastWidth);
    saveColumnWidths();
  }
  scheduleHorizontalScrollbarUpdate();
  columnResizeState = null;
}

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
  if (!WIRED) {
    const wired = wireToolbar();
    if (wired) {
      WIRED = true;
    } else {
      ready(() => {
        if (WIRED) return;
        if (wireToolbar()) {
          WIRED = true;
        }
      });
    }
  }
}

/* ===================== レンダリング ===================== */
function renderList() {
  const container = document.getElementById('list-container');
  if (!container) return;

  const filtered = getFilteredTasks();

  buildAssigneeWorkload(filtered);

  container.innerHTML = '';

  if (filtered.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'empty-message';
    empty.textContent = '該当するタスクはありません。';
    container.appendChild(empty);
    updateHorizontalScrollbar(null);
    updateDueIndicators(filtered);
    return;
  }

  const statusOrder = new Map();
  STATUSES.forEach((s, idx) => { statusOrder.set(s, idx); });

  const table = document.createElement('table');
  table.className = 'task-list';

  const colgroup = document.createElement('colgroup');
  TABLE_COLUMN_CONFIG.forEach(col => {
    const colEl = document.createElement('col');
    colEl.dataset.columnKey = col.key;
    const width = getEffectiveColumnWidth(col.key);
    if (Number.isFinite(width) && width > 0) {
      colEl.style.width = `${width}px`;
    }
    colgroup.appendChild(colEl);
  });
  table.appendChild(colgroup);

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  TABLE_COLUMN_CONFIG.forEach(col => {
    const th = document.createElement('th');
    th.dataset.columnKey = col.key;
    applyColumnBaseStyles(th, col.key);
    if (col.className) th.classList.add(col.className);
    th.textContent = col.label;
    if (col.sortable) {
      th.classList.add('sortable');
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
    th.classList.add('resizable');
    const resizer = document.createElement('div');
    resizer.className = 'column-resizer';
    resizer.title = '列幅を変更';
    if ('PointerEvent' in window) {
      resizer.addEventListener('pointerdown', (ev) => startColumnResize(ev, table, col.key));
    } else {
      resizer.addEventListener('mousedown', (ev) => startColumnResize(ev, table, col.key));
      resizer.addEventListener('touchstart', (ev) => startColumnResize(ev, table, col.key), { passive: false });
    }
    resizer.addEventListener('click', (ev) => ev.stopPropagation());
    th.appendChild(resizer);
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

  const createTaskRow = (task) => {
    const tr = document.createElement('tr');
    tr.dataset.no = String(task.No || '');
    tr.addEventListener('dblclick', () => openEdit(task.No));

    const noTd = document.createElement('td');
    applyColumnBaseStyles(noTd, 'no');
    noTd.textContent = task.No ? `#${task.No}` : '';
    tr.appendChild(noTd);

    const majorTd = document.createElement('td');
    applyColumnBaseStyles(majorTd, 'major');
    if ((task.大分類 || '').trim()) {
      const badge = document.createElement('span');
      badge.className = 'category-pill category-major';
      badge.textContent = task.大分類.trim();
      majorTd.appendChild(badge);
    }
    tr.appendChild(majorTd);

    const minorTd = document.createElement('td');
    applyColumnBaseStyles(minorTd, 'minor');
    if ((task.中分類 || '').trim()) {
      const badge = document.createElement('span');
      badge.className = 'category-pill category-minor';
      badge.textContent = task.中分類.trim();
      minorTd.appendChild(badge);
    }
    tr.appendChild(minorTd);

    const titleTd = document.createElement('td');
    applyColumnBaseStyles(titleTd, 'task');
    titleTd.classList.add('task-cell', 'col-task');
    titleTd.textContent = task.タスク || '(無題)';
    tr.appendChild(titleTd);

    const statusTd = document.createElement('td');
    applyColumnBaseStyles(statusTd, 'status');
    const statusPill = document.createElement('span');
    statusPill.className = 'status-pill';
    statusPill.textContent = normalizeStatusLabel(task.ステータス);
    statusTd.appendChild(statusPill);
    tr.appendChild(statusTd);

    const assigneeTd = document.createElement('td');
    applyColumnBaseStyles(assigneeTd, 'assignee');
    assigneeTd.textContent = (task.担当者 || '').trim();
    tr.appendChild(assigneeTd);

    const priorityTd = document.createElement('td');
    applyColumnBaseStyles(priorityTd, 'priority');
    if (task.優先度 !== undefined && task.優先度 !== null && String(task.優先度).trim() !== '') {
      const pill = document.createElement('span');
      pill.className = 'priority-pill';
      pill.textContent = `優先度: ${task.優先度}`;
      priorityTd.appendChild(pill);
    }
    tr.appendChild(priorityTd);

    const dueTd = document.createElement('td');
    applyColumnBaseStyles(dueTd, 'due');
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
    applyColumnBaseStyles(notesTd, 'notes');
    notesTd.classList.add('notes-cell', 'col-notes');
    notesTd.textContent = task.備考 || '';
    tr.appendChild(notesTd);

    return tr;
  };

  const groupedTasks = buildGroupedTaskList(sortedTasks);
  groupedTasks.forEach(majorGroup => {
    tbody.appendChild(createGroupRow('major', majorGroup.label, majorGroup.count, {
      majorValue: majorGroup.value,
    }));
    majorGroup.minors.forEach(minorGroup => {
      tbody.appendChild(createGroupRow('minor', minorGroup.label, minorGroup.tasks.length, {
        majorValue: majorGroup.value,
        minorValue: minorGroup.value,
      }));
      minorGroup.tasks.forEach(task => {
        tbody.appendChild(createTaskRow(task));
      });
    });
  });

  table.appendChild(tbody);
  container.appendChild(table);

  updateHorizontalScrollbar(table);
  updateDueIndicators(filtered);
}


function buildAssigneeWorkload(tasks) {
  const list = Array.isArray(tasks) ? tasks : [];
  workloadSummary.update(list, {
    total: list.length,
    overall: Array.isArray(TASKS) ? TASKS.length : list.length,
  });
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

  applyFilterCollapsedState();
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
  const addBtn = document.getElementById('btn-add');
  const saveBtn = document.getElementById('btn-save');
  const validationsBtn = document.getElementById('btn-validations');
  const reloadBtn = document.getElementById('btn-reload');

  if (!addBtn || !saveBtn || !validationsBtn || !reloadBtn) {
    return false;
  }

  addBtn.addEventListener('click', () => openCreate());
  saveBtn.addEventListener('click', async () => {
    try {
      const p = await api.save_excel();
      alert('Excelへ保存しました\n' + p);
    } catch (e) {
      alert('保存に失敗: ' + (e?.message || e));
    }
  });
  validationsBtn.addEventListener('click', () => openValidationModal());
  reloadBtn.addEventListener('click', async () => {
    try {
      const payload = await api.reload_from_excel();
      await applyStateFromPayload(payload, { preserveFilters: true, fallbackToApi: true });
    } catch (e) {
      alert('再読込に失敗: ' + (e?.message || e));
    }
  });
  return true;
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
function openCreate(defaults = {}) {
  hideGroupContextMenu();
  CURRENT_EDIT = null;
  openModal({
    No: '',
    ステータス: STATUSES[0] || '未着手',
    大分類: defaults.major !== undefined ? (defaults.major || '') : '',
    中分類: defaults.minor !== undefined ? (defaults.minor || '') : '',
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

window.addEventListener('resize', scheduleHorizontalScrollbarUpdate);

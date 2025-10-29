/* eslint-disable no-console */
(function (global) {
  'use strict';

  const PRIORITY_DEFAULT_OPTIONS = ['高', '中', '低'];
  const DEFAULT_STATUSES = ['未着手', '進行中', '完了', '保留'];
  const UNSET_STATUS_LABEL = 'ステータス未設定';
  const FILTER_PRESET_STORAGE_KEY = 'kanban:filterPresets';

  function clonePlain(value) {
    if (value === null || value === undefined) return value;
    if (Array.isArray(value)) {
      return value.map(item => clonePlain(item));
    }
    if (value instanceof Set) {
      return Array.from(value).map(item => clonePlain(item));
    }
    if (value instanceof Date) {
      const time = value.getTime();
      return Number.isNaN(time) ? null : value.toISOString();
    }
    if (typeof value === 'object') {
      const result = {};
      Object.keys(value).forEach(key => {
        const entry = value[key];
        if (typeof entry === 'function' || entry === undefined) return;
        result[key] = clonePlain(entry);
      });
      return result;
    }
    return value;
  }

  function sanitizePresetEntry(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const name = String(raw.name ?? '').trim();
    if (!name) return null;
    const filters = clonePlain(raw.filters ?? {});
    const updatedAt = Number(raw.updatedAt);
    const lastAppliedAt = Number(raw.lastAppliedAt);
    const preset = { name, filters };
    if (Number.isFinite(updatedAt) && updatedAt > 0) {
      preset.updatedAt = updatedAt;
    } else {
      preset.updatedAt = Date.now();
    }
    if (Number.isFinite(lastAppliedAt) && lastAppliedAt > 0) {
      preset.lastAppliedAt = lastAppliedAt;
    }
    return preset;
  }

  function readFilterPresetStore() {
    let parsed = {};
    let dirty = false;
    try {
      const stored = global.localStorage?.getItem(FILTER_PRESET_STORAGE_KEY) ?? '';
      if (stored) {
        try {
          const json = JSON.parse(stored);
          if (json && typeof json === 'object') {
            parsed = json;
          } else {
            dirty = true;
          }
        } catch (err) {
          console.warn('[kanban] failed to parse filter presets storage', err);
          dirty = true;
          parsed = {};
        }
      }
    } catch (err) {
      console.warn('[kanban] failed to read filter presets storage', err);
      parsed = {};
    }

    const map = {};
    Object.keys(parsed).forEach(key => {
      const list = Array.isArray(parsed[key]) ? parsed[key] : [];
      if (!Array.isArray(parsed[key])) dirty = true;
      const sanitized = [];
      list.forEach(item => {
        const preset = sanitizePresetEntry(item);
        if (preset) {
          sanitized.push(preset);
        } else {
          dirty = true;
        }
      });
      map[key] = sanitized;
      if (sanitized.length !== list.length) dirty = true;
    });

    return { map, dirty };
  }

  function writeFilterPresetStore(store) {
    try {
      global.localStorage?.setItem(FILTER_PRESET_STORAGE_KEY, JSON.stringify(store));
    } catch (err) {
      console.warn('[kanban] failed to store filter presets', err);
    }
  }

  function clonePresetEntry(preset) {
    if (!preset || typeof preset !== 'object') return null;
    const cloned = {
      name: preset.name,
      filters: clonePlain(preset.filters ?? {}),
    };
    if (Number.isFinite(preset.updatedAt)) {
      cloned.updatedAt = preset.updatedAt;
    }
    if (Number.isFinite(preset.lastAppliedAt)) {
      cloned.lastAppliedAt = preset.lastAppliedAt;
    }
    return cloned;
  }

  function clonePresetList(list) {
    if (!Array.isArray(list)) return [];
    return list.map(item => clonePresetEntry(item)).filter(Boolean);
  }

  function loadFilterPresets(viewKey) {
    const key = String(viewKey ?? '').trim();
    if (!key) {
      return { presets: [], lastApplied: null };
    }
    const { map, dirty } = readFilterPresetStore();
    if (dirty) {
      writeFilterPresetStore(map);
    }
    const list = Array.isArray(map[key]) ? map[key] : [];
    const presets = clonePresetList(list);
    let lastApplied = null;
    presets.forEach(preset => {
      if (!preset) return;
      const ts = Number(preset.lastAppliedAt);
      if (!Number.isFinite(ts)) return;
      if (!lastApplied || ts > Number(lastApplied.lastAppliedAt || 0)) {
        lastApplied = preset;
      }
    });
    return { presets, lastApplied };
  }

  function saveFilterPreset(viewKey, presetName, filters, options = {}) {
    const key = String(viewKey ?? '').trim();
    const name = String(presetName ?? '').trim();
    if (!key || !name) {
      return { presets: [], saved: null };
    }
    const { map } = readFilterPresetStore();
    const list = Array.isArray(map[key]) ? map[key] : [];
    const now = Date.now();
    const normalizedFilters = clonePlain(filters ?? {});
    const idx = list.findIndex(item => item?.name === name);
    let entry;
    if (idx >= 0) {
      entry = { ...list[idx], name, filters: normalizedFilters, updatedAt: now };
      if (options.markAsApplied !== false) {
        entry.lastAppliedAt = now;
      }
      list[idx] = entry;
    } else {
      entry = { name, filters: normalizedFilters, updatedAt: now };
      if (options.markAsApplied !== false) {
        entry.lastAppliedAt = now;
      }
      list.push(entry);
    }
    map[key] = list;
    writeFilterPresetStore(map);
    return { presets: clonePresetList(list), saved: clonePresetEntry(entry) };
  }

  function deleteFilterPreset(viewKey, presetName) {
    const key = String(viewKey ?? '').trim();
    const name = String(presetName ?? '').trim();
    if (!key || !name) {
      return { presets: [], removed: false };
    }
    const { map, dirty } = readFilterPresetStore();
    const list = Array.isArray(map[key]) ? map[key] : [];
    const idx = list.findIndex(item => item?.name === name);
    if (idx < 0) {
      if (dirty) {
        writeFilterPresetStore(map);
      }
      return { presets: clonePresetList(list), removed: false };
    }
    list.splice(idx, 1);
    map[key] = list;
    writeFilterPresetStore(map);
    return { presets: clonePresetList(list), removed: true };
  }

  function applyFilterPreset(viewKey, presetName, applyFn) {
    const key = String(viewKey ?? '').trim();
    const name = String(presetName ?? '').trim();
    if (!key || !name || typeof applyFn !== 'function') {
      return { presets: [], applied: null };
    }
    const { map, dirty } = readFilterPresetStore();
    const list = Array.isArray(map[key]) ? map[key] : [];
    const idx = list.findIndex(item => item?.name === name);
    if (idx < 0) {
      if (dirty) {
        writeFilterPresetStore(map);
      }
      return { presets: clonePresetList(list), applied: null };
    }
    const target = list[idx];
    const payload = clonePlain(target.filters ?? {});
    const result = applyFn(payload, clonePresetEntry(target));
    if (result === false) {
      if (dirty) {
        writeFilterPresetStore(map);
      }
      return { presets: clonePresetList(list), applied: null };
    }
    target.lastAppliedAt = Date.now();
    map[key] = list;
    writeFilterPresetStore(map);
    return {
      presets: clonePresetList(list),
      applied: clonePresetEntry(target),
    };
  }

  function ready(fn) {
    if (document.readyState !== 'loading') {
      try {
        fn();
      } catch (err) {
        console.error('[kanban] ready callback failed', err);
      }
      return;
    }
    document.addEventListener('DOMContentLoaded', () => {
      try {
        fn();
      } catch (err) {
        console.error('[kanban] ready callback failed', err);
      }
    }, { once: true });
  }

  function createMockApi() {
    const baseStatuses = ['未着手', '進行中', '完了', '保留'];
    const statusSet = new Set(baseStatuses);
    const pad = (n) => String(n).padStart(2, '0');
    const toISO = (date) => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
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

    const cloneTask = (task) => ({ ...task });

    const sanitizeStatus = (status) => {
      const text = String(status ?? '').trim();
      if (text) return text;
      return baseStatuses[0];
    };

    const toIsoDate = (value) => {
      if (!value) return '';
      const parsed = new Date(value);
      if (Number.isNaN(parsed.getTime())) return '';
      return toISO(parsed);
    };

    const normalizePriority = (value) => {
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
        期限: toIsoDate(payload?.期限),
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
      },
      async get_state_snapshot() {
        return {
          tasks: withSequentialNo(),
          statuses: Array.from(statusSet),
          validations: { ...validations }
        };
      }
    };
  }

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
    rawList.forEach(v => {
      const text = String(v ?? '').trim();
      if (!text || seen.has(text)) return;
      seen.add(text);
      values.push(text);
    });
    return values;
  }

  function createPriorityHelper({ getValidations, defaultOptions = PRIORITY_DEFAULT_OPTIONS } = {}) {
    if (typeof getValidations !== 'function') {
      throw new Error('createPriorityHelper requires getValidations function');
    }

    const getOptions = () => {
      const source = getValidations() || {};
      const base = Array.isArray(source['優先度']) && source['優先度'].length > 0
        ? source['優先度']
        : defaultOptions;
      const seen = new Set();
      const options = [];
      base.forEach((value) => {
        const text = String(value ?? '').trim();
        if (!text || seen.has(text)) return;
        seen.add(text);
        options.push(text);
      });
      if (options.length === 0) {
        defaultOptions.forEach((value) => {
          if (!seen.has(value)) {
            seen.add(value);
            options.push(value);
          }
        });
      }
      return options;
    };

    const getDefaultValue = () => {
      const options = getOptions();
      if (options.includes('中')) return '中';
      return options[0] || '';
    };

    const applyOptions = (selectEl, currentValue, preferDefault = false) => {
      if (!selectEl) return;
      const normalized = currentValue === null || currentValue === undefined
        ? ''
        : String(currentValue).trim();
      const options = getOptions();
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
          selection = getDefaultValue();
        } else if (values.includes('')) {
          selection = '';
        } else {
          selection = getDefaultValue();
        }
      }
      if (!values.includes(selection)) {
        selection = values[0] || '';
      }
      selectEl.value = selection;
    };

    return {
      getOptions,
      getDefaultValue,
      applyOptions,
    };
  }

  function getPriorityLevel(value) {
    const label = String(value ?? '').trim();
    if (!label) return 'unset';
    if (label === '低') return 'low';
    if (label === '中') return 'medium';
    if (label === '高') return 'high';
    return 'custom';
  }

  function setupRuntime({ onInit, onRealtimeUpdate, onApiChanged, mockApiFactory } = {}) {
    const state = {
      api: null,
      runMode: 'mock',
    };

    const getMockApi = () => (typeof mockApiFactory === 'function' ? mockApiFactory() : createMockApi());

    const assignApi = (nextApi, runMode) => {
      if (!nextApi) return;
      state.api = nextApi;
      state.runMode = runMode;
      if (typeof onApiChanged === 'function') {
        try {
          onApiChanged({ api: state.api, runMode: state.runMode });
        } catch (err) {
          console.error('[kanban] onApiChanged callback failed', err);
        }
      }
      if (typeof onInit === 'function') {
        Promise.resolve(onInit({ api: state.api, runMode: state.runMode, force: true }))
          .catch(err => {
            console.error('[kanban] initialization failed', err);
          });
      }
    };

    global.addEventListener('pywebviewready', () => {
      const pyApi = global.pywebview?.api;
      if (pyApi) {
        assignApi(pyApi, 'pywebview');
      }
    });

    ready(() => {
      const pyApi = global.pywebview?.api;
      if (pyApi) {
        assignApi(pyApi, 'pywebview');
      } else {
        assignApi(getMockApi(), 'mock');
      }
    });

    global.__kanban_receive_update = (payload) => {
      if (typeof onRealtimeUpdate !== 'function') return;
      Promise.resolve(onRealtimeUpdate(payload)).catch(err => {
        console.error('[kanban] failed to apply pushed payload', err);
      });
    };

    return {
      get api() {
        return state.api;
      },
      get runMode() {
        return state.runMode;
      }
    };
  }

  function setupDragViewportAutoScroll({ margin = 80, maxStep = 24 } = {}) {
    if (typeof document === 'undefined' || typeof window === 'undefined') {
      return () => {};
    }

    let dragDepth = 0;

    const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

    const handleDragStart = () => {
      dragDepth += 1;
    };

    const reset = () => {
      dragDepth = Math.max(0, dragDepth - 1);
    };

    const handleDragOver = (event) => {
      if (dragDepth <= 0) return;
      if (typeof event?.clientY !== 'number') return;

      const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
      if (!viewportHeight) return;

      let delta = 0;
      if (event.clientY < margin) {
        const intensity = (margin - event.clientY) / margin;
        delta = -Math.ceil(clamp(intensity, 0, 1) * maxStep);
      } else if (event.clientY > viewportHeight - margin) {
        const intensity = (event.clientY - (viewportHeight - margin)) / margin;
        delta = Math.ceil(clamp(intensity, 0, 1) * maxStep);
      }

      if (delta !== 0) {
        window.scrollBy({ top: delta, behavior: 'auto' });
      }
    };

    document.addEventListener('dragstart', handleDragStart, { passive: true });
    document.addEventListener('dragend', reset, { passive: true });
    document.addEventListener('drop', reset, { passive: true });
    document.addEventListener('dragover', handleDragOver, { passive: true });

    return () => {
      document.removeEventListener('dragstart', handleDragStart);
      document.removeEventListener('dragend', reset);
      document.removeEventListener('drop', reset);
      document.removeEventListener('dragover', handleDragOver);
      dragDepth = 0;
    };
  }

  function parseISODate(value) {
    if (!value) return null;
    if (value instanceof Date) {
      const time = value.getTime();
      return Number.isNaN(time) ? null : new Date(time);
    }
    const text = String(value).trim();
    if (!text) return null;
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
    if (!match) return null;
    const dt = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  function formatDateInputValue(date) {
    if (!date) return '';
    const dt = date instanceof Date ? new Date(date.getTime()) : parseISODate(date);
    if (!dt || Number.isNaN(dt.getTime())) return '';
    const year = dt.getFullYear();
    const month = String(dt.getMonth() + 1).padStart(2, '0');
    const day = String(dt.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function startOfWeek(date) {
    const base = date instanceof Date ? new Date(date.getTime()) : parseISODate(date);
    if (!base || Number.isNaN(base.getTime())) return null;
    base.setHours(0, 0, 0, 0);
    base.setDate(base.getDate() - base.getDay());
    return base;
  }

  function endOfWeek(date) {
    const start = startOfWeek(date);
    if (!start) return null;
    const end = new Date(start.getTime());
    end.setDate(start.getDate() + 6);
    return end;
  }

  function getDueFilterPreset(presetName, options = {}) {
    const name = String(presetName ?? '').trim();
    const todayBase = options.today instanceof Date
      ? new Date(options.today.getTime())
      : new Date();
    todayBase.setHours(0, 0, 0, 0);

    const result = { mode: 'none', from: '', to: '' };

    if (name === 'this-week') {
      const weekEnd = endOfWeek(todayBase);
      if (!weekEnd) return null;
      result.mode = 'before';
      result.from = formatDateInputValue(weekEnd);
      return result;
    }

    if (name === 'next-week') {
      const weekStart = startOfWeek(todayBase);
      if (!weekStart) return null;
      const nextWeekStart = new Date(weekStart.getTime());
      nextWeekStart.setDate(weekStart.getDate() + 7);
      const nextWeekEnd = new Date(nextWeekStart.getTime());
      nextWeekEnd.setDate(nextWeekStart.getDate() + 6);
      result.mode = 'before';
      result.from = formatDateInputValue(nextWeekEnd);
      return result;
    }

    return null;
  }

  function isCompletedStatus(value) {
    const text = String(value ?? '').trim();
    if (!text) return false;
    const normalized = text.toLowerCase().replace(/\s+/g, '');
    return normalized === '完了'
      || normalized === '完了済み'
      || normalized === '完了済'
      || normalized === 'done'
      || normalized === 'completed';
  }

  function getDueState(task) {
    if (!task || typeof task !== 'object') return null;
    if (isCompletedStatus(task.ステータス)) return null;

    const dueDate = parseISODate(task.期限 || '');
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
        label,
      };
    }

    return {
      level: 'normal',
      diff: diffDays,
      label,
    };
  }

  function summarizeAssigneeWorkload(tasks, {
    normalizeStatusLabel: normalizeStatus = (value) => {
      const text = String(value ?? '').trim();
      return text || UNSET_STATUS_LABEL;
    },
    unassignedKey = '__UNASSIGNED__',
    unassignedLabel = '（未割り当て）',
    getDueState: dueEvaluator = null,
  } = {}) {
    const list = Array.isArray(tasks) ? tasks : [];
    const map = new Map();
    const assignees = [];
    const statusSet = new Set();

    list.forEach(task => {
      if (!task || typeof task !== 'object') return;
      const rawAssignee = String(task.担当者 ?? '').trim();
      const key = rawAssignee || unassignedKey;
      const label = rawAssignee || unassignedLabel;

      let entry = map.get(key);
      if (!entry) {
        entry = {
          key,
          label,
          total: 0,
          statusCounts: Object.create(null),
          due: { warning: 0, overdue: 0 },
        };
        map.set(key, entry);
        assignees.push(entry);
      }

      entry.total += 1;

      const status = normalizeStatus(task.ステータス);
      const statusKey = String(status ?? '').trim() || UNSET_STATUS_LABEL;
      statusSet.add(statusKey);
      entry.statusCounts[statusKey] = (entry.statusCounts[statusKey] || 0) + 1;

      if (typeof dueEvaluator === 'function') {
        const state = dueEvaluator(task);
        if (state?.level === 'overdue') {
          entry.due.overdue += 1;
        } else if (state?.level === 'warning') {
          entry.due.warning += 1;
        }
      }
    });

    assignees.sort((a, b) => {
      if (b.total !== a.total) return b.total - a.total;
      return a.label.localeCompare(b.label, 'ja');
    });

    return {
      assignees,
      statuses: Array.from(statusSet),
    };
  }

  function createWorkloadSummary({
    container,
    bodySelector = '.workload-summary-body',
    metaSelector = '.workload-summary-meta',
    toggleSelector = '.workload-toggle',
    normalizeStatusLabel: normalizeStatus = (value) => {
      const text = String(value ?? '').trim();
      return text || UNSET_STATUS_LABEL;
    },
    unassignedKey = '__UNASSIGNED__',
    unassignedLabel = '（未割り当て）',
    allKey = '__ALL__',
    getStatuses = () => [],
    getActiveAssignee = () => null,
    getDueState: dueEvaluator = null,
    highlightPredicate = (entry) => (entry?.due?.overdue ?? 0) > 0,
    onSelectAssignee = null,
  } = {}) {
    if (!container) {
      return { update() {} };
    }

    const body = container.querySelector(bodySelector) || container;
    const metaEl = container.querySelector(metaSelector) || null;
    const toggleEl = container.querySelector(toggleSelector) || null;

    let metaTextNode = null;
    let resetButton = null;
    if (metaEl) {
      metaEl.innerHTML = '';
      metaTextNode = document.createElement('span');
      metaTextNode.className = 'workload-summary-text';
      metaEl.appendChild(metaTextNode);
      if (typeof onSelectAssignee === 'function') {
        resetButton = document.createElement('button');
        resetButton.type = 'button';
        resetButton.className = 'assignee-reset';
        resetButton.dataset.assigneeFilter = allKey;
        resetButton.textContent = '全員を表示';
        metaEl.appendChild(resetButton);
      }
    }

    const USER_OVERRIDE_ATTR = 'data-workload-user-toggle';

    const setCollapsed = (collapsed) => {
      if (collapsed) {
        container.classList.add('is-collapsed');
      } else {
        container.classList.remove('is-collapsed');
      }
      if (toggleEl) {
        toggleEl.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
      }
    };

    if (toggleEl) {
      toggleEl.addEventListener('click', () => {
        const collapsed = !container.classList.contains('is-collapsed');
        setCollapsed(collapsed);
        container.setAttribute(USER_OVERRIDE_ATTR, '1');
      });
    }

    const mediaQuery = typeof window.matchMedia === 'function'
      ? window.matchMedia('(max-width: 900px)')
      : null;

    const applyMediaState = () => {
      if (!mediaQuery) return;
      if (mediaQuery.matches) {
        container.classList.add('is-collapsible');
        if (!container.hasAttribute(USER_OVERRIDE_ATTR)) {
          setCollapsed(true);
        }
      } else {
        container.classList.remove('is-collapsible');
        container.removeAttribute(USER_OVERRIDE_ATTR);
        setCollapsed(false);
      }
    };

    if (mediaQuery) {
      if (typeof mediaQuery.addEventListener === 'function') {
        mediaQuery.addEventListener('change', applyMediaState);
      } else if (typeof mediaQuery.addListener === 'function') {
        mediaQuery.addListener(applyMediaState);
      }
      applyMediaState();
    }

    if (typeof onSelectAssignee === 'function') {
      container.addEventListener('click', (ev) => {
        const target = ev.target.closest('[data-assignee-filter]');
        if (!target || !container.contains(target)) return;
        ev.preventDefault();
        const value = target.dataset.assigneeFilter;
        onSelectAssignee(value);
      });
    }

    const highlightFn = typeof highlightPredicate === 'function'
      ? highlightPredicate
      : () => false;

    const update = (tasks, meta = {}) => {
      const list = Array.isArray(tasks) ? tasks : [];
      const activeAssignee = typeof getActiveAssignee === 'function' ? getActiveAssignee() : null;

      if (metaTextNode) {
        const total = Number.isFinite(meta.total) ? meta.total : list.length;
        const overall = Number.isFinite(meta.overall) ? meta.overall : undefined;
        const pieces = [`表示: ${total}件`];
        if (Number.isFinite(overall) && overall !== total) {
          pieces.push(`全体: ${overall}件`);
        }
        metaTextNode.textContent = pieces.join(' / ');
      }

      if (resetButton) {
        const disabled = !activeAssignee || activeAssignee === allKey;
        resetButton.classList.toggle('is-disabled', disabled);
        resetButton.disabled = disabled;
      }

      if (list.length === 0) {
        body.innerHTML = '';
        const empty = document.createElement('div');
        empty.className = 'workload-empty';
        empty.textContent = '担当者別サマリーを表示するカードがありません。';
        body.appendChild(empty);
        return;
      }

      const summary = summarizeAssigneeWorkload(list, {
        normalizeStatusLabel: normalizeStatus,
        unassignedKey,
        unassignedLabel,
        getDueState: dueEvaluator,
      });

      const statusOrder = [];
      const seenStatuses = new Set();
      const baseStatuses = typeof getStatuses === 'function' ? getStatuses() : [];
      if (Array.isArray(baseStatuses)) {
        baseStatuses.forEach(status => {
          const text = String(status ?? '').trim();
          if (!text || seenStatuses.has(text)) return;
          seenStatuses.add(text);
          statusOrder.push(text);
        });
      }
      summary.statuses.forEach(status => {
        const text = String(status ?? '').trim();
        if (!text || seenStatuses.has(text)) return;
        seenStatuses.add(text);
        statusOrder.push(text);
      });

      const fragment = document.createDocumentFragment();
      summary.assignees.forEach(entry => {
        const article = document.createElement('article');
        article.className = 'workload-entry';
        if ((entry.due?.overdue ?? 0) > 0) {
          article.classList.add('has-overdue');
        } else if ((entry.due?.warning ?? 0) > 0) {
          article.classList.add('has-warning');
        }
        if (highlightFn(entry)) {
          article.classList.add('is-heavy');
        }
        if (activeAssignee && activeAssignee === entry.key) {
          article.classList.add('is-active');
        }

        const header = document.createElement('div');
        header.className = 'workload-entry-header';

        const nameBtn = document.createElement('button');
        nameBtn.type = 'button';
        nameBtn.className = 'assignee-button';
        nameBtn.dataset.assigneeFilter = entry.key;
        nameBtn.textContent = entry.label;
        if (activeAssignee && activeAssignee === entry.key) {
          nameBtn.classList.add('is-active');
        }
        header.appendChild(nameBtn);

        const total = document.createElement('span');
        total.className = 'workload-total';
        total.textContent = `${entry.total}件`;
        header.appendChild(total);
        article.appendChild(header);

        const statusesWrap = document.createElement('div');
        statusesWrap.className = 'workload-statuses';
        let statusCount = 0;
        statusOrder.forEach(status => {
          const count = entry.statusCounts?.[status] || 0;
          if (count <= 0) return;
          statusCount += 1;
          const chip = document.createElement('span');
          chip.className = 'status-chip';
          chip.dataset.status = status;
          chip.textContent = `${status} `;
          const countSpan = document.createElement('span');
          countSpan.className = 'status-count';
          countSpan.textContent = `${count}`;
          chip.appendChild(countSpan);
          statusesWrap.appendChild(chip);
        });
        if (statusCount === 0) {
          const chip = document.createElement('span');
          chip.className = 'status-chip';
          chip.textContent = 'ステータスなし';
          statusesWrap.appendChild(chip);
        }
        article.appendChild(statusesWrap);

        const dueWrap = document.createElement('div');
        dueWrap.className = 'workload-due';
        const overdueCount = entry.due?.overdue || 0;
        const warningCount = entry.due?.warning || 0;
        if (overdueCount > 0) {
          const span = document.createElement('span');
          span.className = 'due-overdue';
          span.textContent = `期限超過 ${overdueCount}件`;
          dueWrap.appendChild(span);
        }
        if (warningCount > 0) {
          const span = document.createElement('span');
          span.className = 'due-warning';
          span.textContent = `期限警告 ${warningCount}件`;
          dueWrap.appendChild(span);
        }
        if (!dueWrap.children.length) {
          const span = document.createElement('span');
          span.textContent = '期限警告なし';
          dueWrap.appendChild(span);
        }
        article.appendChild(dueWrap);

        fragment.appendChild(article);
      });

      body.innerHTML = '';
      body.appendChild(fragment);
    };

    return { update };
  }

  global.TaskAppCommon = {
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
    parseISO: parseISODate,
    isCompletedStatus,
    getDueState,
    summarizeAssigneeWorkload,
    createWorkloadSummary,
    loadFilterPresets,
    saveFilterPreset,
    deleteFilterPreset,
    applyFilterPreset,
    PRIORITY_DEFAULT_OPTIONS,
    DEFAULT_STATUSES,
    UNSET_STATUS_LABEL,
    getPriorityLevel,
    getDueFilterPreset,
  };
}(window));

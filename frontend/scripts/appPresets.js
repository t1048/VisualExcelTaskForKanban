(function (global) {
  'use strict';

  const STORAGE_KEY = 'kanban:filterPresets';

  function clonePlain(value) {
    if (value === null || value === undefined) return value;
    if (Array.isArray(value)) {
      return value.map(item => clonePlain(item));
    }
    if (value instanceof Date) {
      const time = value.getTime();
      return Number.isNaN(time) ? null : value.toISOString();
    }
    if (value instanceof Set) {
      return Array.from(value).map(item => clonePlain(item));
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
    const preset = {
      name,
      filters: clonePlain(raw.filters ?? {}),
      updatedAt: Date.now(),
    };
    const updatedAt = Number(raw.updatedAt);
    if (Number.isFinite(updatedAt) && updatedAt > 0) {
      preset.updatedAt = updatedAt;
    }
    const lastAppliedAt = Number(raw.lastAppliedAt);
    if (Number.isFinite(lastAppliedAt) && lastAppliedAt > 0) {
      preset.lastAppliedAt = lastAppliedAt;
    }
    return preset;
  }

  function clonePresetEntry(entry) {
    if (!entry || typeof entry !== 'object') return null;
    const cloned = {
      name: entry.name,
      filters: clonePlain(entry.filters ?? {}),
    };
    if (Number.isFinite(entry.updatedAt)) {
      cloned.updatedAt = entry.updatedAt;
    }
    if (Number.isFinite(entry.lastAppliedAt)) {
      cloned.lastAppliedAt = entry.lastAppliedAt;
    }
    return cloned;
  }

  function clonePresetList(list) {
    if (!Array.isArray(list)) return [];
    return list.map(item => clonePresetEntry(item)).filter(Boolean);
  }

  function readStore() {
    let parsed = {};
    let dirty = false;
    try {
      const stored = global.localStorage?.getItem(STORAGE_KEY) ?? '';
      if (stored) {
        const json = JSON.parse(stored);
        if (json && typeof json === 'object') {
          parsed = json;
        } else {
          dirty = true;
          parsed = {};
        }
      }
    } catch (err) {
      console.warn('[TaskPresets] failed to read storage', err);
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

  function writeStore(store) {
    try {
      global.localStorage?.setItem(STORAGE_KEY, JSON.stringify(store));
    } catch (err) {
      console.warn('[TaskPresets] failed to persist storage', err);
    }
  }

  function load(viewKey) {
    const key = String(viewKey ?? '').trim();
    if (!key) {
      return { presets: [], lastApplied: null };
    }
    const { map, dirty } = readStore();
    if (dirty) {
      writeStore(map);
    }
    const list = Array.isArray(map[key]) ? map[key] : [];
    const presets = clonePresetList(list);
    let lastApplied = null;
    presets.forEach(preset => {
      const ts = Number(preset?.lastAppliedAt);
      if (!Number.isFinite(ts)) return;
      if (!lastApplied || ts > Number(lastApplied.lastAppliedAt || 0)) {
        lastApplied = preset;
      }
    });
    return { presets, lastApplied };
  }

  function save(viewKey, name, filters, options = {}) {
    const key = String(viewKey ?? '').trim();
    const presetName = String(name ?? '').trim();
    if (!key || !presetName) {
      return { presets: [], saved: null };
    }
    const { map } = readStore();
    const list = Array.isArray(map[key]) ? map[key] : [];
    const now = Date.now();
    const normalizedFilters = clonePlain(filters ?? {});
    const idx = list.findIndex(item => item?.name === presetName);
    let entry;
    if (idx >= 0) {
      entry = { ...list[idx], name: presetName, filters: normalizedFilters, updatedAt: now };
      if (options.markAsApplied !== false) {
        entry.lastAppliedAt = now;
      }
      list[idx] = entry;
    } else {
      entry = { name: presetName, filters: normalizedFilters, updatedAt: now };
      if (options.markAsApplied !== false) {
        entry.lastAppliedAt = now;
      }
      list.push(entry);
    }
    map[key] = list;
    writeStore(map);
    return { presets: clonePresetList(list), saved: clonePresetEntry(entry) };
  }

  function remove(viewKey, name) {
    const key = String(viewKey ?? '').trim();
    const presetName = String(name ?? '').trim();
    if (!key || !presetName) {
      return { presets: [], removed: false };
    }
    const { map, dirty } = readStore();
    const list = Array.isArray(map[key]) ? map[key] : [];
    const idx = list.findIndex(item => item?.name === presetName);
    if (idx < 0) {
      if (dirty) {
        writeStore(map);
      }
      return { presets: clonePresetList(list), removed: false };
    }
    list.splice(idx, 1);
    map[key] = list;
    writeStore(map);
    return { presets: clonePresetList(list), removed: true };
  }

  function apply(viewKey, name, applyFn) {
    const key = String(viewKey ?? '').trim();
    const presetName = String(name ?? '').trim();
    if (!key || !presetName || typeof applyFn !== 'function') {
      return { presets: [], applied: null };
    }
    const { map, dirty } = readStore();
    const list = Array.isArray(map[key]) ? map[key] : [];
    const idx = list.findIndex(item => item?.name === presetName);
    if (idx < 0) {
      if (dirty) {
        writeStore(map);
      }
      return { presets: clonePresetList(list), applied: null };
    }
    const target = list[idx];
    const payload = clonePlain(target.filters ?? {});
    const result = applyFn(payload, clonePresetEntry(target));
    if (result === false) {
      if (dirty) {
        writeStore(map);
      }
      return { presets: clonePresetList(list), applied: null };
    }
    target.lastAppliedAt = Date.now();
    map[key] = list;
    writeStore(map);
    return { presets: clonePresetList(list), applied: clonePresetEntry(target) };
  }

  global.TaskPresets = {
    load,
    save,
    remove,
    apply,
  };
}(window));

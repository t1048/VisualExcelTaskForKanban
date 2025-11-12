(function (global) {
  'use strict';

  const constants = {
    ASSIGNEE_FILTER_ALL: '__ALL__',
    ASSIGNEE_FILTER_UNASSIGNED: '__UNASSIGNED__',
    ASSIGNEE_UNASSIGNED_LABEL: '（未割り当て）',
    CATEGORY_FILTER_ALL: '__CATEGORY_ALL__',
    CATEGORY_FILTER_MINOR_ALL: '__CATEGORY_MINOR_ALL__',
  };

  function renderHeaderTemplate(container, options = {}) {
    if (!container) return;
    const presetLabel = options.presetLabel || 'プリセット';
    container.innerHTML = `
      <div class="filter">
        <label for="flt-assignee">担当者</label>
        <select id="flt-assignee">
          <option value="${constants.ASSIGNEE_FILTER_ALL}">（全員）</option>
          <option value="${constants.ASSIGNEE_FILTER_UNASSIGNED}">${constants.ASSIGNEE_UNASSIGNED_LABEL}</option>
        </select>
      </div>

      <div class="filter">
        <label for="flt-major">大分類</label>
        <select id="flt-major">
          <option value="${constants.CATEGORY_FILTER_ALL}">（すべて）</option>
        </select>
      </div>

      <div class="filter">
        <label for="flt-minor">中分類</label>
        <select id="flt-minor" disabled>
          <option value="${constants.CATEGORY_FILTER_MINOR_ALL}">（すべて）</option>
        </select>
      </div>

      <div class="filter" style="min-width:260px;flex:1;">
        <label>ステータス</label>
        <div class="status-checks" id="flt-statuses"></div>
      </div>

      <div class="filter" style="min-width:260px;flex:1;">
        <label for="flt-keyword">キーワード</label>
        <input type="text" id="flt-keyword" placeholder="タスク・備考を検索" />
      </div>

      <div class="filter" style="min-width:320px;">
        <label>期限</label>
        <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
          <select id="flt-date-mode" title="期限フィルター">
            <option value="none">指定なし</option>
            <option value="range">範囲指定</option>
            <option value="before">指定日以前</option>
            <option value="after">指定日以後</option>
          </select>
          <input type="date" id="flt-date-from" />
          <span id="flt-date-sep" style="display:none;">〜</span>
          <input type="date" id="flt-date-to" style="display:none;" />
          <div class="due-quick-buttons" style="display:flex; gap:4px;">
            <button type="button" class="btn btn-ghost" data-due-preset="this-week">今週まで</button>
            <button type="button" class="btn btn-ghost" data-due-preset="next-week">来週まで</button>
          </div>
        </div>
      </div>

      <div class="filters-actions">
        <div class="filter-preset">
          <label for="flt-preset">${presetLabel}</label>
          <div class="preset-actions" style="display:flex; gap:8px; flex-wrap:wrap; align-items:center;">
            <select id="flt-preset" style="min-width:180px;">
              <option value="">（プリセット未選択）</option>
            </select>
            <button id="btn-preset-apply" class="btn" type="button">呼び出し</button>
            <button id="btn-preset-save" class="btn" type="button">保存</button>
            <button id="btn-preset-delete" class="btn btn-danger" type="button">削除</button>
          </div>
        </div>
        <button id="btn-clear-filters" class="btn" type="button">フィルター解除</button>
      </div>
    `;
  }

  function createController(config = {}) {
    const container = config.container;
    if (!container) {
      throw new Error('TaskFilterUI.createController requires container');
    }

    const viewKey = String(config.viewKey || '').trim();
    const onChange = typeof config.onChange === 'function' ? config.onChange : () => {};
    const normalizeStatusLabel = typeof config.normalizeStatusLabel === 'function'
      ? config.normalizeStatusLabel
      : (value => value);
    const parseISO = typeof config.parseISO === 'function' ? config.parseISO : (() => null);
    const getDueFilterPreset = typeof config.getDueFilterPreset === 'function'
      ? config.getDueFilterPreset
      : (() => null);
    const duePresetSelector = config.duePresetSelector || '[data-due-preset]';
    const presetLabel = config.presetLabel || 'プリセット';

    const presetsApi = global.TaskPresets || {};
    const loadFilterPresets = presetsApi.load || (() => ({ presets: [], lastApplied: null }));
    const saveFilterPreset = presetsApi.save || (() => ({ presets: [] }));
    const removeFilterPreset = presetsApi.remove || (() => ({ presets: [], removed: false }));
    const applyFilterPresetInternal = presetsApi.apply || (() => ({ presets: [], applied: null }));
    const UNSET_STATUS_LABEL = global.TaskValidation?.UNSET_STATUS_LABEL || 'ステータス未設定';

    renderHeaderTemplate(container, { presetLabel });

    const state = {
      tasks: [],
      statuses: [],
      validations: {},
    };

    const dateModes = new Set(['none', 'range', 'before', 'after']);

    const createDefaultFilterState = (statuses) => {
      const set = new Set();
      const list = Array.isArray(statuses) ? statuses : [];
      list.forEach(status => set.add(status));
      if (set.size === 0) {
        list.forEach(status => set.add(status));
      }
      return {
        assignee: constants.ASSIGNEE_FILTER_ALL,
        statuses: set,
        keyword: '',
        date: { mode: 'none', from: '', to: '' },
        category: { major: constants.CATEGORY_FILTER_ALL, minor: constants.CATEGORY_FILTER_MINOR_ALL },
      };
    };

    let filters = createDefaultFilterState(state.statuses);
    let presets = [];
    let activePreset = '';
    let presetInitialApplied = false;
    let handlersBound = false;

    const serializeFiltersForPreset = () => {
      const result = {
        assignee: filters.assignee,
        statuses: [],
        keyword: String(filters.keyword ?? ''),
        date: {
          mode: filters.date?.mode || 'none',
          from: filters.date?.from || '',
          to: filters.date?.to || '',
        },
        category: {
          major: filters.category?.major ?? constants.CATEGORY_FILTER_ALL,
          minor: filters.category?.minor ?? constants.CATEGORY_FILTER_MINOR_ALL,
        },
      };

      const seen = new Set();
      const orderedStatuses = Array.isArray(state.statuses) ? state.statuses.slice() : [];
      orderedStatuses.forEach(status => {
        const text = String(status ?? '').trim();
        if (!text || seen.has(text)) return;
        if (filters.statuses.has(text)) {
          result.statuses.push(text);
          seen.add(text);
        }
      });
      if (filters.statuses.has(UNSET_STATUS_LABEL) && !seen.has(UNSET_STATUS_LABEL)) {
        result.statuses.push(UNSET_STATUS_LABEL);
        seen.add(UNSET_STATUS_LABEL);
      }
      filters.statuses.forEach(status => {
        const text = String(status ?? '').trim();
        if (!text || seen.has(text)) return;
        result.statuses.push(text);
        seen.add(text);
      });

      return result;
    };

    const applyPresetFilters = (raw) => {
      const data = raw && typeof raw === 'object' ? raw : {};
      const next = createDefaultFilterState(state.statuses);

      const assigneeRaw = String(data.assignee ?? '').trim();
      if (!assigneeRaw) {
        next.assignee = constants.ASSIGNEE_FILTER_ALL;
      } else if (assigneeRaw === constants.ASSIGNEE_FILTER_UNASSIGNED) {
        next.assignee = constants.ASSIGNEE_FILTER_UNASSIGNED;
      } else {
        next.assignee = assigneeRaw;
      }

      const availableStatuses = new Set(Array.isArray(state.statuses)
        ? state.statuses.map(s => String(s ?? '').trim())
        : []);
      availableStatuses.add(UNSET_STATUS_LABEL);
      const presetStatuses = Array.isArray(data.statuses) ? data.statuses : [];
      const assigned = new Set();
      presetStatuses.forEach(value => {
        const text = String(value ?? '').trim();
        if (!text || assigned.has(text)) return;
        if (availableStatuses.has(text)) {
          assigned.add(text);
        }
      });
      if (assigned.size === 0) {
        availableStatuses.forEach(status => {
          if (status) assigned.add(status);
        });
      }
      next.statuses = assigned;

      next.keyword = String(data.keyword ?? '');

      const dateRaw = data.date && typeof data.date === 'object' ? data.date : {};
      const mode = dateModes.has(dateRaw.mode) ? dateRaw.mode : 'none';
      next.date = {
        mode,
        from: String(dateRaw.from ?? ''),
        to: String(dateRaw.to ?? ''),
      };

      const categoryRaw = data.category && typeof data.category === 'object' ? data.category : {};
      const major = String(categoryRaw.major ?? '').trim() || constants.CATEGORY_FILTER_ALL;
      const minor = String(categoryRaw.minor ?? '').trim() || constants.CATEGORY_FILTER_MINOR_ALL;
      next.category = { major, minor };

      return next;
    };

    const loadPresetsState = () => {
      if (!viewKey) {
        presets = [];
        activePreset = '';
        return;
      }
      try {
        const { presets: stored, lastApplied } = loadFilterPresets(viewKey) || {};
        presets = Array.isArray(stored) ? stored : [];
        activePreset = lastApplied?.name || '';
      } catch (err) {
        console.warn('[filters] failed to load filter presets:', err);
        presets = [];
        activePreset = '';
      }
    };

    const syncFilterStatuses = (prevSelection) => {
      const statuses = Array.isArray(state.statuses) ? state.statuses : [];
      const base = prevSelection instanceof Set ? prevSelection : new Set();
      const next = new Set();
      statuses.forEach(s => {
        if (base.has(s)) next.add(s);
      });
      const hasUnset = statuses.includes(UNSET_STATUS_LABEL);
      if (hasUnset) {
        const hadUnset = base.has(UNSET_STATUS_LABEL);
        const emptyExists = Array.isArray(state.tasks) && state.tasks.some(t => !String(t?.ステータス ?? '').trim());
        if (hadUnset || emptyExists || base.size === 0) {
          next.add(UNSET_STATUS_LABEL);
        }
      }
      if (next.size === 0) {
        statuses.forEach(s => next.add(s));
      }
      filters = {
        ...filters,
        statuses: next,
      };
    };

    const maybeApplyInitialPreset = () => {
      if (presetInitialApplied) return false;
      presetInitialApplied = true;
      if (!activePreset) return false;
      const preset = presets.find(item => item?.name === activePreset);
      if (!preset) {
        activePreset = '';
        return false;
      }
      filters = applyPresetFilters(preset.filters);
      return true;
    };

    const collectCategoryOptions = () => {
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

      if (Array.isArray(state.tasks)) {
        state.tasks.forEach(task => {
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

      const validatedMajors = Array.isArray(state.validations['大分類']) ? state.validations['大分類'] : [];
      validatedMajors.forEach(name => {
        ensureMajor(name);
      });

      const validatedMinors = Array.isArray(state.validations['中分類']) ? state.validations['中分類'] : [];
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
    };

    const uniqAssignees = () => {
      const set = new Set();
      state.tasks.forEach(t => {
        const text = String(t?.担当者 ?? '').trim();
        if (text) set.add(text);
      });
      return Array.from(set).sort((a, b) => a.localeCompare(b, 'ja'));
    };

    const notifyChange = () => {
      try {
        onChange(filters);
      } catch (err) {
        console.error('[filters] onChange callback failed:', err);
      }
    };

    const updatePresetUI = () => {
      const select = container.querySelector('#flt-preset');
      const applyBtn = container.querySelector('#btn-preset-apply');
      const deleteBtn = container.querySelector('#btn-preset-delete');
      ensurePresetHandlers();
      if (!select) return;

      const previousValue = select.value;
      select.innerHTML = '';

      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = '（プリセット未選択）';
      select.appendChild(placeholder);

      presets.forEach(preset => {
        if (!preset || typeof preset.name !== 'string') return;
        const opt = document.createElement('option');
        opt.value = preset.name;
        opt.textContent = preset.name;
        select.appendChild(opt);
      });

      let nextValue = '';
      if (activePreset && presets.some(p => p?.name === activePreset)) {
        nextValue = activePreset;
      } else if (presets.some(p => p?.name === previousValue)) {
        nextValue = previousValue;
        activePreset = previousValue;
      } else {
        activePreset = '';
      }

      select.value = nextValue;
      const hasSelection = Boolean(select.value);
      if (applyBtn) applyBtn.disabled = !hasSelection;
      if (deleteBtn) deleteBtn.disabled = !hasSelection;
    };

    const ensurePresetHandlers = () => {
      if (handlersBound) return;
      const select = container.querySelector('#flt-preset');
      const applyBtn = container.querySelector('#btn-preset-apply');
      const saveBtn = container.querySelector('#btn-preset-save');
      const deleteBtn = container.querySelector('#btn-preset-delete');
      if (!select || !applyBtn || !saveBtn || !deleteBtn) return;
      handlersBound = true;

      const refreshButtonState = () => {
        const selected = select.value;
        const exists = Boolean(selected) && presets.some(preset => preset?.name === selected);
        applyBtn.disabled = !exists;
        deleteBtn.disabled = !exists;
      };

      select.addEventListener('change', () => {
        activePreset = select.value;
        refreshButtonState();
      });

      applyBtn.addEventListener('click', () => {
        const targetName = select.value;
        if (!targetName) {
          alert('プリセットを選択してください。');
          return;
        }
        const result = applyFilterPresetInternal(viewKey, targetName, (payload) => {
          filters = applyPresetFilters(payload);
          return true;
        });
        presets = result.presets;
        if (result.applied) {
          activePreset = result.applied.name;
          presetInitialApplied = true;
          buildFiltersUI();
          notifyChange();
        } else {
          alert('選択したプリセットが見つかりません。');
          updatePresetUI();
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
        const payload = serializeFiltersForPreset();
        const result = saveFilterPreset(viewKey, trimmed, payload);
        presets = result.presets;
        if (result.saved) {
          activePreset = result.saved.name;
          presetInitialApplied = true;
        }
        updatePresetUI();
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
        const result = removeFilterPreset(viewKey, targetName);
        presets = result.presets;
        if (activePreset === targetName) {
          activePreset = '';
        }
        updatePresetUI();
      });

      refreshButtonState();
    };

    const bindDuePresetButtons = () => {
      container.querySelectorAll(duePresetSelector).forEach(btn => {
        if (!btn || btn.dataset.bound === '1') return;
        btn.dataset.bound = '1';
        btn.addEventListener('click', () => {
          const presetName = btn.dataset.duePreset || '';
          const preset = getDueFilterPreset(presetName);
          if (!preset) return;
          filters = {
            ...filters,
            date: {
              mode: preset.mode,
              from: preset.from,
              to: preset.to,
            },
          };
          const modeSel = container.querySelector('#flt-date-mode');
          const fromEl = container.querySelector('#flt-date-from');
          const toEl = container.querySelector('#flt-date-to');
          const sepEl = container.querySelector('#flt-date-sep');
          if (modeSel) modeSel.value = preset.mode;
          if (fromEl) fromEl.value = preset.from;
          if (toEl) toEl.value = preset.to;
          if (sepEl) {
            const isRange = preset.mode === 'range';
            sepEl.style.display = isRange ? '' : 'none';
            toEl.style.display = isRange ? '' : 'none';
          }
          notifyChange();
        });
      });
    };

    const handleClearFilters = () => {
      filters = createDefaultFilterState(state.statuses);
      activePreset = '';
      presetInitialApplied = true;
      buildFiltersUI();
      notifyChange();
    };

    const buildFiltersUI = () => {
      const wrap = container.querySelector('#flt-statuses');
      if (wrap) {
        wrap.innerHTML = '';
        if (filters.statuses.size === 0) {
          state.statuses.forEach(s => filters.statuses.add(s));
        }
        state.statuses.forEach(s => {
          const id = 'st-' + btoa(unescape(encodeURIComponent(s))).replace(/=/g, '');
          const lbl = document.createElement('label');
          const cb = document.createElement('input');
          cb.type = 'checkbox';
          cb.id = id;
          cb.value = s;
          cb.checked = filters.statuses.has(s);
          cb.addEventListener('change', () => {
            if (cb.checked) {
              filters.statuses.add(s);
            } else {
              filters.statuses.delete(s);
            }
            notifyChange();
          });
          const span = document.createElement('span');
          span.textContent = s;
          lbl.appendChild(cb);
          lbl.appendChild(span);
          wrap.appendChild(lbl);
        });
      }

      const majorSel = container.querySelector('#flt-major');
      const minorSel = container.querySelector('#flt-minor');
      if (majorSel && minorSel) {
        const { majorList, minorMap, allMinors } = collectCategoryOptions();
        let currentMajor = filters.category?.major ?? constants.CATEGORY_FILTER_ALL;
        let currentMinor = filters.category?.minor ?? constants.CATEGORY_FILTER_MINOR_ALL;

        if (!majorList.includes(currentMajor)) {
          currentMajor = constants.CATEGORY_FILTER_ALL;
          filters.category.major = constants.CATEGORY_FILTER_ALL;
        }

        const majorOptions = [
          `<option value="${constants.CATEGORY_FILTER_ALL}">（すべて）</option>`
        ].concat(majorList.map(name => `<option value="${name}">${name}</option>`));
        majorSel.innerHTML = majorOptions.join('');
        majorSel.value = currentMajor;

        const renderMinorOptions = ({ preserve = false } = {}) => {
          const isAllMajor = currentMajor === constants.CATEGORY_FILTER_ALL;
          const majorsMinors = isAllMajor ? (allMinors || []) : (minorMap.get(currentMajor) || []);
          const minorOptions = [
            `<option value="${constants.CATEGORY_FILTER_MINOR_ALL}">（すべて）</option>`
          ].concat(majorsMinors.map(name => `<option value="${name}">${name}</option>`));
          minorSel.innerHTML = minorOptions.join('');

          if (!preserve) {
            currentMinor = constants.CATEGORY_FILTER_MINOR_ALL;
          }

          if (currentMinor !== constants.CATEGORY_FILTER_MINOR_ALL && !majorsMinors.includes(currentMinor)) {
            currentMinor = constants.CATEGORY_FILTER_MINOR_ALL;
          }

          if (majorsMinors.length === 0) {
            currentMinor = constants.CATEGORY_FILTER_MINOR_ALL;
            minorSel.disabled = true;
          } else {
            minorSel.disabled = false;
          }

          filters.category.minor = currentMinor;
          minorSel.value = currentMinor;
        };

        renderMinorOptions({ preserve: true });

        majorSel.onchange = () => {
          currentMajor = majorSel.value;
          filters.category.major = currentMajor;
          if (currentMajor !== constants.CATEGORY_FILTER_ALL) {
            currentMinor = constants.CATEGORY_FILTER_MINOR_ALL;
          }
          renderMinorOptions({ preserve: currentMajor === constants.CATEGORY_FILTER_ALL });
          notifyChange();
        };

        minorSel.onchange = () => {
          currentMinor = minorSel.value;
          filters.category.minor = currentMinor;
          notifyChange();
        };
      }

      const assigneeSel = container.querySelector('#flt-assignee');
      if (assigneeSel) {
        const selected = filters.assignee;
        const list = uniqAssignees();
        const options = [
          `<option value="${constants.ASSIGNEE_FILTER_ALL}">（全員）</option>`,
          `<option value="${constants.ASSIGNEE_FILTER_UNASSIGNED}">${constants.ASSIGNEE_UNASSIGNED_LABEL}</option>`
        ].concat(list.map(a => `<option value="${a}">${a}</option>`));
        assigneeSel.innerHTML = options.join('');
        if (selected === constants.ASSIGNEE_FILTER_UNASSIGNED) {
          assigneeSel.value = constants.ASSIGNEE_FILTER_UNASSIGNED;
        } else if (list.includes(selected)) {
          assigneeSel.value = selected;
        } else {
          assigneeSel.value = constants.ASSIGNEE_FILTER_ALL;
          filters.assignee = constants.ASSIGNEE_FILTER_ALL;
        }
        assigneeSel.onchange = () => {
          filters.assignee = assigneeSel.value;
          notifyChange();
        };
      }

      const keywordEl = container.querySelector('#flt-keyword');
      if (keywordEl) {
        keywordEl.value = filters.keyword || '';
        keywordEl.oninput = () => {
          filters.keyword = keywordEl.value;
          notifyChange();
        };
      }

      const modeSel = container.querySelector('#flt-date-mode');
      const fromEl = container.querySelector('#flt-date-from');
      const toEl = container.querySelector('#flt-date-to');
      const sepEl = container.querySelector('#flt-date-sep');

      const updateVisibility = () => {
        const m = modeSel ? modeSel.value : 'none';
        if (!modeSel || !fromEl || !toEl || !sepEl) return;
        if (m === 'range') {
          toEl.style.display = '';
          sepEl.style.display = '';
        } else {
          toEl.style.display = 'none';
          sepEl.style.display = 'none';
        }
      };

      if (modeSel && fromEl && toEl && sepEl) {
        modeSel.value = filters.date.mode || 'none';
        fromEl.value = filters.date.from || '';
        toEl.value = filters.date.to || '';
        updateVisibility();

        modeSel.onchange = () => {
          filters.date.mode = modeSel.value;
          updateVisibility();
          notifyChange();
        };
        fromEl.onchange = () => {
          filters.date.from = fromEl.value;
          notifyChange();
        };
        toEl.onchange = () => {
          filters.date.to = toEl.value;
          notifyChange();
        };
      }

      const clearBtn = container.querySelector('#btn-clear-filters');
      if (clearBtn) {
        clearBtn.onclick = handleClearFilters;
      }

      bindDuePresetButtons();
      updatePresetUI();
    };

    const applyFilters = (list) => {
      const source = Array.isArray(list) ? list : [];
      const assignee = filters.assignee;
      const statuses = filters.statuses;
      const df = filters.date;
      const keyword = (filters.keyword || '').trim().toLowerCase();
      const majorFilter = filters.category?.major ?? constants.CATEGORY_FILTER_ALL;
      const minorFilter = filters.category?.minor ?? constants.CATEGORY_FILTER_MINOR_ALL;

      const shouldFilterMajor = majorFilter !== constants.CATEGORY_FILTER_ALL;
      const shouldFilterMinor = minorFilter !== constants.CATEGORY_FILTER_MINOR_ALL;

      return source.filter(t => {
        if (shouldFilterMajor || shouldFilterMinor) {
          const major = String(t?.大分類 ?? '').trim();
          const minor = String(t?.中分類 ?? '').trim();
          if (shouldFilterMajor && major !== majorFilter) return false;
          if (shouldFilterMinor && minor !== minorFilter) return false;
        }

        const who = String(t?.担当者 ?? '').trim();
        if (assignee === constants.ASSIGNEE_FILTER_UNASSIGNED) {
          if (who) return false;
        } else if (assignee !== constants.ASSIGNEE_FILTER_ALL) {
          if (who !== assignee) return false;
        }

        const normalizedStatus = normalizeStatusLabel(t?.ステータス);
        if (!statuses.has(normalizedStatus)) return false;

        if (keyword) {
          const title = String(t?.タスク ?? '').toLowerCase();
          const note = String(t?.備考 ?? '').toLowerCase();
          if (!title.includes(keyword) && !note.includes(keyword)) return false;
        }

        if (df.mode === 'none') return true;
        const due = parseISO(t?.期限 || '');
        if (!due) return false;

        if (df.mode === 'before') {
          const d = parseISO(df.from);
          if (!d) return true;
          return due.getTime() <= d.getTime();
        }
        if (df.mode === 'after') {
          const d = parseISO(df.from);
          if (!d) return true;
          return due.getTime() >= d.getTime();
        }
        if (df.mode === 'range') {
          const from = parseISO(df.from);
          const to = parseISO(df.to);
          if (from && to) return from.getTime() <= due.getTime() && due.getTime() <= to.getTime();
          if (from && !to) return from.getTime() <= due.getTime();
          if (!from && to) return due.getTime() <= to.getTime();
          return true;
        }
        return true;
      });
    };

    const updateData = ({ tasks, statuses, validations, preserveStatusSelection = true } = {}) => {
      const prevSelection = preserveStatusSelection ? new Set(filters.statuses) : new Set();
      state.tasks = Array.isArray(tasks) ? tasks : [];
      state.statuses = Array.isArray(statuses) ? statuses : [];
      state.validations = validations && typeof validations === 'object' ? validations : {};
      syncFilterStatuses(prevSelection);
      const applied = maybeApplyInitialPreset();
      buildFiltersUI();
      if (applied) {
        notifyChange();
      }
    };

    const setAssignee = (value, options = {}) => {
      const next = (() => {
        if (!value || value === constants.ASSIGNEE_FILTER_ALL) return constants.ASSIGNEE_FILTER_ALL;
        return value;
      })();
      if (filters.assignee === next) return;
      filters.assignee = next;
      const assigneeSel = container.querySelector('#flt-assignee');
      if (assigneeSel) {
        assigneeSel.value = next;
      }
      if (options.silent) return;
      notifyChange();
    };

    loadPresetsState();
    ensurePresetHandlers();
    buildFiltersUI();

    return {
      updateData,
      getFilters: () => filters,
      applyFilters,
      setAssignee,
      getActivePreset: () => activePreset,
      setActivePreset: (name) => { activePreset = String(name || ''); updatePresetUI(); },
      reloadPresetState: () => { loadPresetsState(); updatePresetUI(); },
      markPresetApplied: () => { presetInitialApplied = true; },
      rebuild: buildFiltersUI,
    };
  }

  global.TaskFilterUI = {
    constants,
    renderHeaderTemplate,
    createController,
  };
}(window));

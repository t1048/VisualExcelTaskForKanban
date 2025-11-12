(function (global) {
  const noop = () => {};

  function getPresetApi() {
    const api = global.TaskPresets || {};
    return {
      load: typeof api.load === 'function' ? api.load : () => ({ presets: [], lastApplied: null }),
      save: typeof api.save === 'function' ? api.save : () => ({ presets: [] }),
      remove: typeof api.remove === 'function' ? api.remove : () => ({ presets: [], removed: false }),
      apply: typeof api.apply === 'function' ? api.apply : () => ({ presets: [], applied: null }),
    };
  }

  function resolveElement(selector) {
    if (!selector) return null;
    if (selector instanceof Element) return selector;
    if (typeof selector === 'string') {
      return document.querySelector(selector);
    }
    return null;
  }

  function createFilterPresetManager({ viewKey, selectors = {}, serialize = noop, applyToUI = noop } = {}) {
    if (!viewKey) {
      throw new Error('createFilterPresetManager: viewKey is required');
    }

    const api = getPresetApi();
    let presets = [];
    let activePresetName = '';
    let initialApplied = false;
    let handlersBound = false;

    function loadState() {
      try {
        const { presets: loaded, lastApplied } = api.load(viewKey) || {};
        presets = Array.isArray(loaded) ? loaded : [];
        activePresetName = lastApplied?.name || '';
      } catch (err) {
        console.warn(`[${viewKey}] failed to load filter presets`, err);
        presets = [];
        activePresetName = '';
      }
    }

    loadState();

    function getElements() {
      return {
        select: resolveElement(selectors.select || selectors.dropdown || selectors.preset),
        applyBtn: resolveElement(selectors.apply),
        saveBtn: resolveElement(selectors.save),
        deleteBtn: resolveElement(selectors.delete || selectors.remove),
      };
    }

    function refreshButtonState(selectEl, applyBtn, deleteBtn) {
      const selected = selectEl?.value || '';
      const exists = Boolean(selected) && presets.some(preset => preset?.name === selected);
      if (applyBtn) applyBtn.disabled = !exists;
      if (deleteBtn) deleteBtn.disabled = !exists;
    }

    function updateOptions(selectEl) {
      if (!selectEl) return;
      const previousValue = selectEl.value;
      selectEl.innerHTML = '';

      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = '（プリセット未選択）';
      selectEl.appendChild(placeholder);

      presets.forEach(preset => {
        if (!preset || typeof preset.name !== 'string') return;
        const opt = document.createElement('option');
        opt.value = preset.name;
        opt.textContent = preset.name;
        selectEl.appendChild(opt);
      });

      let nextValue = '';
      if (activePresetName && presets.some(p => p?.name === activePresetName)) {
        nextValue = activePresetName;
      } else if (presets.some(p => p?.name === previousValue)) {
        nextValue = previousValue;
        activePresetName = previousValue;
      } else {
        activePresetName = '';
      }

      selectEl.value = nextValue;
    }

    function ensureHandlers() {
      if (handlersBound) return;
      const { select, applyBtn, saveBtn, deleteBtn } = getElements();
      if (!select || !applyBtn || !saveBtn || !deleteBtn) return;

      const handleSelectionChange = () => {
        activePresetName = select.value || '';
        refreshButtonState(select, applyBtn, deleteBtn);
      };

      select.addEventListener('change', handleSelectionChange);

      applyBtn.addEventListener('click', () => {
        const targetName = select.value || '';
        if (!targetName) {
          alert('プリセットを選択してください。');
          return;
        }
        const result = api.apply(viewKey, targetName, (filters) => {
          applyToUI(filters);
          return true;
        });
        presets = result.presets;
        if (result.applied) {
          activePresetName = result.applied.name;
          initialApplied = true;
          updateUI();
        } else {
          alert('選択したプリセットが見つかりません。');
          loadState();
          updateUI();
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
        const payload = serialize();
        const result = api.save(viewKey, trimmed, payload);
        presets = result.presets;
        if (result.saved) {
          activePresetName = result.saved.name;
          initialApplied = true;
        }
        updateUI();
      });

      deleteBtn.addEventListener('click', () => {
        const targetName = select.value || '';
        if (!targetName) {
          alert('削除するプリセットを選択してください。');
          return;
        }
        if (!window.confirm(`プリセット「${targetName}」を削除しますか？`)) {
          return;
        }
        const result = api.remove(viewKey, targetName);
        presets = result.presets;
        if (activePresetName === targetName) {
          activePresetName = '';
        }
        updateUI();
      });

      handlersBound = true;
      refreshButtonState(select, applyBtn, deleteBtn);
    }

    function updateUI() {
      const { select, applyBtn, deleteBtn } = getElements();
      if (select) {
        updateOptions(select);
      }
      refreshButtonState(select, applyBtn, deleteBtn);
      ensureHandlers();
    }

    function maybeApplyInitialPreset() {
      if (initialApplied) return;
      initialApplied = true;
      if (!activePresetName) return;
      const preset = presets.find(item => item?.name === activePresetName);
      if (!preset) {
        activePresetName = '';
        updateUI();
        return;
      }
      applyToUI(preset.filters);
    }

    return {
      maybeApplyInitialPreset,
      updateUI,
      getActivePresetName: () => activePresetName,
      reload: () => {
        loadState();
        updateUI();
      },
    };
  }

  global.TaskFilterPresets = {
    createFilterPresetManager,
  };
})(window);

(function (global) {
  'use strict';

  const DEFAULT_STATUSES = ['未着手', '進行中', '完了', '保留'];
  const PRIORITY_DEFAULT_OPTIONS = ['高', '中', '低'];
  const UNSET_STATUS_LABEL = 'ステータス未設定';

  function normalizeValidationValues(rawList) {
    if (!Array.isArray(rawList)) return [];
    const seen = new Set();
    const values = [];
    rawList.forEach(value => {
      const text = String(value ?? '').trim();
      if (!text || seen.has(text)) return;
      seen.add(text);
      values.push(text);
    });
    return values;
  }

  function normalizeStatusLabel(value) {
    const text = String(value ?? '').trim();
    return text || UNSET_STATUS_LABEL;
  }

  function denormalizeStatusLabel(value) {
    const text = String(value ?? '').trim();
    return text === UNSET_STATUS_LABEL ? '' : text;
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

  function applyValidationState(snapshot = {}) {
    const result = {
      tasks: Array.isArray(snapshot.tasks) ? snapshot.tasks : [],
      statuses: Array.isArray(snapshot.statuses) ? snapshot.statuses.slice() : [],
      validations: {},
    };

    const rawValidations = snapshot.validations && typeof snapshot.validations === 'object'
      ? snapshot.validations
      : {};

    const normalizedValidations = {};
    Object.keys(rawValidations).forEach(key => {
      const values = normalizeValidationValues(rawValidations[key]);
      if (values.length > 0) {
        normalizedValidations[key] = values;
      }
    });

    if (!Array.isArray(normalizedValidations['ステータス']) || normalizedValidations['ステータス'].length === 0) {
      normalizedValidations['ステータス'] = [...DEFAULT_STATUSES];
    }
    if (!Array.isArray(normalizedValidations['優先度']) || normalizedValidations['優先度'].length === 0) {
      normalizedValidations['優先度'] = [...PRIORITY_DEFAULT_OPTIONS];
    }

    const seenStatuses = new Set();
    let orderedStatuses = [];

    if (Array.isArray(normalizedValidations['ステータス'])) {
      normalizedValidations['ステータス'].forEach(status => {
        const text = String(status ?? '').trim();
        if (!text || seenStatuses.has(text)) return;
        seenStatuses.add(text);
        orderedStatuses.push(text);
      });
    }

    if (Array.isArray(result.statuses)) {
      result.statuses.forEach(status => {
        const text = String(status ?? '').trim();
        if (!text || seenStatuses.has(text)) return;
        seenStatuses.add(text);
        orderedStatuses.push(text);
      });
    }

    let hasUnset = false;
    result.tasks.forEach(task => {
      const text = String(task?.ステータス ?? '').trim();
      if (!text) {
        hasUnset = true;
        return;
      }
      if (seenStatuses.has(text)) return;
      seenStatuses.add(text);
      orderedStatuses.push(text);
    });

    if (orderedStatuses.length === 0) {
      DEFAULT_STATUSES.forEach(status => {
        if (seenStatuses.has(status)) return;
        seenStatuses.add(status);
        orderedStatuses.push(status);
      });
    }

    if (hasUnset) {
      orderedStatuses = [UNSET_STATUS_LABEL, ...orderedStatuses.filter(status => status !== UNSET_STATUS_LABEL)];
    } else {
      orderedStatuses = orderedStatuses.filter((status, index) => status !== UNSET_STATUS_LABEL || orderedStatuses.indexOf(status) === index);
    }

    if (orderedStatuses.length === 0) {
      orderedStatuses = [...DEFAULT_STATUSES];
    }

    result.statuses = orderedStatuses;
    result.validations = normalizedValidations;
    return result;
  }

  global.TaskValidation = {
    DEFAULT_STATUSES,
    PRIORITY_DEFAULT_OPTIONS,
    UNSET_STATUS_LABEL,
    normalizeValidationValues,
    normalizeStatusLabel,
    denormalizeStatusLabel,
    createPriorityHelper,
    applyValidationState,
  };
}(window));

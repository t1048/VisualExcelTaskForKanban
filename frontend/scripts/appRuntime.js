(function (global) {
  'use strict';

  const INITIAL_LOAD_FLAG_KEY = 'kanban:excelLoaded';

  const state = {
    api: null,
    runMode: 'mock',
    handlers: {
      save: null,
      reload: null,
    },
    keydownHandler: null,
  };

  function ready(fn) {
    if (typeof fn !== 'function') return;
    if (document.readyState !== 'loading') {
      try {
        fn();
      } catch (err) {
        console.error('[TaskAppRuntime] ready callback failed', err);
      }
      return;
    }
    document.addEventListener('DOMContentLoaded', () => {
      try {
        fn();
      } catch (err) {
        console.error('[TaskAppRuntime] ready callback failed', err);
      }
    }, { once: true });
  }

  function safeSessionGet(key) {
    try {
      return global.sessionStorage?.getItem(key) ?? null;
    } catch (err) {
      return null;
    }
  }

  function safeSessionSet(key, value) {
    try {
      global.sessionStorage?.setItem(key, value);
    } catch (err) {
      // ignore
    }
  }

  function safeSessionRemove(key) {
    try {
      global.sessionStorage?.removeItem(key);
    } catch (err) {
      // ignore
    }
  }

  function hasInitialExcelLoadFlag() {
    return safeSessionGet(INITIAL_LOAD_FLAG_KEY) === '1';
  }

  function markInitialExcelLoadFlag() {
    safeSessionSet(INITIAL_LOAD_FLAG_KEY, '1');
  }

  function resetInitialExcelLoadFlag() {
    safeSessionRemove(INITIAL_LOAD_FLAG_KEY);
  }

  function bindExcelActions({ onSave, onReload, enableKeyboardShortcuts = true } = {}) {
    state.handlers.save = typeof onSave === 'function' ? onSave : null;
    state.handlers.reload = typeof onReload === 'function' ? onReload : null;

    if (state.keydownHandler) {
      global.removeEventListener('keydown', state.keydownHandler);
      state.keydownHandler = null;
    }

    if (enableKeyboardShortcuts) {
      state.keydownHandler = (event) => {
        if (!event) return;
        const isModifier = event.ctrlKey || event.metaKey;
        if (!isModifier) return;
        const key = String(event.key || '').toLowerCase();
        if (key === 's' && state.handlers.save) {
          event.preventDefault();
          try {
            state.handlers.save(event);
          } catch (err) {
            console.error('[TaskAppRuntime] save handler failed', err);
          }
          return;
        }
        if (key === 'r' && event.shiftKey && state.handlers.reload) {
          event.preventDefault();
          try {
            state.handlers.reload(event);
          } catch (err) {
            console.error('[TaskAppRuntime] reload handler failed', err);
          }
        }
      };
      global.addEventListener('keydown', state.keydownHandler);
    }

    return () => {
      state.handlers.save = null;
      state.handlers.reload = null;
      if (state.keydownHandler) {
        global.removeEventListener('keydown', state.keydownHandler);
        state.keydownHandler = null;
      }
    };
  }

  function setupRuntime({ mockApiFactory, onApiChanged, onInit, onRealtimeUpdate } = {}) {
    const getMockApi = () => {
      if (typeof mockApiFactory === 'function') {
        try {
          return mockApiFactory();
        } catch (err) {
          console.error('[TaskAppRuntime] mockApiFactory failed', err);
        }
      }
      const common = global.TaskAppCommon || {};
      if (typeof common.createMockApi === 'function') {
        return common.createMockApi();
      }
      return null;
    };

    const assignApi = (api, runMode) => {
      if (!api) return;
      state.api = api;
      state.runMode = runMode;
      if (typeof onApiChanged === 'function') {
        try {
          onApiChanged({ api, runMode });
        } catch (err) {
          console.error('[TaskAppRuntime] onApiChanged callback failed', err);
        }
      }
      if (typeof onInit === 'function') {
        Promise.resolve(onInit({ api, runMode, force: true })).catch(err => {
          console.error('[TaskAppRuntime] initialization failed', err);
        });
      }
    };

    const handleRealtime = (payload) => {
      if (typeof onRealtimeUpdate !== 'function') return;
      Promise.resolve(onRealtimeUpdate(payload)).catch(err => {
        console.error('[TaskAppRuntime] failed to apply realtime payload', err);
      });
    };

    global.__kanban_receive_update = (payload) => {
      resetInitialExcelLoadFlag();
      handleRealtime(payload);
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
        const mock = getMockApi();
        if (mock) {
          assignApi(mock, 'mock');
        }
      }
    });

    return {
      get api() {
        return state.api;
      },
      get runMode() {
        return state.runMode;
      },
    };
  }

  global.TaskAppRuntime = {
    setupRuntime,
    hasInitialExcelLoadFlag,
    markInitialExcelLoadFlag,
    resetInitialExcelLoadFlag,
    bindExcelActions,
  };
}(window));

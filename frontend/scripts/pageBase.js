(function () {
  const {
    createMockApi,
    createWorkloadSummary,
    getDueFilterPreset,
    parseISO,
    getDueState,
  } = window.TaskAppCommon || {};

  const {
    setupRuntime,
    markInitialExcelLoadFlag,
    resetInitialExcelLoadFlag,
    bindExcelActions,
  } = window.TaskAppRuntime || {};

  const { normalizeStatusLabel } = window.TaskValidation || {};

  const {
    constants: {
      ASSIGNEE_FILTER_ALL,
      ASSIGNEE_FILTER_UNASSIGNED,
      ASSIGNEE_UNASSIGNED_LABEL,
    } = {},
    createController,
  } = window.TaskFilterUI || {};

  const DEFAULT_WORKLOAD_HIGHLIGHT = () => false;

  function defaultAssigneeSelectHandler(value, context) {
    const controller = context?.filterController;
    if (!controller || typeof controller.getFilters !== 'function') {
      return;
    }

    const current = controller.getFilters().assignee ?? ASSIGNEE_FILTER_ALL;
    let next = ASSIGNEE_FILTER_ALL;
    if (value && value !== ASSIGNEE_FILTER_ALL) {
      next = current === value ? ASSIGNEE_FILTER_ALL : value;
    }
    if (typeof controller.setAssignee === 'function') {
      controller.setAssignee(next);
    }

    const select = document.getElementById('flt-assignee');
    if (select) {
      select.value = next;
    }
  }

  function createTaskPageBase(config = {}) {
    const {
      viewKey,
      headerOptions,
      filtersContainer = document.getElementById('filters-bar'),
      workloadContainer = document.getElementById('workload-summary'),
      onRender,
      onInit,
      onRealtimeUpdate,
      onSave,
      onReload,
      onApiChanged,
      getStatuses,
      getActiveAssignee,
      highlightPredicate,
      onSelectAssignee,
      logLabel = viewKey || 'task-page',
    } = config;

    let apiRef = null;
    let runModeRef = 'mock';

    const context = {
      getApi: () => apiRef,
      getRunMode: () => runModeRef,
      headerController: null,
      filterController: null,
      workloadSummary: null,
    };

    if (window.TaskAppHeader && typeof window.TaskAppHeader.initHeader === 'function') {
      context.headerController = window.TaskAppHeader.initHeader(headerOptions || {});
    }

    if (typeof createController === 'function') {
      context.filterController = createController({
        container: filtersContainer,
        viewKey,
        onChange: () => {
          if (typeof onRender === 'function') {
            onRender(context);
          }
        },
        normalizeStatusLabel,
        parseISO,
        getDueFilterPreset,
      });
    }

    if (typeof createWorkloadSummary === 'function') {
      const highlight = typeof highlightPredicate === 'function'
        ? (entry) => highlightPredicate(entry, context)
        : DEFAULT_WORKLOAD_HIGHLIGHT;

      const resolveActiveAssignee = () => {
        if (typeof getActiveAssignee === 'function') {
          return getActiveAssignee(context);
        }
        const controller = context.filterController;
        if (controller && typeof controller.getFilters === 'function') {
          return controller.getFilters().assignee ?? ASSIGNEE_FILTER_ALL;
        }
        return ASSIGNEE_FILTER_ALL;
      };

      context.workloadSummary = createWorkloadSummary({
        container: workloadContainer,
        getStatuses: () => (typeof getStatuses === 'function' ? getStatuses(context) : []),
        normalizeStatusLabel,
        unassignedKey: ASSIGNEE_FILTER_UNASSIGNED,
        unassignedLabel: ASSIGNEE_UNASSIGNED_LABEL,
        allKey: ASSIGNEE_FILTER_ALL,
        getActiveAssignee: resolveActiveAssignee,
        getDueState,
        highlightPredicate: highlight,
        onSelectAssignee: (value) => {
          if (typeof onSelectAssignee === 'function') {
            onSelectAssignee(value, context);
          } else {
            defaultAssigneeSelectHandler(value, context);
          }
        },
      });
    }

    if (typeof setupRuntime === 'function') {
      setupRuntime({
        mockApiFactory: createMockApi,
        onApiChanged: ({ api, runMode }) => {
          apiRef = api;
          runModeRef = runMode;
          if (typeof onApiChanged === 'function') {
            onApiChanged({ api, runMode }, context);
          } else {
            console.log(`[${logLabel}] run mode:`, runModeRef);
          }
        },
        onInit: async () => {
          try {
            if (typeof onInit === 'function') {
              await onInit(context);
            }
            if (runModeRef === 'pywebview') {
              markInitialExcelLoadFlag?.();
            }
          } catch (err) {
            resetInitialExcelLoadFlag?.();
            throw err;
          }
        },
        onRealtimeUpdate: (payload) => {
          if (typeof onRealtimeUpdate === 'function') {
            onRealtimeUpdate(payload, context);
          }
        },
      });
    }

    if (typeof bindExcelActions === 'function') {
      bindExcelActions({
        onSave: () => {
          if (typeof onSave === 'function') {
            onSave(context);
          }
        },
        onReload: () => {
          if (typeof onReload === 'function') {
            onReload(context);
          }
        },
      });
    }

    return context;
  }

  window.TaskPageBase = {
    createTaskPageBase,
  };
})();

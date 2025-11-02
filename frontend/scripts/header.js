(function (global) {
  'use strict';

  const NAV_ITEMS = [
    { key: 'kanban', label: 'カンバン表示', href: 'index.html' },
    { key: 'list', label: 'リスト表示', href: 'list.html' },
    { key: 'timeline', label: 'タイムライン', href: 'timeline.html' },
    { key: 'calendar', label: 'カレンダー', href: 'calendar.html' },
  ];

  const state = {
    current: null,
  };

  const noop = () => {};

  function createButton(id, className, label) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.id = id;
    btn.className = className;
    btn.textContent = label;
    return btn;
  }

  function createDueSection() {
    const container = document.createElement('div');
    container.className = 'toolbar-due';
    container.id = 'toolbar-due';
    container.setAttribute('aria-live', 'polite');

    const badges = document.createElement('div');
    badges.className = 'toolbar-due-badges';

    const overdue = document.createElement('span');
    overdue.id = 'due-overdue-count';
    overdue.className = 'due-indicator due-overdue';
    overdue.hidden = true;
    overdue.innerHTML = '期限超過 <span class="count">0</span>';

    const warning = document.createElement('span');
    warning.id = 'due-warning-count';
    warning.className = 'due-indicator due-warning';
    warning.hidden = true;
    warning.innerHTML = '期限間近 <span class="count">0</span>';

    const toast = document.createElement('div');
    toast.id = 'due-toast';
    toast.className = 'due-toast';
    toast.hidden = true;

    badges.appendChild(overdue);
    badges.appendChild(warning);
    container.appendChild(badges);
    container.appendChild(toast);

    return { container, overdue, warning, toast };
  }

  function createNav(currentView) {
    const nav = document.createElement('nav');
    nav.className = 'toolbar-nav';
    nav.setAttribute('aria-label', '画面切替');

    NAV_ITEMS.forEach(item => {
      if (item.key === currentView) {
        const current = document.createElement('span');
        current.className = 'btn btn-current';
        current.textContent = item.label;
        current.setAttribute('aria-current', 'page');
        nav.appendChild(current);
        return;
      }
      const link = document.createElement('a');
      link.className = 'btn';
      link.href = item.href;
      link.textContent = item.label;
      nav.appendChild(link);
    });

    return nav;
  }

  function updateDueSummary(tasks, instance = state.current) {
    if (!instance) return;
    const due = instance.due;
    if (!due) return;

    const list = Array.isArray(tasks) ? tasks : [];
    const { getDueState } = global.TaskAppCommon || {};
    if (typeof getDueState !== 'function') return;

    let overdue = 0;
    let warning = 0;

    list.forEach(task => {
      const state = getDueState(task);
      if (!state) return;
      if (state.level === 'overdue') {
        overdue += 1;
      } else if (state.level === 'warning') {
        warning += 1;
      }
    });

    const overdueCount = due.overdue.querySelector('.count');
    const warningCount = due.warning.querySelector('.count');
    if (overdueCount) overdueCount.textContent = overdue;
    if (warningCount) warningCount.textContent = warning;

    due.overdue.hidden = overdue === 0;
    due.warning.hidden = warning === 0;

    if (overdue > 0) {
      due.toast.hidden = false;
      due.toast.textContent = `⚠️ 期限を過ぎたカードが ${overdue} 件あります。`;
    } else if (warning > 0) {
      due.toast.hidden = false;
      due.toast.textContent = `⏰ 期限が近いカードが ${warning} 件あります。`;
    } else {
      due.toast.hidden = true;
      due.toast.textContent = '';
    }

    const hasAlerts = overdue > 0 || warning > 0;
    due.container.classList.toggle('active', hasAlerts);
  }

  function initHeader(options = {}) {
    const mountSelector = options.mount ?? '#app-header';
    const mount = typeof mountSelector === 'string'
      ? document.querySelector(mountSelector)
      : mountSelector;
    if (!mount) return null;

    const root = document.createElement('div');
    root.className = 'toolbar';

    const title = document.createElement('div');
    title.className = 'title';
    title.textContent = options.title || '';
    root.appendChild(title);

    const due = createDueSection();
    root.appendChild(due.container);

    const addButton = createButton('btn-add', 'btn btn-primary', options.addButtonLabel || '＋ 追加');
    const saveButton = createButton('btn-save', 'btn btn-success', options.saveButtonLabel || 'Excelへ保存');
    const validationsButton = createButton('btn-validations', 'btn', options.validationsLabel || '入力規則');
    const reloadButton = createButton('btn-reload', 'btn', options.reloadLabel || '再読込');

    root.appendChild(addButton);
    root.appendChild(saveButton);
    root.appendChild(validationsButton);
    root.appendChild(reloadButton);

    const nav = createNav(options.currentView);
    root.appendChild(nav);

    let hintEl = null;
    if (options.hint) {
      hintEl = document.createElement('span');
      hintEl.className = 'hint';
      hintEl.textContent = options.hint;
      root.appendChild(hintEl);
    }

    mount.innerHTML = '';
    mount.appendChild(root);

    const instance = {
      root,
      due,
      title,
      hintEl,
      buttons: {
        add: addButton,
        save: saveButton,
        validations: validationsButton,
        reload: reloadButton,
      },
    };

    state.current = instance;

    const onAdd = typeof options.onAdd === 'function' ? options.onAdd : noop;
    const onSave = typeof options.onSave === 'function' ? options.onSave : noop;
    const onValidations = typeof options.onValidations === 'function' ? options.onValidations : noop;
    const onReload = typeof options.onReload === 'function' ? options.onReload : noop;

    addButton.addEventListener('click', onAdd);
    saveButton.addEventListener('click', onSave);
    validationsButton.addEventListener('click', onValidations);
    reloadButton.addEventListener('click', onReload);

    return {
      element: root,
      updateDueSummary(tasks) {
        updateDueSummary(tasks, instance);
      },
      setHint(text) {
        if (!instance.hintEl && text) {
          instance.hintEl = document.createElement('span');
          instance.hintEl.className = 'hint';
          instance.hintEl.textContent = text;
          root.appendChild(instance.hintEl);
          return;
        }
        if (!instance.hintEl) return;
        if (!text) {
          instance.hintEl.remove();
          instance.hintEl = null;
        } else {
          instance.hintEl.textContent = text;
        }
      },
      setTitle(text) {
        instance.title.textContent = text || '';
      },
    };
  }

  global.TaskAppHeader = {
    initHeader,
    updateDueSummary(tasks) {
      updateDueSummary(tasks);
    },
  };
}(window));

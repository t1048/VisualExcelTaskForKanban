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

  function injectMenuStyles() {
    if (document.getElementById('app-menu-inline-style')) return;
    const style = document.createElement('style');
    style.id = 'app-menu-inline-style';
    style.textContent = `
.app-menu-trigger {
  display: flex;
  align-items: center;
}
.btn.btn-menu {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-weight: 600;
  padding: 8px 14px;
}
.btn.btn-menu .menu-icon {
  font-size: 16px;
}
.app-menu-backdrop {
  position: fixed;
  inset: 0;
  background: rgba(9, 12, 20, 0.72);
  backdrop-filter: blur(6px);
  display: none;
  z-index: 9999;
}
.app-menu-backdrop.is-open {
  display: block;
}
.app-menu-panel {
  position: fixed;
  top: 0;
  left: 0;
  height: 100vh;
  width: min(320px, calc(100vw - 32px));
  max-width: calc(100vw - 32px);
  background: var(--panel, rgba(15, 23, 42, 0.9));
  color: var(--text, #f8fafc);
  border-right: 1px solid var(--border, rgba(148, 163, 184, 0.24));
  box-shadow: var(--shadow, 0 10px 36px rgba(8, 15, 40, 0.45));
  transform: translateX(-100%);
  transition: transform 0.2s ease;
  display: flex;
  flex-direction: column;
  overflow-x: hidden;
  box-sizing: border-box;
}
.app-menu-backdrop.is-open .app-menu-panel {
  transform: translateX(0);
}
.app-menu-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 14px 16px;
  border-bottom: 1px solid var(--border-color, rgba(148, 163, 184, 0.32));
}
.app-menu-title {
  font-size: 16px;
  font-weight: 600;
  margin: 0;
}
.btn.menu-close {
  background: transparent;
  color: var(--muted, #94a3b8);
  border: none;
  box-shadow: none;
}
.app-menu-content {
  flex: 1;
  overflow-y: auto;
  padding-bottom: 12px;
}
.app-menu-section {
  padding: 16px;
  border-bottom: 1px solid var(--border, rgba(148, 163, 184, 0.24));
}
.app-menu-section:last-of-type {
  border-bottom: none;
}
.app-menu-section h3 {
  margin: 0 0 10px;
  font-size: 13px;
  color: var(--muted, #cbd5f5);
  font-weight: 600;
}
.app-menu-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.app-menu-link.btn {
  justify-content: flex-start;
  width: 100%;
  box-sizing: border-box;
}
.app-menu-footer {
  margin-top: auto;
  padding: 12px 16px;
  font-size: 12px;
  color: var(--muted, #94a3b8);
}
`;
    document.head.appendChild(style);
  }

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
    nav.className = 'app-menu-nav';
    nav.setAttribute('aria-label', '表示（ビュー）');

    const list = document.createElement('div');
    list.className = 'app-menu-list';

    NAV_ITEMS.forEach(item => {
      const link = document.createElement('a');
      link.className = 'btn app-menu-link';
      link.href = item.href;
      link.textContent = item.label;
      if (item.key === currentView) {
        link.classList.add('btn-current');
        link.setAttribute('aria-current', 'page');
      }
      list.appendChild(link);
    });

    nav.appendChild(list);
    return nav;
  }

  function createMenuSection(currentView, validationsButton) {
    injectMenuStyles();

    const existingOverlay = document.getElementById('toolbar-menu-overlay');
    if (existingOverlay) existingOverlay.remove();

    const container = document.createElement('div');
    container.className = 'app-menu-trigger';

    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.id = 'toolbar-menu-toggle';
    toggle.className = 'btn btn-menu';
    toggle.setAttribute('aria-haspopup', 'dialog');
    toggle.setAttribute('aria-expanded', 'false');
    toggle.setAttribute('aria-controls', 'toolbar-menu-panel');
    toggle.setAttribute('aria-label', 'メニューを開く');

    const icon = document.createElement('span');
    icon.className = 'menu-icon';
    icon.textContent = '☰';
    const label = document.createElement('span');
    label.className = 'menu-label';
    label.textContent = ''; //メニュー
    toggle.appendChild(icon);
    // toggle.appendChild(label);
    container.appendChild(toggle);

    const overlay = document.createElement('div');
    overlay.id = 'toolbar-menu-overlay';
    overlay.className = 'app-menu-backdrop';
    overlay.hidden = true;
    overlay.setAttribute('aria-hidden', 'true');

    const panel = document.createElement('aside');
    panel.className = 'app-menu-panel';
    panel.id = 'toolbar-menu-panel';
    panel.tabIndex = -1;
    panel.hidden = true;
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-modal', 'true');
    panel.setAttribute('aria-labelledby', 'toolbar-menu-title');

    const panelHeader = document.createElement('div');
    panelHeader.className = 'app-menu-header';

    const heading = document.createElement('h2');
    heading.id = 'toolbar-menu-title';
    heading.className = 'app-menu-title';
    heading.textContent = 'メニュー';

    const closeButton = document.createElement('button');
    closeButton.type = 'button';
    closeButton.className = 'btn menu-close';
    closeButton.setAttribute('aria-label', 'メニューを閉じる');
    closeButton.textContent = '×';

    panelHeader.appendChild(heading);
    panelHeader.appendChild(closeButton);

    const panelContent = document.createElement('div');
    panelContent.className = 'app-menu-content';

    const viewSection = document.createElement('section');
    viewSection.className = 'app-menu-section';
    const viewTitle = document.createElement('h3');
    viewTitle.textContent = '表示（ビュー）';
    const nav = createNav(currentView);
    viewSection.appendChild(viewTitle);
    viewSection.appendChild(nav);

    const actionSection = document.createElement('section');
    actionSection.className = 'app-menu-section';
    const actionTitle = document.createElement('h3');
    actionTitle.textContent = '操作';
    actionSection.appendChild(actionTitle);
    const actionList = document.createElement('div');
    actionList.className = 'app-menu-list';
    if (validationsButton) {
      validationsButton.classList.add('app-menu-link', 'btn');
      validationsButton.id = 'btn-validations';
      actionList.appendChild(validationsButton);
    }
    actionSection.appendChild(actionList);

    panelContent.appendChild(viewSection);
    panelContent.appendChild(actionSection);

    const footer = document.createElement('div');
    footer.className = 'app-menu-footer';
    footer.textContent = 'Esc または背景クリックで閉じることができます。';

    panel.appendChild(panelHeader);
    panel.appendChild(panelContent);
    panel.appendChild(footer);

    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    return { container, toggle, overlay, panel, closeButton };
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

    if (state.current && typeof state.current.destroy === 'function') {
      state.current.destroy();
    }
    state.current = null;

    const root = document.createElement('div');
    root.className = 'toolbar';

    const validationsButton = createButton('btn-validations', 'btn', options.validationsLabel || '入力規則');
    const menu = createMenuSection(options.currentView, validationsButton);
    root.appendChild(menu.container);

    const title = document.createElement('div');
    title.className = 'title';
    title.textContent = options.title || '';
    root.appendChild(title);

    const due = createDueSection();
    root.appendChild(due.container);

    const addButton = createButton('btn-add', 'btn btn-primary', options.addButtonLabel || '＋ 追加');
    const saveButton = createButton('btn-save', 'btn btn-success', options.saveButtonLabel || 'Excelへ保存');
    const reloadButton = createButton('btn-reload', 'btn', options.reloadLabel || '再読込');

    root.appendChild(addButton);
    root.appendChild(saveButton);
    root.appendChild(reloadButton);

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
      menu,
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

    const setMenuVisibility = open => {
      const shouldOpen = Boolean(open);
      menu.overlay.hidden = !shouldOpen;
      menu.panel.hidden = !shouldOpen;
      menu.overlay.classList.toggle('is-open', shouldOpen);
      menu.overlay.setAttribute('aria-hidden', shouldOpen ? 'false' : 'true');
      menu.toggle.setAttribute('aria-expanded', String(shouldOpen));
      if (shouldOpen) {
        menu.panel.focus();
      } else if (document.activeElement && menu.panel.contains(document.activeElement)) {
        menu.toggle.focus();
      }
    };

    menu.toggle.addEventListener('click', event => {
      event.stopPropagation();
      const willOpen = !menu.overlay.classList.contains('is-open');
      setMenuVisibility(willOpen);
    });

    menu.panel.addEventListener('click', event => {
      const activator = event.target.closest('a, button');
      if (activator) {
        setMenuVisibility(false);
      }
    });

    menu.overlay.addEventListener('click', event => {
      if (event.target === menu.overlay) {
        setMenuVisibility(false);
      }
    });

    menu.closeButton.addEventListener('click', () => {
      setMenuVisibility(false);
    });

    const handleKeydown = event => {
      if (event.key === 'Escape' && menu.overlay.classList.contains('is-open')) {
        setMenuVisibility(false);
      }
    };
    document.addEventListener('keydown', handleKeydown);

    const destroy = () => {
      document.removeEventListener('keydown', handleKeydown);
      if (menu.overlay && menu.overlay.parentNode) {
        menu.overlay.parentNode.removeChild(menu.overlay);
      }
    };
    instance.destroy = destroy;

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
      destroy,
    };
  }

  global.TaskAppHeader = {
    initHeader,
    updateDueSummary(tasks) {
      updateDueSummary(tasks);
    },
  };
}(window));

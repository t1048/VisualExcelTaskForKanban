(function (global) {
  'use strict';

  function getSafeApi(apiAccessor) {
    if (typeof apiAccessor !== 'function') return () => ({});
    return () => {
      try {
        return apiAccessor() || {};
      } catch (err) {
        console.warn('[excelSync] apiAccessor failed:', err);
        return {};
      }
    };
  }

  function createExcelSyncHandlers({ apiAccessor, onAfterValidationSave } = {}) {
    if (typeof apiAccessor !== 'function') {
      throw new Error('createExcelSyncHandlers requires apiAccessor function');
    }

    const resolveApi = getSafeApi(apiAccessor);

    async function handleSaveToExcel() {
      const api = resolveApi();
      if (!api || typeof api.save_excel !== 'function') {
        alert('保存機能が利用できません。');
        return;
      }
      try {
        const result = await api.save_excel();
        const message = result ? `Excelへ保存しました\n${result}` : 'Excelへ保存しました';
        alert(message);
      } catch (err) {
        alert('保存に失敗: ' + (err?.message || err));
      }
    }

    async function handleReloadFromExcel({ onBeforeReload, onAfterReload } = {}) {
      const api = resolveApi();
      if (!api || typeof api.reload_from_excel !== 'function') {
        alert('再読込機能が利用できません。');
        return;
      }
      try {
        if (typeof onBeforeReload === 'function') {
          try {
            onBeforeReload();
          } catch (err) {
            console.warn('[excelSync] onBeforeReload failed:', err);
          }
        }
        const payload = await api.reload_from_excel();
        if (typeof onAfterReload === 'function') {
          await onAfterReload(payload);
        }
      } catch (err) {
        alert('再読込に失敗: ' + (err?.message || err));
      }
    }

    function closeValidationModal() {
      const modal = document.getElementById('validation-modal');
      if (!modal) return;
      modal.classList.remove('open');
      modal.setAttribute('aria-hidden', 'true');
    }

    function openValidationModal({
      columns = [],
      getCurrentValues,
      onAfterRender,
      onBeforeSave,
    } = {}) {
      const modal = document.getElementById('validation-modal');
      const editor = document.getElementById('validation-editor');
      if (!modal || !editor) return;

      const valuesSource = typeof getCurrentValues === 'function' ? getCurrentValues() || {} : {};

      editor.innerHTML = '';

      columns.forEach((column) => {
        const item = document.createElement('div');
        item.className = 'validation-item';

        const label = document.createElement('label');
        const id = 'val-' + btoa(unescape(encodeURIComponent(String(column)))).replace(/=/g, '');
        label.setAttribute('for', id);
        label.textContent = column;

        const textarea = document.createElement('textarea');
        textarea.id = id;
        textarea.dataset.column = column;
        textarea.placeholder = '1 行に 1 候補を入力';
        const existing = valuesSource[column];
        textarea.value = Array.isArray(existing) ? existing.join('\n') : '';
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

      if (typeof onAfterRender === 'function') {
        try {
          onAfterRender({ modal, editor });
        } catch (err) {
          console.warn('[excelSync] onAfterRender failed:', err);
        }
      }

      if (saveBtn) {
        saveBtn.onclick = async () => {
          const payload = {};
          editor.querySelectorAll('textarea[data-column]').forEach((area) => {
            const col = area.dataset.column;
            const lines = area.value.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
            payload[col] = lines;
          });

          try {
            if (typeof onBeforeSave === 'function') {
              try {
                onBeforeSave({ payload });
              } catch (err) {
                console.warn('[excelSync] onBeforeSave failed:', err);
              }
            }

            const api = resolveApi();
            let response;
            if (api && typeof api.update_validations === 'function') {
              response = await api.update_validations(payload);
            }

            let shouldClose = true;
            const closeModal = () => {
              shouldClose = false;
              closeValidationModal();
            };

            if (typeof onAfterValidationSave === 'function') {
              const result = await onAfterValidationSave({ payload, response, closeModal });
              if (result === false) {
                shouldClose = false;
              }
            }

            if (shouldClose) {
              closeValidationModal();
            }
          } catch (err) {
            alert('入力規則の保存に失敗: ' + (err?.message || err));
          }
        };
      }

      modal.classList.add('open');
      modal.setAttribute('aria-hidden', 'false');
    }

    return {
      handleSaveToExcel,
      handleReloadFromExcel,
      openValidationModal,
      closeValidationModal,
    };
  }

  global.TaskExcelSync = {
    createExcelSyncHandlers,
  };
}(window));

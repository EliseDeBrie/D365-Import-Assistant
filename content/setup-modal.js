(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  // Short labels; order matches the real click order on the D365 Import
  // screen. Longer how-to text lives in ROLE_HINTS, shown under the label.
  const ROLE_LABELS = {
    addFileButton: 'Add file button',
    sourceFormatField: 'Source data format — the closed box',
    sourceFormatOption: 'Source data format — one item in the opened list',
    entityNameField: 'Entity name field',
    suggestionItem: 'One row in the entity name suggestions',
    fileTarget: 'File input / drop target',
    uploadButton: 'Upload button',
    entitiesGridRow: 'One row in the uploaded-entities grid',
    runImportButton: 'Page-level Import/Run button'
  };

  const ROLE_HINTS = {
    addFileButton: 'Adds a new row to fill in.',
    sourceFormatField:
      'While the list is still CLOSED: Alt+click anywhere in the box (the value/placeholder text is fine — not specifically the little arrow). Don\'t open it first.',
    sourceFormatOption:
      'Click the box normally (no Alt) to open its list, THEN Alt+click one item inside it, e.g. "Excel". A different element from the box itself. Can be left unbound — the value is then typed into the box instead.',
    entityNameField: 'Pick a format first so this field appears, then Alt+click it.',
    suggestionItem: 'Type a few letters into the entity field to open the list first.',
    fileTarget:
      'Alt+click the visible "Upload data file" box — the hidden input that actually takes the file is found from there automatically.',
    uploadButton:
      'Optional. The file is attached directly, so leave this unbound if the button opens a file-picker dialog ("Browse"). Bind it only if D365 needs an explicit commit click.',
    entitiesGridRow:
      'Add one file manually first so a row exists to click. Used to detect success — falls back to a fixed pause if left unbound.',
    runImportButton: 'Only clicked if auto-run is turned on in Settings; otherwise never touched.'
  };

  const OPTIONAL_ROLES = new Set(['entitiesGridRow', 'runImportButton']);

  let modalEl = null;

  function buildModal() {
    const modal = document.createElement('div');
    modal.id = 'd365ia-setup-modal';
    modal.innerHTML = `
      <div id="d365ia-setup-box">
        <div id="d365ia-setup-header">
          <h3>Bind D365 fields</h3>
          <button id="d365ia-setup-x" title="Close" aria-label="Close">&times;</button>
        </div>
        <p>
          Click <strong>Bind</strong>, use the page normally to reveal the element (click, type,
          open a dropdown), then hold <strong>Alt</strong> and click it to confirm — a plain
          click just reaches the page. <strong>Esc</strong> cancels a bind in progress.
        </p>
        <div id="d365ia-setup-rows"></div>
        <button id="d365ia-setup-close">Done</button>
      </div>
    `;
    document.body.appendChild(modal);

    const rows = modal.querySelector('#d365ia-setup-rows');
    let requiredCount = 0;
    let optionalHeaderAdded = false;

    D365IA.binder.ROLES.forEach((role) => {
      if (OPTIONAL_ROLES.has(role) && !optionalHeaderAdded) {
        const divider = document.createElement('div');
        divider.className = 'd365ia-setup-divider';
        divider.textContent = 'Optional';
        rows.appendChild(divider);
        optionalHeaderAdded = true;
      }

      const row = document.createElement('div');
      row.className = 'd365ia-setup-row';

      const top = document.createElement('div');
      top.className = 'd365ia-setup-row-top';

      const badge = document.createElement('span');
      if (OPTIONAL_ROLES.has(role)) {
        badge.className = 'd365ia-setup-badge d365ia-setup-badge-optional';
        badge.textContent = 'opt';
      } else {
        requiredCount++;
        badge.className = 'd365ia-setup-badge';
        badge.textContent = String(requiredCount);
      }

      const labelWrap = document.createElement('div');
      labelWrap.className = 'd365ia-setup-label-wrap';
      const label = document.createElement('div');
      label.className = 'd365ia-setup-label';
      label.textContent = ROLE_LABELS[role] || role;
      const hint = document.createElement('div');
      hint.className = 'd365ia-setup-hint';
      hint.textContent = ROLE_HINTS[role] || '';
      labelWrap.appendChild(label);
      labelWrap.appendChild(hint);

      const btnGroup = document.createElement('div');
      btnGroup.className = 'd365ia-setup-btn-group';

      const btn = document.createElement('button');
      btn.className = 'd365ia-setup-bind';
      btn.dataset.role = role;
      btn.textContent = 'Bind';

      const clearBtn = document.createElement('button');
      clearBtn.className = 'd365ia-setup-clear';
      clearBtn.dataset.role = role;
      clearBtn.textContent = 'Clear';
      clearBtn.title = 'Unbind this field without touching the others';

      btnGroup.appendChild(btn);
      btnGroup.appendChild(clearBtn);

      top.appendChild(badge);
      top.appendChild(labelWrap);
      top.appendChild(btnGroup);

      const value = document.createElement('div');
      value.className = 'd365ia-setup-value';
      value.id = `d365ia-val-${role}`;
      value.textContent = 'not bound';

      const match = document.createElement('span');
      match.className = 'd365ia-setup-match';
      match.id = `d365ia-match-${role}`;

      row.appendChild(top);
      row.appendChild(value);
      row.appendChild(match);
      rows.appendChild(row);
    });

    rows.querySelectorAll('.d365ia-setup-bind').forEach((btn) => {
      btn.addEventListener('click', () => {
        const prevText = btn.textContent;
        modal.style.display = 'none';
        btn.textContent = 'Alt+click the target...';
        D365IA.binder.startPicking(
          btn.dataset.role,
          async ({ role, selector, isFileInput, isSelect }) => {
            await D365IA.binder.saveBinding(role, selector, { isFileInput, isSelect });
            modal.style.display = 'flex';
            const valueEl = document.getElementById(`d365ia-val-${role}`);
            valueEl.textContent = selector;
            valueEl.title = selector;
            btn.textContent = 'Rebind';
          },
          () => {
            modal.style.display = 'flex';
            btn.textContent = prevText;
          }
        );
      });
    });

    // Unbinds just this one role — for when a bind was misclicked onto the
    // wrong element and shouldn't take the other, correct bindings with it.
    rows.querySelectorAll('.d365ia-setup-clear').forEach((clearBtn) => {
      clearBtn.addEventListener('click', async () => {
        const role = clearBtn.dataset.role;
        await D365IA.binder.clearBinding(role);
        const valueEl = document.getElementById(`d365ia-val-${role}`);
        valueEl.textContent = 'not bound';
        valueEl.title = '';
        const bindBtn = rows.querySelector(`.d365ia-setup-bind[data-role="${role}"]`);
        if (bindBtn) bindBtn.textContent = 'Bind';
      });
    });

    modal.querySelector('#d365ia-setup-close').addEventListener('click', close);
    modal.querySelector('#d365ia-setup-x').addEventListener('click', close);
    modal.addEventListener('click', (e) => {
      if (e.target === modal) close();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && modal.style.display !== 'none') close();
    });

    return modal;
  }

  // Reports what a saved selector actually matches on the page right now.
  // A selector matching nothing (stale) or matching dozens of elements
  // (mis-picked onto something generic) is the usual cause of a run failing,
  // and is otherwise invisible until it errors mid-import.
  function describeMatches(selector, role) {
    let count;
    try {
      count = document.querySelectorAll(selector).length;
    } catch (e) {
      return { text: 'invalid selector', level: 'bad' };
    }
    if (count === 0) return { text: '0 on page now', level: 'warn' };
    if (D365IA.binder.LIST_ROLES.has(role)) {
      return count > 25
        ? { text: `${count} matches — too generic?`, level: 'bad' }
        : { text: `${count} matches`, level: 'ok' };
    }
    return count === 1
      ? { text: '1 match', level: 'ok' }
      : { text: `${count} matches — ambiguous`, level: 'warn' };
  }

  async function refreshValues() {
    const bindings = await D365IA.binder.getBindings();
    D365IA.binder.ROLES.forEach((role) => {
      const el = document.getElementById(`d365ia-val-${role}`);
      if (!el) return;
      const selector = bindings[role] && bindings[role].selector;
      el.textContent = selector || 'not bound';
      el.title = selector || '';

      const status = document.getElementById(`d365ia-match-${role}`);
      if (!status) return;
      if (!selector) {
        status.textContent = '';
        status.className = 'd365ia-setup-match';
        return;
      }
      const match = describeMatches(selector, role);
      status.textContent = match.text;
      status.className = `d365ia-setup-match d365ia-match-${match.level}`;
    });
  }

  async function open() {
    if (!modalEl) modalEl = buildModal();
    modalEl.style.display = 'flex';
    await refreshValues();
  }

  function close() {
    if (modalEl) modalEl.style.display = 'none';
  }

  D365IA.setupModal = { open, close };
})();

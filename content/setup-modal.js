(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  // Order matches the real click order on the D365 Import screen so the
  // list itself is a walkthrough.
  const ROLE_LABELS = {
    addFileButton: '1. "Add file" button (adds a new row to fill in)',
    sourceFormatField: '2. "Source data format" dropdown for that row',
    sourceFormatOption:
      '3. One option in that dropdown’s open list — e.g. click it open, Alt+click "Excel"',
    entityNameField: '4. Entity name field (appears after picking Excel/CSV)',
    suggestionItem: '5. One row in the entity name autocomplete suggestions — type a few letters first to open it',
    fileTarget: '6. File input / drop target for the row',
    uploadButton: '7. "Upload" button for the row',
    entitiesGridRow:
      'Optional: one row in the grid of already-uploaded entities — used to detect success. Add a file manually first so a row exists to click.',
    runImportButton:
      'Optional: the page-level "Import"/Run button — only used if you turn on auto-run in Settings; never clicked otherwise.'
  };

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
          Click "Bind", then use the page normally (click a field, type into it, open its
          dropdown, etc) until the element you want is visible. Hold <strong>Alt</strong> and
          click it to bind — a plain click won't finalize anything. Press Esc to cancel a bind in
          progress, or click outside this box / the &times; to close.
        </p>
        <div id="d365ia-setup-rows"></div>
        <button id="d365ia-setup-close">Done</button>
      </div>
    `;
    document.body.appendChild(modal);

    const rows = modal.querySelector('#d365ia-setup-rows');
    D365IA.binder.ROLES.forEach((role) => {
      const row = document.createElement('div');
      row.className = 'd365ia-setup-row';

      const label = document.createElement('span');
      label.className = 'd365ia-setup-label';
      label.textContent = ROLE_LABELS[role] || role;

      const value = document.createElement('span');
      value.className = 'd365ia-setup-value';
      value.id = `d365ia-val-${role}`;
      value.textContent = 'not bound';

      const btn = document.createElement('button');
      btn.className = 'd365ia-setup-bind';
      btn.dataset.role = role;
      btn.textContent = 'Bind';

      row.appendChild(label);
      row.appendChild(value);
      row.appendChild(btn);
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
            document.getElementById(`d365ia-val-${role}`).textContent = selector;
            btn.textContent = 'Rebind';
          },
          () => {
            modal.style.display = 'flex';
            btn.textContent = prevText;
          }
        );
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

  async function refreshValues() {
    const bindings = await D365IA.binder.getBindings();
    D365IA.binder.ROLES.forEach((role) => {
      const el = document.getElementById(`d365ia-val-${role}`);
      if (el) el.textContent = (bindings[role] && bindings[role].selector) || 'not bound';
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

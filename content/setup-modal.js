(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  const ROLE_LABELS = {
    entityNameField: 'Entity name field (the box you type into)',
    suggestionItem:
      'One suggestion row in the autocomplete dropdown (open the dropdown first, then click a suggestion)',
    fileTarget: 'File input or drop target for the row',
    addRowButton: 'Add new row / new line button (optional)',
    importButton: 'Import / Submit button (optional — only used if you enable auto-submit)'
  };

  let modalEl = null;

  function buildModal() {
    const modal = document.createElement('div');
    modal.id = 'd365ia-setup-modal';
    modal.innerHTML = `
      <div id="d365ia-setup-box">
        <h3>Bind D365 fields</h3>
        <p>Click "Bind" then click the matching element on the D365 page. This is saved per browser and only needs doing once per environment.</p>
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
        modal.style.display = 'none';
        btn.textContent = 'Click target on page...';
        D365IA.binder.startPicking(btn.dataset.role, async ({ role, selector, isFileInput }) => {
          await D365IA.binder.saveBinding(role, selector, { isFileInput });
          modal.style.display = 'flex';
          document.getElementById(`d365ia-val-${role}`).textContent = selector;
          btn.textContent = 'Rebind';
        });
      });
    });

    modal.querySelector('#d365ia-setup-close').addEventListener('click', close);
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

const DEFAULT_UI = {
  showLauncher: true,
  restrictToPages: true,
  urlPatterns: 'mi=DM_DataManagementWorkspaceMenuItem'
};

async function readSettings() {
  const data = await chrome.storage.sync.get('settings');
  const stored = data.settings || {};
  return { stored, ui: Object.assign({}, DEFAULT_UI, stored.ui || {}) };
}

function describeWhereItShows(ui) {
  if (!ui.showLauncher) return 'Currently hidden everywhere.';
  if (!ui.restrictToPages) return 'Currently shows on every D365 page.';
  const patterns = String(ui.urlPatterns || '')
    .split(/[\n,]/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (patterns.length === 0) return 'Currently shows on every D365 page.';
  return `Only on pages matching: ${patterns.join(', ')}`;
}

document.addEventListener('DOMContentLoaded', async () => {
  const { stored, ui } = await readSettings();

  const toggle = document.getElementById('showLauncher');
  const pageHint = document.getElementById('pageHint');
  toggle.checked = !!ui.showLauncher;
  pageHint.textContent = describeWhereItShows(ui);

  toggle.addEventListener('change', async () => {
    const next = Object.assign({}, ui, { showLauncher: toggle.checked });
    // Content scripts watch storage, so open tabs react without a reload.
    await chrome.storage.sync.set({ settings: Object.assign({}, stored, { ui: next }) });
    pageHint.textContent = describeWhereItShows(next);
  });

  const bindingData = await chrome.storage.sync.get('bindings');
  const allBindings = bindingData.bindings || {};
  // Bindings are stored per environment, so count whichever host has the most.
  const perHost = Object.values(allBindings).filter((v) => v && typeof v === 'object');
  const bound = perHost.some(
    (hostBindings) =>
      hostBindings.entityNameField &&
      hostBindings.entityNameField.selector &&
      hostBindings.fileTarget &&
      hostBindings.fileTarget.selector
  );

  const statusEl = document.getElementById('status');
  if (bound) {
    statusEl.textContent = 'Fields bound — ready to use.';
    statusEl.className = 'ready';
  } else {
    statusEl.textContent =
      'Using built-in defaults. If a field fails to resolve, rebind it from Setup fields on the import page.';
    statusEl.className = 'not-ready';
  }

  document.getElementById('open-options').addEventListener('click', () => {
    chrome.runtime.openOptionsPage();
  });
});

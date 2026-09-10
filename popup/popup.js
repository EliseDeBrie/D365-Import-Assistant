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

  await showCurrentEnvironment();

  document.getElementById('open-options').addEventListener('click', () => {
    chrome.runtime.openOptionsPage();
  });
});

// Confirms, right here, that the extension is live on whichever D365
// environment the active tab happens to be on right now -- no URL to type
// in, no per-environment setup. host_permissions is a wildcard
// (https://*.dynamics.com/*), so any sandbox or production tenant, including one
// created after this extension was installed, works the moment you're on it.
async function showCurrentEnvironment() {
  const envEl = document.getElementById('envStatus');
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const host = tab && tab.url && new URL(tab.url).hostname;
    if (host && /\.dynamics\.com$/i.test(host)) {
      envEl.textContent = `Active on: ${host}`;
      envEl.className = 'env-line env-active';
    } else {
      envEl.textContent = 'Not on a D365 page in this tab right now.';
      envEl.className = 'env-line env-inactive';
    }
  } catch (e) {
    // tab.url is only populated when host_permissions covers it, so a
    // non-dynamics.com tab lands here too -- same message either way.
    envEl.textContent = 'Not on a D365 page in this tab right now.';
    envEl.className = 'env-line env-inactive';
  }
}

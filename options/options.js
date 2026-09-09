const CLEANING_FIELDS = [
  'stripLeadingNumbers',
  'stripTrailingNumbers',
  'stripDates',
  'stripVersionSuffixes',
  'separatorsToSpaces',
  'titleCase'
];

const DEFAULT_UI = {
  showLauncher: true,
  restrictToPages: true,
  urlPatterns: 'mi=DM_DataManagementWorkspaceMenuItem'
};

async function load() {
  const data = await chrome.storage.sync.get(['settings', 'bindings']);
  const settings = data.settings || {};
  const rules = Object.assign({}, window.D365IA.matcher.DEFAULT_RULES, settings.rules || {});
  const options = Object.assign(
    { matchThreshold: 0.75, stepDelay: 700, elementTimeout: 5000, uploadTimeout: 300000 },
    settings.options || {}
  );
  const ui = Object.assign({}, DEFAULT_UI, settings.ui || {});

  document.getElementById('showLauncher').checked = !!ui.showLauncher;
  document.getElementById('restrictToPages').checked = !!ui.restrictToPages;
  document.getElementById('urlPatterns').value = ui.urlPatterns;

  CLEANING_FIELDS.forEach((f) => {
    document.getElementById(f).checked = !!rules[f];
  });
  document.getElementById('matchThreshold').value = options.matchThreshold;
  document.getElementById('stepDelay').value = options.stepDelay;
  document.getElementById('elementTimeout').value = options.elementTimeout;
  document.getElementById('uploadTimeout').value = options.uploadTimeout;

  renderBindings(data.bindings || {});
  runTest();
}

function renderBindings(bindings) {
  const list = document.getElementById('bindingsList');
  list.innerHTML = '';
  const roles = Object.keys(bindings);
  if (roles.length === 0) {
    list.innerHTML = '<li>No bindings yet.</li>';
    return;
  }
  roles.forEach((role) => {
    const li = document.createElement('li');
    li.textContent = `${role}: ${bindings[role].selector}`;
    list.appendChild(li);
  });
}

function currentRules() {
  const rules = {};
  CLEANING_FIELDS.forEach((f) => (rules[f] = document.getElementById(f).checked));
  rules.stripExtension = true;
  rules.collapseWhitespace = true;
  return rules;
}

function runTest() {
  const val = document.getElementById('testInput').value || 'Sample';
  document.getElementById('testOutput').textContent = window.D365IA.matcher.cleanFileName(
    val,
    currentRules()
  );
}

CLEANING_FIELDS.forEach((f) => document.getElementById(f).addEventListener('change', runTest));
document.getElementById('testInput').addEventListener('input', runTest);

document.getElementById('save').addEventListener('click', async () => {
  const options = {
    matchThreshold: parseFloat(document.getElementById('matchThreshold').value) || 0.75,
    stepDelay: parseInt(document.getElementById('stepDelay').value, 10) || 700,
    elementTimeout: parseInt(document.getElementById('elementTimeout').value, 10) || 5000,
    uploadTimeout: parseInt(document.getElementById('uploadTimeout').value, 10) || 300000
  };

  const ui = {
    showLauncher: document.getElementById('showLauncher').checked,
    restrictToPages: document.getElementById('restrictToPages').checked,
    urlPatterns: document.getElementById('urlPatterns').value
  };

  await chrome.storage.sync.set({ settings: { rules: currentRules(), options, ui } });
  const status = document.getElementById('saveStatus');
  status.textContent = 'Saved.';
  setTimeout(() => (status.textContent = ''), 1500);
});

document.getElementById('clearBindings').addEventListener('click', async () => {
  await chrome.storage.sync.set({ bindings: {} });
  load();
});

document.addEventListener('DOMContentLoaded', load);

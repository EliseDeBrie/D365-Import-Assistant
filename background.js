const DEFAULT_SETTINGS = {
  rules: {
    stripExtension: true,
    stripLeadingNumbers: true,
    stripTrailingNumbers: true,
    stripDates: true,
    stripVersionSuffixes: false,
    separatorsToSpaces: true,
    collapseWhitespace: true,
    titleCase: true
  },
  options: {
    matchThreshold: 0.75,
    stepDelay: 700,
    elementTimeout: 5000,
    uploadTimeout: 60000
  }
};

chrome.runtime.onInstalled.addListener(async (details) => {
  if (details.reason !== 'install') return;
  const data = await chrome.storage.sync.get('settings');
  if (!data.settings) {
    await chrome.storage.sync.set({ settings: DEFAULT_SETTINGS, bindings: {} });
  }
});

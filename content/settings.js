(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  const DEFAULT_OPTIONS = {
    matchThreshold: 0.75,
    stepDelay: 700,
    elementTimeout: 5000,
    uploadTimeout: 300000
  };

  const DEFAULT_UI = {
    showLauncher: true,
    restrictToPages: true,
    urlPatterns: 'mi=DM_DataManagementWorkspaceMenuItem'
  };

  function asObject(value) {
    return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
  }

  function clamp(value, min, max, fallback) {
    if (typeof value === 'string' && value.trim() === '') return fallback;
    const number = Number(value);
    if (!Number.isFinite(number)) return fallback;
    return Math.min(Math.max(number, min), max);
  }

  function normalize(stored, defaultRules) {
    const source = asObject(stored);
    const storedRules = asObject(source.rules);
    const rules = Object.assign({}, defaultRules || {});
    Object.keys(rules).forEach((key) => {
      if (typeof storedRules[key] === 'boolean') rules[key] = storedRules[key];
    });

    const options = asObject(source.options);
    const ui = asObject(source.ui);
    return {
      rules,
      options: {
        matchThreshold: clamp(options.matchThreshold, 0.1, 1, DEFAULT_OPTIONS.matchThreshold),
        stepDelay: clamp(options.stepDelay, 100, 600000, DEFAULT_OPTIONS.stepDelay),
        elementTimeout: clamp(options.elementTimeout, 250, 600000, DEFAULT_OPTIONS.elementTimeout),
        uploadTimeout: clamp(options.uploadTimeout, 5000, 1800000, DEFAULT_OPTIONS.uploadTimeout)
      },
      ui: {
        showLauncher:
          typeof ui.showLauncher === 'boolean' ? ui.showLauncher : DEFAULT_UI.showLauncher,
        restrictToPages:
          typeof ui.restrictToPages === 'boolean' ? ui.restrictToPages : DEFAULT_UI.restrictToPages,
        urlPatterns:
          typeof ui.urlPatterns === 'string'
            ? ui.urlPatterns.slice(0, 2048)
            : DEFAULT_UI.urlPatterns
      }
    };
  }

  D365IA.settings = { normalize, DEFAULT_OPTIONS, DEFAULT_UI };
})();

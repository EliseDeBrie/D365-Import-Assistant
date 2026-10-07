const { createPage } = require('./harness');

const { D365IA } = createPage();
const { settings, matcher } = D365IA;

test('normalizes settings to the same safe values for every surface', () => {
  const normalized = settings.normalize(
    {
      rules: { stripDates: false, titleCase: 'no' },
      options: {
        matchThreshold: 4,
        stepDelay: -100,
        elementTimeout: 'bad',
        uploadTimeout: 1
      },
      ui: { showLauncher: false, restrictToPages: 'no', urlPatterns: 'x'.repeat(3000) }
    },
    matcher.DEFAULT_RULES
  );

  equal(normalized.rules.stripDates, false);
  equal(normalized.rules.titleCase, matcher.DEFAULT_RULES.titleCase);
  equal(normalized.options, {
    matchThreshold: 1,
    stepDelay: 100,
    elementTimeout: 5000,
    uploadTimeout: 5000
  });
  equal(normalized.ui.showLauncher, false);
  equal(normalized.ui.restrictToPages, true);
  equal(normalized.ui.urlPatterns.length, 2048);
});

test('preserves shipped UI defaults for omitted settings', () => {
  const normalized = settings.normalize({}, matcher.DEFAULT_RULES);
  equal(normalized.ui.showLauncher, true);
  equal(normalized.ui.restrictToPages, true);
  equal(normalized.ui.urlPatterns, 'mi=DM_DataManagementWorkspaceMenuItem');
});

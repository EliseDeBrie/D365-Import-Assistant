const { createPage } = require('./harness');

const HOST_A = 'https://contoso.operations.dynamics.com/?mi=DM_DataManagementWorkspaceMenuItem';
const HOST_B = 'https://fabrikam.sandbox.operations.dynamics.com/?mi=DM_DataManagementWorkspaceMenuItem';

function stubOData(window, names) {
  window.fetch = async () => ({
    ok: true,
    status: 200,
    json: async () => ({ value: names.map((name) => ({ name })) })
  });
}

test('caches the entity list under the environment hostname', async () => {
  const { window, D365IA } = createPage({ url: HOST_A });
  stubOData(window, ['CustomersV3', 'ContosoSecretEntity']);
  await D365IA.entityList.refresh();

  const stored = window.chrome.storage.local.data.entityList;
  equal(Object.keys(stored), ['contoso.operations.dynamics.com']);
  equal(stored['contoso.operations.dynamics.com'].names, ['ContosoSecretEntity', 'CustomersV3']);
});

// A consultant's browser profile holds lists for several customers; one
// tenant's custom entity names must never surface as hints on another's.
test('does not surface another environment\'s entity list', async () => {
  const a = createPage({ url: HOST_A });
  stubOData(a.window, ['ContosoSecretEntity']);
  await a.D365IA.entityList.refresh();

  const b = createPage({ url: HOST_B });
  Object.assign(b.window.chrome.storage.local.data, a.window.chrome.storage.local.data);
  const { count } = await b.D365IA.entityList.load();
  equal(count, 0);
  equal(b.D365IA.entityList.validate('Contoso Secret Entity').status, 'unknown');

  stubOData(b.window, ['FabrikamEntity']);
  await b.D365IA.entityList.refresh();
  const stored = b.window.chrome.storage.local.data.entityList;
  equal(Object.keys(stored).sort(), [
    'contoso.operations.dynamics.com',
    'fabrikam.sandbox.operations.dynamics.com'
  ]);
});

// Earlier versions stored one flat list for every environment. There's no
// way to tell which tenant it came from, so it is dropped, not migrated.
test('drops the legacy single flat list', async () => {
  const { window, D365IA } = createPage({ url: HOST_A });
  window.chrome.storage.local.data.entityList = { names: ['Legacy'], fetchedAt: 1 };
  const { count } = await D365IA.entityList.load();
  equal(count, 0);

  stubOData(window, ['Fresh']);
  await D365IA.entityList.refresh();
  const stored = window.chrome.storage.local.data.entityList;
  assert(!('names' in stored), 'legacy flat list was carried along');
});

test('forgets cached validations when the list is refreshed', async () => {
  const { window, D365IA } = createPage({ url: HOST_A });
  stubOData(window, ['CustomersV3']);
  await D365IA.entityList.refresh();
  equal(D365IA.entityList.validate('Vendors').status, 'none');

  stubOData(window, ['Vendors']);
  await D365IA.entityList.refresh();
  equal(D365IA.entityList.validate('Vendors').status, 'match');
});

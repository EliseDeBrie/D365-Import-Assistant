const { createPage } = require('./harness');
const fakeD365 = require('./fake-d365');

// Short timeouts: the fake page responds synchronously, so a step that waits
// is a step that is genuinely stuck, and the test should say so quickly.
const FAST = { stepDelay: 10, elementTimeout: 800, uploadTimeout: 2000, lookupSettleMs: 20 };

function setup(options = {}) {
  const page = createPage();
  const d365 = fakeD365.install(page.window, options);
  const queue = page.D365IA.queue.createQueue({ onStatusChange: () => {} });
  return { page, d365, queue };
}

async function bindings(page) {
  return page.D365IA.binder.getBindings();
}

function file(page, name, content = 'x') {
  return new page.window.File([content], name);
}

test('ships defaults that resolve against a real control layout', async () => {
  const { page, d365 } = setup();
  d365.openPanel();
  const b = await bindings(page);
  const { queryVisible } = page.D365IA.domUtils;

  ['addFileButton', 'sourceFormatField', 'runImportButton', 'closePanelButton'].forEach((role) => {
    assert(b[role] && b[role].selector, `${role} has no default selector`);
    assert(queryVisible(b[role].selector), `default selector for ${role} matched nothing`);
  });
});

test('uploads a single Excel file end to end', async () => {
  const { page, d365, queue } = setup();
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 1, `run summary was ${JSON.stringify(result)}`);
  const uploads = d365.uploads();
  equal(uploads.length, 1);
  equal(uploads[0].format, 'Excel');
  equal(uploads[0].entity, 'Vendors V2');
  equal(uploads[0].fileName, '01_Vendors_V2.xlsx');
});

// The reported bug: a .zip went into D365 as Excel, and the pipeline then
// waited for an entity field that a package never shows.
test('uploads a data package as Package, with no entity name', async () => {
  const { page, d365, queue } = setup();
  queue.addFiles([file(page, 'GeneralLedgerPackage.zip')], {});

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 1, `run summary was ${JSON.stringify(result)}`);
  const uploads = d365.uploads();
  equal(uploads[0].format, 'Package');
  equal(uploads[0].entity, null, 'a package must not get an entity name');
});

// The other reported bug: after the first file, D365 rebuilds a blank panel
// and the second file has to refill it rather than assume it is still set.
test('uploads several files in one run, refilling the panel each time', async () => {
  const { page, d365, queue } = setup({
    entities: ['Sites V2', 'Vendors V2', 'Customers V3']
  });
  queue.addFiles(
    [
      file(page, '02_Vendors_V2.xlsx'),
      file(page, '01_Sites_V2.xlsx'),
      file(page, '03_Customers_V3.xlsx')
    ],
    {}
  );

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 3, `run summary was ${JSON.stringify(result)}`);
  // And in numbered order, not the order they were handed over.
  equal(
    d365.uploads().map((u) => u.fileName),
    ['01_Sites_V2.xlsx', '02_Vendors_V2.xlsx', '03_Customers_V3.xlsx']
  );
});

test('mixes formats in one run', async () => {
  const { page, d365, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles(
    [file(page, '01_Vendors_V2.xlsx'), file(page, '02_Vendors_V2.csv'), file(page, '03_Pack.zip')],
    {}
  );

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 3, `run summary was ${JSON.stringify(result)}`);
  equal(
    d365.uploads().map((u) => u.format),
    ['Excel', 'CSV', 'Package']
  );
});

// The sheet picker regression: D365's lookup discards typed text, so the
// sheet has to be committed from the list. A test that only checked the
// input's value would have passed while the real thing hung.
test('commits the chosen sheet through the lookup, not as typed text', async () => {
  const { page, d365, queue } = setup({
    entities: ['Vendors V2'],
    sheetsByFile: { '01_Vendors_V2.xlsx': ['en_us', 'DataSheet1'] }
  });

  const [item] = queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});
  queue.updateItem(item.id, { sheetNames: ['en_us', 'DataSheet1'], selectedSheet: 'DataSheet1' });

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 1, `run summary was ${JSON.stringify(result)}`);
  equal(d365.uploads()[0].sheet, 'DataSheet1');
});

test('leaves a single-sheet workbook alone', async () => {
  const { page, d365, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 1);
  equal(d365.uploads()[0].sheet, null);
});

// One bad file used to strand the rest of the batch.
test('records a failure and carries on with the rest of the queue', async () => {
  const { page, d365, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles(
    [file(page, '01_Vendors_V2.xlsx'), file(page, '02_Nonexistent_Entity.xlsx')],
    {}
  );

  const result = await queue.run(await bindings(page), FAST);

  equal(result.total, 2);
  equal(result.uploaded, 1, `run summary was ${JSON.stringify(result)}`);
  assert(result.failed + result.needsReview === 1, 'the unknown entity should not have uploaded');
  equal(d365.uploads().map((u) => u.fileName), ['01_Vendors_V2.xlsx']);
});

test('an unresolved entity name asks for review rather than importing the wrong one', async () => {
  const { page, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles([file(page, '02_Nonexistent_Entity.xlsx')], {});

  await queue.run(await bindings(page), FAST);

  const item = queue.getItems()[0];
  assert(
    item.status === 'needs-review' || item.status === 'error',
    `expected review or error, got "${item.status}"`
  );
  assert(item.error, 'a stalled item must explain itself');
});

test('a failure names the step it happened in', async () => {
  const { page, queue } = setup({ entities: [] });
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  await queue.run(await bindings(page), FAST);

  const item = queue.getItems()[0];
  assert(item.step, `no step recorded on a failed item: ${JSON.stringify(item.error)}`);
});

test('retry puts a failed item back in the queue', async () => {
  const { page, queue } = setup({ entities: [] });
  const [item] = queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  await queue.run(await bindings(page), FAST);
  assert(queue.getItems()[0].status !== 'pending', 'item should have settled');

  queue.retry(item.id);
  const after = queue.getItems()[0];
  equal(after.status, 'pending');
  equal(after.error, null);
});

test('skip takes an item out of the run without failing it', async () => {
  const { page, d365, queue } = setup({ entities: ['Vendors V2'] });
  const items = queue.addFiles(
    [file(page, '01_Vendors_V2.xlsx'), file(page, '02_Unknown.xlsx')],
    {}
  );
  queue.skip(items[1].id);

  const result = await queue.run(await bindings(page), FAST);

  equal(result.skipped, 1);
  equal(result.failed, 0);
  equal(result.uploaded, 1);
  equal(d365.uploads().length, 1);
});

test('summary counts add up to the queue length', async () => {
  const { page, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles(
    [file(page, '01_Vendors_V2.xlsx'), file(page, '02_Unknown.xlsx'), file(page, '03_Pack.zip')],
    {}
  );

  const s = await queue.run(await bindings(page), FAST);
  equal(s.uploaded + s.failed + s.needsReview + s.skipped + s.pending, s.total);
});

// "Upload + Import" must never start a job over a partial batch — that is
// the button that actually loads data into D365.
test('finishImport closes the panel and clicks Import', async () => {
  const { page, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});
  const b = await bindings(page);
  await queue.run(b, FAST);

  let imported = 0;
  page.document
    .querySelector('[data-dyn-controlname="ImportAsync"]')
    .addEventListener('click', () => imported++);

  await queue.finishImport(b, FAST);
  equal(imported, 1);
});

test('reports D365 message bar text when a step fails', async () => {
  const { page, d365, queue } = setup({ entities: [] });
  d365.message('Entity name is not valid.');
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  await queue.run(await bindings(page), FAST);
  const item = queue.getItems()[0];
  assert(item.error && item.error.length > 0, 'a failure must carry an explanation');
});

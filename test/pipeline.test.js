const { createPage } = require('./harness');
const fakeD365 = require('./fake-d365');

// Short timeouts: the fake page responds synchronously, so a step that waits
// is a step that is genuinely stuck, and the test should say so quickly.
//
// uploadSignalTimeout matters most here. The fake makes no /fileUpload
// request, so no completion event ever arrives -- exactly the situation that
// stalled real batches part-way through. With the wait bounded, the run
// carries on; with it unbounded, this suite hangs, which is the point.
const FAST = {
  stepDelay: 10,
  elementTimeout: 800,
  uploadTimeout: 2000,
  uploadSignalTimeout: 150,
  itemTimeout: 30000,
  lookupSettleMs: 20
};

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
  // Committed in the driver's own form. The bare name matches nothing in
  // D365's list, which is why the sheet step used to stall here.
  equal(d365.uploads()[0].sheet, 'DataSheet1$');
});

test('adds the ODBC $ marker to a sheet name that lacks one', () => {
  const { sheetLookupText } = createPage().D365IA.queue;
  equal(sheetLookupText('en_us'), 'en_us$');
  equal(sheetLookupText('DataSheet1'), 'DataSheet1$');
});

test('does not double up a $ the sheet name already has', () => {
  const { sheetLookupText } = createPage().D365IA.queue;
  equal(sheetLookupText('en_us$'), 'en_us$');
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

// The reported regression: a batch that "stops the actions after 2 files".
// The cause was a wait for an upload-completion signal that never arrives in
// some environments, timing out per file, plus a run loop that could die and
// leave the queue permanently busy.
test('runs a batch of six without stalling part-way', async () => {
  const { page, d365, queue } = setup({
    entities: ['Sites V2', 'Vendors V2', 'Customers V3', 'Products V2', 'Terms V1', 'Groups V1']
  });
  queue.addFiles(
    [
      file(page, '01_Sites_V2.xlsx'),
      file(page, '02_Vendors_V2.xlsx'),
      file(page, '03_Customers_V3.xlsx'),
      file(page, '04_Products_V2.xlsx'),
      file(page, '05_Terms_V1.xlsx'),
      file(page, '06_Groups_V1.xlsx')
    ],
    {}
  );

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 6, `run summary was ${JSON.stringify(result)}`);
  equal(result.aborted, null);
  equal(d365.uploads().length, 6);
});

// The page never sends the upload signal here, so the first file establishes
// that and the rest must not each pay the timeout again. Across a 37-file
// batch that difference is the better part of an afternoon.
test('stops waiting for the upload signal once it proves absent', async () => {
  const { page, queue } = setup({ entities: ['Sites V2', 'Vendors V2', 'Customers V3'] });

  const waits = [];
  const { domUtils } = page.D365IA;
  const realWait = domUtils.waitForUploadResponse;
  domUtils.waitForUploadResponse = function (timeout, ownerDocument) {
    waits.push(timeout);
    return realWait.call(this, timeout, ownerDocument);
  };

  queue.addFiles(
    [
      file(page, '01_Sites_V2.xlsx'),
      file(page, '02_Vendors_V2.xlsx'),
      file(page, '03_Customers_V3.xlsx')
    ],
    {}
  );

  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 3, `run summary was ${JSON.stringify(result)}`);
  // The first file waits; once the signal is known to be absent, the rest
  // skip the wait entirely (a zero timeout is never even armed).
  equal(waits[0], FAST.uploadSignalTimeout, 'the first file should wait for the signal');
  equal(
    waits.slice(1),
    [],
    'later files must not wait again for a signal already known to be absent'
  );
});

// A render error used to be able to reach back through the status callback
// and end the batch.
test('a failing panel render does not stop the run', async () => {
  // The queue logs the render failure; that is the intended behaviour, so
  // keep it out of the test output.
  const realError = console.error;
  console.error = () => {};
  const page = createPage();
  const d365 = fakeD365.install(page.window, { entities: ['Sites V2', 'Vendors V2'] });
  const queue = page.D365IA.queue.createQueue({
    onStatusChange: () => {
      throw new Error('render blew up');
    }
  });
  queue.addFiles([file(page, '01_Sites_V2.xlsx'), file(page, '02_Vendors_V2.xlsx')], {});

  const result = await queue.run(await bindings(page), FAST);
  console.error = realError;

  equal(result.uploaded, 2, `run summary was ${JSON.stringify(result)}`);
  equal(d365.uploads().length, 2);
});

// If a run ends badly, the queue has to be usable again — it used to stay
// flagged as running, so every later Upload press silently did nothing.
test('the queue accepts a new run after the previous one ends', async () => {
  const { page, d365, queue } = setup({ entities: ['Vendors V2'] });

  queue.addFiles([file(page, '01_Unknown_Entity.xlsx')], {});
  await queue.run(await bindings(page), FAST);
  assert(queue.isRunning() === false, 'queue stayed flagged as running');

  queue.addFiles([file(page, '02_Vendors_V2.xlsx')], {});
  const result = await queue.run(await bindings(page), FAST);

  equal(result.uploaded, 1, `run summary was ${JSON.stringify(result)}`);
  equal(d365.uploads().map((u) => u.fileName), ['02_Vendors_V2.xlsx']);
});

// D365 stops and asks when two workbooks in one project share a sheet name,
// which happens on every file when they all carry an "en_us" sheet. The
// dialog is modal and defaults to No, so an unanswered one drops the file
// and stalls everything behind it.
test('answers the duplicate-sheet-name confirmation and keeps going', async () => {
  const { page, d365, queue } = setup({
    entities: ['Sites V2', 'Vendors V2'],
    sheetsByFile: {
      '01_Sites_V2.xlsx': ['en_us', 'Data'],
      '02_Vendors_V2.xlsx': ['en_us', 'Data']
    }
  });

  const items = queue.addFiles(
    [file(page, '01_Sites_V2.xlsx'), file(page, '02_Vendors_V2.xlsx')],
    {}
  );
  items.forEach((it) =>
    queue.updateItem(it.id, { sheetNames: ['en_us', 'Data'], selectedSheet: 'en_us' })
  );

  const result = await queue.run(await bindings(page), FAST);

  assert(d365.prompts().includes('sheet-already-mapped'), 'the clash never came up');
  equal(result.uploaded, 2, `run summary was ${JSON.stringify(result)}`);
  equal(d365.uploads().map((u) => u.fileName), ['01_Sites_V2.xlsx', '02_Vendors_V2.xlsx']);
  equal(d365.openDialog(), null, 'the dialog was left on screen');
});

// Auto-clicking anything on an unrecognised confirmation is how an automation
// does real damage: these dialogs are also where delete and overwrite live.
test('leaves an unrecognised dialog alone and reports it', async () => {
  const { page, d365, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  d365.showDialog('Delete all staging data for this project?', [
    { label: 'Yes' },
    { label: 'No' }
  ]);

  const result = await queue.run(await bindings(page), FAST);

  assert(d365.openDialog() !== null, 'an unknown dialog must never be clicked');
  equal(result.uploaded, 0);

  const item = queue.getItems()[0];
  assert(
    item.error && item.error.includes('unanswered dialog'),
    `the failure should name the dialog blocking it, got: ${item.error}`
  );
});

// The reported stall: two files in, the queue sat at [UPLOAD] while D365's
// own message bar already read "'Customer Groups' entity mapping has
// completed successfully". The upload step waited on one signal only -- a new
// row in the entities grid -- which needs an optional binding and doesn't
// move reliably, because D365 renders that grid through React with
// virtualised rows.
async function bindingsWithStaleGrid(page) {
  // A bound selector that matches nothing: the row count starts at zero and
  // never grows, exactly like a stale or virtualised grid binding.
  return Object.assign({}, await bindings(page), {
    entitiesGridRow: { selector: '.entities-grid-row-that-never-appears' }
  });
}

test('a stale entities-grid binding no longer stalls the batch', async () => {
  const { page, d365, queue } = setup({
    entities: ['Sites V2', 'Vendors V2', 'Customers V3']
  });
  queue.addFiles(
    [
      file(page, '01_Sites_V2.xlsx'),
      file(page, '02_Vendors_V2.xlsx'),
      file(page, '03_Customers_V3.xlsx')
    ],
    {}
  );

  const result = await queue.run(await bindingsWithStaleGrid(page), FAST);

  equal(result.uploaded, 3, `run summary was ${JSON.stringify(result)}`);
  equal(d365.uploads().length, 3);
});

test("takes D365's success message as confirmation the file landed", async () => {
  const { page, queue } = setup({ entities: ['Vendors V2'] });
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  const result = await queue.run(await bindingsWithStaleGrid(page), FAST);

  equal(result.uploaded, 1, `run summary was ${JSON.stringify(result)}`);
  equal(queue.getItems()[0].confirmedBy, 'message');
});

// With no success message and no usable grid, the panel clearing itself is
// still proof D365 took the file.
test('takes the panel clearing as confirmation when nothing else says so', async () => {
  const { page, queue } = setup({
    entities: ['Vendors V2'],
    announceSuccess: false,
    resetPanelAfterUpload: true
  });
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  const result = await queue.run(await bindingsWithStaleGrid(page), FAST);

  equal(result.uploaded, 1, `run summary was ${JSON.stringify(result)}`);
  equal(queue.getItems()[0].confirmedBy, 'panel-reset');
});

// And when D365 confirms nothing at all, the step has to say so rather than
// reporting a missing grid row as though that were the whole story.
test('reports a genuinely unconfirmed upload', async () => {
  const { page, queue } = setup({ entities: ['Vendors V2'], announceSuccess: false });
  queue.addFiles([file(page, '01_Vendors_V2.xlsx')], {});

  const impatient = Object.assign({}, FAST, { uploadTimeout: 600 });
  const result = await queue.run(await bindingsWithStaleGrid(page), impatient);

  equal(result.uploaded, 0);
  const item = queue.getItems()[0];
  equal(item.step, 'upload');
  assert(
    item.error.includes('never confirmed the upload'),
    `unhelpful failure text: ${item.error}`
  );
});

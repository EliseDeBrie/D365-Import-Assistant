const fs = require('fs');
const os = require('os');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = path.resolve(__dirname, '..');
const PROFILE = fs.mkdtempSync(path.join(os.tmpdir(), 'd365ia-browser-smoke-'));

async function run() {
  let context;
  try {
    context = await chromium.launchPersistentContext(PROFILE, {
      headless: false,
      args: [`--disable-extensions-except=${ROOT}`, `--load-extension=${ROOT}`]
    });

    await context.route('https://contoso.operations.dynamics.com/**', (route) =>
      route.fulfill({
        status: 200,
        contentType: 'text/html',
        body: '<!doctype html><html><body><h1>Fake D365 import page</h1></body></html>'
      })
    );

    const page = context.pages()[0] || (await context.newPage());
    page.on('pageerror', (error) => console.error(`page error: ${error.message}`));
    page.on('console', (message) => {
      if (message.type() === 'error') console.error(`browser console: ${message.text()}`);
    });
    await page.goto(
      'https://contoso.operations.dynamics.com/?mi=DM_DataManagementWorkspaceMenuItem',
      { waitUntil: 'domcontentloaded' }
    );

    await page.locator('#d365ia-launcher').waitFor({ state: 'visible', timeout: 15000 });
    await page.locator('#d365ia-launcher').click();
    await page.locator('#d365ia-panel').waitFor({ state: 'visible' });

    await page.locator('#d365ia-file-input').setInputFiles([
      { name: 'Orders.csv', mimeType: 'text/csv', buffer: Buffer.from('id,name\n1,one') },
      { name: 'not-supported.exe', mimeType: 'application/octet-stream', buffer: Buffer.from('x') }
    ]);
    await page.getByText('Orders.csv', { exact: false }).waitFor();
    await page.getByText('not-supported.exe (unsupported file type)', { exact: false }).waitFor();
    if ((await page.locator('#d365ia-queue-list > li').count()) !== 1) {
      throw new Error('the extension did not queue exactly the supported file');
    }

    await page.evaluate(() => {
      const transfer = document.createElement('input');
      transfer.type = 'file';
      transfer.id = 'd365ia-file-transfer';
      document.body.appendChild(transfer);

      const target = document.createElement('input');
      target.type = 'file';
      target.id = 'd365ia-hook-target';
      document.body.appendChild(target);
      window.dispatchEvent(new CustomEvent('d365ia:arm-file-hook'));
    });
    await page.locator('#d365ia-file-transfer').setInputFiles({
      name: 'hooked.csv',
      mimeType: 'text/csv',
      buffer: Buffer.from('id\n2')
    });
    const hookedName = await page.evaluate(() => {
      document.getElementById('d365ia-hook-target').click();
      return document.getElementById('d365ia-hook-target').files[0]?.name || '';
    });
    if (hookedName !== 'hooked.csv') {
      throw new Error(`page-world file hook did not transfer the file (got "${hookedName}")`);
    }

    console.log('Chromium extension smoke test passed: launcher, file filtering, and page-world file handoff.');
  } finally {
    if (context) await context.close();
    fs.rmSync(PROFILE, { recursive: true, force: true });
  }
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

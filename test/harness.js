// Loads the extension's content scripts into a jsdom page so they can be
// exercised without a browser or a D365 tenant. The scripts are plain IIFEs
// that hang everything off window.D365IA, so they only need a window, a
// document and a chrome.storage stand-in.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { JSDOM } = require('jsdom');

const ROOT = path.join(__dirname, '..');

// Load order matters: queue.js destructures D365IA.domUtils at definition
// time, so its dependencies have to be in place first. This mirrors the
// order in manifest.json.
const SCRIPTS = [
  'content/dom-utils.js',
  'content/matcher.js',
  'content/xlsx-sheets.js',
  'content/entity-list.js',
  'content/binder.js',
  'content/queue.js'
];

// A minimal chrome.storage: enough for the extension's get/set/onChanged use,
// backed by a plain object so tests can inspect what was written.
function fakeStorageArea() {
  const data = {};
  return {
    data,
    async get(keys) {
      if (keys == null) return Object.assign({}, data);
      const names = Array.isArray(keys) ? keys : [keys];
      const out = {};
      names.forEach((name) => {
        if (name in data) out[name] = data[name];
      });
      return out;
    },
    async set(patch) {
      Object.assign(data, patch);
    },
    async remove(key) {
      delete data[key];
    }
  };
}

function fakeChrome() {
  return {
    storage: {
      local: fakeStorageArea(),
      sync: fakeStorageArea(),
      onChanged: { addListener() {} }
    },
    runtime: { getURL: (p) => `chrome-extension://test/${p}` }
  };
}

function createPage({ url = 'https://contoso.operations.dynamics.com/?cmp=USMF&mi=DM_DataManagementWorkspaceMenuItem', html = '<!doctype html><html><body></body></html>' } = {}) {
  const dom = new JSDOM(html, { url, runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;

  // jsdom has no layout engine, so getClientRects() is always empty and
  // domUtils.isVisible would reject every element. Stand in for layout with
  // the two things the real check is actually protecting against: elements
  // detached from the document, and elements hidden by style.
  window.Element.prototype.getClientRects = function () {
    if (!this.isConnected) return [];
    let node = this;
    while (node && node.nodeType === 1) {
      if (node.hidden) return [];
      const style = node.style || {};
      if (style.display === 'none' || style.visibility === 'hidden') return [];
      node = node.parentNode;
    }
    return [{ width: 100, height: 20, top: 0, left: 0 }];
  };

  // Handing a file to a page means assigning input.files, which jsdom's
  // WebIDL layer rejects for anything that isn't a real FileList (and it
  // exposes no way to build one). Replace it with a plain settable property
  // so both the extension and the fake page can pass files around.
  Object.defineProperty(window.HTMLInputElement.prototype, 'files', {
    configurable: true,
    get() {
      return this.__files || null;
    },
    set(value) {
      this.__files = value;
    }
  });

  // jsdom implements neither of these; the extension needs both.
  if (!window.DataTransfer) {
    window.DataTransfer = class DataTransfer {
      constructor() {
        const list = [];
        this.items = {
          add: (file) => list.push(file)
        };
        Object.defineProperty(this, 'files', { get: () => list });
      }
    };
  }
  // Reading a workbook's sheet names inflates a zip entry through
  // Blob.stream() -> DecompressionStream -> Response. jsdom implements none
  // of those; Node has them, so hand its versions to the page.
  window.DecompressionStream = globalThis.DecompressionStream;
  window.ReadableStream = globalThis.ReadableStream;
  window.Response = globalThis.Response;

  if (!window.Blob.prototype.stream) {
    window.Blob.prototype.stream = function () {
      const blob = this;
      return new globalThis.ReadableStream({
        async start(controller) {
          controller.enqueue(new Uint8Array(await blob.arrayBuffer()));
          controller.close();
        }
      });
    };
  }

  window.chrome = fakeChrome();

  const context = dom.getInternalVMContext();
  SCRIPTS.forEach((rel) => {
    const code = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    vm.runInContext(code, context, { filename: rel });
  });

  return { dom, window, document: window.document, D365IA: window.D365IA };
}

module.exports = { createPage, SCRIPTS };

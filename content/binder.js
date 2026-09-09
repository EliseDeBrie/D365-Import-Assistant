(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  // Roles the extension drives on the D365 page, in the order the Import
  // screen uses them: Add file -> Source data format -> (Excel/CSV only)
  // Entity name -> attach the file -> Upload -> a row appears in the
  // entities grid.
  const ROLES = [
    'addFileButton',
    'sourceFormatField',
    'sourceFormatOption',
    'entityNameField',
    'suggestionItem',
    'fileTarget',
    'uploadButton',
    'sheetSelectField',
    'sheetOption',
    'entitiesGridRow',
    'closePanelButton',
    'runImportButton'
  ];

  // D365 names its controls with data-dyn-controlname, which is stable
  // across sessions and rebuilds (unlike the generated element ids), so the
  // tool ships working selectors and click-to-bind is only needed when a
  // default doesn't resolve in a particular environment.
  //
  // The three list roles are deliberately absent: values are selected by
  // keyboard (type, arrow down, enter), so no dropdown row has to be
  // located in the DOM. They remain bindable as a fallback.
  const DEFAULT_SELECTORS = {
    addFileButton: '[data-dyn-controlname="AddFile"]',
    sourceFormatField: '[data-dyn-controlname="SourceNameControl"] input, [id$="_SourceNameControl_input"]',
    entityNameField: '[data-dyn-controlname="NewEntityNameControl"] input, [id$="_NewEntityNameControl_input"]',
    // The upload control holds two inputs: the box showing the file name and
    // the hidden <input type="file"> the browse button drives. Both match a
    // bare "... input", and since queryVisible takes the last visible match
    // (file inputs always count as visible), the file input would win and the
    // "did D365 accept the file?" check would read an empty value forever.
    fileTarget:
      '[data-dyn-controlname="NewFileUploadControlFileNameDisplay"] input:not([type="file"]), [id$="_NewFileUploadControlFileNameDisplay_input"]',
    uploadButton: '[data-dyn-controlname="NewFileUploadControlBrowseButton"]',
    sheetSelectField:
      '[data-dyn-controlname="NewSheetLookupControl"] input, [id$="_NewSheetLookupControl_input"]',
    closePanelButton: '[data-dyn-controlname="OkButton"]',
    runImportButton: '[data-dyn-controlname="ImportAsync"]'
  };

  // Roles that identify one item in a repeating list rather than a single
  // unique element — these need a selector that generalizes across
  // siblings (see domUtils.getGeneralizedListSelector), not one pinned to
  // the exact element clicked.
  const LIST_ROLES = new Set([
    'sourceFormatOption',
    'suggestionItem',
    'sheetOption',
    'entitiesGridRow'
  ]);

  let pickerActive = false;
  let pickerRole = null;
  let hoverEl = null;
  let onPicked = null;
  let onCancelled = null;
  let hintEl = null;

  function highlight(el) {
    if (hoverEl) hoverEl.classList.remove('d365ia-hover-highlight');
    hoverEl = el;
    if (hoverEl) hoverEl.classList.add('d365ia-hover-highlight');
  }

  function handleMouseOver(e) {
    if (!pickerActive) return;
    highlight(e.target);
  }

  // Plain clicks are left alone so the page behaves normally while picking
  // is active — you can click a field to focus it, type into it, open its
  // dropdown, etc. Only Alt+click finalizes a binding, so reaching a target
  // that's nested behind other interactions (e.g. an autocomplete suggestion
  // that only appears after you've typed something) still works.
  function handleClick(e) {
    if (!pickerActive) return;
    if (!e.altKey) return;
    e.preventDefault();
    e.stopPropagation();
    const el = e.target;
    const role = pickerRole;
    const selector = LIST_ROLES.has(role)
      ? D365IA.domUtils.getGeneralizedListSelector(el)
      : D365IA.domUtils.getSelector(el);
    const cb = onPicked;
    stopPicking();
    if (cb) {
      cb({
        role,
        selector,
        tag: el.tagName,
        isFileInput: el.tagName === 'INPUT' && el.type === 'file',
        isSelect: el.tagName === 'SELECT'
      });
    }
  }

  function handleKeydown(e) {
    if (!pickerActive) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      const cancelCb = onCancelled;
      stopPicking();
      if (cancelCb) cancelCb();
    }
  }

  function showHint() {
    hintEl = document.createElement('div');
    hintEl.id = 'd365ia-picker-hint';
    hintEl.textContent =
      'Interact with the page as usual to reveal the field (click to focus it, type to open a dropdown, etc). Hold Alt and click the element you want to bind it to. Press Esc to cancel.';
    document.body.appendChild(hintEl);
  }

  function hideHint() {
    if (hintEl) {
      hintEl.remove();
      hintEl = null;
    }
  }

  function startPicking(role, callback, onCancel) {
    pickerActive = true;
    pickerRole = role;
    onPicked = callback;
    onCancelled = onCancel || null;
    document.addEventListener('mouseover', handleMouseOver, true);
    // Capture phase so we see the click before the page does, but we only
    // act on it (preventDefault/stopPropagation) when Alt is held — a plain
    // click passes straight through to the page underneath.
    document.addEventListener('click', handleClick, true);
    document.addEventListener('keydown', handleKeydown, true);
    document.body.classList.add('d365ia-picking');
    showHint();
  }

  function stopPicking() {
    pickerActive = false;
    pickerRole = null;
    onPicked = null;
    onCancelled = null;
    document.removeEventListener('mouseover', handleMouseOver, true);
    document.removeEventListener('click', handleClick, true);
    document.removeEventListener('keydown', handleKeydown, true);
    document.body.classList.remove('d365ia-picking');
    highlight(null);
    hideHint();
  }

  const GENERATED_ID_RE = /^\d+_\d+_(.+)$/;

  function unescapeCssIdent(ident) {
    return ident
      .replace(/\\([0-9a-fA-F]{1,6})[ ]?/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
      .replace(/\\(.)/g, '$1');
  }

  // Repairs selectors saved by earlier versions:
  //  - strips this extension's own hover-highlight class, which only exists
  //    while the mouse is over the element during binding, so any selector
  //    containing it could never match again;
  //  - rewrites "#31_5_SourceNameControl_input" to its stable suffix form,
  //    since D365 regenerates those instance counters whenever it rebuilds
  //    the control.
  function migrateSelector(selector) {
    let migrated = selector.replace(/\.d365ia-[\w-]+/g, '').trim();

    // CSS.escape writes a leading digit as a hex escape followed by a space
    // ("#\\33 1_5_Foo_input"), so the id can't be matched with \S+ — unescape
    // first, then confirm it really is a single id and not a compound
    // selector that merely starts with one.
    const idOnly = /^#(.+)$/.exec(migrated);
    if (idOnly) {
      const rawId = unescapeCssIdent(idOnly[1]);
      const generated = GENERATED_ID_RE.exec(rawId);
      if (generated && !/\s/.test(rawId)) migrated = `[id$="_${generated[1]}"]`;
    }
    return migrated;
  }

  function currentHost() {
    return location.hostname || 'default';
  }

  // Selectors are environment-specific (a dev tenant can differ from prod),
  // so they're stored per host. Anything saved by an older version sat in a
  // single flat map — move it under the host it was captured on.
  function migrateStore(store) {
    const roleKeys = Object.keys(store).filter((key) => ROLES.includes(key));
    if (roleKeys.length === 0) return { store, changed: false };

    const host = currentHost();
    const hostBindings = Object.assign({}, store[host]);
    roleKeys.forEach((role) => {
      if (!hostBindings[role]) hostBindings[role] = store[role];
      delete store[role];
    });
    store[host] = hostBindings;
    return { store, changed: true };
  }

  async function readStore() {
    const data = await chrome.storage.sync.get('bindings');
    const { store, changed: migratedShape } = migrateStore(data.bindings || {});

    const host = currentHost();
    const hostBindings = store[host] || {};
    let changed = migratedShape;

    Object.keys(hostBindings).forEach((role) => {
      const binding = hostBindings[role];
      if (!binding || !binding.selector) return;
      const migrated = migrateSelector(binding.selector);
      if (migrated !== binding.selector) {
        binding.selector = migrated;
        changed = true;
      }
    });

    store[host] = hostBindings;
    if (changed) await chrome.storage.sync.set({ bindings: store });
    return store;
  }

  // What the automation actually uses: shipped defaults, with any binding
  // saved for this environment taking precedence. `source` lets the setup
  // dialog show which is which.
  async function getBindings() {
    const store = await readStore();
    const custom = store[currentHost()] || {};
    const effective = {};

    ROLES.forEach((role) => {
      if (custom[role] && custom[role].selector) {
        effective[role] = Object.assign({}, custom[role], { source: 'custom' });
      } else if (DEFAULT_SELECTORS[role]) {
        effective[role] = { selector: DEFAULT_SELECTORS[role], source: 'default' };
      }
    });
    return effective;
  }

  async function saveBinding(role, selector, meta) {
    const store = await readStore();
    const host = currentHost();
    store[host] = Object.assign({}, store[host], {
      [role]: Object.assign({ selector }, meta || {})
    });
    await chrome.storage.sync.set({ bindings: store });
    return getBindings();
  }

  // Clearing a custom binding falls back to the shipped default rather than
  // leaving the role unusable.
  async function clearBinding(role) {
    const store = await readStore();
    const host = currentHost();
    const hostBindings = Object.assign({}, store[host]);
    delete hostBindings[role];
    store[host] = hostBindings;
    await chrome.storage.sync.set({ bindings: store });
    return getBindings();
  }

  D365IA.binder = {
    ROLES,
    LIST_ROLES,
    DEFAULT_SELECTORS,
    startPicking,
    stopPicking,
    getBindings,
    saveBinding,
    clearBinding,
    currentHost
  };
})();

(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  // Roles the extension needs bound to real elements on the D365 page.
  // These are picked once by the user (click-to-bind) because D365's DOM
  // varies by version/customization and can't be safely hardcoded.
  //
  // Matches the real Import screen flow: click Add file -> pick a Source
  // data format -> (for Excel/CSV) type an Entity name and pick a
  // suggestion -> attach the file -> click Upload -> a new row appears in
  // the entities grid once it succeeds.
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

  async function getBindings() {
    const data = await chrome.storage.sync.get('bindings');
    const bindings = data.bindings || {};
    let changed = false;

    Object.keys(bindings).forEach((role) => {
      const binding = bindings[role];
      if (!binding || !binding.selector) return;
      const migrated = migrateSelector(binding.selector);
      if (migrated !== binding.selector) {
        binding.selector = migrated;
        changed = true;
      }
    });

    if (changed) await chrome.storage.sync.set({ bindings });
    return bindings;
  }

  async function saveBinding(role, selector, meta) {
    const bindings = await getBindings();
    bindings[role] = Object.assign({ selector }, meta || {});
    await chrome.storage.sync.set({ bindings });
    return bindings;
  }

  async function clearBinding(role) {
    const bindings = await getBindings();
    delete bindings[role];
    await chrome.storage.sync.set({ bindings });
    return bindings;
  }

  D365IA.binder = {
    ROLES,
    LIST_ROLES,
    startPicking,
    stopPicking,
    getBindings,
    saveBinding,
    clearBinding
  };
})();

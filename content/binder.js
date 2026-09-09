(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  // Roles the extension needs bound to real elements on the D365 page.
  // These are picked once by the user (click-to-bind) because D365's DOM
  // varies by version/customization and can't be safely hardcoded.
  const ROLES = ['entityNameField', 'suggestionItem', 'fileTarget', 'addRowButton', 'importButton'];

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
    const selector = D365IA.domUtils.getSelector(el);
    const role = pickerRole;
    const cb = onPicked;
    stopPicking();
    if (cb) {
      cb({
        role,
        selector,
        tag: el.tagName,
        isFileInput: el.tagName === 'INPUT' && el.type === 'file'
      });
    }
  }

  function handleKeydown(e) {
    if (!pickerActive) return;
    if (e.key === 'Escape') {
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

  async function getBindings() {
    const data = await chrome.storage.sync.get('bindings');
    return data.bindings || {};
  }

  async function saveBinding(role, selector, meta) {
    const bindings = await getBindings();
    bindings[role] = Object.assign({ selector }, meta || {});
    await chrome.storage.sync.set({ bindings });
    return bindings;
  }

  D365IA.binder = {
    ROLES,
    startPicking,
    stopPicking,
    getBindings,
    saveBinding
  };
})();

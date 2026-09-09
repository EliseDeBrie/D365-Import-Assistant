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

  function highlight(el) {
    if (hoverEl) hoverEl.classList.remove('d365ia-hover-highlight');
    hoverEl = el;
    if (hoverEl) hoverEl.classList.add('d365ia-hover-highlight');
  }

  function handleMouseOver(e) {
    if (!pickerActive) return;
    highlight(e.target);
  }

  function handleClick(e) {
    if (!pickerActive) return;
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

  function startPicking(role, callback) {
    pickerActive = true;
    pickerRole = role;
    onPicked = callback;
    document.addEventListener('mouseover', handleMouseOver, true);
    document.addEventListener('click', handleClick, true);
    document.body.classList.add('d365ia-picking');
  }

  function stopPicking() {
    pickerActive = false;
    pickerRole = null;
    onPicked = null;
    document.removeEventListener('mouseover', handleMouseOver, true);
    document.removeEventListener('click', handleClick, true);
    document.body.classList.remove('d365ia-picking');
    highlight(null);
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

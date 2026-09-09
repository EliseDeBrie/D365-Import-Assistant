(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  function setNativeValue(element, value) {
    const proto =
      element.tagName === 'TEXTAREA'
        ? window.HTMLTextAreaElement.prototype
        : window.HTMLInputElement.prototype;
    const descriptor = Object.getOwnPropertyDescriptor(proto, 'value');
    if (descriptor && descriptor.set) {
      descriptor.set.call(element, value);
    } else {
      element.value = value;
    }
  }

  function fireEvent(element, type, opts) {
    if (opts && opts.isKeyboard) {
      const event = new KeyboardEvent(type, {
        bubbles: true,
        cancelable: true,
        key: opts.key || '',
        code: opts.code || opts.key || ''
      });
      // keyCode/which are read-only legacy getters that can't be set through
      // the constructor, and D365's older controls still test them.
      if (opts.keyCode) {
        Object.defineProperty(event, 'keyCode', { get: () => opts.keyCode });
        Object.defineProperty(event, 'which', { get: () => opts.keyCode });
      }
      element.dispatchEvent(event);
      return;
    }
    element.dispatchEvent(new Event(type, { bubbles: true, cancelable: true }));
  }

  // Simulates a user typing into a field so frameworks that listen for
  // input/change (D365's control framework included) pick up the value.
  // Accepts a control wrapper as well as the input itself, since that's
  // what a user is likely to have clicked when binding.
  function typeIntoField(el, text) {
    const inputEl = resolveTextInput(el);
    inputEl.focus();
    setNativeValue(inputEl, '');
    fireEvent(inputEl, 'input');
    setNativeValue(inputEl, text);
    fireEvent(inputEl, 'input');
    fireEvent(inputEl, 'keyup', { isKeyboard: true, key: text.slice(-1) });
    fireEvent(inputEl, 'change');
    return inputEl;
  }

  // Commits a typed value in a D365 combo/lookup without picking from the
  // list, for controls that resolve what you typed on blur.
  function commitField(el) {
    const inputEl = resolveTextInput(el);
    pressKey(inputEl, { key: 'Enter', keyCode: 13 });
    fireEvent(inputEl, 'change');
    inputEl.blur();
  }

  const KEYS = {
    Enter: 13,
    ArrowDown: 40,
    ArrowUp: 38,
    Escape: 27,
    Tab: 9
  };

  function pressKey(el, { key, keyCode }) {
    const code = keyCode || KEYS[key];
    ['keydown', 'keypress', 'keyup'].forEach((type) =>
      fireEvent(el, type, { isKeyboard: true, key, code: key, keyCode: code })
    );
  }

  // Picks a value from a D365 lookup the way the keyboard does: type to
  // filter, arrow down to highlight the first row, Enter to commit.
  //
  // This mirrors how the control actually works — D365 opens its lookup from
  // the input's keyDown handler (Edit.js autoPresentLookup) and renders the
  // list in a separate popup host, which is exactly why trying to find and
  // click a row in the DOM was so fragile.
  //
  // Committing a value round-trips to the server, so timings vary; the waits
  // below poll for the value to settle rather than assuming a fixed delay.
  async function selectByKeyboard(el, desiredText, { settleMs = 400 } = {}) {
    const inputEl = typeIntoField(el, desiredText);
    // How long to leave between keystrokes and between polls. Derived from
    // settleMs so the whole interaction scales together: the shipped default
    // keeps the 120ms that D365 needs, while tests can drive it much faster.
    const tick = Math.max(20, Math.min(120, Math.round(settleMs / 3)));

    // Give the lookup time to open and filter, but stop early once the
    // control has clearly reacted.
    await waitFor(() => fieldText(inputEl) !== '', { timeout: settleMs, interval: 80 }).catch(
      () => null
    );

    pressKey(inputEl, { key: 'ArrowDown' });
    await new Promise((r) => setTimeout(r, tick));
    pressKey(inputEl, { key: 'Enter' });
    fireEvent(inputEl, 'change');

    // The commit is a server round-trip; wait for the field to stop changing
    // rather than guessing how long that takes.
    let previous = fieldText(inputEl);
    let stableFor = 0;
    while (stableFor < 2) {
      await new Promise((r) => setTimeout(r, tick));
      const current = fieldText(inputEl);
      stableFor = current === previous ? stableFor + 1 : 0;
      previous = current;
    }

    // Typed text is not a value. A D365 lookup only holds what was chosen
    // from its list; anything merely typed is thrown away the moment the
    // field loses focus. Reading the field while it still has focus can't
    // tell the two apart — which is how a sheet name that was only typed,
    // and an entity name D365 never recognised, both read back as "set" and
    // then failed later for some unrelated-looking reason. Blurring forces
    // the control to show what it actually holds.
    inputEl.blur();
    fireEvent(inputEl, 'blur');
    // focusout is the bubbling half of losing focus, and the one D365's
    // control framework binds on the wrapper rather than the input.
    fireEvent(inputEl, 'focusout');
    await new Promise((r) => setTimeout(r, tick));

    return fieldText(inputEl);
  }

  // Resolves when D365 reports its /fileUpload POST finished (see
  // page-hook.js), or null if nothing arrives in time.
  //
  // This is a corroborating signal, never a gate. The event is dispatched by
  // the page-world hook in whichever frame ran the upload, and can legitimately
  // never arrive -- the request may go through fetch() rather than XHR, or the
  // frame may be cross-origin and carry no hook at all. So it listens in the
  // upload's own frame as well as the top one, and the caller must treat a
  // null as "no information", not as failure. Waiting minutes on a signal that
  // may not be coming is what stalled a batch part-way through.
  function waitForUploadResponse(timeout, ownerDocument) {
    const views = [window];
    const view = ownerDocument && ownerDocument.defaultView;
    if (view && view !== window) views.push(view);

    return new Promise((resolve) => {
      function cleanup() {
        views.forEach((v) => {
          try {
            v.removeEventListener('d365ia:file-upload-finished', onFinished);
          } catch (e) {
            // Frame went away mid-upload; nothing to detach.
          }
        });
      }

      const timer = setTimeout(() => {
        cleanup();
        resolve(null);
      }, timeout);

      function onFinished(event) {
        clearTimeout(timer);
        cleanup();
        resolve((event && event.detail) || { ok: true });
      }

      views.forEach((v) => {
        try {
          v.addEventListener('d365ia:file-upload-finished', onFinished);
        } catch (e) {
          // Cross-origin frame -- unreachable, and the top listener stands.
        }
      });
    });
  }

  // D365 states failures plainly in its own message bar ("Entity not found",
  // "Excel sheet lookup value is mandatory"). Reading it is far more reliable
  // than inferring what went wrong from DOM state. Several selectors are
  // tried because the bar's markup differs across versions.
  const MESSAGE_SELECTORS = [
    '[data-dyn-controlname="MessageBar"]',
    '.messageBar-messageText',
    '.messageBar-message',
    '.messageBar',
    '.sysMessageArea',
    '[role="alert"]'
  ];

  function readMessages() {
    const seen = new Set();
    MESSAGE_SELECTORS.forEach((selector) => {
      queryAllVisible(selector).forEach((el) => {
        const text = (el.textContent || '').replace(/\s+/g, ' ').trim();
        if (text && text.length < 500) seen.add(text);
      });
    });
    return Array.from(seen);
  }

  // Messages that appeared since a previous snapshot — i.e. what this step
  // caused, rather than whatever was already on screen.
  function newMessagesSince(before) {
    const previous = new Set(before || []);
    return readMessages().filter((text) => !previous.has(text));
  }

  // The text a control currently shows, whether it's an input or a rendered
  // combo box wrapper.
  function fieldText(el) {
    if (!el) return '';
    const inputEl = resolveTextInput(el);
    if (inputEl && (inputEl.tagName === 'INPUT' || inputEl.tagName === 'TEXTAREA')) {
      return (inputEl.value || '').trim();
    }
    return (el.textContent || '').trim();
  }

  // Attaches a real File object to a native <input type="file">, as if the
  // user had picked it, then fires the events the page listens for.
  function dropFileOnInput(fileInputEl, file) {
    const dt = new DataTransfer();
    dt.items.add(file);
    fileInputEl.files = dt.files;
    fireEvent(fileInputEl, 'input');
    fireEvent(fileInputEl, 'change');
  }

  function waitFor(checkFn, { timeout = 4000, interval = 100 } = {}) {
    return new Promise((resolve, reject) => {
      const start = Date.now();
      const timer = setInterval(() => {
        let result;
        try {
          result = checkFn();
        } catch (e) {
          result = null;
        }
        if (result) {
          clearInterval(timer);
          resolve(result);
        } else if (Date.now() - start > timeout) {
          clearInterval(timer);
          reject(new Error('Timed out waiting for condition'));
        }
      }, interval);
    });
  }

  // D365 stamps instance counters into element ids ("31_5_SourceNameControl_input")
  // and regenerates them every time it rebuilds a control — the Add file panel
  // is torn down and rebuilt on every open — so a raw #id selector goes stale
  // immediately. Match on the stable suffix instead.
  const GENERATED_ID_RE = /^\d+_\d+_(.+)$/;

  // Classes this extension itself adds while picking. They must never end up
  // in a saved selector: d365ia-hover-highlight only exists while the mouse
  // is over the element during binding, so any selector containing it can
  // never match again afterwards.
  const OWN_CLASS_RE = /^d365ia-/;

  function usefulClasses(el) {
    if (!el.className || typeof el.className !== 'string') return [];
    return el.className
      .trim()
      .split(/\s+/)
      .filter((c) => c && !OWN_CLASS_RE.test(c));
  }

  function attrSelector(name, value) {
    return `[${name}="${String(value).replace(/"/g, '\\"')}"]`;
  }

  function classSelector(el) {
    const classes = usefulClasses(el);
    if (!classes.length) return null;
    return el.tagName.toLowerCase() + classes.map((c) => '.' + CSS.escape(c)).join('');
  }

  // A short path from a control wrapper down to the element inside it. The
  // wrapper's internal structure survives re-instantiation even though its
  // generated ids don't, so this stays valid across panel rebuilds.
  function descendantSelector(ancestor, el) {
    const tag = el.tagName.toLowerCase();
    const candidates = [];
    if (el.getAttribute('name')) candidates.push(tag + attrSelector('name', el.getAttribute('name')));
    if (el.getAttribute('role')) candidates.push(tag + attrSelector('role', el.getAttribute('role')));
    if (el.tagName === 'INPUT' && el.type) candidates.push(`${tag}[type="${el.type}"]`);
    const cls = classSelector(el);
    if (cls) candidates.push(cls);
    candidates.push(tag);

    for (const candidate of candidates) {
      const matches = ancestor.querySelectorAll(candidate);
      if (matches.length === 1 && matches[0] === el) return candidate;
    }

    const parts = [];
    let node = el;
    let depth = 0;
    while (node && node !== ancestor && depth < 10) {
      const parent = node.parentElement;
      if (!parent) return null;
      parts.unshift(`:nth-child(${Array.prototype.indexOf.call(parent.children, node) + 1})`);
      node = parent;
      depth++;
    }
    return node === ancestor && parts.length ? '> ' + parts.join(' > ') : null;
  }

  function nearestControlName(el) {
    let node = el;
    let depth = 0;
    while (node && node.nodeType === 1 && depth < 10) {
      const name = node.getAttribute && node.getAttribute('data-dyn-controlname');
      if (name) return { node, name };
      node = node.parentElement;
      depth++;
    }
    return null;
  }

  // Builds a selector for a single, specific element, strongly preferring
  // handles that survive D365 rebuilding the control: its own
  // data-dyn-controlname, then the nearest ancestor's, then an id matched by
  // its stable suffix. A raw id or DOM path is the last resort.
  function getSelector(el) {
    const own = el.getAttribute && el.getAttribute('data-dyn-controlname');
    if (own) return attrSelector('data-dyn-controlname', own);

    const control = nearestControlName(el);
    if (control) {
      const base = attrSelector('data-dyn-controlname', control.name);
      const rel = descendantSelector(control.node, el);
      const combined = rel ? `${base} ${rel}` : base;
      if (document.querySelector(combined) === el) return combined;
      return base;
    }

    if (el.id) {
      const generated = GENERATED_ID_RE.exec(el.id);
      if (generated) {
        const suffixSelector = `[id$="_${generated[1]}"]`;
        if (document.querySelector(suffixSelector) === el) return suffixSelector;
      } else {
        return `#${CSS.escape(el.id)}`;
      }
    }

    if (el.getAttribute && el.getAttribute('name')) {
      const named = el.tagName.toLowerCase() + attrSelector('name', el.getAttribute('name'));
      if (document.querySelector(named) === el) return named;
    }

    const path = [];
    let node = el;
    let depth = 0;
    while (node && node.nodeType === 1 && depth < 6) {
      let selector = classSelector(node) || node.tagName.toLowerCase();
      const parent = node.parentElement;
      if (parent) {
        const siblings = Array.from(parent.children).filter((c) => c.tagName === node.tagName);
        if (siblings.length > 1) {
          selector += `:nth-of-type(${siblings.indexOf(node) + 1})`;
        }
      }
      path.unshift(selector);
      node = parent;
      depth++;
    }
    return path.join(' > ');
  }

  // Builds a selector for one item in a repeating list (an autocomplete
  // suggestion, a dropdown option, a grid row) that's meant to match ALL of
  // that item's siblings, not just the one clicked — deliberately omits any
  // nth-of-type/nth-child indexing, unlike getSelector() above. Without
  // this, binding "the 2nd suggestion in the list" would only ever match
  // that exact position and never generalize to picking a different item.
  function getGeneralizedListSelector(el) {
    const cls = classSelector(el);
    if (cls && document.querySelectorAll(cls).length >= 1) return cls;

    if (el.getAttribute && el.getAttribute('role')) {
      return `${el.tagName.toLowerCase()}[role="${el.getAttribute('role')}"]`;
    }

    const control = nearestControlName(el);
    if (control && control.node !== el) {
      return `${attrSelector('data-dyn-controlname', control.name)} ${el.tagName.toLowerCase()}`;
    }

    const parent = el.parentElement;
    if (parent) {
      const pClass = usefulClasses(parent)[0];
      if (pClass) return `.${CSS.escape(pClass)} > ${el.tagName.toLowerCase()}`;
    }
    return el.tagName.toLowerCase();
  }

  function isVisible(el) {
    if (!el) return false;
    // File inputs are deliberately hidden by the page; they're still usable.
    if (el.tagName === 'INPUT' && el.type === 'file') return true;
    if (el.getClientRects().length === 0) return false;
    const view = (el.ownerDocument && el.ownerDocument.defaultView) || window;
    const style = view.getComputedStyle(el);
    return style.visibility !== 'hidden' && style.display !== 'none';
  }

  // D365 renders parts of its UI inside same-origin iframes, and a content
  // script's document.querySelector only ever sees its own document. Rather
  // than injecting a copy of the whole tool into every frame (which
  // duplicates the panel), search the top document plus any same-origin
  // iframe documents from one place. Cross-origin frames throw on access and
  // are skipped.
  function searchableDocuments() {
    const docs = [document];
    const frames = document.querySelectorAll('iframe, frame');
    frames.forEach((frame) => {
      try {
        const doc = frame.contentDocument;
        if (doc && doc.documentElement) docs.push(doc);
      } catch (e) {
        // Cross-origin frame — nothing we can read, and nothing to do.
      }
    });
    return docs;
  }

  function queryAllAcrossFrames(selector) {
    const found = [];
    searchableDocuments().forEach((doc) => {
      try {
        found.push(...doc.querySelectorAll(selector));
      } catch (e) {
        // Invalid selector — treated as matching nothing.
      }
    });
    return found;
  }

  // D365 can leave a torn-down copy of a panel in the DOM behind the live
  // one, so prefer the last visible match rather than the first match.
  function queryVisible(selector) {
    const all = queryAllAcrossFrames(selector);
    const visible = all.filter(isVisible);
    return visible.length ? visible[visible.length - 1] : null;
  }

  function queryAllVisible(selector) {
    return queryAllAcrossFrames(selector).filter(isVisible);
  }

  const LIST_ITEM_EXCLUDED_TAGS = new Set(['INPUT', 'TEXTAREA', 'SELECT']);

  // Elements in a repeating list (a suggestion, a dropdown option, a grid
  // row) are never themselves form controls. A selector that resolves to
  // inputs was bound onto a filter/search box or the field itself rather
  // than an actual list item — clicking one would hit something arbitrary,
  // so this is the one filter both the automation and the setup dialog's
  // health check apply, so what the dialog reports matches what will
  // actually work at runtime.
  function queryListCandidates(selector) {
    return queryAllVisible(selector).filter((el) => !LIST_ITEM_EXCLUDED_TAGS.has(el.tagName));
  }

  function resolveTextInput(el) {
    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') return el;
    return el.querySelector('input, textarea') || el;
  }

  function resolveSelect(el) {
    if (el.tagName === 'SELECT') return el;
    return el.querySelector('select');
  }

  // Finds the real <input type="file"> behind an upload control. Users bind
  // the visible box or button, but the input that actually accepts a file is
  // usually a hidden sibling. Never matches this extension's own inputs.
  function resolveFileInput(el) {
    const SELECTOR =
      'input[type="file"]:not(#d365ia-file-input):not(#d365ia-file-transfer)';
    if (el.tagName === 'INPUT' && el.type === 'file') return el;
    const inside = el.querySelector && el.querySelector(SELECTOR);
    if (inside) return inside;

    let node = el.parentElement;
    let depth = 0;
    while (node && depth < 8) {
      const found = node.querySelector(SELECTOR);
      if (found) return found;
      node = node.parentElement;
      depth++;
    }

    const all = queryAllAcrossFrames(SELECTOR);
    return all.length === 1 ? all[0] : null;
  }

  // D365's client framework binds to pointer events, not just mouse events,
  // so a mousedown/mouseup/click trio alone can be ignored entirely. Events
  // are constructed from the element's own window so they stay valid when it
  // lives in an iframe.
  function clickElement(el) {
    const view = (el.ownerDocument && el.ownerDocument.defaultView) || window;
    const base = { bubbles: true, cancelable: true, composed: true, view, button: 0 };
    const pointer = Object.assign({ pointerType: 'mouse', isPrimary: true }, base);

    if (view.PointerEvent) {
      el.dispatchEvent(new view.PointerEvent('pointerdown', Object.assign({ buttons: 1 }, pointer)));
    }
    el.dispatchEvent(new view.MouseEvent('mousedown', Object.assign({ buttons: 1 }, base)));
    if (view.PointerEvent) {
      el.dispatchEvent(new view.PointerEvent('pointerup', Object.assign({ buttons: 0 }, pointer)));
    }
    el.dispatchEvent(new view.MouseEvent('mouseup', base));
    el.dispatchEvent(new view.MouseEvent('click', base));
  }

  // Sets a native <select>'s value by matching one of its options' visible
  // text (case-insensitive, trimmed) and dispatches change.
  function selectNativeOption(selectEl, optionText) {
    const options = Array.from(selectEl.options || []);
    const match = options.find(
      (o) => o.textContent.trim().toLowerCase() === optionText.trim().toLowerCase()
    );
    if (!match) return false;
    selectEl.value = match.value;
    fireEvent(selectEl, 'input');
    fireEvent(selectEl, 'change');
    return true;
  }

  // Hands a file to the page-world hook (content/page-hook.js) so that when
  // D365 tries to open the OS file picker, it receives this file instead.
  // The File itself travels through a hidden input in the shared DOM,
  // because it can't be passed in a cross-world event payload.
  let hookFired = false;
  function markHookFired() {
    hookFired = true;
  }
  window.addEventListener('d365ia:file-hook-fired', markHookFired);

  // The hook lives in the page world of a specific frame, so both the
  // transfer input and the arm signal have to land in the frame that owns
  // the upload control — arming the top window does nothing for an input
  // inside an iframe.
  function armFileHook(file, ownerDocument) {
    const doc = ownerDocument || document;
    const view = doc.defaultView || window;

    let transfer = doc.getElementById('d365ia-file-transfer');
    if (!transfer) {
      transfer = doc.createElement('input');
      transfer.type = 'file';
      transfer.id = 'd365ia-file-transfer';
      transfer.style.display = 'none';
      doc.body.appendChild(transfer);
    }
    const dt = new (view.DataTransfer || DataTransfer)();
    dt.items.add(file);
    transfer.files = dt.files;

    hookFired = false;
    if (view !== window) view.addEventListener('d365ia:file-hook-fired', markHookFired);
    view.dispatchEvent(new view.CustomEvent('d365ia:arm-file-hook'));
  }

  function disarmFileHook(ownerDocument) {
    const doc = ownerDocument || document;
    const view = doc.defaultView || window;
    view.dispatchEvent(new view.CustomEvent('d365ia:disarm-file-hook'));
    const transfer = doc.getElementById('d365ia-file-transfer');
    if (transfer) transfer.value = '';
  }

  function fileHookFired() {
    return hookFired;
  }

  D365IA.domUtils = {
    setNativeValue,
    fireEvent,
    typeIntoField,
    commitField,
    fieldText,
    pressKey,
    selectByKeyboard,
    waitForUploadResponse,
    readMessages,
    newMessagesSince,
    dropFileOnInput,
    waitFor,
    searchableDocuments,
    getSelector,
    getGeneralizedListSelector,
    clickElement,
    selectNativeOption,
    isVisible,
    queryVisible,
    queryAllVisible,
    queryListCandidates,
    resolveTextInput,
    resolveSelect,
    resolveFileInput,
    armFileHook,
    disarmFileHook,
    fileHookFired
  };
})();

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
      element.dispatchEvent(
        new KeyboardEvent(type, { bubbles: true, cancelable: true, key: opts.key || '' })
      );
      return;
    }
    element.dispatchEvent(new Event(type, { bubbles: true, cancelable: true }));
  }

  // Simulates a user typing into a field so frameworks that listen for
  // input/change (D365's control framework included) pick up the value.
  function typeIntoField(inputEl, text) {
    inputEl.focus();
    setNativeValue(inputEl, '');
    fireEvent(inputEl, 'input');
    setNativeValue(inputEl, text);
    fireEvent(inputEl, 'input');
    fireEvent(inputEl, 'keyup', { isKeyboard: true, key: text.slice(-1) });
    fireEvent(inputEl, 'change');
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

  // For UI that only accepts drag-and-drop (no reachable file input), this
  // replays the drag/drop sequence with a real DataTransfer carrying the file.
  function dropFileOnDropTarget(targetEl, file) {
    const dt = new DataTransfer();
    dt.items.add(file);
    ['dragenter', 'dragover', 'drop'].forEach((type) => {
      const evt = new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: dt });
      targetEl.dispatchEvent(evt);
    });
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

  // Builds a selector that survives page reloads: prefers D365's own
  // data-dyn-controlname attribute, falls back to id, then a short DOM path.
  function getSelector(el) {
    if (el.getAttribute && el.getAttribute('data-dyn-controlname')) {
      return `[data-dyn-controlname="${el.getAttribute('data-dyn-controlname')}"]`;
    }
    if (el.id) return `#${CSS.escape(el.id)}`;

    const path = [];
    let node = el;
    let depth = 0;
    while (node && node.nodeType === 1 && depth < 6) {
      let selector = node.tagName.toLowerCase();
      if (node.className && typeof node.className === 'string') {
        const cls = node.className.trim().split(/\s+/).slice(0, 2).join('.');
        if (cls) selector += '.' + CSS.escape(cls).replace(/\\\./g, '.');
      }
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

  D365IA.domUtils = {
    setNativeValue,
    fireEvent,
    typeIntoField,
    dropFileOnInput,
    dropFileOnDropTarget,
    waitFor,
    getSelector
  };
})();

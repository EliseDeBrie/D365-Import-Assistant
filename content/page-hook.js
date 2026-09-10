// Runs in the PAGE's JavaScript world (see manifest "world": "MAIN"), not
// the extension's isolated world, so it can replace the methods D365 itself
// calls. Everything else in this extension runs isolated and cannot do this.
//
// D365's "Upload and add" button opens the operating system's file picker,
// which no extension can drive or dismiss. Instead of letting it open, this
// intercepts the call and hands over the file the user already dropped onto
// the panel — the same outcome as picking it in the dialog, minus the dialog.
(function () {
  const TRANSFER_ID = 'd365ia-file-transfer';
  let armed = false;

  window.addEventListener('d365ia:arm-file-hook', () => {
    armed = true;
    installUploadWatchers();
  });
  window.addEventListener('d365ia:disarm-file-hook', () => {
    armed = false;
  });

  // The file is handed across worlds through a hidden <input type="file"> in
  // the shared DOM — File objects can't be passed in an event payload.
  function takePendingFile() {
    const transfer = document.getElementById(TRANSFER_ID);
    if (!transfer || !transfer.files || !transfer.files.length) return null;
    return transfer.files[0];
  }

  function injectInto(inputEl) {
    const file = takePendingFile();
    if (!file) return false;

    const dt = new DataTransfer();
    dt.items.add(file);
    inputEl.files = dt.files;
    armed = false;

    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    window.dispatchEvent(new CustomEvent('d365ia:file-hook-fired'));
    return true;
  }

  function shouldIntercept(el) {
    return armed && el && el.type === 'file' && el.id !== TRANSFER_ID;
  }

  const nativeClick = HTMLInputElement.prototype.click;
  HTMLInputElement.prototype.click = function () {
    if (shouldIntercept(this) && injectInto(this)) return;
    return nativeClick.apply(this, arguments);
  };

  // D365 uploads through jquery.fileupload, which POSTs the file to
  // /fileUpload and only then reports success (FileUpload.js
  // UploadWasSuccessful). Watching that request finish is a far more direct
  // completion signal than waiting for a row to appear in the entities grid,
  // which is rendered asynchronously well afterwards. XHR can only be
  // observed from the page's own world, which is why this lives here.
  //
  // These wrappers are installed the first time a run arms the hook in this
  // frame, never at load. This script matches every *.dynamics.com page --
  // including Dynamics apps that are not Finance & Operations -- and leaving
  // a permanent wrapper around fetch and XMLHttpRequest on pages where the
  // extension is never used is not something a reviewer should have to take
  // on trust. Arming always precedes the upload it is watching for (see
  // armFileHook in content/dom-utils.js), so nothing is missed.
  let watchersInstalled = false;

  function installUploadWatchers() {
    if (watchersInstalled) return;
    watchersInstalled = true;

    const nativeOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function (method, url) {
      try {
        this.__d365iaIsFileUpload = typeof url === 'string' && /\/fileUpload\b/i.test(url);
      } catch (e) {
        this.__d365iaIsFileUpload = false;
      }
      return nativeOpen.apply(this, arguments);
    };

    // The upload can run inside one of D365's same-origin iframes, while the
    // extension's queue is watching from the top frame. Announce it in both, so
    // the queue hears it wherever the request was actually made.
    function announceUploadFinished(detail) {
      const targets = [window];
      try {
        if (window.top && window.top !== window) targets.push(window.top);
      } catch (e) {
        // Cross-origin top -- unreachable, and the local listener stands.
      }
      targets.forEach((target) => {
        try {
          target.dispatchEvent(new CustomEvent('d365ia:file-upload-finished', { detail }));
        } catch (e) {
          // Frame is gone or not reachable from here; nothing to do.
        }
      });
    }

    const nativeSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function () {
      if (this.__d365iaIsFileUpload) {
        this.addEventListener('loadend', () => {
          announceUploadFinished({
            status: this.status,
            ok: this.status >= 200 && this.status < 300
          });
        });
      }
      return nativeSend.apply(this, arguments);
    };

    // Newer D365 builds send some uploads through fetch() rather than XHR.
    // Without this, the completion signal simply never arrives for those.
    const nativeFetch = window.fetch;
    if (typeof nativeFetch === 'function') {
      window.fetch = function (input, init) {
        let isUpload = false;
        try {
          const url = typeof input === 'string' ? input : input && input.url;
          isUpload = typeof url === 'string' && /\/fileUpload\b/i.test(url);
        } catch (e) {
          isUpload = false;
        }

        const result = nativeFetch.apply(this, arguments);
        if (!isUpload) return result;

        return result.then(
          (response) => {
            announceUploadFinished({ status: response.status, ok: response.ok });
            return response;
          },
          (error) => {
            announceUploadFinished({ status: 0, ok: false });
            throw error;
          }
        );
      };
    }
  }

  if (HTMLInputElement.prototype.showPicker) {
    const nativeShowPicker = HTMLInputElement.prototype.showPicker;
    HTMLInputElement.prototype.showPicker = function () {
      if (shouldIntercept(this) && injectInto(this)) return;
      return nativeShowPicker.apply(this, arguments);
    };
  }
})();

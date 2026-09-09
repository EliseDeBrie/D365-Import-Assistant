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
  const nativeOpen = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function (method, url) {
    try {
      this.__d365iaIsFileUpload = typeof url === 'string' && /\/fileUpload\b/i.test(url);
    } catch (e) {
      this.__d365iaIsFileUpload = false;
    }
    return nativeOpen.apply(this, arguments);
  };

  const nativeSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function () {
    if (this.__d365iaIsFileUpload) {
      this.addEventListener('loadend', () => {
        window.dispatchEvent(
          new CustomEvent('d365ia:file-upload-finished', {
            detail: { status: this.status, ok: this.status >= 200 && this.status < 300 }
          })
        );
      });
    }
    return nativeSend.apply(this, arguments);
  };

  if (HTMLInputElement.prototype.showPicker) {
    const nativeShowPicker = HTMLInputElement.prototype.showPicker;
    HTMLInputElement.prototype.showPicker = function () {
      if (shouldIntercept(this) && injectInto(this)) return;
      return nativeShowPicker.apply(this, arguments);
    };
  }
})();

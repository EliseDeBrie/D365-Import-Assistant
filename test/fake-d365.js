// A stand-in for D365's "Import" form: the Add file panel, its two lookups
// and the upload box, wired to behave the way the real one does. It is
// deliberately faithful about the behaviours that broke in practice:
//
//  - the panel is built only when Add file is clicked;
//  - the entity name field only appears for Excel/CSV, never for a package;
//  - the upload box only appears once an entity name has been committed;
//  - the lookups reject typed text — a value only sticks if it was chosen
//    from the list (type, ArrowDown, Enter), which is what D365 does and
//    what the sheet picker bug was;
//  - the panel clears itself after an upload, so the second file in a queue
//    starts from a blank form rather than the previous one's values.
function install(window, options = {}) {
  const { document } = window;
  const state = {
    uploads: [],
    panelOpen: false,
    sheetsByFile: options.sheetsByFile || {},
    // Entity names this fake tenant knows. A lookup will not commit anything
    // outside this list, exactly like the real one.
    entities: options.entities || ['Inventory Adjustment Journal Names', 'Vendors V2'],
    formats: options.formats || ['Excel', 'CSV', 'Package', 'XML-Element'],
    // Sheet names already mapped in this import project, and the prompts
    // raised as a result.
    mappedSheets: new Set(),
    prompts: []
  };

  document.body.innerHTML = `
    <div class="page">
      <button data-dyn-controlname="AddFile">Add file</button>
      <div data-dyn-controlname="MessageBar"></div>
      <div id="panel-host"></div>
      <button data-dyn-controlname="ImportAsync">Import</button>
    </div>
  `;

  const host = document.getElementById('panel-host');
  const addFileBtn = document.querySelector('[data-dyn-controlname="AddFile"]');
  const messageBar = document.querySelector('[data-dyn-controlname="MessageBar"]');

  // D365's modal message box. It blocks the page: while one is up, nothing
  // else responds, which is what makes an unanswered dialog stall a batch
  // rather than merely slow it down.
  let openDialog = null;

  function showDialog(text, buttons) {
    const dialog = document.createElement('div');
    dialog.setAttribute('role', 'dialog');
    dialog.className = 'dialog-popup';

    const body = document.createElement('div');
    body.textContent = text;
    dialog.appendChild(body);

    buttons.forEach(({ label, onPick }) => {
      const btn = document.createElement('button');
      btn.setAttribute('data-dyn-controlname', `${label}Button`);
      btn.textContent = label;
      btn.addEventListener('click', () => {
        dialog.remove();
        openDialog = null;
        if (onPick) onPick();
      });
      dialog.appendChild(btn);
    });

    document.body.appendChild(dialog);
    openDialog = dialog;
    return dialog;
  }

  // Everything the page does goes through here, so a pending dialog really
  // does block it.
  function blocked() {
    return openDialog !== null;
  }

  function message(text) {
    const line = document.createElement('div');
    line.className = 'messageBar-messageText';
    line.textContent = text;
    messageBar.appendChild(line);
  }

  // D365's lookups are text inputs backed by a flyout list. Typing filters
  // the list but does NOT set the value; ArrowDown highlights a row and
  // Enter commits it. Anything left as raw typed text is discarded on blur.
  //
  // `exact` models the sheet lookup specifically: it will only commit a value
  // that matches a list entry outright. D365 reaches Excel through the ODBC
  // driver, which names each worksheet "Sheet$", so typing the bare sheet name
  // read out of the workbook matches nothing and never commits -- which is
  // exactly what left the sheet step stuck.
  function makeLookup(controlName, getOptions, { exact = false } = {}) {
    const wrap = document.createElement('div');
    wrap.setAttribute('data-dyn-controlname', controlName);
    const input = document.createElement('input');
    input.type = 'text';
    input.id = `31_5_${controlName}_input`;
    wrap.appendChild(input);

    const flyout = document.createElement('div');
    flyout.className = 'lookup-flyout';
    flyout.style.display = 'none';
    wrap.appendChild(flyout);

    let highlighted = -1;
    let committed = '';

    function currentMatches() {
      const typed = input.value.trim().toLowerCase();
      if (exact) return getOptions().filter((o) => o.toLowerCase() === typed);
      return getOptions().filter((o) => o.toLowerCase().includes(typed));
    }

    function renderFlyout() {
      flyout.innerHTML = '';
      currentMatches().forEach((name, i) => {
        const row = document.createElement('div');
        row.className = 'lookup-row';
        row.textContent = name;
        if (i === highlighted) row.classList.add('highlighted');
        row.addEventListener('click', () => commit(name));
        flyout.appendChild(row);
      });
      flyout.style.display = flyout.children.length ? 'block' : 'none';
    }

    function commit(value) {
      committed = value;
      input.value = value;
      highlighted = -1;
      flyout.style.display = 'none';
      input.dispatchEvent(new window.Event('change', { bubbles: true }));
      if (options.onCommit) options.onCommit(controlName, value);
    }

    input.addEventListener('input', () => {
      highlighted = -1;
      renderFlyout();
    });

    input.addEventListener('keydown', (e) => {
      const matches = currentMatches();
      if (e.key === 'ArrowDown') {
        highlighted = Math.min(highlighted + 1, matches.length - 1);
        if (highlighted < 0 && matches.length) highlighted = 0;
        renderFlyout();
      } else if (e.key === 'Enter') {
        if (highlighted >= 0 && matches[highlighted]) commit(matches[highlighted]);
      }
    });

    // Uncommitted text is not a value: leaving the field throws it away.
    input.addEventListener('blur', () => {
      input.value = committed;
    });

    Object.defineProperty(wrap, 'committedValue', { get: () => committed });
    return { wrap, input, commit, getCommitted: () => committed };
  }

  let sourceLookup = null;
  let entityLookup = null;
  let sheetLookup = null;

  function clearPanel() {
    host.innerHTML = '';
    sourceLookup = null;
    entityLookup = null;
    sheetLookup = null;
    state.panelOpen = false;
  }

  function buildUploadBox() {
    const wrap = document.createElement('div');
    wrap.setAttribute('data-dyn-controlname', 'NewFileUploadControlFileNameDisplay');
    const display = document.createElement('input');
    display.type = 'text';
    display.readOnly = true;
    wrap.appendChild(display);

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.style.display = 'none';
    wrap.appendChild(fileInput);

    const browse = document.createElement('button');
    browse.setAttribute('data-dyn-controlname', 'NewFileUploadControlBrowseButton');
    browse.textContent = 'Upload';
    // The real button opens the OS file picker via the hidden input; the
    // extension's page hook intercepts exactly this call.
    browse.addEventListener('click', () => fileInput.click());
    wrap.appendChild(browse);

    fileInput.addEventListener('change', () => {
      if (blocked()) return;
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      display.value = file.name;

      const sheets = state.sheetsByFile[file.name] || [];
      if (sheets.length > 1 && !sheetLookup) {
        // D365 lists worksheets in the driver's "Sheet$" form, not the bare
        // names that come out of xl/workbook.xml.
        sheetLookup = makeLookup(
          'NewSheetLookupControl',
          () => sheets.map((name) => `${name}$`),
          { exact: true }
        );
        host.appendChild(sheetLookup.wrap);
      }

      function recordUpload() {
        state.uploads.push({
          fileName: file.name,
          format: sourceLookup ? sourceLookup.getCommitted() : null,
          entity: entityLookup ? entityLookup.getCommitted() : null,
          sheet: () => (sheetLookup ? sheetLookup.getCommitted() : null)
        });
        message(`Uploaded ${file.name}.`);
      }

      // Two workbooks in one project using the same sheet name -- every file
      // having an "en_us" sheet, say -- makes D365 stop and ask. It defaults
      // to No, so an unanswered one silently drops the file.
      const clash = sheets.find((name) => state.mappedSheets.has(name));
      if (clash) {
        state.prompts.push('sheet-already-mapped');
        showDialog(
          'The sheet with the same name is already mapped in this project. Do you still want to continue?',
          [
            { label: 'Yes', onPick: () => { sheets.forEach((n) => state.mappedSheets.add(n)); recordUpload(); } },
            { label: 'No', onPick: () => message('Upload cancelled.') }
          ]
        );
        return;
      }

      sheets.forEach((name) => state.mappedSheets.add(name));
      recordUpload();
    });

    host.appendChild(wrap);
  }

  function onFormatCommitted(value) {
    // A package has no entity name and no sheet — its manifest supplies them.
    if (value === 'Package') {
      buildUploadBox();
      return;
    }
    if (entityLookup) return;
    entityLookup = makeLookup('NewEntityNameControl', () => state.entities);
    host.appendChild(entityLookup.wrap);
  }

  function openPanel() {
    if (state.panelOpen) return;
    state.panelOpen = true;
    sourceLookup = makeLookup('SourceNameControl', () => state.formats);
    host.appendChild(sourceLookup.wrap);

    const ok = document.createElement('button');
    ok.setAttribute('data-dyn-controlname', 'OkButton');
    ok.textContent = 'Close';
    ok.addEventListener('click', clearPanel);
    host.appendChild(ok);
  }

  addFileBtn.addEventListener('click', () => {
    if (blocked()) return;
    openPanel();
  });

  options.onCommit = (controlName, value) => {
    if (controlName === 'SourceNameControl') onFormatCommitted(value);
    // The upload box appears only once a real entity has been chosen.
    if (controlName === 'NewEntityNameControl' && value) buildUploadBox();
  };

  return {
    state,
    showDialog,
    prompts: () => state.prompts.slice(),
    openDialog: () => openDialog,
    openPanel,
    clearPanel,
    message,
    // The panel D365 leaves behind after a successful upload: same controls,
    // all values blank. This is what tripped the second file in a batch.
    resetForNextFile() {
      clearPanel();
      openPanel();
    },
    uploads: () => state.uploads.map((u) => ({ ...u, sheet: u.sheet() }))
  };
}

module.exports = { install };

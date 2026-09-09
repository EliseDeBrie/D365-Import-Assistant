(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  function createPanel({ queue, getBindings, getSettings }) {
    const panel = document.createElement('div');
    panel.id = 'd365ia-panel';
    panel.innerHTML = `
      <div id="d365ia-header">
        <span id="d365ia-title">D365 Import Assistant</span>
        <div id="d365ia-header-actions">
          <button id="d365ia-minimize" title="Minimize">_</button>
          <button id="d365ia-close" title="Close">x</button>
        </div>
      </div>
      <div id="d365ia-body">
        <div id="d365ia-dropzone">
          Drag Excel files here<br /><small>or click to browse</small>
          <input type="file" id="d365ia-file-input" multiple accept=".xlsx,.xls,.xlsm,.csv" style="display:none" />
        </div>
        <div id="d365ia-entity-bar">
          <span id="d365ia-entity-status">Entity list: not loaded</span>
          <button id="d365ia-entity-refresh">Load</button>
        </div>
        <div id="d365ia-status-bar"></div>
        <ul id="d365ia-queue-list"></ul>
        <div id="d365ia-actions">
          <button id="d365ia-run" title="Upload every queued file as an entity, then stop">Upload</button>
          <button id="d365ia-run-import" title="Upload everything, close the Add file panel, then start the import job">Upload + Import</button>
        </div>
        <div id="d365ia-actions-secondary">
          <button id="d365ia-clear">Clear</button>
          <button id="d365ia-setup">Setup fields</button>
        </div>
      </div>
    `;
    document.body.appendChild(panel);

    const dz = panel.querySelector('#d365ia-dropzone');
    const fileInput = panel.querySelector('#d365ia-file-input');
    const list = panel.querySelector('#d365ia-queue-list');
    const statusBar = panel.querySelector('#d365ia-status-bar');
    const runBtn = panel.querySelector('#d365ia-run');
    const runImportBtn = panel.querySelector('#d365ia-run-import');
    const clearBtn = panel.querySelector('#d365ia-clear');
    const setupBtn = panel.querySelector('#d365ia-setup');
    const entityStatus = panel.querySelector('#d365ia-entity-status');
    const entityRefreshBtn = panel.querySelector('#d365ia-entity-refresh');

    function showEntityCount(count) {
      entityStatus.textContent = count
        ? `Entity list: ${count.toLocaleString()} entities`
        : 'Entity list: not loaded';
      entityRefreshBtn.textContent = count ? 'Refresh' : 'Load';
    }

    // Pulled from this environment's own OData service document, then used
    // to flag file names that don't correspond to a real entity before a
    // long run starts.
    entityRefreshBtn.addEventListener('click', async () => {
      entityStatus.textContent = 'Entity list: loading...';
      try {
        const { count } = await D365IA.entityList.refresh();
        showEntityCount(count);
        render(queue.getItems());
      } catch (e) {
        entityStatus.textContent = `Entity list: ${e.message}`;
      }
    });

    D365IA.entityList.load().then(({ count }) => {
      showEntityCount(count);
      if (count) render(queue.getItems());
    });

    let dragCounter = 0;

    dz.addEventListener('dragenter', (e) => {
      e.preventDefault();
      dragCounter++;
      dz.classList.add('d365ia-dragover');
    });
    dz.addEventListener('dragover', (e) => e.preventDefault());
    dz.addEventListener('dragleave', () => {
      dragCounter--;
      if (dragCounter <= 0) {
        dragCounter = 0;
        dz.classList.remove('d365ia-dragover');
      }
    });
    // The actual file-adding on drop is handled once, globally, in
    // content.js (setupGlobalDropCapture) — that's what lets you drop
    // anywhere on the page, not just this small box. This listener only
    // resets the visual highlight.
    dz.addEventListener('drop', (e) => {
      e.preventDefault();
      dragCounter = 0;
      dz.classList.remove('d365ia-dragover');
    });
    dz.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
      const settings = await getSettings();
      queue.addFiles(fileInput.files, settings.rules);
      setStatus(`Added ${fileInput.files.length} file(s) to the queue.`);
      fileInput.value = '';
    });

    const REQUIRED_ROLES = ['addFileButton', 'sourceFormatField', 'entityNameField', 'fileTarget'];

    // Two modes: upload the entities and stop, or upload and then actually
    // kick off the import. The import step is deliberately a separate button
    // rather than a setting, since it's the one that loads data into D365.
    async function startRun(alsoImport) {
      const bindings = await getBindings();
      const missing = REQUIRED_ROLES.filter((role) => !bindings[role] || !bindings[role].selector);
      if (alsoImport && (!bindings.runImportButton || !bindings.runImportButton.selector)) {
        missing.push('runImportButton');
      }
      if (missing.length > 0) {
        setStatus(`Bind these fields first (Setup fields): ${missing.join(', ')}.`, 'warn');
        return;
      }

      // A multi-sheet workbook makes D365 stop and ask which sheet to
      // import. With nothing bound to answer that, every such file stalls
      // until it times out, so say so now rather than after the first one.
      const multiSheet = queue
        .getItems()
        .filter((it) => it.status === 'pending' && it.sheetNames && it.sheetNames.length > 1);
      const sheetBound = bindings.sheetSelectField && bindings.sheetSelectField.selector;
      if (multiSheet.length > 0 && !sheetBound) {
        setStatus(
          `${multiSheet.length} file(s) have multiple sheets, so D365 will ask which sheet to import and nothing can answer it. Bind "Sheet picker" in Setup fields first.`,
          'warn'
        );
        return;
      }

      const settings = await getSettings();
      setStatus('Running...');
      await queue.run(bindings, settings.options);

      if (queue.isPaused()) {
        const stuck = queue.getItems().find((it) => it.status === 'needs-review' || it.status === 'error');
        setStatus(
          stuck && stuck.status === 'error'
            ? 'Stopped — a file failed, see the red item below for why.'
            : 'Paused — an item needs your review below.',
          'warn'
        );
        return;
      }

      const items = queue.getItems();
      const allUploaded = items.length > 0 && items.every((it) => it.status === 'filled');

      if (!alsoImport) {
        setStatus(allUploaded ? 'All files uploaded.' : 'Done.');
        return;
      }

      // Never start an import over a partial batch.
      if (!allUploaded) {
        setStatus('Not every file uploaded — import not started.', 'warn');
        return;
      }

      try {
        await queue.finishImport(bindings, settings.options);
        setStatus('All files uploaded. Import started.');
      } catch (e) {
        setStatus(`Uploaded, but couldn't start the import: ${e.message}`, 'warn');
      }
    }

    runBtn.addEventListener('click', () => startRun(false));
    runImportBtn.addEventListener('click', () => startRun(true));

    clearBtn.addEventListener('click', () => {
      queue.clear();
      setStatus('');
    });

    setupBtn.addEventListener('click', () => {
      window.dispatchEvent(new CustomEvent('d365ia:open-setup'));
    });

    function setStatus(text, level) {
      statusBar.textContent = text || '';
      statusBar.className = level || '';
    }

    function render(items) {
      list.innerHTML = '';
      items.forEach((item, index) => {
        const li = document.createElement('li');
        li.className = `d365ia-item d365ia-status-${item.status}`;

        const nameDiv = document.createElement('div');
        nameDiv.className = 'd365ia-item-name';
        nameDiv.title = item.rawName;

        // Upload order is significant, so make each file's position explicit.
        const seq = document.createElement('span');
        seq.className = 'd365ia-item-seq';
        seq.textContent = index + 1;
        nameDiv.appendChild(seq);
        nameDiv.appendChild(document.createTextNode(item.rawName));

        const cleanDiv = document.createElement('div');
        cleanDiv.className = 'd365ia-item-clean';
        cleanDiv.textContent = `→ ${item.cleanedName}${item.matchedEntity ? ' (' + item.matchedEntity + ')' : ''}`;

        const check = D365IA.entityList.validate(item.cleanedName);
        if (check.status !== 'unknown') {
          const badge = document.createElement('span');
          badge.className = `d365ia-entity-badge d365ia-entity-${check.status}`;
          if (check.status === 'match') {
            badge.textContent = 'entity found';
          } else if (check.status === 'close') {
            badge.textContent = `entity list: ${check.name}?`;
          } else {
            badge.textContent = check.name ? `entity list: ${check.name}?` : 'not in entity list';
          }
          badge.title =
            'This is a hint, not an error: the entity list uses OData\'s technical names ' +
            '("OperationalSitesV2"), while D365\'s own lookup shows display labels ("Sites V2") ' +
            '— those are often legitimately different strings for the same entity. Check what ' +
            'D365 itself shows in the Entity name field before assuming this is wrong.';
          cleanDiv.appendChild(badge);
        }

        const statusDiv = document.createElement('div');
        statusDiv.className = 'd365ia-item-status';
        statusDiv.textContent = item.status + (item.error ? ': ' + item.error : '');

        li.appendChild(nameDiv);
        li.appendChild(cleanDiv);

        // Only workbooks with more than one sheet get a picker — D365 asks
        // which sheet to import for those, and defaults to the first here.
        if (item.sheetNames && item.sheetNames.length > 1) {
          const sheetRow = document.createElement('div');
          sheetRow.className = 'd365ia-item-sheet';

          const label = document.createElement('span');
          label.textContent = 'Sheet:';

          const sheetSelect = document.createElement('select');
          item.sheetNames.forEach((name) => {
            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            if (name === item.selectedSheet) opt.selected = true;
            sheetSelect.appendChild(opt);
          });
          sheetSelect.addEventListener('change', () => queue.setSheet(item.id, sheetSelect.value));

          sheetRow.appendChild(label);
          sheetRow.appendChild(sheetSelect);
          li.appendChild(sheetRow);
        }

        li.appendChild(statusDiv);

        if (item.status === 'needs-review' && item.suggestions && item.suggestions.length) {
          const select = document.createElement('select');
          const blank = document.createElement('option');
          blank.textContent = 'Pick entity...';
          blank.value = '';
          select.appendChild(blank);
          item.suggestions.forEach((s) => {
            const opt = document.createElement('option');
            opt.value = s;
            opt.textContent = s;
            select.appendChild(opt);
          });
          select.addEventListener('change', async () => {
            if (!select.value) return;
            const bindings = await getBindings();
            const settings = await getSettings();
            queue.resumeAfterReview(item.id, select.value, bindings, settings.options);
          });

          const skipBtn = document.createElement('button');
          skipBtn.textContent = 'Skip';
          skipBtn.addEventListener('click', async () => {
            queue.skip(item.id);
            const bindings = await getBindings();
            const settings = await getSettings();
            queue.run(bindings, settings.options);
          });

          li.appendChild(select);
          li.appendChild(skipBtn);
        }

        // An errored item just sits there otherwise — nothing in the UI
        // says whether clicking Upload again does anything with it (it
        // silently moves on to the next file; this one stays failed).
        if (item.status === 'error') {
          const retryBtn = document.createElement('button');
          retryBtn.textContent = 'Retry';
          retryBtn.addEventListener('click', async () => {
            queue.updateItem(item.id, { status: 'pending', error: null });
            const bindings = await getBindings();
            const settings = await getSettings();
            queue.run(bindings, settings.options);
          });

          const skipBtn = document.createElement('button');
          skipBtn.textContent = 'Skip';
          skipBtn.addEventListener('click', async () => {
            queue.skip(item.id);
            const bindings = await getBindings();
            const settings = await getSettings();
            queue.run(bindings, settings.options);
          });

          li.appendChild(retryBtn);
          li.appendChild(skipBtn);
        }

        list.appendChild(li);
      });
    }

    panel.querySelector('#d365ia-minimize').addEventListener('click', () => {
      panel.classList.toggle('d365ia-minimized');
    });
    panel.querySelector('#d365ia-close').addEventListener('click', () => {
      panel.remove();
    });

    makeDraggablePanel(panel, panel.querySelector('#d365ia-header'));

    return { render, setStatus, el: panel };
  }

  function makeDraggablePanel(panel, handle) {
    let dragging = false;
    let offsetX = 0;
    let offsetY = 0;
    handle.addEventListener('mousedown', (e) => {
      if (e.target.tagName === 'BUTTON') return;
      dragging = true;
      const rect = panel.getBoundingClientRect();
      offsetX = e.clientX - rect.left;
      offsetY = e.clientY - rect.top;
      e.preventDefault();
    });
    document.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      panel.style.left = `${e.clientX - offsetX}px`;
      panel.style.top = `${e.clientY - offsetY}px`;
      panel.style.right = 'auto';
    });
    document.addEventListener('mouseup', () => {
      dragging = false;
    });
  }

  D365IA.dropzone = { createPanel };
})();

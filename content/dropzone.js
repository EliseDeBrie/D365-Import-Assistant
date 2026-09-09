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
        <div id="d365ia-status-bar"></div>
        <ul id="d365ia-queue-list"></ul>
        <div id="d365ia-actions">
          <button id="d365ia-run">Start</button>
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
    const clearBtn = panel.querySelector('#d365ia-clear');
    const setupBtn = panel.querySelector('#d365ia-setup');

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

    runBtn.addEventListener('click', async () => {
      const bindings = await getBindings();
      if (!bindings.entityNameField || !bindings.fileTarget) {
        setStatus('Bind the Entity name field and File target first (Setup fields).', 'warn');
        return;
      }
      const settings = await getSettings();
      setStatus('Running...');
      await queue.run(bindings, settings.options);
      if (queue.isPaused()) {
        setStatus('Paused — an item needs your review below.', 'warn');
      } else {
        setStatus('Done.');
      }
    });

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
      items.forEach((item) => {
        const li = document.createElement('li');
        li.className = `d365ia-item d365ia-status-${item.status}`;

        const nameDiv = document.createElement('div');
        nameDiv.className = 'd365ia-item-name';
        nameDiv.title = item.rawName;
        nameDiv.textContent = item.rawName;

        const cleanDiv = document.createElement('div');
        cleanDiv.className = 'd365ia-item-clean';
        cleanDiv.textContent = `→ ${item.cleanedName}${item.matchedEntity ? ' (' + item.matchedEntity + ')' : ''}`;

        const statusDiv = document.createElement('div');
        statusDiv.className = 'd365ia-item-status';
        statusDiv.textContent = item.status + (item.error ? ': ' + item.error : '');

        li.appendChild(nameDiv);
        li.appendChild(cleanDiv);
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

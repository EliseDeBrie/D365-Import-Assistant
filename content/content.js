(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  async function getSettings() {
    const data = await chrome.storage.sync.get('settings');
    const defaults = D365IA.matcher.DEFAULT_RULES;
    const stored = data.settings || {};
    return {
      rules: Object.assign({}, defaults, stored.rules || {}),
      options: Object.assign(
        { matchThreshold: 0.75, stepDelay: 700, suggestionTimeout: 2500, autoAddRow: true },
        stored.options || {}
      )
    };
  }

  function createLauncher(onOpen) {
    const btn = document.createElement('button');
    btn.id = 'd365ia-launcher';
    btn.textContent = 'Import Assist';
    btn.addEventListener('click', onOpen);
    document.body.appendChild(btn);
    return btn;
  }

  function hasFilesBeingDragged(e) {
    return !!(e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files'));
  }

  // Catches a file drop anywhere on the page — not just the small box inside
  // the panel — so dropping onto the launcher button (or anywhere else)
  // still queues the files instead of silently no-opping. Runs in the
  // capture phase so it sees the event before the D365 page's own handlers
  // have a chance to swallow it.
  function setupGlobalDropCapture(onFilesDropped) {
    let overlay = null;
    let dragDepth = 0;

    function showOverlay() {
      if (overlay) return;
      overlay = document.createElement('div');
      overlay.id = 'd365ia-drop-overlay';
      overlay.textContent = 'Drop Excel files to queue them in D365 Import Assistant';
      document.body.appendChild(overlay);
    }

    function hideOverlay() {
      dragDepth = 0;
      if (overlay) {
        overlay.remove();
        overlay = null;
      }
    }

    document.addEventListener(
      'dragenter',
      (e) => {
        if (!hasFilesBeingDragged(e)) return;
        e.preventDefault();
        dragDepth++;
        showOverlay();
      },
      true
    );

    document.addEventListener(
      'dragover',
      (e) => {
        if (!hasFilesBeingDragged(e)) return;
        e.preventDefault();
      },
      true
    );

    document.addEventListener(
      'dragleave',
      (e) => {
        if (!hasFilesBeingDragged(e)) return;
        dragDepth--;
        if (dragDepth <= 0) hideOverlay();
      },
      true
    );

    document.addEventListener(
      'drop',
      (e) => {
        if (!hasFilesBeingDragged(e)) return;
        e.preventDefault();
        hideOverlay();
        onFilesDropped(e.dataTransfer.files);
      },
      true
    );
  }

  function init() {
    let panel = null;

    const queueApi = D365IA.queue.createQueue({
      onStatusChange: (items) => {
        if (panel) panel.render(items);
      }
    });

    function openPanel() {
      if (panel) {
        panel.el.classList.remove('d365ia-minimized');
        return panel;
      }
      panel = D365IA.dropzone.createPanel({
        queue: queueApi,
        getBindings: D365IA.binder.getBindings,
        getSettings
      });
      return panel;
    }

    createLauncher(openPanel);

    setupGlobalDropCapture(async (files) => {
      const p = openPanel();
      const settings = await getSettings();
      queueApi.addFiles(files, settings.rules);
      p.setStatus(`Added ${files.length} file(s) to the queue.`);
    });

    window.addEventListener('d365ia:open-setup', () => {
      D365IA.setupModal.open();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();

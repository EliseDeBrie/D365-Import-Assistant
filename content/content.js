(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  const DEFAULT_UI = {
    showLauncher: true,
    restrictToPages: true,
    // D365 puts the current form in the "mi" query parameter, so this keeps
    // the button off every other screen in the app.
    urlPatterns: 'mi=DM_DataManagementWorkspaceMenuItem'
  };

  const DEFAULT_OPTIONS = {
    matchThreshold: 0.75,
    stepDelay: 700,
    elementTimeout: 5000,
    uploadTimeout: 300000
  };

  async function getSettings() {
    const data = await chrome.storage.sync.get('settings');
    const stored = data.settings || {};
    return {
      rules: Object.assign({}, D365IA.matcher.DEFAULT_RULES, stored.rules || {}),
      options: Object.assign({}, DEFAULT_OPTIONS, stored.options || {}),
      ui: Object.assign({}, DEFAULT_UI, stored.ui || {})
    };
  }

  function parsePatterns(text) {
    return String(text || '')
      .split(/[\n,]/)
      .map((part) => part.trim())
      .filter(Boolean);
  }

  function urlMatches(ui) {
    if (!ui.restrictToPages) return true;
    const patterns = parsePatterns(ui.urlPatterns);
    if (patterns.length === 0) return true;
    return patterns.some((pattern) => location.href.includes(pattern));
  }

  async function loadLauncherPosition() {
    const data = await chrome.storage.local.get('launcherPosition');
    return data.launcherPosition || null;
  }

  // Lets the button be dragged out of the way, and keeps it where it was put.
  // A drag must not also fire the click that opens the panel, so movement past
  // a few pixels suppresses the click that follows mouseup.
  function makeLauncherDraggable(btn, onClick) {
    let dragging = false;
    let moved = false;
    let startX = 0;
    let startY = 0;
    let offsetX = 0;
    let offsetY = 0;

    btn.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      dragging = true;
      moved = false;
      startX = e.clientX;
      startY = e.clientY;
      const rect = btn.getBoundingClientRect();
      offsetX = e.clientX - rect.left;
      offsetY = e.clientY - rect.top;
      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      if (!moved && Math.abs(e.clientX - startX) < 4 && Math.abs(e.clientY - startY) < 4) return;
      moved = true;

      const left = Math.min(Math.max(0, e.clientX - offsetX), window.innerWidth - btn.offsetWidth);
      const top = Math.min(Math.max(0, e.clientY - offsetY), window.innerHeight - btn.offsetHeight);
      btn.style.left = `${left}px`;
      btn.style.top = `${top}px`;
      btn.style.right = 'auto';
      btn.style.bottom = 'auto';
    });

    document.addEventListener('mouseup', () => {
      if (!dragging) return;
      dragging = false;
      if (moved) {
        chrome.storage.local.set({
          launcherPosition: { left: btn.style.left, top: btn.style.top }
        });
      }
    });

    btn.addEventListener('click', (e) => {
      if (moved) {
        moved = false;
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      onClick();
    });
  }

  function createLauncher(onOpen, position) {
    const btn = document.createElement('button');
    btn.id = 'd365ia-launcher';
    btn.textContent = 'Import Assistant';
    btn.title = 'Click to open. Drag to move. Hide it in the extension popup.';
    if (position && position.left && position.top) {
      btn.style.left = position.left;
      btn.style.top = position.top;
      btn.style.right = 'auto';
      btn.style.bottom = 'auto';
    }
    makeLauncherDraggable(btn, onOpen);
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
  // have a chance to swallow it. isActive() keeps it from hijacking drops on
  // screens where the tool isn't meant to be running at all.
  function setupGlobalDropCapture(onFilesDropped, isActive) {
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
        if (!isActive() || !hasFilesBeingDragged(e)) return;
        e.preventDefault();
        dragDepth++;
        showOverlay();
      },
      true
    );

    document.addEventListener(
      'dragover',
      (e) => {
        if (!isActive() || !hasFilesBeingDragged(e)) return;
        e.preventDefault();
      },
      true
    );

    document.addEventListener(
      'dragleave',
      (e) => {
        if (!isActive() || !hasFilesBeingDragged(e)) return;
        dragDepth--;
        if (dragDepth <= 0) hideOverlay();
      },
      true
    );

    document.addEventListener(
      'drop',
      (e) => {
        if (!isActive() || !hasFilesBeingDragged(e)) return;
        e.preventDefault();
        hideOverlay();
        onFilesDropped(e.dataTransfer.files);
      },
      true
    );
  }

  async function init() {
    let panel = null;
    let launcher = null;
    let ui = DEFAULT_UI;

    const queueApi = D365IA.queue.createQueue({
      onStatusChange: (items) => {
        if (panel) panel.render(items);
      }
    });

    function isActive() {
      return ui.showLauncher && urlMatches(ui);
    }

    function openPanel() {
      if (panel) {
        panel.el.classList.remove('d365ia-minimized');
        panel.el.hidden = false;
        return panel;
      }
      panel = D365IA.dropzone.createPanel({
        queue: queueApi,
        getBindings: D365IA.binder.getBindings,
        getSettings
      });
      return panel;
    }

    function applyVisibility() {
      const active = isActive();
      if (launcher) launcher.hidden = !active;
      // The panel keeps its queue while hidden, so navigating away and back
      // doesn't lose what was already dropped in.
      if (panel) panel.el.hidden = !active;
    }

    const settings = await getSettings();
    ui = settings.ui;
    launcher = createLauncher(openPanel, await loadLauncherPosition());
    applyVisibility();

    setupGlobalDropCapture(async (files) => {
      const p = openPanel();
      const current = await getSettings();
      queueApi.addFiles(files, current.rules);
      p.setStatus(`Added ${files.length} file(s) to the queue.`);
    }, isActive);

    window.addEventListener('d365ia:open-setup', () => {
      D365IA.setupModal.open();
    });

    // D365 is a single-page app: it swaps forms with history.pushState, which
    // an isolated content script can't hook (that call happens in the page's
    // own JS world). Comparing the URL on a timer is the reliable way to
    // notice the user moved to or away from the data management screens.
    let lastHref = location.href;
    setInterval(() => {
      if (location.href === lastHref) return;
      lastHref = location.href;
      applyVisibility();
    }, 750);

    chrome.storage.onChanged.addListener(async (changes, area) => {
      if (area !== 'sync' || !changes.settings) return;
      ui = (await getSettings()).ui;
      applyVisibility();
    });
  }

  // The manifest injects into every frame so the automation can reach
  // controls D365 renders inside same-origin iframes (domUtils queries across
  // them). The user interface must exist exactly once, though — one launcher,
  // one panel, one queue — so only the top frame builds it. Sub-frames still
  // load page-hook.js, which is what they're actually needed for.
  function start() {
    if (window.top !== window.self) return;
    init();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
})();

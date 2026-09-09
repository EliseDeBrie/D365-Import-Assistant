document.addEventListener('DOMContentLoaded', async () => {
  const data = await chrome.storage.sync.get('bindings');
  const bindings = data.bindings || {};
  const bound = ['entityNameField', 'fileTarget'].every(
    (r) => bindings[r] && bindings[r].selector
  );

  const statusEl = document.getElementById('status');
  if (bound) {
    statusEl.textContent = 'Fields bound — ready to use.';
    statusEl.className = 'ready';
  } else {
    statusEl.textContent = 'Fields not bound yet. Open the import page, click Import Assist, then Setup fields.';
    statusEl.className = 'not-ready';
  }

  document.getElementById('open-options').addEventListener('click', () => {
    chrome.runtime.openOptionsPage();
  });
});

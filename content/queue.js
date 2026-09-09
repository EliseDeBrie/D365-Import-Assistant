(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  // Drives the batch: for each dropped file, type the cleaned name into the
  // bound entity field, wait for D365's own autocomplete suggestions, pick
  // the best match (or pause for the user to pick one), then attach the file.
  function createQueue({ onStatusChange }) {
    let items = [];
    let running = false;
    let paused = false;

    function emit() {
      if (onStatusChange) onStatusChange(items.slice());
    }

    function addFiles(fileList, rules) {
      const newItems = Array.from(fileList).map((file) => ({
        id: `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        file,
        rawName: file.name,
        cleanedName: D365IA.matcher.cleanFileName(file.name, rules),
        status: 'pending', // pending | matching | needs-review | matched | filled | error | skipped
        matchedEntity: null,
        suggestions: [],
        error: null
      }));
      items = items.concat(newItems);
      emit();
      return newItems;
    }

    function getItems() {
      return items.slice();
    }

    function updateItem(id, patch) {
      items = items.map((it) => (it.id === id ? Object.assign({}, it, patch) : it));
      emit();
    }

    function removeItem(id) {
      items = items.filter((it) => it.id !== id);
      emit();
    }

    function clear() {
      items = [];
      paused = false;
      emit();
    }

    async function processItem(id, bindings, options) {
      const item = items.find((it) => it.id === id);
      if (!item) return null;
      updateItem(id, { status: 'matching' });

      try {
        const entityFieldEl = document.querySelector(bindings.entityNameField.selector);
        const fileTargetEl = document.querySelector(bindings.fileTarget.selector);
        if (!entityFieldEl || !fileTargetEl) {
          throw new Error('Bound fields not found on this page — re-check Setup fields.');
        }

        D365IA.domUtils.typeIntoField(entityFieldEl, item.cleanedName);

        const suggestionSelector = bindings.suggestionItem && bindings.suggestionItem.selector;
        let suggestions = [];
        if (suggestionSelector) {
          try {
            await D365IA.domUtils.waitFor(
              () => document.querySelectorAll(suggestionSelector).length > 0,
              { timeout: options.suggestionTimeout || 2500 }
            );
            suggestions = Array.from(document.querySelectorAll(suggestionSelector)).map((el) =>
              el.textContent.trim()
            );
          } catch (e) {
            suggestions = [];
          }
        }

        if (suggestions.length > 0) {
          const { candidate, score } = D365IA.matcher.bestMatch(item.cleanedName, suggestions);
          if (score >= (options.matchThreshold || 0.75)) {
            const picked = selectSuggestion(suggestionSelector, candidate);
            if (picked) {
              updateItem(id, { matchedEntity: candidate });
            } else {
              updateItem(id, { status: 'needs-review', suggestions });
              paused = true;
              return { needsReview: true };
            }
          } else {
            updateItem(id, { status: 'needs-review', suggestions });
            paused = true;
            return { needsReview: true };
          }
        }

        attachFile(fileTargetEl, item.file);

        if (bindings.addRowButton && bindings.addRowButton.selector && options.autoAddRow) {
          const addBtn = document.querySelector(bindings.addRowButton.selector);
          if (addBtn) addBtn.click();
        }

        updateItem(id, { status: 'filled' });
        return { filled: true };
      } catch (err) {
        updateItem(id, { status: 'error', error: err.message });
        paused = true;
        return { error: err.message };
      }
    }

    function selectSuggestion(suggestionSelector, candidateText) {
      const els = Array.from(document.querySelectorAll(suggestionSelector));
      const el = els.find((e) => e.textContent.trim() === candidateText);
      if (!el) return false;
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      return true;
    }

    function attachFile(fileTargetEl, file) {
      if (fileTargetEl.tagName === 'INPUT' && fileTargetEl.type === 'file') {
        D365IA.domUtils.dropFileOnInput(fileTargetEl, file);
      } else {
        D365IA.domUtils.dropFileOnDropTarget(fileTargetEl, file);
      }
    }

    async function run(bindings, options) {
      if (running) return;
      running = true;
      paused = false;
      while (!paused) {
        const next = items.find((it) => it.status === 'pending');
        if (!next) break;
        const result = await processItem(next.id, bindings, options || {});
        if (!result) break;
        await new Promise((r) => setTimeout(r, (options && options.stepDelay) || 700));
      }
      running = false;
      emit();
    }

    function pause() {
      paused = true;
    }

    // Called from the UI once the user manually picks a suggestion for an
    // item stuck at "needs-review"; resumes the run afterwards.
    function resumeAfterReview(id, chosenSuggestionText, bindings, options) {
      const item = items.find((it) => it.id === id);
      if (!item) return;
      const suggestionSelector = bindings.suggestionItem && bindings.suggestionItem.selector;
      if (suggestionSelector) {
        selectSuggestion(suggestionSelector, chosenSuggestionText);
      }
      const fileTargetEl = document.querySelector(bindings.fileTarget.selector);
      if (fileTargetEl) attachFile(fileTargetEl, item.file);
      if (bindings.addRowButton && bindings.addRowButton.selector && options && options.autoAddRow) {
        const addBtn = document.querySelector(bindings.addRowButton.selector);
        if (addBtn) addBtn.click();
      }
      updateItem(id, { status: 'filled', matchedEntity: chosenSuggestionText });
      paused = false;
      run(bindings, options);
    }

    function skip(id) {
      updateItem(id, { status: 'skipped' });
      paused = false;
    }

    return {
      addFiles,
      getItems,
      updateItem,
      removeItem,
      clear,
      run,
      pause,
      resumeAfterReview,
      skip,
      isPaused: () => paused,
      isRunning: () => running
    };
  }

  D365IA.queue = { createQueue };
})();

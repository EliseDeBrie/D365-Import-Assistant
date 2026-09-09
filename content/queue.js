(function () {
  const D365IA = (window.D365IA = window.D365IA || {});
  const { domUtils, matcher } = D365IA;

  const CSV_EXTENSION_RE = /\.csv$/i;

  function sourceFormatFor(fileName) {
    return CSV_EXTENSION_RE.test(fileName) ? 'CSV' : 'Excel';
  }

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  // Drives the batch through D365's real per-file sequence: Add file ->
  // Source data format -> Entity name (autocomplete) -> attach file ->
  // Upload -> wait for a new row in the entities grid. For each file, the
  // entity name is matched against D365's own suggestion list rather than
  // typed in blind, and the queue pauses for you to pick manually whenever
  // that match is ambiguous.
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
        cleanedName: matcher.cleanFileName(file.name, rules),
        sourceFormat: sourceFormatFor(file.name),
        // pending | matching | needs-review | uploading | filled | error | skipped
        status: 'pending',
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

    // Waits for the bound element to show up rather than checking once —
    // D365 re-renders the row after every click (Add file, picking a
    // format, etc), so the next element in the sequence often doesn't
    // exist in the DOM yet at the instant we go looking for it.
    async function requireBoundEl(bindings, role, options) {
      const binding = bindings[role];
      if (!binding || !binding.selector) {
        throw new Error(`"${role}" isn't bound yet — open Setup fields.`);
      }
      try {
        return await domUtils.waitFor(() => document.querySelector(binding.selector), {
          timeout: (options && options.elementTimeout) || 5000
        });
      } catch (e) {
        throw new Error(`Bound element for "${role}" didn't show up on the page in time.`);
      }
    }

    // Opens a dropdown/select control and picks the option whose text best
    // matches desiredText. Used for the Source data format field, but not
    // tied to it specifically.
    async function pickFromDropdown(fieldEl, optionSelector, desiredText, options) {
      if (domUtils.isNativeSelect(fieldEl)) {
        if (!domUtils.selectNativeOption(fieldEl, desiredText)) {
          throw new Error(`No option matching "${desiredText}" in the dropdown.`);
        }
        return desiredText;
      }

      domUtils.clickElement(fieldEl);
      await domUtils.waitFor(() => document.querySelectorAll(optionSelector).length > 0, {
        timeout: options.elementTimeout || 5000
      });
      const optionEls = Array.from(document.querySelectorAll(optionSelector));
      const texts = optionEls.map((el) => el.textContent.trim());
      const { candidate, score } = matcher.bestMatch(desiredText, texts);
      if (!candidate || score < 0.5) {
        throw new Error(`Couldn't find a "${desiredText}" option in the dropdown.`);
      }
      const chosenEl = optionEls.find((el) => el.textContent.trim() === candidate);
      domUtils.clickElement(chosenEl);
      return candidate;
    }

    function selectSuggestion(suggestionSelector, candidateText) {
      const els = Array.from(document.querySelectorAll(suggestionSelector));
      const el = els.find((e) => e.textContent.trim() === candidateText);
      if (!el) return false;
      domUtils.clickElement(el);
      return true;
    }

    function attachFile(fileTargetEl, file) {
      if (fileTargetEl.tagName === 'INPUT' && fileTargetEl.type === 'file') {
        domUtils.dropFileOnInput(fileTargetEl, file);
      } else {
        domUtils.dropFileOnDropTarget(fileTargetEl, file);
      }
    }

    function countGridRows(bindings) {
      const binding = bindings.entitiesGridRow;
      if (!binding || !binding.selector) return null;
      return document.querySelectorAll(binding.selector).length;
    }

    // Fills in the Entity name field and resolves it against D365's own
    // autocomplete. Returns { needsReview: true } if the match isn't
    // confident enough to proceed unattended.
    async function matchEntityName(id, item, bindings, options) {
      const entityFieldEl = await requireBoundEl(bindings, 'entityNameField', options);
      domUtils.typeIntoField(entityFieldEl, item.cleanedName);

      const suggestionSelector = bindings.suggestionItem && bindings.suggestionItem.selector;
      let suggestions = [];
      if (suggestionSelector) {
        try {
          await domUtils.waitFor(() => document.querySelectorAll(suggestionSelector).length > 0, {
            timeout: options.elementTimeout || 5000
          });
          suggestions = Array.from(document.querySelectorAll(suggestionSelector)).map((el) =>
            el.textContent.trim()
          );
        } catch (e) {
          suggestions = [];
        }
      }

      if (suggestions.length === 0) return { needsReview: false };

      const { candidate, score } = matcher.bestMatch(item.cleanedName, suggestions);
      if (score >= (options.matchThreshold || 0.75) && selectSuggestion(suggestionSelector, candidate)) {
        updateItem(id, { matchedEntity: candidate });
        return { needsReview: false };
      }

      updateItem(id, { status: 'needs-review', suggestions });
      paused = true;
      return { needsReview: true };
    }

    // Runs one file through the full Add file -> Source format -> Entity
    // name -> File -> Upload sequence, then waits for the entities grid to
    // confirm success (or falls back to a fixed delay if that row isn't
    // bound). Stops and flags for review if the entity match is ambiguous;
    // stops and flags an error if anything else goes wrong.
    async function processItem(id, bindings, options) {
      const item = items.find((it) => it.id === id);
      if (!item) return null;
      updateItem(id, { status: 'matching' });

      try {
        const baselineCount = countGridRows(bindings);

        domUtils.clickElement(await requireBoundEl(bindings, 'addFileButton', options));

        const formatFieldEl = await requireBoundEl(bindings, 'sourceFormatField', options);
        const formatOptionSelector = bindings.sourceFormatOption && bindings.sourceFormatOption.selector;
        if (!formatOptionSelector && !domUtils.isNativeSelect(formatFieldEl)) {
          throw new Error(
            '"sourceFormatOption" isn\'t bound — open Setup fields, open the Source data format dropdown, and Alt+click one option (e.g. "Excel") to bind it.'
          );
        }
        await pickFromDropdown(formatFieldEl, formatOptionSelector, item.sourceFormat, options);

        const matchResult = await matchEntityName(id, item, bindings, options);
        if (matchResult.needsReview) return { needsReview: true };

        attachFile(await requireBoundEl(bindings, 'fileTarget', options), item.file);

        updateItem(id, { status: 'uploading' });
        domUtils.clickElement(await requireBoundEl(bindings, 'uploadButton', options));

        if (baselineCount !== null) {
          await domUtils.waitFor(() => countGridRows(bindings) > baselineCount, {
            timeout: options.uploadTimeout || 60000,
            interval: 500
          });
        } else {
          await sleep(options.stepDelay || 700);
        }

        updateItem(id, { status: 'filled' });
        return { filled: true };
      } catch (err) {
        updateItem(id, { status: 'error', error: err.message });
        paused = true;
        return { error: err.message };
      }
    }

    async function run(bindings, options) {
      if (running) return;
      running = true;
      paused = false;
      let completedAll = false;
      while (!paused) {
        const next = items.find((it) => it.status === 'pending');
        if (!next) {
          completedAll = true;
          break;
        }
        const result = await processItem(next.id, bindings, options || {});
        if (!result) break;
        await sleep((options && options.stepDelay) || 700);
      }
      running = false;

      if (completedAll && options && options.autoRunImport && bindings.runImportButton && bindings.runImportButton.selector) {
        const runBtn = document.querySelector(bindings.runImportButton.selector);
        if (runBtn) domUtils.clickElement(runBtn);
      }

      emit();
    }

    function pause() {
      paused = true;
    }

    // Called from the UI once the user manually picks a suggestion for an
    // item stuck at "needs-review"; finishes that file's remaining steps
    // (file attach, upload, wait for grid) and resumes the run afterwards.
    async function resumeAfterReview(id, chosenSuggestionText, bindings, options) {
      const item = items.find((it) => it.id === id);
      if (!item) return;

      try {
        const suggestionSelector = bindings.suggestionItem && bindings.suggestionItem.selector;
        if (suggestionSelector) selectSuggestion(suggestionSelector, chosenSuggestionText);

        const baselineCount = countGridRows(bindings);
        attachFile(await requireBoundEl(bindings, 'fileTarget', options), item.file);
        updateItem(id, { status: 'uploading', matchedEntity: chosenSuggestionText });
        domUtils.clickElement(await requireBoundEl(bindings, 'uploadButton', options));

        if (baselineCount !== null) {
          await domUtils.waitFor(() => countGridRows(bindings) > baselineCount, {
            timeout: (options && options.uploadTimeout) || 60000,
            interval: 500
          });
        } else {
          await sleep((options && options.stepDelay) || 700);
        }

        updateItem(id, { status: 'filled' });
      } catch (err) {
        updateItem(id, { status: 'error', error: err.message });
        return;
      }

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

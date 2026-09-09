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
        return await domUtils.waitFor(() => domUtils.queryVisible(binding.selector), {
          timeout: (options && options.elementTimeout) || 5000
        });
      } catch (e) {
        throw new Error(
          `Bound element for "${role}" didn't show up on the page in time (selector: ${binding.selector}). Re-bind it in Setup fields.`
        );
      }
    }

    function findBoundEl(bindings, role) {
      const binding = bindings[role];
      if (!binding || !binding.selector) return null;
      return domUtils.queryVisible(binding.selector);
    }

    // Elements in a suggestion/option list are never form controls. A
    // binding that resolves to inputs was mis-picked (it typically matches
    // every field on the page), and acting on it would click something
    // arbitrary — so ignore those and let the caller fall back to typing.
    function listCandidates(selector) {
      if (!selector) return [];
      return domUtils
        .queryAllVisible(selector)
        .filter((el) => !['INPUT', 'TEXTAREA', 'SELECT'].includes(el.tagName));
    }

    // Opens the Add file panel, retrying with a native click if D365 ignored
    // the synthetic event sequence.
    async function openAddFilePanel(bindings, options) {
      const addFileEl = await requireBoundEl(bindings, 'addFileButton', options);
      domUtils.clickElement(addFileEl);

      const opened = await domUtils
        .waitFor(() => findBoundEl(bindings, 'sourceFormatField'), {
          timeout: (options && options.elementTimeout) || 5000
        })
        .catch(() => null);

      if (!opened && typeof addFileEl.click === 'function') addFileEl.click();
    }

    // Sets a dropdown to desiredText. Native <select> is set directly;
    // otherwise it opens the control and clicks the best-matching option if
    // an option selector is bound, and falls back to typing the value into
    // the combo box (D365 combos resolve what you type) when it isn't.
    async function pickFromDropdown(fieldEl, optionSelector, desiredText, options) {
      const selectEl = domUtils.resolveSelect(fieldEl);
      if (selectEl) {
        if (!domUtils.selectNativeOption(selectEl, desiredText)) {
          throw new Error(`No option matching "${desiredText}" in the dropdown.`);
        }
        return desiredText;
      }

      domUtils.clickElement(fieldEl);

      if (optionSelector) {
        try {
          await domUtils.waitFor(() => listCandidates(optionSelector).length > 0, {
            timeout: options.elementTimeout || 5000
          });
          const optionEls = listCandidates(optionSelector);
          const texts = optionEls.map((el) => el.textContent.trim());
          const { candidate, score } = matcher.bestMatch(desiredText, texts);
          if (candidate && score >= 0.5) {
            const chosenEl = optionEls.find((el) => el.textContent.trim() === candidate);
            domUtils.clickElement(chosenEl);
            return candidate;
          }
        } catch (e) {
          // Fall through to typing the value instead.
        }
      }

      domUtils.typeIntoField(fieldEl, desiredText);
      domUtils.commitField(fieldEl);
      return desiredText;
    }

    function selectSuggestion(suggestionSelector, candidateText) {
      const el = listCandidates(suggestionSelector).find(
        (e) => e.textContent.trim() === candidateText
      );
      if (!el) return false;
      domUtils.clickElement(el);
      return true;
    }

    // The visible "Upload data file" box isn't the input that accepts a
    // file — the real <input type="file"> is hidden nearby, so bind either
    // and let this find the usable one.
    function attachFile(fileTargetEl, file) {
      const fileInput = domUtils.resolveFileInput(fileTargetEl);
      if (fileInput) {
        domUtils.dropFileOnInput(fileInput, file);
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
          await domUtils.waitFor(() => listCandidates(suggestionSelector).length > 0, {
            timeout: options.elementTimeout || 5000
          });
          suggestions = listCandidates(suggestionSelector).map((el) => el.textContent.trim());
        } catch (e) {
          suggestions = [];
        }
      }

      // No suggestion list to pick from — commit what was typed and let D365
      // validate it rather than leaving the field half-filled.
      if (suggestions.length === 0) {
        domUtils.commitField(entityFieldEl);
        return { needsReview: false };
      }

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

        // Clicking "Add file" while its panel is already open closes it
        // again, so only click when the panel isn't showing.
        if (!findBoundEl(bindings, 'sourceFormatField')) {
          await openAddFilePanel(bindings, options);
        }

        const formatFieldEl = await requireBoundEl(bindings, 'sourceFormatField', options);
        const formatOptionSelector = bindings.sourceFormatOption && bindings.sourceFormatOption.selector;
        await pickFromDropdown(formatFieldEl, formatOptionSelector, item.sourceFormat, options);

        const matchResult = await matchEntityName(id, item, bindings, options);
        if (matchResult.needsReview) return { needsReview: true };

        attachFile(await requireBoundEl(bindings, 'fileTarget', options), item.file);

        updateItem(id, { status: 'uploading' });
        // Optional: the file is attached to the input directly, so this is
        // only needed where D365 waits for an explicit commit click.
        const uploadEl = findBoundEl(bindings, 'uploadButton');
        if (uploadEl) domUtils.clickElement(uploadEl);

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
        const uploadEl = findBoundEl(bindings, 'uploadButton');
        if (uploadEl) domUtils.clickElement(uploadEl);

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

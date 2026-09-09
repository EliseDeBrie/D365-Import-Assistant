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
        sheetNames: [],
        selectedSheet: null,
        error: null
      }));
      items = items.concat(newItems);
      sortPending();
      emit();

      // Sheet names come from reading the workbook itself, which is async —
      // the queue is usable meanwhile and the rows fill in as they resolve.
      newItems.forEach((item) => {
        D365IA.xlsxSheets.readSheetNames(item.file).then((sheetNames) => {
          if (sheetNames.length < 2) return;
          updateItem(item.id, { sheetNames, selectedSheet: sheetNames[0] });
        });
      });

      return newItems;
    }

    // A dropped selection arrives in whatever order the OS hands it over,
    // but these files depend on each other and must import in their numbered
    // order. Files already processed keep their place.
    function sortPending() {
      const settled = items.filter((it) => it.status !== 'pending');
      const pending = items
        .filter((it) => it.status === 'pending')
        .sort((a, b) => matcher.compareNaturally(a.rawName, b.rawName));
      items = settled.concat(pending);
    }

    function setSheet(id, sheetName) {
      updateItem(id, { selectedSheet: sheetName });
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

    function listCandidates(selector) {
      return selector ? domUtils.queryListCandidates(selector) : [];
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

    // Hands the file to D365 by driving its own upload flow: arm the page
    // hook, then click the button that would normally open the OS file
    // picker. The hook answers that picker with the dropped file, so D365
    // runs its real upload path and no dialog appears. Falls back to
    // writing straight into a reachable file input if the hook never fires.
    async function attachFile(bindings, item, options) {
      let fileTargetEl;
      try {
        fileTargetEl = await requireBoundEl(bindings, 'fileTarget', options);
      } catch (e) {
        throw new Error(
          'The upload box never appeared. D365 only shows it once a valid entity name is selected, so check what the Entity name field on the page actually holds.'
        );
      }
      domUtils.armFileHook(item.file);

      const browseEl = findBoundEl(bindings, 'uploadButton') || fileTargetEl;
      domUtils.clickElement(browseEl);

      const fired = await domUtils
        .waitFor(() => domUtils.fileHookFired(), { timeout: 2000, interval: 100 })
        .catch(() => false);

      if (!fired) {
        const fileInput = domUtils.resolveFileInput(fileTargetEl);
        if (fileInput) {
          domUtils.dropFileOnInput(fileInput, item.file);
        } else {
          domUtils.disarmFileHook();
          throw new Error(
            'Couldn\'t hand the file to D365 — no file input was reachable and the Upload button didn\'t ask for one. Bind "Upload button" to the "Upload and add" button in Setup fields.'
          );
        }
      }

      domUtils.disarmFileHook();
    }

    function countGridRows(bindings) {
      const binding = bindings.entitiesGridRow;
      if (!binding || !binding.selector) return null;
      return document.querySelectorAll(binding.selector).length;
    }

    // A workbook with more than one sheet makes D365 ask which one to
    // import. Best-effort: if the sheet control is bound and shows up, set
    // it to the sheet chosen in the panel; otherwise leave it to the user.
    async function applySheetSelection(bindings, item, options) {
      if (!item.selectedSheet) return;
      const binding = bindings.sheetSelectField;
      if (!binding || !binding.selector) return;

      const sheetEl = await domUtils
        .waitFor(() => domUtils.queryVisible(binding.selector), {
          timeout: (options && options.elementTimeout) || 5000
        })
        .catch(() => null);
      if (!sheetEl) return;

      if (domUtils.fieldText(sheetEl).toLowerCase() === item.selectedSheet.toLowerCase()) return;
      const optionSelector = bindings.sheetOption && bindings.sheetOption.selector;
      await pickFromDropdown(sheetEl, optionSelector, item.selectedSheet, options);
    }

    // D365 re-renders the Add file panel after an upload and clears the
    // entity name. Starting the next file mid-render loses whatever is typed,
    // so wait for the field to come back empty before moving on.
    async function waitForPanelReset(bindings, options) {
      const binding = bindings.entityNameField;
      if (!binding || !binding.selector) return;
      await domUtils
        .waitFor(
          () => {
            const el = domUtils.queryVisible(binding.selector);
            return el && !domUtils.fieldText(el);
          },
          { timeout: (options && options.elementTimeout) || 5000 }
        )
        .catch(() => null);
    }

    async function waitForGridRow(bindings, baselineCount, options) {
      const timeout = (options && options.uploadTimeout) || 300000;
      try {
        await domUtils.waitFor(() => countGridRows(bindings) > baselineCount, {
          timeout,
          interval: 500
        });
      } catch (e) {
        throw new Error(
          `The file was handed to D365 but no new row appeared in the entities grid within ${Math.round(
            timeout / 1000
          )}s. If this workbook has several sheets, D365 is probably still waiting for a sheet to be picked — bind the sheet picker in Setup fields. Otherwise the upload may just be slow (D365 warns that Excel imports can queue for the Excel driver), so raise "Max wait for a file to finish uploading" in Settings.`
        );
      }
    }

    // Types the name and commits it, then reports back what the field
    // actually holds — D365 can re-render the panel underneath us after the
    // previous upload and silently discard what was typed.
    async function typeAndCommit(entityFieldEl, name) {
      domUtils.typeIntoField(entityFieldEl, name);
      await sleep(300);
      domUtils.commitField(entityFieldEl);
      await sleep(300);
      return domUtils.fieldText(entityFieldEl);
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

      // The environment's own entity list (Load/Refresh in the panel) uses
      // OData's technical entity names ("OperationalSitesV2"), while D365's
      // own lookup shows display labels ("Sites V2") — often genuinely
      // different strings for the same entity, not a sign anything is
      // wrong. So this is only ever an extra signal when picking among
      // live suggestions that D365 itself is already offering, never a
      // reason to reject what D365 has actually put in the field.
      const validated = D365IA.entityList.validate(item.cleanedName);
      const trustedName =
        validated.status === 'match' || !validated.name ? item.cleanedName : validated.name;

      // No usable suggestion list — either unbound, or bound onto something
      // that isn't a real list of rows (see queryListCandidates). Best
      // effort: commit what was typed and trust D365 kept it if the field
      // is non-empty. There's no reliable local signal for "D365 silently
      // rejected this" short of the field going empty, which is already
      // handled below.
      if (suggestions.length === 0) {
        let text = await typeAndCommit(entityFieldEl, item.cleanedName);

        if (!text) {
          const retryEl = await requireBoundEl(bindings, 'entityNameField', options);
          text = await typeAndCommit(retryEl, item.cleanedName);
        }

        if (!text) {
          throw new Error(
            `D365 didn't accept the entity name "${item.cleanedName}" — it may not match an entity in this environment. Set it by hand to check the exact name, or bind "one row in the entity name suggestions" so a match can be picked from the list.`
          );
        }

        updateItem(id, { matchedEntity: text });
        return { needsReview: false };
      }

      const threshold = options.matchThreshold || 0.75;
      const direct = matcher.bestMatch(item.cleanedName, suggestions);
      const trusted = trustedName === item.cleanedName ? direct : matcher.bestMatch(trustedName, suggestions);
      const best = trusted.score >= direct.score ? trusted : direct;

      if (best.score >= threshold && selectSuggestion(suggestionSelector, best.candidate)) {
        updateItem(id, { matchedEntity: best.candidate });
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

        // From the second file on, the panel stays open and keeps the last
        // format, so only touch the dropdown when it needs changing.
        const formatFieldEl = await requireBoundEl(bindings, 'sourceFormatField', options);
        if (domUtils.fieldText(formatFieldEl).toLowerCase() !== item.sourceFormat.toLowerCase()) {
          const formatOptionSelector =
            bindings.sourceFormatOption && bindings.sourceFormatOption.selector;
          await pickFromDropdown(formatFieldEl, formatOptionSelector, item.sourceFormat, options);
        }

        const matchResult = await matchEntityName(id, item, bindings, options);
        if (matchResult.needsReview) return { needsReview: true };

        updateItem(id, { status: 'uploading' });
        await attachFile(bindings, item, options);
        await applySheetSelection(bindings, item, options);

        if (baselineCount !== null) {
          await waitForGridRow(bindings, baselineCount, options);
        } else {
          await sleep(options.stepDelay || 700);
        }

        await waitForPanelReset(bindings, options);
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
      emit();
      return completedAll;
    }

    // Closes the Add file panel and starts D365's import job. Only reached
    // from the explicit "Upload + Import" button, never from a plain upload
    // run — this is the step that actually loads data into D365.
    async function finishImport(bindings, options) {
      const closeEl = findBoundEl(bindings, 'closePanelButton');
      if (closeEl) {
        domUtils.clickElement(closeEl);
        await sleep((options && options.stepDelay) || 700);
      }

      const importEl = await requireBoundEl(bindings, 'runImportButton', options);
      domUtils.clickElement(importEl);
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
        updateItem(id, { status: 'uploading', matchedEntity: chosenSuggestionText });
        await attachFile(bindings, item, options || {});

        if (baselineCount !== null) {
          await waitForGridRow(bindings, baselineCount, options);
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
      setSheet,
      clear,
      run,
      finishImport,
      pause,
      resumeAfterReview,
      skip,
      isPaused: () => paused,
      isRunning: () => running
    };
  }

  D365IA.queue = { createQueue };
})();

(function () {
  const D365IA = (window.D365IA = window.D365IA || {});
  const { domUtils, matcher } = D365IA;

  // D365's Source data format for each kind of file. A data package is a zip
  // holding its own manifest, so it has no entity name and no sheet to pick —
  // treating one as Excel (the old behaviour) made those steps fail.
  const FORMAT_BY_EXTENSION = [
    { pattern: /\.csv$/i, format: 'CSV' },
    { pattern: /\.zip$/i, format: 'Package' },
    { pattern: /\.(xlsx|xlsm|xls)$/i, format: 'Excel' }
  ];

  function sourceFormatFor(fileName) {
    const match = FORMAT_BY_EXTENSION.find((entry) => entry.pattern.test(fileName));
    return match ? match.format : 'Excel';
  }

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function stepError(step, message, messages) {
    const detail = messages && messages.length ? ` D365 said: "${messages.join(' | ')}"` : '';
    const error = new Error(`[${step}] ${message}${detail}`);
    error.step = step;
    return error;
  }

  function normalizeText(text) {
    return String(text || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '');
  }

  // D365 reaches Excel through the ODBC driver, which exposes each worksheet
  // as "SheetName$" -- that trailing marker is what distinguishes the whole
  // sheet from a named range over part of it. The sheet names read out of
  // xl/workbook.xml have no "$", so typing one of those filters D365's sheet
  // lookup to nothing and the value never commits. Add the marker when
  // driving the lookup; the queue and the panel keep the plain name.
  function sheetLookupText(sheetName) {
    const name = String(sheetName || '');
    return name.endsWith('$') ? name : `${name}$`;
  }

  // Drives the batch through D365's real per-file sequence as a list of named
  // steps. Each step verifies its own outcome, so a failure names the step
  // that actually failed instead of surfacing two steps later as something
  // unrelated, and D365's own message bar is quoted when it has something to
  // say.
  function createQueue({ onStatusChange }) {
    let items = [];
    let running = false;
    let paused = false;
    // null = not known yet, true = the page hook's upload signal reaches us,
    // false = it doesn't in this environment. Learned from the first file so
    // a batch never pays the timeout more than once: 20s x 37 files of
    // waiting for a signal that was never coming is most of an afternoon.
    let uploadSignalWorks = null;

    // Notifying the UI must never be able to break the run: a render error
    // reaching back into a step would fail the file it was working on, and
    // one reaching run() itself would end the batch.
    function emit() {
      if (!onStatusChange) return;
      try {
        onStatusChange(items.slice());
      } catch (e) {
        console.error('[D365 Import Assistant] panel render failed', e);
      }
    }

    function addFiles(fileList, rules) {
      const newItems = Array.from(fileList).map((file) => ({
        id: `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        file,
        rawName: file.name,
        cleanedName: matcher.cleanFileName(file.name, rules),
        sourceFormat: sourceFormatFor(file.name),
        // pending | running | needs-review | filled | error | skipped
        status: 'pending',
        step: null,
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
        if (item.sourceFormat === 'Package') return;
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

    function summary() {
      const count = (status) => items.filter((it) => it.status === status).length;
      return {
        total: items.length,
        uploaded: count('filled'),
        failed: count('error'),
        needsReview: count('needs-review'),
        skipped: count('skipped'),
        pending: count('pending')
      };
    }

    // ---------------------------------------------------------------- DOM

    async function requireBoundEl(bindings, role, options, step) {
      const binding = bindings[role];
      if (!binding || !binding.selector) {
        throw stepError(step, `"${role}" isn't bound and has no built-in default.`);
      }
      try {
        return await domUtils.waitFor(() => domUtils.queryVisible(binding.selector), {
          timeout: (options && options.elementTimeout) || 5000
        });
      } catch (e) {
        throw stepError(
          step,
          `couldn't find "${role}" on the page (selector: ${binding.selector}). Rebind it in Setup fields.`,
          domUtils.readMessages()
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

    function countGridRows(bindings) {
      const binding = bindings.entitiesGridRow;
      if (!binding || !binding.selector) return null;
      return domUtils.queryAllVisible(binding.selector).length;
    }

    // Sets a lookup/dropdown to desiredText. Keyboard first (type, arrow
    // down, enter) because that's how D365's own lookups are designed to be
    // driven and needs no knowledge of the dropdown's markup; falling back to
    // clicking a bound option row only if the keyboard path doesn't take.
    async function setLookupValue(fieldEl, optionSelector, desiredText, options) {
      const selectEl = domUtils.resolveSelect(fieldEl);
      if (selectEl && domUtils.selectNativeOption(selectEl, desiredText)) {
        return { value: desiredText, via: 'select' };
      }

      const afterKeyboard = await domUtils.selectByKeyboard(fieldEl, desiredText, {
        settleMs: (options && options.lookupSettleMs) || 400
      });
      if (normalizeText(afterKeyboard) === normalizeText(desiredText)) {
        return { value: afterKeyboard, via: 'keyboard' };
      }

      if (optionSelector) {
        const optionEls = listCandidates(optionSelector);
        if (optionEls.length) {
          const texts = optionEls.map((el) => el.textContent.trim());
          const { candidate, score } = matcher.bestMatch(desiredText, texts);
          if (candidate && score >= 0.5) {
            // A bare sheet name can match both a named range and the
            // whole-sheet "name$" form; prefer the latter.
            const tied = texts.filter((t) => matcher.scoreMatch(desiredText, t) >= score);
            const preferred = tied.find((t) => t.endsWith('$')) || candidate;
            const chosen = optionEls.find((el) => el.textContent.trim() === preferred);
            domUtils.clickElement(chosen);
            await sleep(200);
            return { value: preferred, via: 'click' };
          }
        }
      }

      return { value: domUtils.fieldText(fieldEl), via: 'none' };
    }

    // ---------------------------------------------------------------- steps

    async function openPanel(ctx) {
      if (findBoundEl(ctx.bindings, 'sourceFormatField')) return;

      const addFileEl = await requireBoundEl(ctx.bindings, 'addFileButton', ctx.options, 'add-file');
      domUtils.clickElement(addFileEl);

      const opened = await domUtils
        .waitFor(() => findBoundEl(ctx.bindings, 'sourceFormatField'), {
          timeout: ctx.options.elementTimeout || 5000
        })
        .catch(() => null);

      // D365 sometimes ignores the synthetic sequence; its own click handler
      // still works.
      if (!opened && typeof addFileEl.click === 'function') {
        addFileEl.click();
        await domUtils
          .waitFor(() => findBoundEl(ctx.bindings, 'sourceFormatField'), {
            timeout: ctx.options.elementTimeout || 5000
          })
          .catch(() => null);
      }

      if (!findBoundEl(ctx.bindings, 'sourceFormatField')) {
        throw stepError('add-file', 'the Add file panel never opened.', domUtils.readMessages());
      }
    }

    async function setSourceFormat(ctx) {
      const fieldEl = await requireBoundEl(
        ctx.bindings,
        'sourceFormatField',
        ctx.options,
        'source-format'
      );

      // The panel keeps the previous file's format, so only touch it when it
      // actually differs.
      if (normalizeText(domUtils.fieldText(fieldEl)) === normalizeText(ctx.item.sourceFormat)) {
        return;
      }

      const before = domUtils.readMessages();
      const optionSelector =
        ctx.bindings.sourceFormatOption && ctx.bindings.sourceFormatOption.selector;
      await setLookupValue(fieldEl, optionSelector, ctx.item.sourceFormat, ctx.options);

      const current = await domUtils
        .waitFor(
          () => {
            const el = findBoundEl(ctx.bindings, 'sourceFormatField');
            const text = domUtils.fieldText(el);
            return normalizeText(text) === normalizeText(ctx.item.sourceFormat) ? text : null;
          },
          { timeout: ctx.options.elementTimeout || 5000 }
        )
        .catch(() => null);

      if (!current) {
        throw stepError(
          'source-format',
          `couldn't set the format to "${ctx.item.sourceFormat}" (field shows "${domUtils.fieldText(
            findBoundEl(ctx.bindings, 'sourceFormatField')
          )}").`,
          domUtils.newMessagesSince(before)
        );
      }
    }

    async function setEntityName(ctx) {
      // A data package carries its own manifest, so D365 shows no entity
      // field for it at all.
      if (ctx.item.sourceFormat === 'Package') return;

      const fieldEl = await requireBoundEl(
        ctx.bindings,
        'entityNameField',
        ctx.options,
        'entity-name'
      );
      const before = domUtils.readMessages();

      const suggestionSelector =
        ctx.bindings.suggestionItem && ctx.bindings.suggestionItem.selector;
      const result = await setLookupValue(
        fieldEl,
        suggestionSelector,
        ctx.item.cleanedName,
        ctx.options
      );

      const shown = domUtils.fieldText(findBoundEl(ctx.bindings, 'entityNameField'));
      if (!shown) {
        throw stepError(
          'entity-name',
          `D365 didn't accept "${ctx.item.cleanedName}" as an entity name — the field is empty.`,
          domUtils.newMessagesSince(before)
        );
      }

      // It resolved to something, but not what was asked for: let the user
      // confirm rather than importing the wrong entity.
      if (result.via === 'none' || normalizeText(shown) !== normalizeText(ctx.item.cleanedName)) {
        const suggestions = listCandidates(suggestionSelector).map((el) => el.textContent.trim());
        ctx.needsReview = {
          reason: `The entity field shows "${shown}" rather than "${ctx.item.cleanedName}".`,
          suggestions
        };
        return;
      }

      ctx.update({ matchedEntity: shown });
    }

    async function attachFile(ctx) {
      const fileTargetEl = await requireBoundEl(
        ctx.bindings,
        'fileTarget',
        ctx.options,
        'attach-file'
      ).catch(() => {
        throw stepError(
          'attach-file',
          'the upload box never appeared. For Excel/CSV, D365 only shows it once a valid entity name is selected.',
          domUtils.readMessages()
        );
      });

      const ownerDocument = fileTargetEl.ownerDocument || document;
      const before = domUtils.readMessages();
      domUtils.armFileHook(ctx.item.file, ownerDocument);

      // D365 POSTs the file to /fileUpload; page-hook.js reports when that
      // finishes. It corroborates the filename check below, so it gets a short
      // grace window of its own -- NOT uploadTimeout, which is the multi-minute
      // budget for the Excel driver to produce a grid row. The event can
      // legitimately never arrive (fetch instead of XHR, a cross-origin frame),
      // and blocking minutes per file on it is what stopped a batch part-way.
      const signalTimeout =
        uploadSignalWorks === false ? 0 : ctx.options.uploadSignalTimeout || 20000;
      const uploadResponse =
        signalTimeout > 0
          ? domUtils.waitForUploadResponse(signalTimeout, ownerDocument)
          : Promise.resolve(null);

      const browseEl = findBoundEl(ctx.bindings, 'uploadButton') || fileTargetEl;
      domUtils.clickElement(browseEl);

      const fired = await domUtils
        .waitFor(() => domUtils.fileHookFired(), { timeout: 2000, interval: 100 })
        .catch(() => false);

      if (!fired) {
        const fileInput = domUtils.resolveFileInput(fileTargetEl);
        if (fileInput) {
          domUtils.dropFileOnInput(fileInput, ctx.item.file);
        } else {
          domUtils.disarmFileHook(ownerDocument);
          throw stepError(
            'attach-file',
            'couldn\'t hand the file to D365 — no file input was reachable and the Upload button didn\'t ask for one.',
            domUtils.newMessagesSince(before)
          );
        }
      }

      domUtils.disarmFileHook(ownerDocument);

      // The filename box should now show the file; if it doesn't, the upload
      // silently didn't take and later steps would fail for the wrong reason.
      const accepted = await domUtils
        .waitFor(
          () => {
            const text = domUtils.fieldText(findBoundEl(ctx.bindings, 'fileTarget'));
            return text ? text : null;
          },
          { timeout: ctx.options.elementTimeout || 5000 }
        )
        .catch(() => null);

      if (!accepted) {
        throw stepError(
          'attach-file',
          'the file was handed over but D365 never showed a file name in the upload box.',
          domUtils.newMessagesSince(before)
        );
      }

      // A null here means the signal never came, which proves nothing either
      // way -- D365 already showed the file name, so the run carries on.
      const response = await uploadResponse;
      if (response) uploadSignalWorks = true;
      else if (uploadSignalWorks === null) uploadSignalWorks = false;

      if (response && !response.ok) {
        throw stepError(
          'attach-file',
          `D365 rejected the upload (HTTP ${response.status}).`,
          domUtils.newMessagesSince(before)
        );
      }
      ctx.uploadConfirmed = !!response;
    }

    async function setSheetSelection(ctx) {
      if (ctx.item.sourceFormat === 'Package') return;
      if (!ctx.item.selectedSheet) return;

      const sheetEl = await domUtils
        .waitFor(() => findBoundEl(ctx.bindings, 'sheetSelectField'), {
          timeout: ctx.options.elementTimeout || 5000
        })
        .catch(() => null);

      // Single-sheet workbooks never prompt, so a missing picker is normal.
      if (!sheetEl) return;

      if (normalizeText(domUtils.fieldText(sheetEl)) === normalizeText(ctx.item.selectedSheet)) {
        return;
      }

      const before = domUtils.readMessages();
      const optionSelector = ctx.bindings.sheetOption && ctx.bindings.sheetOption.selector;
      await setLookupValue(sheetEl, optionSelector, sheetLookupText(ctx.item.selectedSheet), ctx.options);

      // Sheet names are read from this very file, so unlike an entity name
      // there's no label/technical mismatch to excuse a miss.
      const settled = await domUtils
        .waitFor(
          () => {
            const text = domUtils.fieldText(findBoundEl(ctx.bindings, 'sheetSelectField'));
            return normalizeText(text) === normalizeText(ctx.item.selectedSheet) ? text : null;
          },
          { timeout: ctx.options.elementTimeout || 5000 }
        )
        .catch(() => null);

      if (!settled) {
        throw stepError(
          'sheet',
          `couldn't set the sheet to "${sheetLookupText(
            ctx.item.selectedSheet
          )}" — D365's sheet lookup needs a real selection from its list, not just typed text.`,
          domUtils.newMessagesSince(before)
        );
      }
    }

    async function waitForUpload(ctx) {
      const timeout = ctx.options.uploadTimeout || 300000;
      if (ctx.baselineGridRows === null) {
        await sleep(ctx.options.stepDelay || 700);
        return;
      }

      try {
        await domUtils.waitFor(() => countGridRows(ctx.bindings) > ctx.baselineGridRows, {
          timeout,
          interval: 500
        });
      } catch (e) {
        throw stepError(
          'upload',
          `no new row appeared in the entities grid within ${Math.round(
            timeout / 1000
          )}s. D365 warns Excel imports can queue for the Excel driver, so this may just need a longer wait in Settings.`,
          domUtils.readMessages()
        );
      }
    }

    async function waitForPanelReset(ctx) {
      const binding = ctx.bindings.entityNameField;
      if (!binding || !binding.selector) return;
      // D365 clears the panel after an upload; typing into it mid-render
      // loses the value, so let it settle before the next file.
      await domUtils
        .waitFor(
          () => {
            const el = domUtils.queryVisible(binding.selector);
            return el && !domUtils.fieldText(el);
          },
          { timeout: ctx.options.elementTimeout || 5000 }
        )
        .catch(() => null);
    }

    const STEPS = [
      { name: 'add-file', run: openPanel },
      { name: 'source-format', run: setSourceFormat },
      { name: 'entity-name', run: setEntityName },
      { name: 'attach-file', run: attachFile },
      { name: 'sheet', run: setSheetSelection },
      { name: 'upload', run: waitForUpload },
      { name: 'panel-reset', run: waitForPanelReset }
    ];

    // ------------------------------------------------------------- running

    async function runSteps(item, bindings, options, fromStep) {
      const ctx = {
        item,
        bindings,
        options,
        baselineGridRows: countGridRows(bindings),
        needsReview: null,
        update: (patch) => updateItem(item.id, patch)
      };

      const startIndex = fromStep ? STEPS.findIndex((s) => s.name === fromStep) : 0;
      for (let i = Math.max(0, startIndex); i < STEPS.length; i++) {
        const step = STEPS[i];
        updateItem(item.id, { status: 'running', step: step.name });
        // Steps read the item fresh, since sheet choice can change mid-queue.
        ctx.item = items.find((it) => it.id === item.id) || item;
        await step.run(ctx);

        if (ctx.needsReview) {
          updateItem(item.id, {
            status: 'needs-review',
            step: step.name,
            error: ctx.needsReview.reason,
            suggestions: ctx.needsReview.suggestions || []
          });
          return { needsReview: true };
        }
      }

      updateItem(item.id, { status: 'filled', step: null, error: null });
      return { filled: true };
    }

    // After a failure the Add file panel can be left half-filled, which would
    // derail the next file too. Closing it makes the next item start clean.
    async function recoverPanel(bindings, options) {
      try {
        const closeEl = findBoundEl(bindings, 'closePanelButton');
        if (!closeEl) return;
        domUtils.clickElement(closeEl);
      } catch (e) {
        // The panel may already be gone, or the element torn down mid-click.
        // Either way the next file re-opens it; this must not end the batch.
      }
      await sleep((options && options.stepDelay) || 700);
    }

    // Every wait inside a step is bounded, but a step that somehow never
    // returns would hang the whole batch with no way back. The watchdog is
    // the backstop: deliberately generous (past even the slowest legitimate
    // upload), it guarantees the queue always moves on to the next file.
    function withWatchdog(promise, options) {
      const opts = options || {};
      const limit = opts.itemTimeout || (opts.uploadTimeout || 300000) + 120000;
      let timer;
      const watchdog = new Promise((_, reject) => {
        timer = setTimeout(() => {
          reject(
            stepError(
              'watchdog',
              `this file made no progress for ${Math.round(
                limit / 1000
              )}s, so the run moved on to the next one.`
            )
          );
        }, limit);
      });
      return Promise.race([promise, watchdog]).finally(() => clearTimeout(timer));
    }

    async function processItem(id, bindings, options, fromStep) {
      const item = items.find((it) => it.id === id);
      if (!item) return null;

      try {
        return await withWatchdog(runSteps(item, bindings, options, fromStep), options);
      } catch (err) {
        updateItem(id, { status: 'error', error: err.message, step: err.step || null });
        return { error: err.message };
      }
    }

    // One bad file shouldn't strand the other 36: a failure is recorded
    // against that item and the run moves on, with a summary at the end.
    //
    // The whole loop is guarded. Anything thrown outside processItem's own
    // catch -- a stale element in recoverPanel, a render error reaching back
    // through emit -- used to escape here and leave `running` stuck true, so
    // every later press of Upload returned immediately and did nothing. The
    // batch looked like it simply stopped part-way, permanently, until the
    // page was reloaded.
    async function run(bindings, options) {
      if (running) return summary();
      running = true;
      paused = false;
      let aborted = null;

      try {
        while (!paused) {
          const next = items.find((it) => it.status === 'pending');
          if (!next) break;

          const result = await processItem(next.id, bindings, options || {});
          if (result && (result.error || result.needsReview)) {
            await recoverPanel(bindings, options);
          }
          await sleep((options && options.stepDelay) || 700);
        }
      } catch (err) {
        // Record it against the file it happened on, so the panel shows why
        // the run ended rather than just going quiet.
        aborted = err;
        const current = items.find((it) => it.status === 'running');
        if (current) {
          updateItem(current.id, {
            status: 'error',
            error: `the run stopped unexpectedly: ${err.message}`
          });
        }
      } finally {
        running = false;
        emit();
      }

      return Object.assign(summary(), { aborted: aborted ? aborted.message : null });
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

      const importEl = await requireBoundEl(bindings, 'runImportButton', options, 'import');
      domUtils.clickElement(importEl);
    }

    function pause() {
      paused = true;
    }

    // The user resolved an ambiguous entity name: apply their choice, then
    // re-enter the same pipeline at the following step rather than repeating
    // the tail of it here (which is how the sheet step once went missing).
    async function resumeAfterReview(id, chosenText, bindings, options) {
      const item = items.find((it) => it.id === id);
      if (!item) return;

      const fieldEl = findBoundEl(bindings, 'entityNameField');
      if (fieldEl) {
        await domUtils.selectByKeyboard(fieldEl, chosenText, {
          settleMs: (options && options.lookupSettleMs) || 400
        });
      }
      updateItem(id, { matchedEntity: chosenText, error: null, suggestions: [] });

      await processItem(id, bindings, options || {}, 'attach-file');
      run(bindings, options);
    }

    function retry(id) {
      updateItem(id, { status: 'pending', error: null, step: null });
    }

    function skip(id) {
      updateItem(id, { status: 'skipped', error: null });
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
      retry,
      skip,
      summary,
      isPaused: () => paused,
      isRunning: () => running
    };
  }

  D365IA.queue = { createQueue, sourceFormatFor, sheetLookupText };
})();

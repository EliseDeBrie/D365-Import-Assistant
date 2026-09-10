# D365 Import Assistant

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

An Edge/Chromium extension for the Dynamics 365 Finance & Operations
**Data management → Import** screen. Drag a folder of Excel files onto it and
it auto-fills the D365 "Entity name" field for each one, guessing the entity
from the cleaned-up file name, then attaches the file — no typing per file.
![Uploading import-demo.gif…]()

> **Status: beta.** Test it in a sandbox environment before you use it
> against production, and see *Known limitations* below.

Licensed under [MIT](LICENSE). See [`PRIVACY.md`](PRIVACY.md) for exactly what the
extension can see and where it's stored (short version: nothing leaves your
browser except calls to your own D365 tenant).

### Designed to work across D365 F&O environments

There is no tenant URL to configure, and nothing to change when a sandbox is
rebuilt or you move to a new tenant. The extension's permissions
(`host_permissions` in `manifest.json`) are a wildcard —
`https://*.dynamics.com/*` — so it activates on whichever D365 environment
you're on, the moment you're on it. Field bindings are stored per environment
hostname (see *Why bindings prefer `data-dyn-controlname`* below), so a new
environment starts on the shipped defaults until you rebind anything for it —
no manual setup, no per-tenant install. Open the toolbar popup on any D365
page and it shows *Active on: \<hostname\>*, confirming it picked up the new
environment.

The cached entity list is *not* per environment: it's stored once per browser
profile, so after switching tenants press **Load** again in the panel to
refresh it for the environment you're now on.

It drives D365's standard user interface, so a customised environment or a
future Microsoft interface change can require rebinding — see *Known
limitations*.

## How it works

1. A floating **Import Assistant** button appears on the Data management pages
   only — by default, URLs containing `mi=DM_DataManagementWorkspaceMenuItem`
   — so it stays out of the way the rest of the time. Drag it anywhere on
   screen and it stays there; change which pages it appears on, or switch it
   off entirely, in the options page under *Where the button appears*. Click
   it to open the drop panel.
2. Drag one or many files anywhere onto the page (the whole page is a
   drop target, not just the small box — dropping on the launcher button
   works too, and it'll open the panel for you). Or click the panel's box to
   browse instead.

   The queue is sorted by the file names' numbering, not by the arbitrary
   order the OS hands the drop over, and each row shows its position. Numeric
   runs sort numerically, so `10_` comes after `2_`, and `04.12PMD-` before
   `80.ALOG.WM-`. Import order matters when the files depend on each other.

   Workbooks are read locally to list their sheets (an `.xlsx` is a zip
   holding `xl/workbook.xml`). Anything with more than one sheet gets a sheet
   picker on its row, defaulting to the first sheet — that's what D365 will
   be told to import.

   The picker shows plain sheet names, but D365 reaches Excel through the
   ODBC driver, which names every worksheet `Sheet$`. That trailing `$` is
   what separates the whole sheet from a named range over part of it, and
   D365's sheet lookup will not accept a name without it. The extension adds
   it when driving the lookup, so you never have to.
3. For each file, the file name is cleaned up (leading/trailing sequence
   numbers, dates/timestamps, and `_`/`-` separators are stripped — see
   *Cleaning rules* below) to produce a guessed entity name, and the source
   format is picked from the extension: `.csv` → **CSV**, `.zip` → **Package**,
   `.xlsx`/`.xlsm`/`.xls` → **Excel**. Each row shows the format it detected.
   A data package is handled differently throughout — it carries its own
   manifest, so D365 asks it for neither an entity name nor a sheet, and the
   extension skips both steps for it.
4. Click **Upload** (or **Upload + Import**). For each queued file, the
   extension replays D365's own click sequence as a named pipeline —
   `add-file`, `source-format`, `entity-name`, `attach-file`, `sheet`,
   `upload`, `panel-reset` — where every step verifies its own outcome, so a
   failure names the step that actually failed and quotes D365's message bar:
   - clicks **Add file** — but only if that panel isn't already open, since
     clicking it again would close it. From the second file onward the panel
     stays open, so this step is skipped,
   - checks the **Source data format** and only opens the dropdown if it
     needs changing (it keeps the previous file's value),
   - types the cleaned name into the **Entity name** field, picks D365's
     best-scoring autocomplete suggestion where one is available, and
     confirms the field actually kept the value,
   - attaches the Excel file — in the currently supported D365 import flow,
     the upload control becomes available once the entity has been selected,
   - waits for D365 to confirm the file landed (see *How the extension knows
     a file landed* below), then for the panel to reset, before moving on.
5. If no suggestion is a confident match, that file is marked
   **needs-review** — pick the right entity from a dropdown (built from
   D365's own suggestions) or skip it. Nothing is ever typed in blind without
   going through D365's real autocomplete, so a bad guess can't silently
   attach the wrong entity.

   The rest of the batch keeps going regardless. One file that fails or needs
   a decision never strands the other 36; each row carries its own status and
   Retry/Skip buttons, and the run ends with a count of what happened.

### Dialogs D365 raises mid-run

D365 interrupts with modal message boxes, and while one is up nothing else on
the page responds — so an unanswered dialog doesn't slow a batch down, it
stops it, and every step behind it times out reporting something that isn't
the real problem.

While a run is in progress the extension watches for these and answers the
ones it recognises. Right now that is one:

| Dialog | Answer | Why |
| --- | --- | --- |
| *"The sheet with the same name is already mapped in this project. Do you still want to continue?"* | **Yes** | Raised when two workbooks in the project share a sheet name — which is every file when they all carry an `en_us` sheet. Continuing is the point of the batch, and the dialog defaults to **No**. |

**Nothing else is ever clicked.** Auto-answering a confirmation you don't
recognise is how an automation does real damage — these same dialogs are
where *delete*, *overwrite* and *publish* live. An unrecognised dialog is
left exactly as it is, and its text is attached to the failure of whatever
step stalled behind it, so the panel says *"unanswered dialog: …"* rather
than something misleading about a missing field.

To teach it another prompt, add an entry to `KNOWN_PROMPTS` in
`content/queue.js` — a regex for the text and the button label to press.

### How the extension knows a file landed

Whichever of these arrives first, per file:

1. **D365's `/fileUpload` response** — the POST returning 2xx, seen by
   `page-hook.js` in whichever frame made the request.
2. **D365's own message bar** — *"'Customer Groups' entity mapping has
   completed successfully"*.
3. **The Add file panel clearing itself**, which D365 only does once it has
   taken the file — counted only for a file that actually put a value in the
   entity field, since an empty panel is also how it starts out. Some D365
   versions rebuild it blank, others tear it down; both count.
4. **A new row in the entities grid**, if `entitiesGridRow` is bound.

It used to wait on (4) alone. That needs an optional binding, and D365 renders
that grid through React with virtualised rows, so the count often doesn't move
even when the upload plainly succeeded — leaving a batch parked at `[UPLOAD]`
for the full five-minute timeout per file while the message bar on screen
already said the file was in. Each row now shows which signal confirmed it,
so a stall is diagnosable rather than mysterious.

### Why a batch can't stall part-way

Three separate guards, because a 37-file run that quietly stops after two is
worse than one that fails loudly:

- **Nothing waits without a deadline.** The completion signal from
  `page-hook.js` is corroboration, not a gate — it can legitimately never
  arrive (the request went through `fetch` rather than XHR, or ran in a
  cross-origin frame), so it gets a short window of its own rather than the
  multi-minute budget meant for the Excel driver. If the first file shows the
  signal isn't reaching us in this environment, later files skip that wait
  entirely instead of each paying it again.
- **The run loop can't die.** Anything thrown outside a step's own handling
  is caught, recorded against the file it happened on, and the queue is
  released. It used to escape and leave the queue flagged as running, so
  every later press of Upload returned immediately and did nothing — the
  batch looked permanently stopped until the page was reloaded.
- **A per-file watchdog.** Deliberately generous (past even the slowest
  legitimate upload), it exists only so a file that somehow makes no progress
  is failed and stepped over rather than hanging the batch.

## Two run buttons

- **Upload** — uploads every queued file as an entity and stops there,
  leaving the Add file panel open and the import unstarted.
- **Upload + Import** — does the same, then closes the Add file panel and
  clicks the page-level Import button, so a whole batch can be started and
  left to run.

The import step is a separate button rather than a setting, because it's the
one that actually loads data into D365. It only fires when *every* queued
file uploaded successfully — a batch that paused for review, errored, or had
a file skipped never triggers an import — and needs `runImportButton` bound
(plus `closePanelButton` to close the panel first).

## One-time setup: bind fields

**You normally don't need to do this.** The extension ships working selectors
for every field, built on D365's own `data-dyn-controlname` attributes, which
are stable across sessions and rebuilds. Binding is an *override* for an
environment where a shipped default doesn't resolve — Setup fields shows which
rows are running on a default and which on your own binding, and clearing a
binding falls back to the default rather than leaving the row unusable.

Bindings are stored per environment host, so a dev tenant and production can
differ. If you do need to bind, the setup list is in the same order as the
real click sequence:

1. Open the Data management → Import screen you use.
2. Click **Import Assistant → Setup fields**.
3. Click **Bind** next to a row. The setup dialog hides itself and the page
   behaves completely normally again — click into fields, type, open
   dropdowns, whatever it takes to reveal the actual element you want. When
   the right element is visible, hold **Alt** and click it to bind it (a
   plain click does nothing to the picker — it only reaches the page). Press
   **Esc** to cancel a bind in progress, or close the dialog with the **×**,
   **Done**, or by clicking outside it.
   - **Add file button** — click it once yourself first so a row exists,
     then Alt+click the button itself.
   - **Source data format dropdown** — while it's still *closed*, Alt+click
     anywhere in the box (the value or placeholder text is fine — not
     specifically the little arrow icon). Don't open it first for this one.
   - **Source data format option** — a *different* element from the box
     above. First click the box normally (no Alt) so its list opens, *then*
     Alt+click one option (e.g. "Excel") inside that open list. This needs
     to match *any* option in the list, not just the one you clicked — if
     picking a different file later lands on the wrong option, rebind on an
     option in a different position and see **Notes on binding a list**
     below.
   - **Entity name field** — pick a format first so the field appears, then
     Alt+click it directly.
   - **Suggestion row** — type a few letters into the entity field so D365's
     autocomplete opens, *then* Alt+click one suggestion row in that list.
   - **File target** — the file input or drop target for that row.
   - **Upload button** — the row's Upload button.
   - **Entities grid row** *(optional)* — add one file manually first so a
     row exists in the grid of already-uploaded entities, then Alt+click that
     row. This is one more upload-confirmation signal, not the only one: the
     extension also confirms a file landed through D365's `/fileUpload`
     response, its message bar, or the Add file panel clearing (see *How the
     extension knows a file landed* above), so this binding is normally not
     needed.
   - **Import/Run button** *(optional)* — the page-level button that actually
     runs the import job. Needed only for **Upload + Import**; a plain
     **Upload** run never touches it. Bind **Close panel button** too, since
     the panel is closed before the import starts.
4. Bindings are saved (`chrome.storage.sync`) and reused automatically next
   time — you shouldn't need to redo this unless D365 changes its layout.

### Why bindings prefer `data-dyn-controlname`

D365 stamps instance counters into element ids — `31_5_SourceNameControl_input`
— and regenerates them every time it rebuilds a control. The "Add file" panel
is torn down and rebuilt on every open, so a plain `#id` selector captured
during setup is stale by the next run and matches nothing. Selector
generation therefore prefers, in order:

1. the element's own `data-dyn-controlname`,
2. the nearest ancestor's `data-dyn-controlname` plus a short path down to it,
3. an id matched by its stable suffix (`[id$="_SourceNameControl_input"]`),
4. `name`, then a plain id, then a DOM path.

If a binding still goes stale, check what it saved in Options → Field
bindings: anything that looks like a bare `#12_3_Something` is the fragile
case, and rebinding by clicking slightly higher up (on the control's box
rather than deep inside it) usually lands on an element with a control name.

### Notes on binding a list (suggestion row / dropdown option / grid row)

These three roles all need to match *every* item in a repeating list, not
just the one you clicked — otherwise the extension could only ever pick the
first suggestion, or the first dropdown option, no matter what it should
actually be. The picker handles this by generalizing the clicked element's
own CSS class to match its siblings, rather than pinning to its position.
That works well as long as list items share a class D365 doesn't reuse
elsewhere on the page. If matching seems to only ever land on the item you
originally bound, rebind by Alt+clicking a *different* item in that same
list — if it still gets the same result, the class is being shared with
something unrelated on the page and the selector needs a manual edit
(current selectors are visible in Options → Field bindings).

## Checking names against the real entity list

The panel's **Load** button fetches `<your environment>/data` — D365's own
OData service document, which lists every entity this tenant exposes. It's
same-origin, so the session already open in the tab authenticates it, and the
list is right for *this* environment rather than a hardcoded one that drifts
with version and customizations. It's cached in `chrome.storage.local`
(thousands of names exceed `storage.sync`'s per-item quota).

Each queued file then shows a hint badge for whether its cleaned name
corresponds to a real entity. **Treat it as a hint, not an error report**:
OData's service document lists *technical* entity names (`OperationalSitesV2`),
while D365's own Entity name lookup shows *display labels* (`Sites V2`) —
often legitimately different strings for the same entity. A "not in entity
list" badge is common and frequently harmless; what matters is what D365
itself shows in the Entity name field once a file has run.

What gets *typed* into D365 is always what the queue shows — the OData
spelling is never substituted in. The one place the entity list actually
affects automation: when D365 *is* offering a live list of suggestions to
pick from, a validated real name is tried as an extra candidate alongside the
typed guess, and whichever scores better against those live suggestions
wins. It never overrides or rejects what D365 itself already put in the
field — the technical/label mismatch means it isn't reliable enough for that.

## Cleaning rules

Configurable in the extension's **Options** page (right-click the toolbar
icon → Options, or from the popup):

- Strip leading numbers and code prefixes — `01_Customers.xlsx` → `Customers`,
  `02B.03SYS-Inventory adjustment journal names.xlsx` → `Inventory adjustment
  journal names`. Any leading token starting with a digit is treated as a
  numbering/classification code.
- Strip trailing numbers — `Customers_01.xlsx` → `Customers`
- Strip dates/timestamps — `Customers_20240115.xlsx` → `Customers`
- Strip version words (`v2`, `(2)`, `copy`, `final`, `draft`) — off by default
- Turn `_`/`-` into spaces — `Customer_Groups.xlsx` → `Customer Groups`
- Title Case the result

Rules only strip tokens from the **edges** of the file name, never from the
middle — so `CustomersV3` keeps its `3` (it's part of a word, not a separate
numeric token) and `Address2` is untouched. If a junk token like `final` sits
after a date and "strip version words" is off, the date stays put since it's
no longer at the edge — turn that rule on if your files follow that pattern.

The Options page has a live "Try it" box to test a file name against your
current rules before running a real import.

## Loading the extension in Edge (unpacked)

1. Go to `edge://extensions`.
2. Turn on **Developer mode** (bottom-left toggle).
3. Click **Load unpacked** and select this folder.
4. Pin the extension if you want quick access to the popup/status.

## Building a package to distribute

```
npm run package
```

Writes `dist/D365ImportAssistant-v<version>.zip`, ready to upload to the Edge
Add-ons or Chrome Web Store.

### Versioning, and the beta marker

`manifest.json` carries two version fields, because it has to:

- **`version`** (`0.1.0`) is what the store orders releases by. It must be
  one to four dot-separated integers — the browser rejects a manifest whose
  `version` contains anything else, so `0.1.0-beta` is not an option here.
- **`version_name`** (`0.1.0 beta`) is free text, and is what the browser
  actually shows in `edge://extensions`. This is where the beta marker lives.

The packaged file name is derived from `version_name` when it's present, so a
beta build can't end up in a file called plain `v0.1.0` — which is exactly
how a hand-named archive and the manifest inside it drift apart.

The number starts at `0.1.0` because that is what this is: a first public
beta. Earlier builds carried higher numbers, but those counted development
iterations rather than releases, and shipping a first public version as
`1.12.0` would imply eleven releases that never happened.

**To ship a stable release:** delete the `version_name` line, bump `version`,
and drop the *Status: beta* notice at the top of this file. Note that the
store rejects re-uploading a version number it has already seen, so `version`
has to increase for every submission, beta or not.

Use this rather than zipping the folder by hand. A hand-made archive puts
everything under a top-level folder, and **both stores reject a zip whose
`manifest.json` isn't at the archive root** — while *Load unpacked* accepts
it either way, so the mistake survives local testing and only shows up at
submission. The script writes paths relative to the extension root, so the
shape is correct by construction.

It also ships only what the extension needs (`manifest.json`, `background.js`,
`content/`, `icons/`, `options/`, `popup/`, `styles/`, plus `LICENSE` and
`PRIVACY.md`) — the tests, this README and `package.json` stay out. The list
is an allow-list in `scripts/package.js`, so anything added to the repo later
has to be named there before it reaches users. Before writing the zip it
checks that every file `manifest.json` references is actually present, and
fails the build if one isn't.

## Project layout

```
manifest.json          Manifest V3 config
background.js          Sets default settings on install
content/
  dom-utils.js          Low-level DOM helpers (native value setting, file
                         attachment via DataTransfer, drag/drop replay,
                         list-item vs. single-element selector generation)
  matcher.js             Filename cleaning + suggestion/option scoring
  binder.js               Click-to-bind element picker (Alt+click to
                            confirm) + storage
  queue.js                 Batch queue: replays Add file -> Source format
                            -> Entity name -> File -> Upload per file,
                            verifies each step, pauses on ambiguity, waits
                            for whichever upload-confirmation signal arrives
  dropzone.js               Floating drop panel UI
  setup-modal.js             "Bind D365 fields" modal
  content.js                  Wires it all together, injects the launcher,
                                catches drops anywhere on the page
options/                Settings page (cleaning rules, matching, bindings)
popup/                  Toolbar popup (status + link to settings)
styles/dropzone.css     All injected UI styling
icons/                  Toolbar/extension icons
test/                   jsdom test suite (see below)
scripts/package.js      Builds the store-ready zip (npm run package)
```

## Running the tests

```
npm install
npm test
```

The suite runs the real content scripts in [jsdom](https://github.com/jsdom/jsdom)
against `test/fake-d365.js` — a stand-in for the Import form that copies the
behaviours that actually caused bugs here:

- the entity name field only exists for Excel/CSV, never for a data package;
- in the currently supported D365 import flow, the upload box only appears
  once an entity name has been *committed*;
- the lookups **discard typed text** — a value only sticks if it was chosen
  from the list, which is what broke the sheet picker;
- the panel comes back blank after each upload, which is what broke the
  second file in a batch.

So `npm test` covers a real multi-file run, a `.zip` going in as `Package`,
sheet selection, upload ordering, and per-file failure recovery, without a
browser or a D365 tenant. Zip parsing is tested against actual zip bytes
built in `test/xlsx-sheets.test.js`, including the streaming-mode layout that
leaves the sizes out of the local header.

jsdom has no layout engine and implements neither `Blob.stream()` nor
`DataTransfer`; `test/harness.js` stands in for those and says why in each
case. Nothing in `content/` is stubbed.

## Notes on how file attachment works

**You never tell the extension where your files live, and there's no folder
to configure.** Browsers won't let a script read a file from disk on its own —
the user has to supply it. Dragging files onto the panel *is* that supply
step: the browser hands over the real `File` objects, contents included. The
queue then feeds them to D365 one per Add-file cycle, in drop order.

The obstacle is that D365's "Upload and add" button opens the operating
system's file picker, which no extension can drive or dismiss. So
`content/page-hook.js` runs in the **page's** JavaScript world (via the
manifest's `"world": "MAIN"`, unlike every other script here, which runs
isolated) and replaces `HTMLInputElement.prototype.click` and `showPicker`.
When D365 asks for the picker, the hook answers with the already-dropped
file instead of opening it — same code path D365 would run after you picked
the file by hand, minus the dialog.

The file itself crosses from the extension's isolated world into the page
world through a hidden `<input type="file">` in the shared DOM, since a
`File` can't be passed in a cross-world event payload.

If the hook doesn't fire (a D365 version that opens the picker some other
way), the queue falls back to writing the file straight into a reachable
`<input type="file">`, and reports a clear error if neither works.

## Known limitations

- It's a beta. Test it in a sandbox before running it against production.
- It drives D365's standard user interface rather than an API, so a
  customised environment or a future Microsoft interface change can require
  rebinding a field.
- The success message it watches for on D365's message bar is matched in
  **English**. On a differently-localised D365 that signal never fires, and
  confirmation falls back to the upload response, the panel reset, or the
  entities-grid row.
- Entity matching is only as good as your file names plus the suggestions
  D365 itself returns. Nothing is typed in blind: an unconfident match is
  parked as **needs-review** for you to resolve.
- The entity-name hint compares against OData's *technical* names
  (`OperationalSitesV2`), while the import form shows *display labels*
  (`Sites V2`), so a "not in entity list" badge is common and often harmless.
- The cached entity list is stored once per browser profile, not per
  environment — press **Load** again after switching tenants.
- Data packages (`.zip`) carry their own manifest, so entity and sheet
  selection are skipped for them.
- Only the one documented duplicate-sheet confirmation is answered
  automatically; any other D365 dialog is left untouched and reported.
- Distributed as an unpacked Edge/Chromium extension, which a managed device
  may block. Validate it on the browser you intend to use.
- It prepares — and optionally starts — an import. It does not validate that
  the data itself is correct.

## License

[MIT](LICENSE).

In practice, this means:

- **You can use, study, modify, redistribute and sell this code freely**,
  including inside a closed-source product. Install it across your whole
  organization, fork it, change it to fit your own D365 environment — the
  only condition is that the copyright notice in `LICENSE` travels with it.
- **There is no warranty.** The license is provided AS IS, per the text in
  `LICENSE`.

MIT was chosen deliberately over a copyleft license. This is a client-side
browser extension, so the "hosted service" clause that an AGPL license exists
to enforce has nothing to attach to here — while the license itself would be
enough for some organizations' policies to block installing it at all. The
point of this project is that people can actually run it.

See [`CONTRIBUTING.md`](CONTRIBUTING.md) if you'd like to submit changes; a
pull request is accepted under the same license as the rest of the repo, no
separate agreement needed. Contributors keep the copyright in their own
changes, which are licensed to everyone — including the maintainers — under
the MIT terms like the rest of the project.

## Privacy

See [`PRIVACY.md`](PRIVACY.md) for exactly what the extension can access
and where it stores it. Short version: it only runs on `*.dynamics.com`
pages, nothing it reads or stores ever leaves your browser except a call to
your own tenant's OData service (for the entity-name hint) and the upload
D365 itself was already going to do — there's no backend, no analytics, and
nothing sent to the developer.

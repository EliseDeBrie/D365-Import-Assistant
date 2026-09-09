# D365 Import Assistant

An Edge/Chromium extension for the Dynamics 365 Finance & Operations
**Data management → Import** screen. Drag a folder of Excel files onto it and
it auto-fills the D365 "Entity name" field for each one, guessing the entity
from the cleaned-up file name, then attaches the file — no typing per file.

## How it works

1. A floating **Import Assist** button appears on the Data management pages
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
3. For each file, the file name is cleaned up (leading/trailing sequence
   numbers, dates/timestamps, and `_`/`-` separators are stripped — see
   *Cleaning rules* below) to produce a guessed entity name, and the source
   format is picked from the extension: `.csv` → **CSV**, `.zip` → **Package**,
   `.xlsx`/`.xlsm`/`.xls` → **Excel**. Each row shows the format it detected.
   A data package is handled differently throughout — it carries its own
   manifest, so D365 asks it for neither an entity name nor a sheet, and the
   extension skips both steps for it.
4. Click **Start**. For each queued file, the extension replays D365's own
   click sequence:
   - clicks **Add file** — but only if that panel isn't already open, since
     clicking it again would close it. From the second file onward the panel
     stays open, so this step is skipped,
   - checks the **Source data format** and only opens the dropdown if it
     needs changing (it keeps the previous file's value),
   - types the cleaned name into the **Entity name** field, picks D365's
     best-scoring autocomplete suggestion where one is available, and
     confirms the field actually kept the value,
   - attaches the Excel file — the upload box only renders once a valid
     entity is selected,
   - waits for a new row to appear in the entities grid (if that's bound),
     then for the panel to reset, before moving to the next file.
5. If no suggestion is a confident match, that file is marked
   **needs-review** and the queue pauses — pick the right entity from a
   dropdown (built from D365's own suggestions) or skip the file. Nothing is
   ever typed in blind without going through D365's real autocomplete, so a
   bad guess can't silently attach the wrong entity.

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

D365's DOM differs by version, environment and customization, so instead of
hardcoding CSS selectors that could silently break, the extension asks you to
point at the real elements once. The setup list is in the same order as the
real click sequence:

1. Open the Data management → Import screen you use.
2. Click **Import Assist → Setup fields**.
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
   - **Entities grid row** *(optional but recommended)* — add one file
     manually first so a row exists in the grid of already-uploaded
     entities, then Alt+click that row. This is how the extension knows a
     file finished uploading; without it, it just waits a fixed pause
     (configurable in Settings) and assumes success.
   - **Import/Run button** *(optional)* — the page-level button that actually
     runs the import job. Only clicked if you turn on auto-run in Settings;
     otherwise it's never touched.
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
                            waits for the entities grid, pauses on ambiguity
  dropzone.js               Floating drop panel UI
  setup-modal.js             "Bind D365 fields" modal
  content.js                  Wires it all together, injects the launcher,
                                catches drops anywhere on the page
options/                Settings page (cleaning rules, matching, bindings)
popup/                  Toolbar popup (status + link to settings)
styles/dropzone.css     All injected UI styling
icons/                  Toolbar/extension icons
test/                   jsdom test suite (see below)
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
- the upload box only appears once an entity name has been *committed*;
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

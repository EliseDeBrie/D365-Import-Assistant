# D365 Import Assistant

An Edge/Chromium extension for the Dynamics 365 Finance & Operations
**Data management → Import** screen. Drag a folder of Excel files onto it and
it auto-fills the D365 "Entity name" field for each one, guessing the entity
from the cleaned-up file name, then attaches the file — no typing per file.

## How it works

1. A floating **Import Assist** button appears on any `*.dynamics.com` page.
   Click it to open the drop panel.
2. Drag one or many Excel files onto the panel (or click it to browse).
3. For each file, the file name is cleaned up (leading/trailing sequence
   numbers, dates/timestamps, and `_`/`-` separators are stripped — see
   *Cleaning rules* below) to produce a guessed entity name.
4. Click **Start**. For each queued file, the extension:
   - types the cleaned name into D365's own Entity name field,
   - waits for D365's real autocomplete suggestions to appear,
   - picks the best-scoring suggestion (exact/near match) and selects it,
   - attaches the Excel file to that row,
   - optionally clicks "Add row" and moves to the next file.
5. If no suggestion is a confident match, that file is marked
   **needs-review** and the queue pauses — pick the right entity from a
   dropdown (built from D365's own suggestions) or skip the file. Nothing is
   ever typed in blind without going through D365's real autocomplete, so a
   bad guess can't silently attach the wrong entity.

The extension never auto-clicks D365's final **Import**/submit button —
you always review the filled-in rows and trigger the actual import yourself.

## One-time setup: bind fields

D365's DOM differs by version, environment and customization, so instead of
hardcoding CSS selectors that could silently break, the extension asks you to
point at the real elements once:

1. Open the Data management → Import screen you use.
2. Click **Import Assist → Setup fields**.
3. Click **Bind** next to each row, then click the matching element on the
   page:
   - **Entity name field** — the box you type the entity into.
   - **Suggestion row** — open the autocomplete dropdown first (type
     anything into the entity field), then click one suggestion row.
   - **File target** — the file input or drop target for that grid row.
   - **Add row button** *(optional)* — the button that adds a new import line.
   - **Import/Submit button** *(optional)* — only needed if you turn on
     auto-submit later; not auto-clicked otherwise.
4. Bindings are saved (`chrome.storage.sync`) and reused automatically next
   time — you shouldn't need to redo this unless D365 changes its layout.

## Cleaning rules

Configurable in the extension's **Options** page (right-click the toolbar
icon → Options, or from the popup):

- Strip leading numbers — `01_Customers.xlsx` → `Customers`
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
                         attachment via DataTransfer, drag/drop replay)
  matcher.js             Filename cleaning + suggestion scoring
  binder.js               Click-to-bind element picker + storage
  queue.js                 Batch queue: matches, fills, pauses on ambiguity
  dropzone.js               Floating drop panel UI
  setup-modal.js             "Bind D365 fields" modal
  content.js                  Wires it all together, injects the launcher
options/                Settings page (cleaning rules, matching, bindings)
popup/                  Toolbar popup (status + link to settings)
styles/dropzone.css     All injected UI styling
icons/                  Toolbar/extension icons
```

## Notes on how file attachment works

Browsers won't let a script silently pick a file from disk — the user has to
supply it. Here, you already supplied it via the real drag-and-drop gesture
onto the panel. The extension then re-attaches that same `File` object to
D365's own file input/drop target using a `DataTransfer` object (the same
mechanism browsers use internally for drag-and-drop), so D365 sees it exactly
as if you'd dropped or browsed to it yourself.

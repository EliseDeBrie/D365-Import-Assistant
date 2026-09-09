# D365 Import Assistant

An Edge/Chromium extension for the Dynamics 365 Finance & Operations
**Data management → Import** screen. Drag a folder of Excel files onto it and
it auto-fills the D365 "Entity name" field for each one, guessing the entity
from the cleaned-up file name, then attaches the file — no typing per file.

## How it works

1. A floating **Import Assist** button appears on any `*.dynamics.com` page.
   Click it to open the drop panel.
2. Drag one or many Excel files anywhere onto the page (the whole page is a
   drop target, not just the small box — dropping on the launcher button
   works too, and it'll open the panel for you). Or click the panel's box to
   browse instead.
3. For each file, the file name is cleaned up (leading/trailing sequence
   numbers, dates/timestamps, and `_`/`-` separators are stripped — see
   *Cleaning rules* below) to produce a guessed entity name, and the source
   format is picked from the extension (`.csv` → CSV, otherwise → Excel).
4. Click **Start**. For each queued file, the extension replays D365's own
   click sequence:
   - clicks **Add file**,
   - opens the **Source data format** dropdown and picks Excel/CSV,
   - types the cleaned name into the **Entity name** field once it appears,
   - waits for D365's real autocomplete suggestions and picks the
     best-scoring one,
   - attaches the Excel file,
   - clicks **Upload**, then waits for a new row to appear in the entities
     grid (if that's bound) before moving to the next file.
5. If no suggestion is a confident match, that file is marked
   **needs-review** and the queue pauses — pick the right entity from a
   dropdown (built from D365's own suggestions) or skip the file. Nothing is
   ever typed in blind without going through D365's real autocomplete, so a
   bad guess can't silently attach the wrong entity.

The per-row **Upload** is automated (it's just staging that one file), but the
extension never auto-clicks the page-level **Import/Run** button that actually
loads everything into D365 — that stays a manual, deliberate step unless you
explicitly turn on auto-run in Settings.

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
```

## Notes on how file attachment works

Browsers won't let a script silently pick a file from disk — the user has to
supply it. Here, you already supplied it via the real drag-and-drop gesture
onto the panel. The extension then re-attaches that same `File` object to
D365's own file input/drop target using a `DataTransfer` object (the same
mechanism browsers use internally for drag-and-drop), so D365 sees it exactly
as if you'd dropped or browsed to it yourself.

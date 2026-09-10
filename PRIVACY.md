# Privacy Policy — D365 Import Assistant

_Last updated: 2026-09-09_

This is a browser extension, not a service — there is no backend server
operated by the developer, and nothing described below is sent to the
developer, to D365Solutions, or to any third party. Everything it does
happens inside your own browser, on your own machine.

## What the extension can see

The extension only runs on pages under `https://*.dynamics.com` (its
`host_permissions` in `manifest.json` are scoped to that domain and nothing
else — it cannot read or act on any other site you visit).

That domain covers more than Finance & Operations, so the scripts also load
on other Dynamics apps hosted there. The **Import Assistant** button is limited
to the Data management pages (configurable under *Where the button appears*),
but `content/page-hook.js` — the small script that lets D365's own upload
control accept a dropped file — loads on every `dynamics.com` page. It takes
no action unless a run you started has armed it, and it reports only whether
an upload request to D365 finished.

On those pages, it can:

- **Read and fill in D365's own form fields** on the Data management →
  Import screen (source format, entity name, sheet, file upload), the same
  way a person clicking through the screen would.
- **Read the files you drag onto it.** Excel workbooks are opened locally
  (in your browser) just far enough to list their sheet names: the workbook's
  index (`xl/workbook.xml` inside the `.xlsx` zip) is read, and nothing else.
  The cell data is never read, transmitted, or stored. The file
  itself goes only where you were already sending it: into D365's own
  upload control, via D365's own upload mechanism.
- **Call your own tenant's OData service document** (`<your D365
  origin>/data`) to fetch the list of entity names your environment
  exposes, using the session you're already signed into. This is used only
  to show a hint ("does this name look right?") next to each file — never
  to silently change what gets typed into D365.

## What it stores, and where

Everything the extension remembers is kept in your browser's own
`chrome.storage`, which is:

- **Local to your browser profile** — never sent to a server the developer
  runs, because there isn't one.
- **Only synced the way your browser already syncs your other extension
  data** — `chrome.storage.sync` rides on your own Microsoft/Google
  browser-profile sync, if you have that turned on. That's the browser
  vendor's sync, not the extension's.

What's stored:

| Data | Where | Purpose |
| --- | --- | --- |
| Cleaning rules, run options, UI preferences (whether the launcher shows, which pages it appears on) | `chrome.storage.sync` | Remembering your settings |
| Field-selector bindings, per D365 host | `chrome.storage.sync` | Letting the extension find the right controls on your environment |
| Cached entity name list (one list per browser profile, not per environment) | `chrome.storage.local` | Avoiding re-fetching the list on every run |
| Launcher button position | `chrome.storage.local` | Remembering where you dragged it |

None of this is a file's contents, a password, a session token, or anything
that identifies you personally beyond the D365 hostnames you've used it on.

## What it does not do

- No analytics, telemetry, crash reporting, or usage tracking of any kind.
- No advertising, and no data is sold or shared — there's no data collected
  to sell in the first place.
- No network requests to anywhere other than your own D365 origin (for the
  entity list) and whatever D365's own page already does when you upload a
  file (D365's own upload endpoint, not one added by this extension).
- No code execution outside the `*.dynamics.com` pages it's scoped to.

## Source code

The full source is available at
<https://github.com/D365Solutions/D365-Data-management-helper-tool> — every
claim above can be checked directly against `content/entity-list.js` (the
only `fetch` call the extension makes) and the `chrome.storage` calls
throughout `content/` and `background.js`.

## Changes to this policy

If what the extension accesses or stores ever changes, this file will be
updated alongside that change in the same commit, so its version history
in Git is the change history of this policy too.

## Contact

Questions or concerns: open an issue at
<https://github.com/D365Solutions/D365-Data-management-helper-tool/issues>.

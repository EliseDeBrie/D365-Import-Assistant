# Contributing

This project is open to outside contributions — pull requests, bug reports,
and D365-environment quirks you've hit that the extension should handle
better are all welcome.

## License note

This repository is licensed under **MIT** (see `LICENSE`).
By submitting a pull request you agree your contribution is licensed under
the same terms as the rest of the project. There's no separate CLA to sign
— the license itself is the agreement.

## Getting set up

```
git clone https://github.com/D365Solutions/D365-Data-management-helper-tool
cd D365-Data-management-helper-tool
npm install
npm test
```

To try your changes against real D365: `edge://extensions` (or
`chrome://extensions`) → enable Developer mode → **Load unpacked** → select
the repo folder. Reload the extension after each change.

## Before opening a pull request

- **`npm test` has to pass.** The suite runs the real content scripts
  against a fake D365 Import form in jsdom (see `test/`), so it catches
  regressions in the upload pipeline without needing a live tenant. If
  you're fixing a bug, add a test that fails before your fix and passes
  after — most of the existing tests exist because a real batch stalled or
  mis-imported and the fix needed something to hold it in place.
- **Keep comments load-bearing.** This codebase explains *why*, not *what*
  — a comment earns its place by recording a decision, a D365 quirk, or a
  failure mode that isn't obvious from the code alone. Match the existing
  density rather than adding narration to every line, and don't strip
  existing comments unless the reasoning they record no longer applies.
- **Small, focused changes.** A PR that fixes one thing is far easier to
  verify against a real D365 environment than one that reshuffles several
  files at once.
- **New selectors or D365 behaviour**: if you're adding support for a D365
  version or control layout this doesn't already handle, explain what
  changed in D365 and why the old approach didn't cover it — that context
  is what keeps `content/binder.js`'s defaults trustworthy for everyone
  else.

## Reporting a bug

Open an issue with: what you dropped in, what you expected, what actually
happened (screenshot of the panel and, if you can, the browser console),
and which step the failing row was on (`content/queue.js`'s pipeline names
each step, and the panel shows it — e.g. `error [attach-file]`). That step
name is usually most of the diagnosis.

## Project layout

See the *Project layout* section in `README.md` for what lives where.

# Product reliability

Intact's primary contract is preservation: opening a document must not silently change content or package parts the user did not edit. Reliability work therefore starts with evidence and regression gates, not a broad rewrite of the editor.

## Scope and acceptance criteria

| Area | Acceptance criteria | Existing evidence / next gate |
|---|---|---|
| DOCX round-trip fidelity | An unchanged document has equivalent document XML and preserves every other package part; a one-location edit changes only the intended part/block; a second save is stable. | `tests/roundtrip.spec.js` and `tests/fixtures.spec.js` already exercise these contracts against generated and third-party files. Keep them as release blockers; expand fixtures only with anonymized documents and their source provenance. |
| Corruption prevention and recovery | Invalid/truncated packages fail visibly; a failed open or save never replaces the source file; recovery never silently discards edits. | Existing package-soundness assertions validate XML, content types and internal relationships. Add explicit malformed-package and interrupted-save coverage before changing ZIP loading or persistence behavior. |
| Security and privacy | Production CSP remains restrictive; opening, editing and saving a local file cause no document-content network requests; secrets/passwords stay in the browser. | `vite.config.js` defines the production CSP. `tests/deployment.spec.js` checks the built policy, a whole editing session against it and the requests it makes. Server response headers (including `frame-ancestors`) remain a deployment responsibility. |
| Cross-browser compatibility | The same reliability suite runs in Chromium, Firefox and WebKit; engine-specific exclusions are explicit, narrow and explained. | `playwright.config.js` runs all three engines. Do not expand WebKit exclusions to hide product regressions. |
| Complex document structures | Real files from Word, LibreOffice and other producers retain relationships, styles, numbering, sections, notes, tables, fields and embedded media not explicitly edited. | Third-party fixture suite and package-integrity checks are the baseline. Add fixtures for any newly discovered structure regression. |
| Real-world fixture coverage | Each fixture has a known producer, anonymized content and a short feature inventory; fixtures are exercised in the same cross-browser suite. | See `tests/fixtures/README.md`; current corpus includes Word, LibreOffice, pandoc and python-docx output. |
| Large-document performance | Representative large files can open, edit and save without dropped content, browser errors or pathological slowdowns. | `tests/reliability.spec.js` edits one paragraph of 400 and checks every other unchanged, recording open and save times. Collect timings before setting hardware-sensitive thresholds. |
| Editor accessibility | Core actions are keyboard reachable, controls have accessible names, focus is visible and dialogs manage/return focus; document accessibility checks remain distinct from editor UI accessibility. | Existing tests cover keyboard opening and document accessibility findings. Expand editor-level keyboard and screen-reader checks separately; do not mistake the document checker for an audit of the editor UI. |

## The tests that carry it

- **Fidelity:** `tests/roundtrip.spec.js` (unchanged documents, a second save, one paragraph edited
  among others), `tests/fixtures.spec.js` (every real-world fixture, edited in every kind of place,
  its package checked for soundness).
- **Large documents:** `tests/reliability.spec.js` opens a 400-paragraph document, edits one
  paragraph and checks it exactly and every other one unchanged, and records how long opening and
  saving took (as test annotations, in the report), so baselines collect before any budget is set.
- **Privacy:** `tests/deployment.spec.js` checks the built Content Security Policy as a whole, that a
  whole editing session breaks none of its rules and makes no request to another origin, and that
  the page can't send anything to another site; `tests/feedback.spec.js` that the feedback link
  sends nothing.
- **Browsers:** every spec runs in Chromium, Firefox and WebKit, as parallel jobs in CI
  (`.github/workflows/test.yml`); WebKit's pending specs are listed, each with why, in
  `playwright.config.js`.

## Release gate

A change touching DOCX parsing, serialization, package relationships, encryption, save flows or editor rendering must pass:

1. `npm ci`
2. `npm run build`
3. `npm test` (Chromium, Firefox and WebKit)

A failure in one browser is a regression to investigate, not a reason to add a skip. Any intentional engine limitation must name the missing browser capability and have a focused test.

## What comes next

The work still to do in each area is in [VISION.md](../VISION.md)'s Direction (Now: reliability),
where it is ordered with everything else; this document says what each must show when it is done.

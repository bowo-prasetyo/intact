# P0 product reliability

Intact's primary contract is preservation: opening a document must not silently change content or package parts the user did not edit. Reliability work therefore starts with evidence and regression gates, not a broad rewrite of the editor.

## Scope and acceptance criteria

| Area | Acceptance criteria | Existing evidence / next gate |
|---|---|---|
| DOCX round-trip fidelity | An unchanged document has equivalent document XML and preserves every other package part; a one-location edit changes only the intended part/block; a second save is stable. | `tests/roundtrip.spec.js` and `tests/fixtures.spec.js` already exercise these contracts against generated and third-party files. Keep them as release blockers; expand fixtures only with anonymized documents and their source provenance. |
| Corruption prevention and recovery | Invalid/truncated packages fail visibly; a failed open or save never replaces the source file; recovery never silently discards edits. | Existing package-soundness assertions validate XML, content types and internal relationships. Add explicit malformed-package and interrupted-save coverage before changing ZIP loading or persistence behavior. |
| Security and privacy | Production CSP remains restrictive; opening, editing and saving a local file cause no document-content network requests; secrets/passwords stay in the browser. | `vite.config.js` defines the production CSP. The new browser test verifies the built policy and checks observed requests while opening and saving. Server response headers (including `frame-ancestors`) remain a deployment responsibility. |
| Cross-browser compatibility | The same reliability suite runs in Chromium, Firefox and WebKit; engine-specific exclusions are explicit, narrow and explained. | `playwright.config.js` runs all three engines. Do not expand WebKit exclusions to hide product regressions. |
| Complex document structures | Real files from Word, LibreOffice and other producers retain relationships, styles, numbering, sections, notes, tables, fields and embedded media not explicitly edited. | Third-party fixture suite and package-integrity checks are the baseline. Add fixtures for any newly discovered structure regression. |
| Real-world fixture coverage | Each fixture has a known producer, anonymized content and a short feature inventory; fixtures are exercised in the same cross-browser suite. | See `tests/fixtures/README.md`; current corpus includes Word, LibreOffice, pandoc and python-docx output. |
| Large-document performance | Representative large files can open, edit and save without dropped content, browser errors or pathological slowdowns. | New stress regression creates a 400-paragraph document, edits one paragraph and proves every other paragraph is unchanged. Record timings before setting hardware-sensitive thresholds. |
| Editor accessibility | Core actions are keyboard reachable, controls have accessible names, focus is visible and dialogs manage/return focus; document accessibility checks remain distinct from editor UI accessibility. | Existing tests cover keyboard opening and document accessibility findings. Expand editor-level keyboard and screen-reader checks separately; do not mistake the document checker for an audit of the editor UI. |

## Changes in this PR

- Add a large-document regression that checks content completeness and single-paragraph isolation after save.
- Add a production-browser privacy regression that asserts the deployed CSP contains the restrictive directives and that opening and saving a local document make no cross-origin requests.
- Run both tests through the existing Chromium, Firefox and WebKit Playwright projects.
- Keep the existing real-world fixture corpus and round-trip tests intact; this PR does not rewrite the DOCX serializer or claim complete compatibility with every WordprocessingML extension.

## Release gate

A change touching DOCX parsing, serialization, package relationships, encryption, save flows or editor rendering must pass:

1. `npm ci`
2. `npm run build`
3. `npm test` (Chromium, Firefox and WebKit)

A failure in one browser is a regression to investigate, not a reason to add a skip. Any intentional engine limitation must name the missing browser capability and have a focused test.

## Follow-up work deliberately not hidden by this PR

- Add malformed ZIP/CRC and interrupted-save recovery tests before introducing archive limits or changing JSZip loading options.
- Add independently sourced fixtures for charts, SmartArt, OLE objects, custom XML, alternate content, nested tables, unusual numbering and complex section/header combinations.
- Establish performance baselines on a documented CI runner and a representative multi-megabyte DOCX before enforcing timing budgets.
- Audit editor UI semantics, focus order, zoom/reflow, contrast, reduced motion and screen-reader announcements with automated checks plus manual assistive-technology passes.
- Verify deployed HTTP security headers, including `frame-ancestors`, `X-Content-Type-Options` and an appropriate referrer policy; a CSP meta element cannot set every response header.

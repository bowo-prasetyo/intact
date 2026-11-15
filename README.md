# Office

Fix a Word document in seconds, privately: no install, no account, no upload — and nothing you
didn't touch changes. Open a `.docx`, edit its text, and download it again. The file is unzipped,
parsed and rebuilt entirely in your browser.

Where this is heading, and what it deliberately won't do, is in [VISION.md](VISION.md).

Built with [Vue 3](https://vuejs.org/), [Vue Router](https://router.vuejs.org/) (hash mode, so it
works on GitHub Pages without rewrites), [Vite](https://vite.dev/) and [JSZip](https://stuk.github.io/jszip/).

## Features

- Open `.docx` (file picker or drag & drop) or start a blank document
- Edit text in paragraphs, headings, lists and table cells
- Bold, italic, underline, strikethrough; paragraph styles (Normal, Title, Heading 1–3); alignment
- Download the result as `.docx` (Ctrl/Cmd+S)
- Faithful round-trip: everything the editor doesn't understand is written back untouched —
  fonts, colours, sizes, spacing, numbering, images, table layout, headers/footers, footnotes,
  fields, content controls, equations, bookmarks, section settings and tracked changes (shown
  read-only with their author, so they can still be accepted or rejected in Word). The test suite
  checks this on every pull request (see [Testing](#testing))
- Undo/redo for every edit, including paragraph style changes; list numbers update as you edit

## How it works

`src/docx/DocxDocument.js` is the core:

1. **Load** – JSZip opens the package; `_rels/.rels` locates the main part (`word/document.xml`),
   and `styles.xml`, `theme1.xml` and `numbering.xml` are read for display.
2. **Render** – each `<w:p>` becomes an HTML block (`<p>` or `<h1>`–`<h6>`). Its `<w:pPr>` and each
   run's `<w:rPr>` are kept as the original XML in a side table, referenced by `data-ppr` /
   `data-rpr`. Bold/italic/underline/strike become HTML tags so the browser can edit them.
   Unsupported content becomes a non-editable placeholder (`data-raw`) holding the original XML.
   Tables keep their XML structure; only cell contents are editable.
3. **Save** – the editor DOM is walked back into WordprocessingML: stored properties are cloned,
   toggles are merged in (respecting the schema order of `CT_RPr`/`CT_PPr`), adjacent runs with the
   same formatting are merged, and placeholders are written back verbatim. Only
   `word/document.xml` (and `styles.xml`, if a heading style had to be added) change.

## Development

```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # production build in dist/
npm test         # round-trip and editing tests in headless Chromium
```

## Testing

`tests/` drives the production build in a real browser with [Playwright](https://playwright.dev/),
because editing depends on the browser's `contenteditable`. Fixtures are built from XML in the
tests themselves (`tests/helpers.js`). The core checks are:

- opening and downloading a document **without edits** gives back the same document;
- saving a saved file again changes nothing;
- editing one paragraph changes **only** that paragraph.

"The same" is judged by meaning, not bytes (`canonicalize` in `tests/helpers.js` lists the
normalisations, such as merging adjacent runs with identical formatting). A new feature that touches
the document should add its case to the kitchen-sink fixture in `tests/roundtrip.spec.js`.
First run: `npx playwright install chromium`.

## Deploying to GitHub Pages

The workflow in `.github/workflows/deploy.yml` builds and deploys on every push to `main`.
Enable it once under **Settings → Pages → Build and deployment → Source: GitHub Actions**.
The build uses relative asset paths (`base: './'`), so it works under any repository sub-path.

## Limitations

- Can't insert images, tables or lists, or change table structure
- Pasting is plain text only
- Tracked changes, content controls, fields and footnotes are preserved but can't be edited
- Inline `customXml`, `dir`/`bdo` (bidirectional override) and smart-tag wrappers are dropped on save
  (their text is kept)
- Only Chromium is covered by the tests so far; Firefox and Safari are untested
- Legacy binary `.doc` files are not supported

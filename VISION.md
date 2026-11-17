# Vision

> **Fix a Word document in seconds, privately: no install, no account, no upload — and nothing you
> didn't touch changes.**

This document says who the product is for, what it promises, and what it deliberately won't do.
Use it to decide feature requests: a change that doesn't serve the job below, or that weakens a
promise, needs a very good reason.

## The problem

People regularly get a `.docx` that needs a small change: a typo in a contract, a name in a
template, a date in a letter, a sentence in a CV. The usual options all have a cost:

- **Microsoft Word** needs a licence and an install, which many people don't have on the device in
  front of them (Chromebooks, Linux, borrowed or locked-down machines).
- **Google Docs and Word for the web** need an account and an upload. That's a real barrier for
  confidential documents such as contracts, HR files, medical or legal letters.
- **LibreOffice and other converters** often change things you didn't touch: fonts, spacing,
  numbering, tracked changes. That's a problem when the file goes back to someone who uses Word.

## Who it's for

1. **Privacy-conscious professionals** (legal, HR, finance, healthcare, public sector) who
   mustn't upload a document but need to correct it.
2. **People without Word** on the device they're using.
3. **Anyone sending a file back** to a Word user who will notice if anything else changed.

## The job

Open a Word file, change some text and basic formatting, and get back a file that is the original
with **only those edits**. Success is when the recipient can't tell another tool was used.

## Promises

These are product requirements, not aspirations. A release that breaks one is a bug.

1. **Private.** The document never leaves the device. No server, no account, no analytics, and no
   third-party requests while a document is open.
2. **Faithful.** Anything the user did not edit is written back with the same meaning: formatting,
   images, tables, fields, content controls, comments, footnotes and tracked changes. Content the
   editor can't edit is shown read-only and kept, never silently changed or dropped. The round-trip
   tests (`tests/roundtrip.spec.js`) enforce this; every new capability adds cases to them.
3. **Instant.** Opening the page and a document takes seconds, with nothing to install or set up.
4. **Honest.** When something can't be edited or preserved, the user is told. The README and About
   page list known limitations, and the headline claims stay true.

## Non-goals

Saying no to these keeps the product small enough to keep its promises:

- **Writing documents from scratch.** A blank document exists for quick notes, but this is not a
  word processor. Page layout, styles management and advanced typesetting are out of scope.
- **Collaboration and cloud storage.** No accounts, sharing, comment threads or sync: those need
  a server, and they conflict with "private".
- **Other formats.** No spreadsheets, slides, PDF or legacy `.doc`. Do one format well.
- **Pixel-perfect pagination.** The editor shows content faithfully, not exact Word page breaks.

## Principles for deciding what's next

- **Fidelity over features.** A feature that risks the round-trip waits until it can be tested.
- **Unlock what's locked before adding what's new.** The most valuable work is making preserved
  content editable where it matters for the job (for example the text inside content controls in
  templates), not adding new kinds of content.
- **Real documents decide.** Grow the test fixtures with documents written by Word, Google Docs and
  LibreOffice, and check output in Microsoft Word before each release.

## Direction

**Now — make the promises true and provable**
- Round-trip test suite in CI, gating deploys ✔
- Tracked changes preserved (read-only) ✔
- Undo/redo for all edits; live list numbering ✔
- Close the remaining known losses (inline `customXml`, `dir`/`bdo` wrappers)
- Deleting text never silently drops tracked changes or leaves orphaned comments ✔
- Read-only content says what it is when clicked or typed into, instead of ignoring the keys ✔
- Test in Firefox ✔ (automated and by hand); output checked in Microsoft Word ✔ (by hand)
- Add real-world fixture documents (started: a letter written in Word ✔; more from Google Docs and
  LibreOffice welcome); test Safari

**Next — do the core job better**
- Edit text inside content controls (templates and forms); tick boxes, pick from lists and dates ✔
- Start and end bulleted and numbered lists; Word's Enter, Backspace and Tab behaviour in lists and tables ✔
- Paste with basic formatting (bold, italic, lists) instead of plain text only ✔
- Show list and table spacing as Word does ✔
- Find and replace ✔
- Show comments, headers, footers and notes read-only, in plain language ✔

**Later — only if the job demands it**
- Track your own edits as tracked changes ✔; review your changes before downloading ✔; Word's
  keyboard shortcuts ✔
- Superscript, subscript and highlight ✔; show formatting marks (¶) ✔
- Accept or reject tracked changes ✔, including formatting changes and table rows ✔, and table, row
  and cell property changes ✔
- Add a comment ✔; reply to, edit, resolve and delete comments ✔; add a table row with Tab in the
  last cell ✔
- Comments in headers, footers and notes: reply, edit, resolve, delete ✔; tracked changes in locked
  content: accept and reject ✔
- Insert tables and pictures ✔; insert and delete table rows and columns ✔, tracked ✔; merge and
  split cells ✔, tracked ✔, field cells too ✔; tables whose rows or cells are
  fields ✔; resize, wrap and position pictures ✔; crop, rotate and add effects ✔; picture styles, 3-D and artistic effects ✔; 3-D lighting and all of Word's artistic effects ✔; wrap text to a
  picture's outline ✔
- Edit footnotes and endnotes ✔; show equations ✔
- Apply the document's own paragraph styles ✔; say when fields Word updates may be out of date ✔
- Name reference managers' citations ✔; keep pasted links and insert links (DOIs too) ✔; count the
  words selected ✔; check spelling in the document's language ✔
- Edit headers and footers ✔; edit the text in text boxes and shapes ✔; show designed documents as
  Word lays them out (columns, paragraph borders and shading, drop caps, character spacing, theme
  colours, embedded fonts) and name fonts shown with a substitute ✔
- Font, size, colour and character style from the document's own ✔; clear formatting ✔; track
  formatting changes ✔; smart quotes and dashes ✔; count characters ✔
- Respect Word's restrictions: Restrict Editing (read only, comments, forms, tracked changes),
  read-only recommended, digital signatures (kept valid when unchanged, removed when edited, as in
  Word), with an explicit "Edit anyway" ✔; say what encrypted, rights-managed and `.doc` files are ✔
- Installable offline app (PWA)

## How we'll know it's working

- **Fidelity:** zero known round-trip regressions; the fixture set keeps growing.
- **Trust:** users report that recipients didn't notice anything unusual about the file.
- **Speed:** a user goes from opening the page to downloading an edited file in under a minute.

## Name

The product is called **Intact**, with the descriptor **"private .docx editor"** wherever there is
room for it ("Intact — private .docx editor"). The name states the "Faithful" promise: the document
comes back intact, with only your edits. It replaces "Office", which suggested a suite of
applications and invited trademark questions. Use this one name everywhere: the header, the page
title, the package name and the `Application` written into new documents.

Call it an **editor**, never a "word processor". A word processor is for writing documents from
scratch, which is a non-goal above; "editor" says the job is changing a document someone else
made. It also keeps the descriptor clear of Word.

## Brand colour

The brand colour is **Intact Teal**, `#0e6467` (`--brand` and `--accent` in `src/style.css`, and
the icon). Teal says calm, safe and checked, and it sits apart from Word's blue and from the colours
that mean something in a document here: green and red (others' insertions and deletions), purple
(your tracked changes), coral (comments) and yellow (find). Don't use teal for any of those, and
don't bring Word's blue back into the interface. Document content keeps its own colours, such as
Word's hyperlink blue and highlight colours.

| Token | Light | Dark | Use |
|---|---|---|---|
| `--brand` | `#0e6467` | `#0b3f42` | Header bar, with white text |
| `--accent` | `#0e6467` | `#5ec4bd` | Links, selected and focused controls |
| `--active` | `#e0efee` | `#1d3b3c` | Selected item backgrounds |
| `--hover` | `#eef3f3` | `#263033` | Hover backgrounds |
| `--sheet-accent` / `--sheet-tint` / `--sheet-line` | `#0e6467` / `#e3f1f0` / `#9ccfcb` | same | Marks on the always-white page |

Every text pairing meets WCAG AA (white on `--brand` is 6.9:1 in light and 11.6:1 in dark).

## Logo

The mark is a page with a text cursor: the cursor is where you change what you need, the page is the
rest, kept whole. The cursor is also the I of Intact. The wordmark is Inter Display SemiBold,
outlined, so the logo needs no font.

| File | Use |
|---|---|
| `public/logo.svg` / `public/logo-dark.svg` | Mark and wordmark, on light / dark backgrounds |
| `public/favicon.svg` | Browser tab; follows the browser's light or dark theme |
| `public/favicon.ico` | 16, 32 and 48 px, for browsers and tools without SVG icons |
| `public/apple-touch-icon.png` | Home-screen icon, 180 px, white page on teal |
| `.github/social-preview.png` | Link previews (Settings → Social preview), 1280 × 640 |

In the app header the mark is drawn inline, reversed (white page, teal cursor). Keep clear space of
at least the cursor's width around the logo; don't recolour it outside the brand colours, add
effects, or put Word's "W" or blue near it.

## Talking about Word

Lead with what Intact does, not with Word. The start page headline is "Change what you need. Keep
the rest intact." Mention Microsoft Word only to describe compatibility, in plain text and never in
a heading or the name: "`.docx` files, the format Microsoft Word uses", not "Edit Word documents".

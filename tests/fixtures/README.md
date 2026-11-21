# Real-world fixture documents

Documents written by other applications, run through the core promises by `tests/fixtures.spec.js`:
everything shown; no edits (or an edit undone) gives back every part byte for byte; a second save
changes nothing; typing in each kind of place (paragraph, heading, list item, table cell, header,
footer, text box, note), Enter and bold change only that block; the saved package stays sound; no
errors in the page. Add a `.docx` here and it is tested automatically.

The generated fixtures are rebuilt from `sources/` with `sources/make.sh` (fictional content only).
The files in git are the ones tested: other versions of these applications write other files.

The repository is public, so fixtures must contain nothing personal or identifying: replace names
(of people and organizations), phone numbers, e-mail and postal addresses and web sites (including text
split across runs), the author and "last modified by" in `docProps/core.xml` (Word writes the current
user there on every save), and local file paths such as an attached template in
`word/_rels/settings.xml.rels`; swap pictures for blank ones of the same size. Keep everything else as
the application wrote it.

| File | Written by | Features | Notes |
|---|---|---|---|
| `word-letter-id.docx` | Microsoft Word (Japanese UI) | Indonesian letter from a `.dotx` template: header with logo, footer with address, four pictures, numbering part, footnote/endnote parts, theme fonts | Contributed and re-saved in Word by the maintainer; anonymized (people's and organization's names, web site, phone numbers, e-mail, address, author, template path; pictures blanked). |
| `pandoc-report.docx` | pandoc 3.1 (from `sources/report.md`) | Title block, headings, nested and numbered lists, footnotes, a table with a caption, a picture, a block quote, code, a definition list, strikeout, sub/superscript, a line break, links | Generated |
| `python-docx-memo.docx` | python-docx (`sources/memo.py`) | Two sections (one landscape), different first-page header, footer, title and heading styles, numbered and bulleted lists, a table with merged cells, a page break | Generated |
| `libreoffice-minutes.docx` | LibreOffice Writer 24.2 (from `sources/minutes.fodt`) | Track Changes on, tracked insertion and deletion, a comment, a table of contents, page and page-count fields, a bookmark and cross-reference, a text frame, a two-column section, a footnote and an endnote, a header row, merged cells | Generated |
| `libreoffice-report.docx` | LibreOffice Writer 24.2 (the report via ODT) | As the pandoc report, written the LibreOffice way | Generated |
| `libreoffice-memo.docx` | LibreOffice Writer 24.2 (the memo re-saved) | As the python-docx memo, with every header and footer variant written out | Generated |
| `libreoffice-letter.docx` | LibreOffice Writer 24.2 (the Word letter re-saved) | As the Word letter; LibreOffice puts the footer's picture and text box in one run | Generated from the anonymized letter |

## Encrypted with a password (`encrypted/`)

Kept apart from the fixtures above (which `tests/fixtures.spec.js` opens without a password), and
tested by `tests/encryption.spec.js`. Each is a fixture above, encrypted by another program, so the
decryption is checked against other implementations than Intact's own:

| File | Encrypted by | Encryption | Password |
|---|---|---|---|
| `encrypted/report-password.docx` | msoffcrypto-tool 5 (`OOXMLFile.encrypt`) | Agile (4.4): AES-256, SHA-512, 100,000 rounds, as Word 2010 and later | `Rahasia 2026!` |
| `encrypted/libreoffice-memo-password.docx` | LibreOffice Writer 24.2 (Save with password, Word 2007–365) | Standard (3.2): AES-128, SHA-1, as Office 2007 | `Memo-Pass 7` |

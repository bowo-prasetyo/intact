# Real-world fixture documents

Documents written by other applications, run through the core promises by `tests/fixtures.spec.js`
(no-edit round trip with every other part byte-identical, second save, single-paragraph edit). Add a
`.docx` here and it is tested automatically.

The repository is public, so fixtures must contain nothing personal or identifying: replace names
(of people and organizations), phone numbers, e-mail and postal addresses and web sites (including text
split across runs), the author and "last modified by" in `docProps/core.xml` (Word writes the current
user there on every save), and local file paths such as an attached template in
`word/_rels/settings.xml.rels`; swap pictures for blank ones of the same size. Keep everything else as
the application wrote it.

| File | Written by | Features | Notes |
|---|---|---|---|
| `word-letter-id.docx` | Microsoft Word (Japanese UI) | Indonesian letter from a `.dotx` template: header with logo, footer with address, four pictures, numbering part, footnote/endnote parts, theme fonts | Contributed and re-saved in Word by the maintainer; anonymized (people's and organization's names, web site, phone numbers, e-mail, address, author, template path; pictures blanked). |

# Real-world fixture documents

Documents written by other applications, run through the core promises by `tests/fixtures.spec.js`
(no-edit round trip with every other part byte-identical, second save, single-paragraph edit). Add a
`.docx` here and it is tested automatically.

The repository is public, so fixtures must contain nothing personal: replace names, phone numbers,
e-mail and postal addresses (including text split across runs) and the author in `docProps/core.xml`,
and swap pictures for blank ones of the same size, keeping everything else as the application wrote it.

| File | Written by | Features | Notes |
|---|---|---|---|
| `word-letter-id.docx` | Microsoft Word (Japanese UI) | Indonesian letter from a `.dotx` template: header with logo, footer with address, four pictures, numbering part, footnote/endnote parts, theme fonts | Contributed by the maintainer; anonymized (people's names, phone numbers, e-mail, address, author; pictures blanked). |

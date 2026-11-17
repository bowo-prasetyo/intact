<h1>
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="public/logo-dark.svg" />
    <img src="public/logo.svg" alt="Intact" height="56" />
  </picture>
</h1>

**Intact — private .docx editor.** Change what you need. Keep the rest intact.

Fix a Word document in seconds, privately: no install, no account, no upload — and nothing you
didn't touch changes. Open a `.docx`, edit its text, and download it again. The file is unzipped,
parsed and rebuilt entirely in your browser.

Where this is heading, and what it deliberately won't do, is in [VISION.md](VISION.md).

Built with [Vue 3](https://vuejs.org/), [Vue Router](https://router.vuejs.org/) (hash mode, so it
works on GitHub Pages without rewrites), [Vite](https://vite.dev/) and [JSZip](https://stuk.github.io/jszip/).

## Features

- Open `.docx` (file picker, by mouse or keyboard, or by dropping it anywhere on the start page; a file dropped where the editor can't take it is refused, never opened by the browser in the app's place) or start a blank document; Word templates (`.dotx`) and
  files with macros (`.docm`, `.dotm`) open too and are saved as what they are (the header shows the
  extension). Macros are never run: a bar says the file contains them and that they are kept, as
  they are, and "Before you release" lists them
- Right-to-left text, as Word shows it: paragraphs (`w:bidi`, directly or from their style) are laid
  out and edited right to left; Word's left and right alignment in them are logical (start and end),
  as Word and LibreOffice read them, and the alignment buttons align as shown; runs marked
  right-to-left (`w:rtl`) and right-to-left tables (`w:bidiVisual`, first column at the right) too.
  Direction embeddings and overrides (`w:dir`, `w:bdo`), smart tags and inline custom XML are kept
  around their text, also when it is edited, with their properties
- Line numbers in the margin, as Word prints them for sections with line numbering
  (`w:lnNumType`): every line of the text but tables, notes, headers, footers, text boxes and
  paragraphs that suppress them, every n-th number (count by), from the starting number, restarting
  each section or never, or each page, where Word last began one (its `w:lastRenderedPageBreak`
  marks, page breaks and "page break before"), so the numbers match Word's printed pages while the
  edits don't move its page breaks. Drawn over the page, never in it, and kept up to date as you type
- Edit text in paragraphs, headings, lists and table cells
- Bold, italic, underline, strikethrough, superscript and subscript (each turns the other off, as in
  Word), and highlight in Word's 15 highlight colours (the last colour chosen is remembered; existing
  highlight shows and can be removed); paragraph styles (Normal, Title, Heading 1–3, and the document's own: headings 4–9 it defines, the
  styles in Word's style gallery that aren't hidden, such as Quote and Caption, and any style its text
  uses; a paragraph in another style shows that style's name); alignment
- Bulleted and numbered lists that behave like Word's: the Bullets and Numbering buttons start and
  end lists; Enter on an empty item (or Backspace at its start) ends the list; Tab and Shift+Tab at
  the start of an item change its level
- Tab and Shift+Tab move between table cells; Tab in the last cell adds a row, as in Word (a copy of
  the last row's cells, empty: widths, merged columns, shading and paragraph formatting, but not a
  header row's repeat or vertical merges), tracked while tracking changes
- Insert tables (toolbar grid, or type any size up to Word's 63 columns) at the caret, as Word does:
  the paragraph is split there, the table spans the page in Word's Table Grid style (added to
  `styles.xml` if the file lacks it), and while tracking its rows are tracked insertions
- Change a table's layout from the Table layout menu, as on Word's Layout tab: insert rows above or
  below (as many as are selected; copies of the row, carrying vertical merges and header repeat on
  where they fall inside them), insert columns left or right (cells spanning the place are widened,
  the grid and widths follow, and a table that fitted the page narrows to fit), delete rows, columns
  or the whole table. Rows and columns are tracked while tracking, as Word tracks them: a row
  inserted or deleted is marked on the row (`trPr`), a column on each of its cells (`cellIns`,
  `cellDel`), deleted text struck through, and each row or column is accepted or rejected as one
  change (rejecting an inserted column, or accepting a deleted one, takes its grid column out and
  narrows cells spanning it). Deleting a row or column you inserted removes it; rows and columns
  holding someone's tracked change aren't deleted untracked. Rows and cells that are content
  controls change too: a row added in a repeating section is a new item of it (with its own ids,
  as Word's + adds one), a row added in a plain field stays in it, a field left empty goes, and new
  cells next to a field cell are ordinary cells; locked fields and fields linked to data stay as
  they are
- Merge and split cells from the same menu, as Word does: Merge cells joins the selected cells
  (grown to whole cells where a merged cell sticks out) into one spanning them across (`gridSpan`)
  and down (`vMerge`), their content one after another in it; Split cells… divides a cell into
  columns of equal width (adding grid columns that the other rows span where needed) and rows (the
  row's other cells merged down), and a merged cell splits back into the cells it spans. Cells that
  are text fields merge too (the merged cell is the first one's field; a placeholder isn't moved);
  locked fields, fields linked to data and other kinds of field don't, and the status bar says
  why. While tracking, merging and splitting are tracked as a Word document records them: cells
  merged across a row with `w:hMerge`, each cell's old properties kept in a tracked property change
  (`tcPrChange`); cells merged down a column (or split back) with `w:cellMerge`; the text moved
  shown inserted in the merged cell and deleted where it was; cells split into columns as inserted
  cells, the old grid kept (`tblGridChange`) so rejecting gives the widths back; into rows as
  inserted rows. Each is one change to accept or reject; accepting a merge across a row leaves one
  cell spanning the columns (`gridSpan`), as Word writes a merge. Cells an older Word file merged
  with `w:hMerge` show merged
- Insert pictures (PNG, JPEG, GIF or BMP) from the toolbar, by pasting a picture on its own (such as
  a screenshot) or by dropping picture files on the page: in line with the text at the caret, at
  their size or narrowed to fit between the margins, tracked while tracking. The picture's file,
  relationship and content type are added when saving, only if it is still in the document
- Lay out pictures, as Word does: a click selects a picture (a frame with handles); corners resize
  it in proportion, sides stretch it; a floating picture is dragged to a new place (measured from
  the margin and its paragraph, so it moves with the text). Layout Options (the button beside a
  selected picture) set its text wrapping (in line, square, tight, through, top and bottom, behind
  or in front of text), its position (left, centred or right of the margins or at a distance;
  below its paragraph or at the top, middle or bottom of the margins), its exact size in
  centimetres (keeping proportions, or back to its own size) and its alt text. The picture's
  DrawingML is changed where it must be (`wp:inline` ↔ `wp:anchor`, extents, position, wrap,
  `docPr/@descr`) and the rest kept; the page shows the layout as near as a web page can. Resizing
  and moving aren't tracked; the review lists them. Charts, shapes and other drawings stay as they are
- Crop, rotate and add effects to pictures, as Word's Picture Format tab does: crop by numbers or
  with the handles (`a:srcRect`; what shows keeps its scale, the frame follows), rotate with the
  round handle or by 90° and flip (`a:xfrm` `rot`, `flipH`, `flipV`), and set a border, shadow,
  glow, soft edges, reflection, greyscale, brightness and contrast (`a:ln`, `a:effectLst`,
  `a:grayscl`, `a:lum`); effects Intact doesn't set are kept. Tight and Through wrapping follow the
  picture's outline, as in Word: the outline of what isn't transparent is written as the wrap
  polygon, and the page wraps the text around it (CSS `shape-outside`), as it does for a polygon
  Word drew
- Picture styles, shapes, 3-D and artistic effects, as on Word's Picture Format tab: a gallery of
  Word's picture styles (each written out as Word writes it: shape, border, shadow, glow, soft
  edges, reflection and 3-D format; Word keeps no reference to a style), a shape to cut the picture
  to (`a:prstGeom`), Word's 3-D Format (a bevel with its width and height, a material, a lighting
  rig and its angle) and 3-D Rotation (a preset, or X, Y and Z rotation and a perspective;
  `a:sp3d`, `a:scene3d`, `a:camera/a:rot`, `a:lightRig`), and all 22 of Word's artistic effects
  (Marker, Pencil Sketch, Paint Strokes, Watercolor Sponge, Mosaic Bubbles, Glass, Plastic Wrap and
  the rest), each with an amount, drawn into the picture in the browser and saved as a new picture,
  as Word saves them, with the original kept beside it in an extension of the picture (`a:extLst`),
  so the effect can be changed or taken off here (Word shows the drawn picture; Word's own record of
  an effect needs the original in Microsoft's HD Photo format, which a browser can't write). A bevel
  is lit per pixel, as a 3-D renderer lights it: its profile gives the surface's slope near the
  edge, and the light rig (from its direction and angle) and the material give light and shade by
  Lambert's and Phong's laws (`src/docx/lighting.js`); a 3-D rotation is shown with the file's
  angles and field of view. The page shows every effect with the file's own values: a shadow's distance,
  direction, blur, colour and transparency, a glow's size and colour, soft edges' width, a
  reflection's distance and fade, a bevel's width, Word's brightness and contrast (tones moved and
  stretched about the middle, not CSS's multiply) and greyscale, the picture's shape; a turned
  picture in line with the text takes the room of its turned outline, on the page and in the file
  (`wp:effectExtent`), as in Word
- Add comments (toolbar button or Ctrl+Alt+M) on the selected text, or the word at the caret, with
  your name (asked once, shared with Track Changes): they appear in the comments panel and are
  written to `comments.xml` (created if the document has none) next to the comments already there.
  Every comment can be replied to, edited, resolved (or reopened) and deleted in the panel, as in
  Word: a reply is written as a real reply (linked to its comment in `commentsExtended.xml`, which is
  created, with paragraph ids, if the file has none), resolving sets Word's "done" flag, and Undo
  covers all of it. Deleting a comment's text deletes it too. Comments anchored in a header, footer
  or note work the same way (the panel says where they are): a reply is written into that part next
  to its comment, inside its range, and a deleted comment's marks go from it
- Fill in templates and forms: text fields (content controls) are editable, with placeholder text
  replaced as you type; check boxes tick on click; drop-down lists and date pickers offer their
  choices. From the keyboard, Alt+↓ (or Space on a selected field) ticks a box or opens a list or
  date picker for the field at the caret. Fields around whole table cells or rows work too. Locked
  fields and fields linked to the document's data stay read-only and say why
- Edit footnotes and endnotes in place: they are listed below the text (as in Word's draft view),
  numbered in the order the text refers to them, and edited like the text: typing, formatting,
  Undo, Find and Replace and Track Changes all work in them. Clicking a note's number in the text
  goes to the note, and its number in the note goes back. A note can't be deleted with the text
  around it (nor its number); deleting its reference in the text hides it and renumbers the others,
  as in Word. Only notes that changed are rewritten in `footnotes.xml` / `endnotes.xml`; untouched,
  the parts come back byte for byte. Comments in notes stay with the comments panel
- Edit headers and footers in place, in the page's margins, as in Word: the first page's, other
  pages', even pages' and each section's, chosen with tabs above and below the page (a click in one
  edits it; it is dimmed until then). Typing, formatting, Undo, Find and Replace and Track Changes
  work in them as in the text, and lists in them are numbered on their own. Only the parts that
  changed are rewritten (`header1.xml`, `footer2.xml`…); untouched, they come back byte for byte.
  Comments in them stay with the comments panel, and go with their text when it is deleted
- Edit the text in text boxes and shapes, shown where Word puts them: in line with the text or
  floating left, right or centred with the text wrapping around them, at their size, with their fill
  and outline (theme colours too), inner margins and vertical alignment. Both kinds Word writes are
  edited: DrawingML shapes (`wps:wsp`, Word 2010 and later) and legacy VML text boxes. The text is
  edited like the body (formatting, Undo, Track Changes); on save only the text changes, written into
  the shape and into its copy for older readers (`mc:Fallback`), whose tracked changes get ids of
  their own. A text box left alone is written back exactly as it was
- See a designed document as Word lays it out: a section set in columns (`w:cols`, with a line
  between them when it has one) is shown in columns; paragraph borders (`w:pBdr`, with their style,
  width, colour and distance from the text) and shading (`w:shd`, theme colours too); drop caps
  (`w:framePr w:dropCap`) beside the lines after them; character spacing (`w:spacing`) and raised or
  lowered text (`w:position`); and theme colours on text with their tint or shade
  (`w:themeColor`, `w:themeShade`), worked out from the document's theme as Word does. None of this
  changes the file
- Fonts: the fonts embedded in a document (`word/fonts/*.odttf`, listed in `fontTable.xml`) are
  read from it, unscrambled as the standard says (ECMA-376 Part 1, 17.8.1) and shown with it, as Word
  does; nothing is fetched. A font the text is set in that is neither embedded nor on this device is
  named in the status bar ("Font shown with a substitute: Playfair Display"), so you know the page
  shows another in its place; the document keeps it
- See equations: Word's equations (Office Math) are drawn as MathML by the browser (fractions,
  roots, scripts, sums and integrals, brackets, matrices, functions, accents, and Word's script,
  double-struck and bold letters), on their own line or in the line of text. They stay read-only
  and are written back exactly as they were
- Find and replace (Ctrl/Cmd+F, Ctrl/Cmd+H), with match case and whole words. Matches are found
  across formatting but not across paragraphs; a replacement takes the formatting of the match's
  first character, as in Word. Footnotes, endnotes, headers, footers and text boxes are searched and
  replaced too (a match in another header or footer shows its tab). Comments and locked fields are
  searched, but their matches are read-only and never replaced
- See what you can't edit, read-only and in plain language: comments with their replies and resolved
  state in a side panel (click one to highlight its text); fields labelled by what they
  are ("Page number – Word updates it"); citations and bibliographies by the reference manager they
  come from ("Zotero citation – Zotero updates it in Word; edits made here are lost when it does";
  also EndNote, Mendeley, Citavi and Word's own), and typing in one says so in the status bar. Clicking read-only content (a tracked change, chart,
  equation, locked field or cell) says in the status bar what it is and where to edit it, and puts
  the caret beside it, as Word does; typing in it, or in a header, footer or comment, says why
  nothing happens instead of doing nothing (or typing where the caret was before)
- Show formatting marks (¶ button, Ctrl+Shift+8), as in Word: · for spaces, ° for non-breaking
  spaces, → for tabs, ↵ for line breaks and ¶ at each paragraph end. They are drawn over the page,
  never in it, so they can't change the document; the choice is remembered in this browser
- Paste with basic formatting, as Word's "merge formatting": bold, italic, underline,
  strikethrough, superscript, subscript, paragraphs, line breaks, bulleted/numbered lists and links
  to web and e-mail addresses come through from web pages, Google Docs and Word; fonts, colours,
  styles, other links and pictures in copied text don't (the status bar says when pictures were
  skipped; a picture copied on its own is inserted)
- Links (toolbar button or Ctrl+K), as in Word: the selected text, or new text at the caret,
  becomes a link to a web address, an e-mail address or a DOI (`10.1000/xyz` becomes
  `https://doi.org/10.1000/xyz`); in a link, the same dialog changes its address or removes it. Links
  are written as Word writes them (`w:hyperlink` with an external relationship in the part the link
  is in, so links in notes work too, and Word's Hyperlink style, added if missing); a link inserted
  while tracking is a tracked insertion. Relationships of links that are deleted again are taken out
- Fonts, sizes, colours and character styles from the document itself, as a brand template needs:
  the Font menu offers the theme's heading and body fonts (written as Word writes theme fonts,
  `w:asciiTheme`) and the fonts the document uses; Font size the sizes it uses, then Word's usual
  ones; Font colour (a split button that applies the last colour) Automatic, the theme's ten colours
  (`w:themeColor`) and the colours its text uses; Character style its own character styles (such as
  Strong or Emphasis). Nothing else is offered, so text stays on brand. The selected text's run
  properties are changed and the text is rebuilt as runs with them, so the page shows what is saved.
  Clear formatting (button or Ctrl+Space) takes direct character formatting and the character style
  off, as Word does; highlight, language and a link's style stay
- Formatting changed while tracking is tracked, as Word records it: bold, italic, underline,
  strikethrough, superscript, subscript, highlight, font, size, colour, character style and clear
  formatting record the text's old run properties (`w:rPrChange`); paragraph style, alignment and
  lists the paragraph's old properties (`w:pPrChange`). They show as tracked formatting changes and
  are accepted or rejected here or in Word; formatting changed twice keeps the first, and formatting
  changed back records nothing. The review lists them as tracked
- Smart quotes and dashes as you type, as Word's AutoFormat: " and ' become curly quotes in the
  text's language (“ ” ‘ ’ in English, „ “ in German, « » with no-break spaces in French, and others),
  "--" between words an em dash, " - " an en dash and "..." an ellipsis. Typed as a key press would
  be, so they are tracked while tracking. A button in the status bar turns this off (remembered in
  this browser)
- The status bar counts the words and characters in the document, and those selected ("12 of 340
  words · 64 of 2104 characters"; with and without spaces in its tooltip); headers and footers aren't
  counted, as in Word; the
  page takes the document's language (and a run in another language its own), so browsers that
  choose a spelling dictionary by language (Firefox) check the text in it
- Track Changes for your own edits (toolbar button or Ctrl+Shift+E), as in Word: typing, deleting,
  Enter, Shift+Enter, paste, Replace and page breaks are recorded as tracked changes with your name
  and the date, shown in colour as you work (deleted text struck through and kept, a deleted
  paragraph break shown as a struck ¶ without merging). Deleting your own insertion removes it.
  Word's setting is saved too, so the recipient keeps tracking; documents saved with Track Changes
  on open with it on. Your name is remembered in this browser only
- Accept and reject tracked changes, anyone's and your own, as in Word's Review tab: when the
  document has tracked changes, the toolbar shows Previous, Next, Accept and Reject. Accept and
  Reject act on the change at the cursor (just click it) or every change in a selection, then go
  to the next; their menus also accept or reject without moving on, or all changes at once. An
  insertion accepted (or a deletion rejected) becomes ordinary text; a paragraph break removed
  joins its paragraph with the next; a formatting change rejected restores the old formatting
  (run or paragraph: style, alignment, list); an inserted row rejected (or a deleted one accepted)
  goes. Tracked paragraph breaks show as a coloured ¶, formatting changes with a dashed underline
  or a bar in the margin, tracked rows tinted. Tracked changes to a table's, row's or cell's
  properties (`tblPrChange`, `tblGridChange`, `trPrChange`, `tcPrChange`) are changes like the
  others: accepted, the record goes; rejected, the old properties come back. Tracked changes inside
  content kept as it is (a locked field, a field linked to data, a locked cell, a field such as
  PAGE) are one change per piece: Accept or Reject resolves all of them in its XML at once, as
  Word's Accept All in it would (insertions, deletions, moves, paragraph breaks, formatting, rows,
  cells and merges), and the piece stays locked or linked
- Review your changes before downloading: Download (or Ctrl/Cmd+S) first lists every paragraph that
  changed, with the words deleted and inserted, which changes are tracked, a warning when some
  aren't, and anything else that changes (comments removed with their text, Track Changes switched,
  styles or lists added). An untouched document says "No changes". When the text changed and the
  document has fields Word works out (a table of contents, cross-references, figure and table
  numbers), it says they may be out of date and how to update them in Word (Ctrl+A, F9), and offers
  to have Word update them when the file is opened (`w:updateFields` in `settings.xml`); when
  citations from Word or a reference manager (Zotero, EndNote, Mendeley) were added or removed and
  a bibliography lists them, it says the bibliography needs refreshing (`src/docx/fields.js`). The
  review can be skipped, and opened any time from the toolbar
- Download the result as `.docx` (Ctrl/Cmd+S)
- Respect the restrictions Word puts on a document, as Word does: with Restrict Editing on
  (`w:documentProtection`, enforced) it opens read-only, comments only (comments can be added and
  answered), filling in forms only (its fields can be filled in) or with every edit tracked (Track
  Changes stays on, and tracked changes can't be accepted or rejected); a document marked "read-only
  recommended" or with a password to modify (`w:writeProtection`) opens read-only. A bar above the
  toolbar says what applies, and typing or clicking in locked text says why nothing changes. "Edit
  anyway" lifts it for the session after saying what it means (going around Restrict Editing, which
  only someone allowed to should do); the restriction stays in the file, so Word still applies it,
  and the review says so. Protection set up but not enforced restricts nothing, as in Word
- Digitally signed documents (`_xmlsignatures`) open read-only too, as Word opens them. Downloaded
  unchanged, the file is the one opened, byte for byte, so its signatures stay valid; edited (after
  "Edit anyway"), the signatures are removed from the download, as Word removes them when a signed
  document is edited, rather than sent along invalid; the review says which will happen
- See the document's sensitivity label beside its name, as Office shows it: the Microsoft Purview
  label (`MSIP_Label_*` custom properties, or `docMetadata/LabelInfo.xml`, which doesn't store the
  name), with who or what applied it, when, and the marking it adds, or another tool's classification
  property (such as a protective marking). It is never changed: the file keeps it, and the review
  says so
- Check a document before you release it (for example under Freedom of Information): the review lists
  what the file carries beyond the page, as Word's Document Inspector does, and removes what can be
  removed here. Comments (Remove all comments), tracked changes, whose deleted text anyone can still
  read (Accept all changes), hidden text (`w:vanish`; Remove hidden text, untracked as in Word) and
  personal information: the author, last saved by, company and manager properties and the names on
  comments and tracked changes. "Remove it from the downloaded file" takes those out (names become
  "Author", initials "A", `people.xml` loses the accounts behind them) and sets Word's
  `w:removePersonalInformation`, so Word keeps doing so; unticked, the parts go back as they were.
  Other document and custom properties, custom XML data, embedded files and macros are listed, for
  you to check in Word. Removing is one undoable step each, and not offered while the document is
  restricted
- Saving an edit records it in the document's properties, as Word does: modified now, last modified
  by you (the name given for comments and Track Changes; without one, the last name is kept and the
  review says so) and the revision number one up, worked out from the file as opened so saving twice
  counts once. Nothing changed, nothing is recorded: the properties come back byte for byte. The
  review shows what is recorded and can switch it off (remembered in this browser); with personal
  information removed, no name is recorded
- Check accessibility before publishing, as Word's Accessibility Checker does (Section 508,
  EN 301 549, WCAG): the review lists errors (pictures without alt text, unless marked decorative;
  tables without a header row), warnings (skipped heading levels, no document language, text whose
  colour is below WCAG's contrast minimum) and tips (unclear link text such as "click here", blank
  paragraphs used for spacing, pictures not in line with text, no title). "Show" goes to each place
  in turn; for a picture it opens its alt text, and a missing title can be added there and is saved
  in the document's properties
- Files that aren't `.docx` packages say what they are, instead of failing to unzip: encrypted with a
  password, protected by rights management (a sensitivity label that encrypts, or IRM), a Word
  97–2003 `.doc` with a `.docx` name, or something else. Intact doesn't decrypt files
- Word's keyboard shortcuts: Ctrl+B/I/U, Ctrl+Shift+= / Ctrl+= (superscript / subscript),
  Ctrl+Alt+H (highlight), Ctrl+Space (clear character formatting; highlight stays, as in Word), Ctrl+Shift+8 (formatting marks), Ctrl+L/E/R/J (alignment), Ctrl+Shift+N (Normal), Ctrl+Alt+1–3 (headings),
  Ctrl+Shift+L (bullets), Ctrl+M / Ctrl+Shift+M (list level), Ctrl+Enter (page break), Shift+Enter
  (line break), Ctrl+Shift+Space (non-breaking space), Ctrl+Shift+E (Track Changes), Ctrl+K (link), Ctrl+F/H,
  Ctrl+Z/Y, Ctrl+S; Cmd on a Mac. Some browsers keep a few keys for themselves
- The toolbar works from the keyboard, as WAI-ARIA's toolbar pattern and Word have it: Alt+F10 goes to
  it, it is one Tab stop, the arrow keys (and Home and End) move along it and Esc goes back to the text
  where the caret was. Its buttons are named and say whether they are on (bold, alignment, lists…) for
  screen readers. Its controls stay in place: Accept, Reject and the count of tracked changes are always
  there, turned off while there is nothing to accept, and keep their labels on wide and laptop screens.
  On a phone it is one row that scrolls sideways, faded at the edge where more controls are, with
  touch-sized buttons, and the status bar is one line
- Help (Ctrl+/, or the toolbar's Help button) lists the keyboard shortcuts and opens the review's
  checks without downloading: Review your changes, Before you release and Check accessibility, each
  opening the review at its part
- The review starts with what needs a look: a line saying how many things there are to check, then
  warnings and choices (such as recording the edit in the properties). Before you release,
  Accessibility, the changed paragraphs (more than five) and other notes are folded behind their
  counts, and open by themselves when they hold something to check
- Unsaved changes survive a reload, a crash or a closed tab: while a document has changes not
  downloaded, a copy is kept in this browser (IndexedDB), on this device only, a moment after each
  edit ("Changes kept on this device" in the status bar), and the start page offers to restore or
  discard it. The copy holds the edits only, so the download made from it records the edit in the
  document's properties once; a signed document's copy has no signatures, as its changed download
  wouldn't. It goes when the document is downloaded or closed, when it is discarded, and after 7
  days; the five most recent are kept. Where the browser refuses storage (some private windows), editing works
  as before, without it
- Faithful round-trip: everything the editor doesn't understand is written back untouched —
  fonts, colours, sizes, spacing, numbering, images, table layout, headers/footers, footnotes,
  fields, equations, bookmarks, section settings and tracked changes (shown with their author until
  accepted or rejected, here or in Word). The test suite
  checks this on every pull request (see [Testing](#testing))
- Undo/redo for every edit, including paragraph style changes; list numbers update as you edit
- Deleting never quietly takes what you can't edit with it. An edit that would remove a tracked
  change (someone's insertion, deletion or move, a tracked paragraph break, a table with tracked
  rows) is refused, and the status bar says whose change it is and what to do. Deleting the text a
  comment is attached to deletes the comment, as in Word: the comments panel and status bar say so,
  Undo brings it back, and on save it is removed completely (with its replies, from every comment
  part). A comment whose mark survives keeps its range, shrunk to the text that is left. Backspace
  and Delete next to invisible markers (comment ranges, bookmarks, field codes) delete the text
  beyond them, not the marker

## How it works

`src/docx/DocxDocument.js` is the core:

1. **Load** – JSZip opens the package; `_rels/.rels` locates the main part (`word/document.xml`),
   and `styles.xml`, `theme1.xml` and `numbering.xml` are read for display.
2. **Render** – each `<w:p>` becomes an HTML block (`<p>` or `<h1>`–`<h6>`). Its `<w:pPr>` and each
   run's `<w:rPr>` are kept as the original XML in a side table, referenced by `data-ppr` /
   `data-rpr`. Bold/italic/underline/strike become HTML tags so the browser can edit them.
   Unsupported content becomes a non-editable placeholder (`data-raw`) holding the original XML.
   Tables keep their XML structure; cell contents are editable, and rows and columns change through
   `src/docx/tables.js`, which edits a copy of the table's XML by grid column while the rendered
   rows and cells (with any edits in them) are kept. A section in columns is wrapped in a
   `div.docx-columns` for the page only; saving reads the blocks inside it.
3. **Save** – the editor DOM is walked back into WordprocessingML: stored properties are cloned,
   toggles are merged in (respecting the schema order of `CT_RPr`/`CT_PPr`), adjacent runs with the
   same formatting are merged, and placeholders are written back verbatim. Only
   `word/document.xml` changes, plus `styles.xml` if a heading or the Table Grid style had to be
   added, `numbering.xml` if a new list was started, `footnotes.xml` / `endnotes.xml` if a note was
   edited, a header or footer part if it was edited, and for inserted pictures their files in
   `word/media`, relationships and content types.
4. **Content controls** – a `<w:sdt>` is shown as a frame around its rendered content. Text and
   combo box fields hold an editable `.sdt-content` island; check boxes, drop-downs and dates
   change through clicks. On save its `<w:sdtContent>` is rebuilt and its properties kept, updating
   only what the user changed (`showingPlcHdr`, `w14:checked`, `lastValue`, `fullDate`).
   Placeholder text left alone or cleared is written back verbatim. Locked fields, fields bound to
   custom XML (Word would overwrite edits from the data on open), tables of contents, citations,
   pictures and groups stay read-only and are written back verbatim.
5. **Footnotes and endnotes** – the notes the body refers to are rendered after it, inside the
   editable page, as a `section.docx-notes` (not editable itself) holding one editing host per note
   (`.docx-note`), so typing, formatting, Undo, Find and Track Changes work there as in the body.
   `serializeBlocks` skips the section; `serializeNotes` saves each note that no longer saves the
   same way as when it was rendered into its part, and `writeNotes` writes the part only then. The
   notes, and each note's number, are `data-protected`: `src/editor/protect.js` refuses an edit
   that would remove one with the text around it.
6. **Equations** – `src/docx/omml.js` turns Office Math (`m:oMath`, `m:oMathPara`) into MathML for
   display. The equation itself stays a read-only placeholder holding its original XML.
7. **Find and replace** – `src/editor/find.js` searches runs of text nodes in the same paragraph
   and editing host, and highlights matches with the CSS Custom Highlight API, which doesn't
   change the DOM. Replacing edits the text nodes directly as one undoable step.
8. **Companions** – `src/docx/annotations.js` loads comments (replies and resolved state from
   `commentsExtended.xml`), footnote/endnote text (for references' tooltips) and finds the headers
   and footers each section uses (edited in the page: see 10). Comment ranges are highlighted with the CSS Custom Highlight API between
   the hidden `commentRangeStart` / `commentRangeEnd` markers.
9. **Lists** – while editing, a paragraph's list and level are `data-num-id` / `data-ilvl`
   attributes (set from its own `<w:numPr>`; otherwise its style's numbering applies). A paragraph's
   `<w:numPr>` is only rewritten when the user changed its list or level; its direct indentation is
   then dropped so the list level's applies, as in Word. New lists get Word's default nine-level
   bullet or numbering definitions.
10. **Headers and footers** – rendered into the page, before the text and after the notes, as
   `section.docx-hf-band` bands (not editable themselves, like the notes' section) holding one
   editing host per header or footer part (`.docx-hf-part`), each rendered with its part's
   relationships (`doc.inPart`). The view shows the variant chosen in its tabs with a stylesheet, so
   the choice survives Undo; the header band is in the flow at the top of the page, the footer band
   is placed at its foot, with room kept for it below the text. `serializeBlocks` skips the bands;
   `serializeParts` saves each part whose content no longer saves the same way as when it was
   rendered, and `writeParts` writes it only then.
11. **Text boxes** – `src/docx/textboxes.js` finds a run's text box (DrawingML in `w:drawing` or
   `mc:AlternateContent`, or VML in `w:pict`) and its size, place, fill, outline and insets. It is
   rendered as a `span.textbox` frame (not editable, placed as a picture is) around a
   `.textbox-content` editing host. On save the run is cloned from the raw store; if the text
   changed, the new blocks replace its `w:txbxContent`, and in `update` (once revision ids are known)
   the copies in `mc:Fallback` get the same blocks, their revisions their own ids. The page nests
   paragraphs inside the paragraph the box is anchored in, which the HTML parser doesn't allow, so
   Undo (`src/editor/history.js`) restores copies of the page's nodes rather than parsing its HTML.
12. **Layout and fonts** – `StyleSheet.pPrCss` / `rPrCss` (`src/docx/styles.js`) give borders,
   shading, character spacing, raised text and theme colours as CSS; drop caps float. Embedded fonts
   are loaded with the `FontFace` API (`src/docx/fonts.js`), and fonts that are missing are found by
   measuring text on a canvas, all in the browser.
13. **Protected deletions** – everything that is or holds a tracked change is labelled when rendered
   (`data-tracked`, see `src/docx/revisions.js`); `src/editor/protect.js` checks what each edit would
   remove (`beforeinput` target ranges, or the selection) and refuses edits that would take a
   labelled element away. Comments follow their mark: `src/docx/commentRemoval.js` removes, on save,
   comments whose mark was deleted from `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`
   and `commentsExtensible.xml`, and those parts are only rewritten when a comment was removed.

## Development

```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # production build in dist/
npm test         # round-trip and editing tests in headless Chromium and Firefox
```

## Testing

`tests/` drives the production build in real browsers (Chromium and Firefox) with
[Playwright](https://playwright.dev/), because editing depends on the browser's `contenteditable`,
which differs between engines. Fixtures are built from XML in the
tests themselves (`tests/helpers.js`). The core checks are:

- opening and downloading a document **without edits** gives back the same document;
- saving a saved file again changes nothing;
- editing one paragraph changes **only** that paragraph.

"The same" is judged by meaning, not bytes (`canonicalize` in `tests/helpers.js` lists the
normalisations, such as merging adjacent runs with identical formatting). A new feature that touches
the document should add its case to the kitchen-sink fixture in `tests/roundtrip.spec.js`.

Real documents written by Word and other applications live in `tests/fixtures/` (anonymized; see the
README there) and go through the same promises, with every part other than `document.xml` required
to come back byte for byte. A `.docx` added there is tested automatically.
First run: `npx playwright install chromium firefox`. To run one browser:
`npm test -- --project=chromium` (or `firefox`).

Checked by hand as well: documents edited in Intact open correctly in Microsoft Word, and the
editor works in Firefox (November 2026).

## Deploying to GitHub Pages

The workflow in `.github/workflows/deploy.yml` builds and deploys on every push to `main`.
Enable it once under **Settings → Pages → Build and deployment → Source: GitHub Actions**.
The build uses relative asset paths (`base: './'`), so it works under any repository sub-path.

## Self-hosting

Organisations that can't use a public site can serve Intact from their own web server, or on a
network with no internet at all: it is static files with no server code and no external requests.

1. Download `intact-<version>.zip`, its `.sbom.cdx.json` and `SHA256SUMS` from a
   [release](../../releases) (made by `.github/workflows/release.yml` when a `v*` tag is pushed,
   after the tests pass).
2. Check them:

   ```bash
   sha256sum -c SHA256SUMS
   ```

   ```bash
   gh attestation verify intact-<version>.zip --repo bowo-prasetyo/intact
   ```

   The attestation proves the zip was built by this repository's release workflow from the tagged
   commit. The SBOM (CycloneDX) lists what the build contains (Vue, Vue Router, JSZip).
3. Unzip it into any directory your server serves, at any path. Serve `.webmanifest` as
   `application/manifest+json`.
4. Recommended response headers for its path (the page already carries the Content Security
   Policy in a `<meta>` tag; a header adds `frame-ancestors`, which a `<meta>` can't set):

   ```nginx
   location /intact/ {
     add_header Content-Security-Policy "default-src 'none'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' blob: data:; font-src 'self' blob: data:; media-src 'self' blob: data:; connect-src 'none'; worker-src 'self'; manifest-src 'self'; base-uri 'none'; form-action 'none'; object-src 'none'; frame-src 'none'; frame-ancestors 'none'" always;
     add_header Referrer-Policy "no-referrer" always;
     add_header X-Content-Type-Options "nosniff" always;
     add_header Permissions-Policy "camera=(), microphone=(), geolocation=()" always;
   }
   ```

**What the browser enforces.** The Content Security Policy (`vite.config.js`, put into the built
`index.html`) lets the page load only Intact's own files and connect nowhere (`connect-src 'none'`):
even a script that got into the page couldn't send a document anywhere. Links opened from a
document send no referrer. The tests check that a whole editing session breaks no rule of the policy
and that a request to another site is refused (`tests/deployment.spec.js`).

**Offline and installable.** Served over HTTPS (or from `localhost`), the built app registers a
service worker (`src/sw-template.js`, written out as `sw.js` with the list of the build's files)
that keeps a copy of the app's own files, so once opened it also works offline, and the browser
offers to install it as an app (`manifest.webmanifest`), which can open `.docx` files from the
desktop where the system supports it. The service worker caches nothing else: documents are opened
from the device's files and never fetched, cached or stored. A new version replaces the old copy.

The About page shows the version.

## Limitations

- After restoring a draft, the text you had changed while tracking shows as tracked changes kept from
  the file (accept or reject them; they can't be edited inside), and "Review your changes" lists only
  what changed since restoring. A draft is kept per browser: another browser or device doesn't see it
- Tracked changes inside locked content are accepted or rejected together, not one by one; edits
  inside form fields and text moved by dragging aren't tracked (cut and paste is), nor is ending a list
  with Enter on an empty item, or a picture's layout
- Fonts, sizes and colours are chosen from those the document uses (and Word's usual sizes), not
  from every font installed or any colour; a theme colour's tints and shades aren't offered. A
  font, size or colour applies to selected text, not to the word at the caret

- 3-D lighting is worked out for a bevel's surface, not for a picture's depth (extrusion) or contour;
  Word's preset 3-D rotations are shown with angles as near as their look (a rotation set by angles
  is exact); artistic effects are drawn with Intact's own versions of Word's filters, so they look
  alike rather than identical, and in Word they are part of the picture; a wrap outline isn't
  turned with its picture. Field cells that are locked, linked to data or not text
  fields can't be merged
- Lists that define only one level (such as Word's built-in "List Bullet" style) can't be indented
- Pasting keeps only bold, italic, underline, strikethrough, superscript, subscript, paragraphs,
  lists and links to web and e-mail addresses (not highlight); pictures within copied text and other
  links (their text is kept) are not added. Making existing text a link isn't tracked
- Chrome and Edge check spelling in the languages chosen in their settings, whatever the document's
  language
- Text that holds someone's tracked change can't be deleted together with it; accept or reject the
  change first, or select around it
- Deleting a footnote or endnote reference leaves its note in the notes part (Word ignores it);
  new footnotes can't be added, and pictures, tables, page breaks and new comments can't go in a
  note (comments already in notes can be replied to, edited, resolved and deleted in the panel)
- Equations are shown, not edited; they need a math font to look as in Word (Windows and macOS have
  one: Cambria Math, STIX Two Math; on Linux install one such as Latin Modern Math). Equation
  arrays are not aligned at their `&` points
- Fields are preserved but can't be edited (their results can, until Word updates them); editing a
  comment's text replaces any formatting within it
- Pictures, tables, page breaks and new comments can't go in a header, footer or text box (the
  status bar says so); a note reference in a header stays as it is
- A text box's size, position and look are kept, not changed here, and it grows to show all its text
  where Word would cut it off; text wrapping around it follows its rectangle. Groups of shapes,
  canvases, shapes without text, charts and SmartArt stay as they are. A text box laid out on the page
  (rather than beside its paragraph) is shown near its paragraph
- With no pages, columns are balanced (as Word balances them before a continuous section break), each
  header and footer is shown once, and adjacent paragraphs with the same borders get a border each
  (Word draws one box around them); a drop cap in the margin is shown like one in the text
- Content controls: locked and data-bound ones, and anything but text fields around table cells
  or rows, are read-only
- The tests cover Chromium and Firefox; Safari is untested (find and comment highlighting need
  Safari 17.2 or later)
- Line numbers restart where Word last broke the pages, not where the edits move the breaks to; a
  section in columns has one column of numbers, at the left; right-to-left sections (`w:bidi` in a
  section) keep their numbers at the left and their columns in order
- Legacy binary `.doc` files are not supported, nor files encrypted with a password or by rights
  management (sensitivity labels that encrypt, IRM): they say so when opened
- Restrict Editing applies to the whole document: the parts Word lets everyone (or named people)
  edit in a read-only document (`w:permStart`), and sections left unprotected in a form
  (`w:formProt`), are locked here too; "Edit anyway" opens everything. Formatting restrictions
  (allowed styles only) aren't enforced. Under "tracked changes only", the edits Intact doesn't track
  (in form fields, a picture's layout) aren't prevented in fields, and pictures can't be laid out
- "Before you release" doesn't remove dates from comments and tracked changes, document properties
  other than personal ones, custom XML data, embedded files or macros (it lists them), nor hidden
  text set by a character style or inside content Intact keeps as it is (it says so). A sensitivity
  label can't be applied or changed here
- The accessibility check reads the page as shown: it doesn't check reading order, colour alone
  conveying meaning, or charts, equations and objects kept as they are; a table's header row, the
  document's language and a picture's "decorative" mark are set in Word. Without a core properties
  part, a title can't be added here
- Signatures are recognised, not verified: Intact doesn't check that a signature is valid or who
  signed

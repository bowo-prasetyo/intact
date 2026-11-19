<h1>
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="public/logo-dark.svg" />
    <img src="public/logo.svg" alt="Intact" height="56" />
  </picture>
</h1>

**Intact — private .docx editor.** Change what you need. Keep the rest intact.

Fix a Word document in seconds, privately: no install, no account, no upload — and nothing you
didn't touch changes. Open a `.docx`, edit its text, and download it again. Or start a new one, from
a blank page, and write it here: it is saved as Word writes it, so it opens in Word as it looks
here. The file is unzipped, parsed and rebuilt entirely in your browser.

Where this is heading, and what it deliberately won't do, is in [VISION.md](VISION.md).

Built with [Vue 3](https://vuejs.org/), [Vue Router](https://router.vuejs.org/) (hash mode, so it
works on GitHub Pages without rewrites), [Vite](https://vite.dev/) and [JSZip](https://stuk.github.io/jszip/).

## Features

- Open `.docx` (file picker, by mouse or keyboard, or by dropping it anywhere on the start page; a document dropped on the editor's page opens in a new tab, and a file dropped where the editor can't take it is refused, never opened by the browser in the app's place) or start a blank document; Word templates (`.dotx`) and
  files with macros (`.docm`, `.dotm`) open too and are saved as what they are (the header shows the
  extension). Macros are never run: a bar says the file contains them and that they are kept, as
  they are, and "Before you release" lists them
- Start a new document from a template, as Word does when a template is opened to write with: a
  template opens with a bar that says downloading saves the template itself, and "Start a new
  document" makes it download as a `.docx` instead (a `.dotm` as a `.docm`, macros kept), only the
  main part's content type changed, so the quote, letter or invoice written from it never replaces
  the template. "Edit the template instead" goes back. The template file opened is never changed
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
- Lists started by typing, as Word's AutoFormat: "1. ", "1) " or "(1) ", "a. ", "A) ", "(i) " and the
  like, "* " or "- " typed at the start of a paragraph make it a numbered list in that style (starting
  from the number typed: "3. " starts at 3) or a bulleted list (Word's bullet, or a dash), written as
  Word writes such a list. A number that follows on from a numbered list just above, in its style,
  continues it. Ctrl+Z straight after gives back the text as typed; the AutoFormat button in the
  status bar turns this off, with smart quotes. Tracked while tracking
- Line and paragraph spacing and indents, as Word's Line and Paragraph Spacing menu and Paragraph
  dialog: line spacing (1.0, 1.15, 1.5, 2.0, 2.5, 3.0; Ctrl+1, Ctrl+5, Ctrl+2), add or remove space
  before (Ctrl+0) and after the paragraph, and in the Paragraph dialog left and right indents, a
  first-line or hanging indent (as Word's dialog shows one, Left is where the first line starts:
  "Left 0, Hanging 1.27 cm" is the usual hanging indent, `w:left="720" w:hanging="720"`, and a list
  item's Left is where its number is), space before and after, and line spacing single, 1.5 lines, double,
  at least, exactly or multiple. The dialog shows what the paragraph has (from its style when not its
  own) and writes only what is changed in it, into the paragraph's own `w:spacing` and `w:ind`, with
  everything else in its properties kept (and Word's "lines" values, which would win, taken off what
  is changed). Increase and Decrease Indent (Ctrl+M, Ctrl+Shift+M) move a paragraph's left indent to
  the next half inch, or a list item's level. Tracked while tracking (`w:pPrChange`)
- Lengths in centimetres or inches, as Word shows them: inches for a document on North American
  paper (Letter, Legal, Executive, Tabloid), centimetres for the others, in Page setup, the
  Paragraph dialog (indents and tab stops) and a picture's Layout Options. The Units choice in Page
  setup and the Paragraph dialog switches them, converting what the dialog shows (a length left as
  it was still counts as unchanged, so it isn't rewritten), and is remembered in this browser for
  every document (`src/editor/units.js`)
- Tab stops, as Word's Tabs dialog, in the Paragraph dialog: the paragraph's stops (its styles' and
  list level's, and its own) are listed with their alignment (left, centre, right, decimal, bar) and
  leader (dots, dashes, line); Set adds one at a position in centimetres or inches (replacing one there),
  × removes one and Clear all takes them all off. Only the paragraph's own `w:tabs` is written: a
  stop for each one it doesn't inherit, and a clear (`w:val="clear"`) for a style's stop taken off,
  as Word writes them; tracked while tracking. On the page, tabs go to Word's default stops
  (`w:defaultTabStop`, half an inch unless the file says otherwise), or, where a paragraph's first
  stop is a left one, to that stop, so a letter's "Nomor:" and "Perihal:" lines line up as in Word
- Print the document (Ctrl+P, or the browser's Print): only the page is printed, not the editor
  around it, on the document's paper size and in its margins, with a new page at each page break,
  "page break before" and section break; tracked changes and highlights print as they show.
  Headers and footers print on every page, in the margins as Word places them (the text starts
  below a header taller than its margin): one that is the same on every page repeats as it shows,
  pictures and tables too; page numbers (`PAGE`, `NUMPAGES`, also as simple fields) are numbered on
  each page, "Page 2 of 3", and a different first page or different even pages get their own, in
  CSS page-margin boxes (Chrome, Edge), which hold a header's or footer's text, laid out by its
  alignment and Word's centre and right tab stops, in its font and colour
- Save as PDF from Download's menu (▾): a short dialog says that the browser makes the PDF from
  its print window, on this device (Save as PDF, or Microsoft Print to PDF, as the destination), what
  it will be called (the file's name: the page's title is the file's name while printing), about
  how many pages it has, and what it will show that you may not want sent: tracked changes and
  highlights, which print as they show (accept or reject the changes first for a clean copy).
  Comments and the marks drawn over the page don't print
- About how many pages, as Word would print it: the status bar says "≈ 2 pages", and a dashed line
  across the page, numbered in the margin, shows about where each page begins. Worked out from the
  paper size and margins (and a header or footer taller than its margin), line by line as Word fills
  a page, with Word's widow and orphan control, headings kept with the text after them, table rows
  kept whole, and a new page at each page break, "page break before" and section break that starts
  one; kept up to date as you type. Zoomed in or out, where the browser lays the text out at that
  size, the pages are worked out on a hidden copy at 100%, as printed, so the count and the breaks
  don't change with the zoom. Drawn over the page, never in it; clicking the count shows or
  hides the lines (remembered in this browser). The count matches the printout; Word may break a
  line or two differently with its own fonts
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
  they are. Only the rows and cells a change touches count, so a locked row (a form's total, say)
  or a cell linked to data doesn't stop changes elsewhere: rows are still inserted and deleted
  around it, and columns beside a locked cell. A change that would alter such a field (a column
  adding a cell to a locked row, deleting the column of a linked cell) is refused, and the status
  bar names the field
- Shading and borders for table cells, from the Table layout menu (Shading and borders…), as Word's
  Shading and Borders buttons: a fill colour for the selected cells (the theme's colours, lightened
  80% and 40% as Word offers them, the colours the text uses, your colours, any colour by its code,
  or No colour) and All borders, Outside borders (round the selected cells as a block), Bottom
  border (a heading row or a total) or No border (which hides the table style's too). Written in
  each cell's properties (`w:shd`, `w:tcBorders`), one undoable step, the selection kept; tracked
  while tracking as a change of the cells' formatting (`w:tcPrChange`). A cell's own borders in the
  file are drawn as Word draws them
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
- A logo in the header or footer (a letterhead): with the caret in a header or footer, Insert
  picture (or pasting or dropping one) puts it there, in line at the caret, with the picture file's
  relationship in that part's own relationships (`header1.xml.rels`), as Word writes it; the text's
  part is not changed. Undo takes it out, and its file and relationship with it
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
- Fill in legacy forms, as Word's older form fields (Developer → Legacy Tools) that school, medical
  and club forms still use: check boxes (`FORMCHECKBOX`) are shown as Word draws them and tick on a
  click, drop-down lists (`FORMDROPDOWN`) show their chosen entry and offer the others, and text
  fields (`FORMTEXT`) are typed in, Word's blank (five spaces) selected on a click so typing replaces
  it. Ticking or choosing changes only the field's settings (`w:ffData`: `w:checked`, `w:result`),
  as Word does; the box or entry is drawn by the page, never added to the text. A form protected for
  filling in (Restrict Editing) is filled in as in Word: only its fields take typing, Tab and
  Shift+Tab (and Enter in a text field) go from field to field, and Space, Enter or Alt+↓ ticks a box
  or opens a list. Fields turned off in Word stay as they are and say so
- Sign a form (the toolbar's Draw signature): draw a signature with a finger, pen or mouse, in black
  or blue ink, or type your name in a handwriting font, and it goes in at the caret, or in place of
  the selected "______" line, as a picture in line with the text: cropped to the ink, with a
  transparent background, about 6 cm across, with "Signature" as its alt text. Drawn in this page
  only: it isn't stored or sent anywhere. Tracked while tracking, undone with Ctrl+Z
- Insert a check box or today's date (the toolbar's Insert button), as Word's Developer tab and
  Insert Date and Time do: the check box is Word's check box content control (`w14:checkbox`, ☐ and
  ☒ in MS Gothic), ticked with a click here and in Word; the date goes at the caret in one of a few
  forms in the text's language ("19 November 2026", "Thursday, 19 November 2026", "19/11/2026",
  "2026-11-19"): as text, so it stays the day the form was filled in, or, with Update automatically
  ticked (as in Word's Date and Time dialog, and remembered in this browser), as Word's DATE field
  in the same format (`DATE \@ "d MMMM yyyy"`, its picture worked out from how the date is written
  in that language), showing today's date, which Word changes to the day it opens or prints the
  file: a quote's or letter's date that is always current. Both are tracked while tracking
- Edit footnotes and endnotes in place: they are listed below the text (as in Word's draft view),
  numbered in the order the text refers to them, and edited like the text: typing, formatting,
  Undo, Find and Replace and Track Changes all work in them. Clicking a note's number in the text
  goes to the note, and its number in the note goes back. A note can't be deleted with the text
  around it (nor its number); deleting its reference in the text hides it and renumbers the others,
  as in Word. Only notes that changed are rewritten in `footnotes.xml` / `endnotes.xml`; untouched,
  the parts come back byte for byte. Comments in notes stay with the comments panel
- Insert footnotes and endnotes (the toolbar, or Ctrl+Alt+F and Ctrl+Alt+D, as in Word): a
  reference at the caret (after the selection) and its note below the text with the caret in it,
  numbered in the text's order among the document's own notes (endnotes i, ii, iii), in Word's
  Footnote (Endnote) Text and Reference styles (added to `styles.xml` if the file lacks them). A
  document without such notes gets `footnotes.xml` or `endnotes.xml` (with Word's separators), its
  relationship and content type. A note is saved only while its reference is in the text: Undo, or
  deleting the reference, takes a new one away. The reference is tracked while tracking, and so is
  the note's text
- References, from the Insert menu, as on Word's References tab:
  - Table of contents: "Contents" in Word's TOC Heading style and Word's `TOC` field over headings
    1–3 (\o "1-3" \h \z \u), built at once from the headings as Update fields builds one, linked to
    them; Word fills in the page numbers when it opens the file (`w:updateFields`)
  - Caption…: "Figure 1: text" (or Table, Equation) in Word's Caption style, numbered by a `SEQ`
    field among the others, below a picture or paragraph or above a table, as Word places them
  - Cross-reference…: to a heading (its text) or a caption (its label and number, or all of it), or
    "above"/"below", as a `REF` field to a `_Ref` bookmark added on it, as Word adds one
  - Equation…: typed as in LaTeX (fractions, scripts, roots, sums and integrals with limits,
    brackets that grow, Greek letters, the usual symbols, functions upright), previewed, and written
    as Word's own equations (Office Math, `m:oMath`; on its own line, `m:oMathPara`, in an empty
    paragraph) that Word edits as its own (`src/editor/mathInput.js`)
  - Citation…: a source the document has, or a new one (a book, a journal article or a website), in
    APA (7th edition) or MLA (9th edition) style, with the page cited if wanted, as Word's `CITATION`
    field. Sources are kept as Word keeps them, in its Source Manager's part (`customXml/itemN.xml`,
    `b:Sources`, with its properties part and relationships), so Word's Manage Sources lists them;
    the part is written only when a source or the style is changed, else it goes back byte for byte
  - Bibliography: "References" (APA) or "Works Cited" (MLA) in Heading 1, and Word's `BIBLIOGRAPHY`
    field listing the document's sources in that style, one paragraph each in Word's Bibliography
    style with a half-inch hanging indent (`src/docx/references.js`). A new source or another
    style brings the citations and bibliographies up to date at once
- Headings (the status bar's button), as Word's Navigation pane: the document's headings by level,
  the one the caret is under marked; a click goes to a heading. Kept up to date as you type, and
  shown or hidden as last time in this browser
- Update fields (the toolbar, or F9, as Word's for the whole text): caption numbers (`SEQ`, with
  Word's number formats, \c repeat, \r reset and \s per chapter) are renumbered in the text's
  order; cross-references (`REF`) show the bookmarked text again, or its paragraph's number
  (\r, \n, \w) and "above"/"below" (\p); note numbers in cross-references (`NOTEREF`) follow the
  notes; citations (`CITATION`) and bibliographies (`BIBLIOGRAPHY`, in the text) from the
  document's sources, in its citation style; and tables of contents (`TOC`, in Word's content control or in the text) are rebuilt from
  the headings (\o outline levels, \u paragraphs' own levels, \t styles), as Word builds one:
  each entry in the table's own formatting for its level (Word's TOC 1–9 styles where it has none),
  linked to its heading with \h, through Word's `_Toc` bookmark (added to a heading without one).
  Page numbers are Word's: an entry keeps the page number it had (a new one takes the number above
  it), and the file asks Word to update fields when it opens it (`w:updateFields`). One undoable
  step; the status bar says what was updated, and the review says the fields were updated here
- Page numbers, as Word's Insert Page Number: Word's page number field (`PAGE`) at the caret in a
  header or footer. From the text: into the footer if it is empty (centred), else the caret goes to
  the footer to say where, so a footer's design stays the document's; a document without a footer
  gets one for all pages with the number centred, in Word's Footer style: `footer1.xml`, its
  relationship, content type and the first section's reference, written only while it is in the
  page (Undo takes it away). Tracked while tracking. Page number's menu (▾) puts it where Word's
  Top of Page and Bottom of Page do: at the left, centre or right of the header or footer (in a
  new header for all pages, `header1.xml`, if there is none: a student's "Surname 1" at the top
  right, with the name typed before the number); a header or footer with text keeps its design, and
  the caret goes there to say where. Edit header and Edit footer, in the same menu, go to the end of
  the header or footer shown, or give a document without one an empty one for all pages to type in
  (Undo takes it away)
- Page setup for the whole document, as Word's: paper size (A4, Letter, Legal, A5, A3, Executive),
  orientation and margins (in centimetres or inches, with Word's Normal, Narrow, Moderate and Wide), set in
  every section's `w:pgSz` and `w:pgMar`. Only what is changed in the dialog is written; the rest of
  each section (headers and footers, columns, gutters) is kept. The page and the printout follow,
  and Ctrl+Z puts it back
- A watermark, as Word's Design > Watermark, in Page setup: DRAFT, CONFIDENTIAL, DO NOT COPY,
  SAMPLE, URGENT, ASAP or your own text, diagonal or horizontal, in Word's semi-transparent silver.
  Written as Word writes one, as WordArt in every header behind the text (a VML text path named
  `PowerPlusWaterMarkObject`, centred between the margins), at the start of each header's first
  paragraph; the rest of the header is kept. A document without a header gets one for all pages,
  in Word's Header style (`header1.xml`, its relationship, content type and the first section's
  reference), saved only while it is in the page. Word's own text watermark shows, and can be
  changed or taken off (a header made only for it goes with it). It is drawn behind the text in the
  middle of each page the page count works out, and printed on every page; Ctrl+Z takes it back
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
  their own. A text box left alone is written back exactly as it was. Text boxes and pictures that
  share a run with other content, as LibreOffice writes them, are shown and edited the same way
  (written back as a run each, which means the same)
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
  and are written back exactly as they were (new ones can be typed: see References)
- Fill in placeholders (the toolbar's Fill in button, which shows how many there are, or the
  template bar): text a template leaves to fill in, in brackets — `[Client name]`, `{{date}}`,
  `«Amount»` or `<<Amount>>` — is listed once each, with how many places it appears; focusing one
  highlights them all, and what is typed goes in everywhere it appears, as one undoable edit
  (tracked while tracking), taking each placeholder's formatting as Replace does. Citations and
  references (`[1]`, `[Smith 2020]`), check boxes (`[x]`) and editorial notes (`[sic]`) aren't
  placeholders; those in read-only content are listed and left as they are
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
- Paste tables from Excel, Google Sheets, LibreOffice Calc, Word or a web page as tables, as Word
  pastes them: the rows and columns, cells merged across and down (`gridSpan`, `vMerge`), the
  columns' widths in proportion to the copied ones (`<col width>`, or the cells'), and each cell's
  text with the basic formatting above (Excel's bold heading cells, written by class in its
  clipboard HTML, too). The table goes where Insert table puts one, in Word's Table Grid style
  across the page; text copied around it pastes before and after it. One cell copied on its own
  pastes as its text. Tracked while tracking (its rows and their text are insertions); one step for
  Undo. Where a table can't go (a header, footer, note, text box, field or another table's cell),
  and for one wider than Word's 63 columns, its rows are pasted as lines with tabs between the
  cells, and the status bar says so
- Links (toolbar button or Ctrl+K), as in Word: the selected text, or new text at the caret,
  becomes a link to a web address, an e-mail address or a DOI (`10.1000/xyz` becomes
  `https://doi.org/10.1000/xyz`); in a link, the same dialog changes its address or removes it. Links
  are written as Word writes them (`w:hyperlink` with an external relationship in the part the link
  is in, so links in notes work too, and Word's Hyperlink style, added if missing); a link inserted
  while tracking is a tracked insertion. Relationships of links that are deleted again are taken out.
  A web or e-mail address typed in the text becomes a link when a space or Enter follows it, as
  Word's AutoFormat makes it ("Internet and network paths with hyperlinks"): `https://…`,
  `http://…` and `www.…` addresses, `mailto:` links and e-mail addresses, without the punctuation
  around them ("(see www.example.com)." links www.example.com; a bracket the address opened stays in
  it). Ctrl+Z straight after gives the address back as plain text; the AutoFormat button turns it
  off; typed while tracking, the link goes with the tracked insertion. Not in a link already, a
  plain-text field or a field
- Fonts, sizes, colours and character styles from the document itself, as a brand template needs:
  the Font menu offers the theme's heading and body fonts (written as Word writes theme fonts,
  `w:asciiTheme`) and the fonts the document uses; Font size the sizes it uses, then Word's usual
  ones; Font colour (a split button that applies the last colour) Automatic, the theme's ten colours
  (`w:themeColor`) and the colours its text uses, then More colours: a colour picker and a box for a
  colour's code (`#0E6467`, or short as `#0E6`), for a brand colour; colours chosen that way are
  remembered in this browser and offered again as Your colours, in any document (cell shading
  offers them too). Character style its own character styles (such as Strong or Emphasis). The
  document's fonts come first, so text stays on brand, then, under Other fonts, Word's usual ones
  that Windows and macOS have (Arial, Calibri, Cambria, Times New Roman and others), written by name,
  for a blank document or one with few fonts of its own. The selected text's run
  properties are changed and the text is rebuilt as runs with them, so the page shows what is saved.
  Clear formatting (button or Ctrl+Space) takes direct character formatting and the character style
  off, as Word does; highlight, language and a link's style stay. Grow Font and Shrink Font (the
  buttons beside the size, Ctrl+Shift+> and Ctrl+Shift+<) step the selected text through Word's
  sizes (… 10, 11, 12, 14, 16 … 72, then by 10), each run from its own size, and Ctrl+] and Ctrl+[
  change it by 1 pt
- Format Painter (the brush button, or Ctrl+Shift+C and Ctrl+Shift+V), as in Word: a click copies
  the character formatting at the caret (or of the selection's first character) and paints it on
  the text selected next, or the word clicked; a double-click keeps it on for several places, until
  Esc or another click on it. The text's run properties become the copied ones (a tracked
  formatting change while tracking); painting whole paragraphs gives them the copied paragraph
  style too. Ctrl+Shift+V is the browser's paste as plain text until formatting has been copied
- Change Case (the Aa menu, or Shift+F3), as in Word: Sentence case., lowercase, UPPERCASE,
  Capitalize Each Word and tOGGLE cASE for the selection, or the word at the caret; Shift+F3 goes
  round UPPERCASE, lowercase and Capitalize Each Word, the text staying selected. Each run keeps its
  formatting; the text's language decides the letters (Turkish dotted i, for example). While
  tracking, a changed run is a tracked deletion and insertion, as Word records it; text already
  deleted isn't changed
- A right-click menu, as Word's (also the Menu key or Shift+F10): Cut, Copy, Paste and Paste as text
  (the browser may ask to let the page read the clipboard; Ctrl+V never needs it), Accept and Reject
  on a tracked change, Link… or, in a link, Open, Edit and Remove link, New comment and Paragraph…,
  and in a table Insert rows above or below and columns left or right, Merge and Split cells,
  Shading and borders and Delete row, column or table. Dialogs opened from it open where it was.
  Cut deletes as a key press would, so it is tracked while tracking. Locked text offers Copy (and a
  comment where comments are allowed). Shift+right-click, and a long press on a touch screen, give
  the browser's own menu, with its spelling suggestions
- Insert Symbol (the toolbar's Ω Symbol), as Word's Symbol dialog: ©, €, ✓, →, ½, accented letters,
  Greek, arrows, maths and the spaces Word offers (non-breaking, thin, en and em), by group or by
  name ("euro", "tick"), or any character by its Unicode code (`20AC`, `U+00BD`, as Word's
  Character code). It goes in at the caret as typed (tracked while tracking); the symbols used last
  are offered first, remembered in this browser
- Formatting changed while tracking is tracked, as Word records it: bold, italic, underline,
  strikethrough, superscript, subscript, highlight, font, size, colour, character style and clear
  formatting record the text's old run properties (`w:rPrChange`); paragraph style, alignment and
  lists the paragraph's old properties (`w:pPrChange`). They show as tracked formatting changes and
  are accepted or rejected here or in Word; formatting changed twice keeps the first, and formatting
  changed back records nothing. The review lists them as tracked
- Smart quotes and dashes as you type, as Word's AutoFormat (lists too: see above): " and ' become
  curly quotes in the text's language (“ ” ‘ ’ in English, „ “ in German, « » with no-break spaces
  in French, and others), "--" between words an em dash, " - " an en dash and "..." an ellipsis;
  and Word's AutoCorrect symbols: (c) ©, (r) ®, (tm) ™, (e) €, --> →, <-- ←, ==> ⇒, <== ⇐ and
  <=> ⇔, replaced as the last character is typed (Ctrl+Z straight after gives back what was typed;
  "(c)" stays as it is where it is an item's letter: at the start of a paragraph, or after "(b)").
  Typed as a key press would be, so they are tracked while tracking. The AutoFormat button in the
  status bar turns this off (remembered in this browser)
- Zoom, as in Word's status bar: − and + step the page through 50% to 200%, the percentage goes
  back to 100%, and Ctrl+mouse wheel (or a trackpad's pinch) on the page zooms it. Only how the page
  shows changes, not the document, and it prints at its size; remembered in this browser. Formatting
  marks, line numbers and a picture's frame and handles follow the zoom
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
- Review your changes before downloading: for a document started here (a blank one, or a new one
  from a template) it doesn't ask for the edits to be tracked, and the CV screening check shows only
  for a CV (or from Help). Download (or Ctrl/Cmd+S) first lists every paragraph that
  changed, with the words deleted and inserted, which changes are tracked, a warning when some
  aren't, and anything else that changes (comments removed with their text, Track Changes switched,
  styles or lists added). An untouched document says "No changes". When the text changed and the
  document has fields Word works out (a table of contents, cross-references, figure and table
  numbers), it says they may be out of date and how to update them in Word (Ctrl+A, F9), and offers
  to have Word update them when the file is opened (`w:updateFields` in `settings.xml`); when
  citations from Word or a reference manager (Zotero, EndNote, Mendeley) were added or removed and
  a bibliography lists them, it says the bibliography needs refreshing (`src/docx/fields.js`). The
  review can be skipped, and opened any time from the toolbar
- Download the result as `.docx` (Ctrl/Cmd+S), or save a PDF (see above)
- Share… from Download's menu (▾), where the browser can share files (Safari on iPhone and iPad,
  among others): after "Review your changes", the saved file goes to the device's share sheet, to
  send it back by mail or a messaging app without looking for it in Downloads. It goes from the page
  to the app chosen, on the device; Intact uploads nothing. Where the browser asks for one more tap
  (saving took a moment), a small dialog offers it. Where files can't be shared, the item isn't shown
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
  A CV or letter made from someone else's file or template carries their name as its author: where
  your name is known (given for comments or Track Changes), the review says when the author isn't
  you. "Do this for every file I change that has no comments or tracked changes" makes removing it
  the default, remembered in this browser: then it is chosen for each file you have changed whose
  properties name someone (not a placeholder such as "Author"), with no comments or tracked changes
  (whose names it would change), no restriction, and the edit recorded in its properties. The
  review opens there and says what goes and why; unticked, the file keeps it. Without that choice,
  nothing is removed for you, so the properties of a file you didn't ask to change stay as they were
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
- Check a CV for employers' screening software (applicant tracking systems), which read a CV's text
  before anyone does and miss or mix up what isn't in plain paragraphs: the review warns of contact
  details only in a header or footer, no e-mail address, text in text boxes or columns, and charts
  or diagrams, and gives tips on text laid out side by side in tables, pictures (no text is read in
  them) and a file called just "CV". "Show" goes to each place. A document reads as a CV by its name
  (CV, résumé, curriculum vitae, Lebenslauf) or two section headings such as Experience, Education
  and Skills, in English, Indonesian, German, French, Spanish or Portuguese; for any other the check
  is there, folded and not counted
- Files that aren't `.docx` packages say what they are, instead of failing to unzip: encrypted with a
  password, protected by rights management (a sensitivity label that encrypts, or IRM), a Word
  97–2003 `.doc` with a `.docx` name, or something else. Intact doesn't decrypt files
- Word's keyboard shortcuts: Ctrl+B/I/U, Ctrl+Shift+= / Ctrl+= (superscript / subscript),
  Ctrl+Alt+H (highlight), Shift+F3 (change case), Ctrl+Shift+C / Ctrl+Shift+V (copy and paint formatting), Ctrl+Space (clear character formatting; highlight stays, as in Word), Ctrl+Shift+8 (formatting marks), Ctrl+L/E/R/J (alignment), Ctrl+Shift+N (Normal), Ctrl+Alt+1–3 (headings),
  Ctrl+Shift+L (bullets), Ctrl+M / Ctrl+Shift+M (list level), Ctrl+Enter (page break), Shift+Enter
  (line break), Ctrl+Shift+Space (non-breaking space), Ctrl+Shift+E (Track Changes), Ctrl+K (link), Ctrl+F/H,
  Ctrl+Z/Y, Ctrl+S; Cmd on a Mac. Some browsers keep a few keys for themselves
- The toolbar works from the keyboard, as WAI-ARIA's toolbar pattern and Word have it: Alt+F10 goes to
  it, it is one Tab stop, the arrow keys (and Home and End) move along it and Esc goes back to the text
  where the caret was. Its buttons are named and say whether they are on (bold, alignment, lists…) for
  screen readers. Its controls stay in place: Accept, Reject and the count of tracked changes are always
  there, turned off while there is nothing to accept, and keep their labels on wide and laptop screens.
  The icon buttons get short text labels (Insert, Symbol, Picture, Sign, Table, Layout, Footnote, Endnote,
  Update, Page no., Page setup, Find, Fill in) where the toolbar has room for them: each in turn, as
  long as it fits without the toolbar taking another row, worked out again when the window's width
  changes, so the toolbar never grows; their names for screen readers stay as they were
  On a phone it is one row that scrolls sideways, faded at the edge where more controls are, with
  touch-sized buttons. The tools a phone is used for come first (Undo, Redo, Bold, Italic,
  Underline, Highlight, the lists, Insert, Draw signature, Insert picture and New comment), the rest a
  swipe away, and the arrow keys follow that order. The status bar is one line, and what just happened
  ("Check box inserted", "Shared letter.docx") comes first in it, rather than out of sight at its
  end; zoom is left to the phone's pinch. Nothing makes the document scroll sideways
  there: a picture wider than the screen is shown smaller (its size in the file unchanged), and what
  lies past the page's edge is cut off at it, as on paper
- Help (Ctrl+/, or the toolbar's Help button) lists the keyboard shortcuts and opens the review's
  checks without downloading: Review your changes, Before you release, Check accessibility and
  Check a CV for screening software, each opening the review at its part
- The review starts with what needs a look: a line saying how many things there are to check, then
  warnings and choices (such as recording the edit in the properties). Before you release,
  Accessibility, the changed paragraphs (more than five) and other notes are folded behind their
  counts, and open by themselves when they hold something to check
- Several documents at once, each in a tab of its own (a CV in one, its cover letter in the next):
  the Open button in the header, or Ctrl+O (Word's Open, which in the browser would leave the app),
  opens the start page in a new tab, and this document stays open. A document dropped on the page
  is offered to open in a new tab, and handed to it in memory (a `BroadcastChannel` between the two
  tabs: nothing is stored or sent). Each tab holds a Web Lock for its document's draft, so the start
  page in another tab shows that draft as "Open in another tab" instead of offering to restore a
  second copy whose edits would overwrite each other's; closed there, it can be restored
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
   `word/document.xml` changes (and only if the text did: saved as it was when rendered, the
   part goes back byte for byte, so an edit to a header alone, or one undone, leaves it alone), plus `styles.xml` if a heading or the Table Grid style had to be
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
   then dropped so the list level's applies, as in Word, unless it was set since in the Paragraph
   dialog (`data-ind-set`). Spacing and indents set here are written into the paragraph's stored
   `<w:pPr>` straight away, so undo, tracking (`w:pPrChange`) and saving see them as they are. New
   lists get Word's default nine-level bullet or numbering definitions.
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

Real documents written by Word, LibreOffice, pandoc and python-docx live in `tests/fixtures/`
(anonymized; see the README there, and `sources/` for how the generated ones are made). Each goes
through every promise in `tests/fixtures.spec.js`: all its text (body, headers, footers, text
boxes, notes) and pictures are shown; saved without edits, or with an edit undone, every part comes
back byte for byte; typing in a paragraph, heading, list item, table cell, header, footer, text box
or note, pressing Enter or making a word bold changes only that block of that part (as a tracked
change where the document has Track Changes on), the typed text in the formatting around it; the
saved package stays sound (well-formed XML, content types, relationship targets); and the page logs
no errors. A `.docx` added there is tested automatically.
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
- Fonts are chosen from those the document uses and Word's usual ones, and sizes from the document's
  and Word's usual sizes, not from every font installed; a font not on this device is shown with a
  substitute (the status bar names it); a colour by its code is written as that colour, not as a theme colour, and a theme
  colour's tints and shades aren't offered for text. A font, size or colour applies to selected
  text, not to the word at the caret. "Your colours" are kept per browser (another browser or device
  doesn't have them)

- 3-D lighting is worked out for a bevel's surface, not for a picture's depth (extrusion) or contour;
  Word's preset 3-D rotations are shown with angles as near as their look (a rotation set by angles
  is exact); artistic effects are drawn with Intact's own versions of Word's filters, so they look
  alike rather than identical, and in Word they are part of the picture; a wrap outline isn't
  turned with its picture. Field cells that are locked, linked to data or not text
  fields can't be merged
- Printing prints headers and footers once, at the start and the end, as the page shows them, where
  they can't be put on the right pages: a document whose sections have headers or footers of their
  own, and a page number beside a picture or table, or in a browser without CSS page-margin boxes
  (Firefox, Safari), where a page number would otherwise print the same on every page. Page numbers
  print in Arabic numerals, from 1; line numbers and comments in the margin aren't printed. Browsers that switch tabs with Ctrl+1 to Ctrl+9 may
  keep Ctrl+1, Ctrl+2 and Ctrl+5 for themselves (the installed app, in its own window, gets them):
  the Line and Paragraph Spacing menu does the same. Ctrl+T and Ctrl+Shift+T (hanging indent) are
  the browser's: use the Paragraph dialog. Indents are set in centimetres or inches and spacing in points
- Lists that define only one level (such as Word's built-in "List Bullet" style) can't be indented
- Tab stops on the page: only a paragraph's first stop, if it is a left one, and evenly spaced stops
  after it are where Word puts them; tabs to a right, centre or decimal stop, and leaders, show at
  the default stops (they are kept, and Word lays them out)
- A watermark is text only (Word's picture watermarks are kept, not made or shown here), in Word's
  silver, the same in every header; it isn't tracked. It is drawn in the middle of each page as
  the page count works them out, so it may sit a little higher or lower than in Word
- The right-click menu's Paste reads the clipboard through the browser, which may ask first or
  refuse (then use Ctrl+V)
- The Format Painter copies character formatting and a paragraph's style, not its alignment,
  spacing or indents; Ctrl+Shift+C opens the developer tools in some browsers instead (use the
  button)
- Pasting keeps only bold, italic, underline, strikethrough, superscript, subscript, paragraphs,
  lists, tables and links to web and e-mail addresses (not highlight); pictures within copied text and other
  links (their text is kept) are not added. A pasted table's cell shading, borders, number alignment
  and row heights aren't kept (it takes Table Grid's), and a table pasted in another table's cell is
  pasted as text, not as a nested table or into the cells there. Making existing text a link isn't tracked
- Chrome and Edge check spelling in the languages chosen in their settings, whatever the document's
  language
- Text that holds someone's tracked change can't be deleted together with it; accept or reject the
  change first, or select around it
- Deleting a footnote or endnote reference leaves its note in the notes part (Word ignores it), but
  for a note added here; pictures, tables, page breaks, new comments and notes can't go in a note (comments already in notes can be replied to, edited,
  resolved and deleted in the panel)
- A page number (`PAGE`) shows # here, as each page has its own (it prints numbered: see Print;
  Word numbers the pages when it opens the file), and the number of pages (`NUMPAGES`,
  `SECTIONPAGES`) about how many there are, as the status bar says (on a phone, the number in the
  file): drawn over the numbers the file holds, which are kept as they are and aren't typed into
  (End and Shift+End go past one at the end of a line, as past a footnote's or endnote's
  reference, to type or select after it; Home and Shift+Home go before one at the start of a line,
  so text typed there goes before the number, not into it). A
  footer made here is for all pages of the first section and those after it without their own; a
  first page set to have a different footer, and even pages where they differ, don't show it, as in
  Word. Page setup applies to every section alike (not to one section) and isn't tracked; columns
  and gutters stay as they are
- Equations already in a document are shown, not edited (new ones are typed as LaTeX, a part of it:
  no matrices, accents or equation arrays); they need a math font to look as in Word (Windows and macOS have
  one: Cambria Math, STIX Two Math; on Linux install one such as Latin Modern Math). Equation
  arrays are not aligned at their `&` points
- Fields are preserved; their instructions can't be edited (their results can, until they are
  updated). Update fields doesn't work out page numbers (`PAGE`, `PAGEREF`, `NUMPAGES`: Word does
  when it opens the file), nor dates, document properties or mail-merge fields, nor citations whose
  source the document doesn't have or citations and bibliographies in Word's content controls (Word
  updates them), nor
  fields in notes, headers and footers or written as simple fields (`w:fldSimple`); it isn't
  tracked. Editing a comment's text replaces any formatting within it
- Citations and bibliographies are formatted here in APA or MLA style as their guides set them out,
  close to but not always the same as Word's own styles, which Word applies again when it updates
  them; a citation cites one source, books, journal articles and websites are the kinds of source
  offered, and Word's other styles (Chicago, Harvard, IEEE…) aren't. A table of contents,
  bibliography or caption isn't tracked (it can't be inserted while tracking), and its paragraphs go in
  the text itself, not in a table, note, header, footer or text box. Undo takes a citation out of
  the text, not its source out of the document's sources (as in Word, where a source stays in the list)
- Tables, page breaks and new comments can't go in a header, footer or text box, nor pictures in a
  text box or note (the status bar says so); a note reference in a header stays as it is
- Cell borders are a single ½ pt line in the automatic colour, as Word's Borders button draws them;
  their style, width and colour are set in Word. Borders and shading go on cells, not on the whole
  table or on paragraphs
- A text box's size, position and look are kept, not changed here, and it grows to show all its text
  where Word would cut it off; text wrapping around it follows its rectangle. Groups of shapes,
  canvases, shapes without text, charts and SmartArt stay as they are. A text box laid out on the page
  (rather than beside its paragraph) is shown near its paragraph
- With no pages, columns are balanced (as Word balances them before a continuous section break), each
  header and footer is shown once, and adjacent paragraphs with the same borders get a border each
  (Word draws one box around them); a drop cap in the margin is shown like one in the text
- Legacy form fields: a text field's maximum length, type (number, date) and format, and a field's
  macros and calculations aren't applied; Word applies them when the form is next filled in there.
  Filling in fields isn't tracked. A signature is a picture of a signature, not a digital signature:
  it doesn't prove who signed (Word's signature lines and digital signatures are kept, not made here)
- Placeholders are found in the page (its headers, text, notes and footers), not in comments, and
  only on one line and within one paragraph; a placeholder written as underscores (`________`)
  or as "XXX" isn't recognised. A new document made from a template doesn't record which template
  it came from (Word's attached template)
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
- The page count is worked out, not Word's: the browser's fonts and line breaking, the last
  section's paper and margins for every section, footnotes not taking room at the bottom of a page,
  and rows kept whole where Word may split them. A paragraph's "keep with next" and "keep lines
  together" other than headings' aren't followed. On a phone, where the page takes the screen's
  width, pages aren't counted
- The PDF is the browser's: its fonts (a font the device lacks is substituted, as on screen), and
  links and a document outline as the browser makes them. Saving the PDF is the print window's
- Typed addresses become links as Word makes them: an address without `www.` or a scheme (such as
  linkedin.com/in/…) stays text, as in Word; use Ctrl+K for it. Network paths (`\\server\share`)
  aren't made links
- A document dropped on the page is handed to the new tab only while both are open, and a browser
  may block the new tab (allow pop-ups for the site); without `BroadcastChannel` the new tab opens
  on the start page, to open the file there. Without Web Locks (older browsers), a draft open in
  another tab isn't recognised as such
- Share… needs a browser that can share `.docx` files (the Web Share API with files): Safari on
  iPhone, iPad and Mac can; Chrome shares only some kinds of file, so where it can't share a `.docx`,
  Share… isn't offered and Download keeps the file on the device. A date that updates shows the day
  it was inserted here (Word updates it when it opens the file); dates already in a file as fields
  aren't updated here
- The screening check knows what such software commonly misses, not what a given employer's does;
  it doesn't rate how well a CV matches a job

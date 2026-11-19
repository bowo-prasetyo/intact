<template>
  <div>
    <AppHeader />
    <main class="about">
      <h1>About</h1>
      <p>
        A small, client-only editor for <code>.docx</code> files, the format Microsoft Word uses. The file is unzipped and
        parsed in your browser; no data is sent to any server.
      </p>
      <h2>What you can edit</h2>
      <ul>
        <li>Text of paragraphs, headings and table cells; Tab in the last cell adds a row</li>
        <li>Footnotes and endnotes, below the text: click a note's number in the text to go to it; insert a footnote (Ctrl+Alt+F) or an endnote (Ctrl+Alt+D)</li>
        <li>Update fields (F9): the table of contents, cross-references, caption and note numbers; Word updates page numbers when it opens the file</li>
        <li>Page numbers in a header or footer, or a new footer with the page number; page setup (paper size, orientation, margins) for the whole document</li>
        <li>Headers and footers, in the page's margins: the first page's, other pages', even pages' and each section's, chosen with the tabs above and below the page</li>
        <li>The text in text boxes and shapes, shown where Word puts them with their fill and outline</li>
        <li>Insert tables and pictures (from the toolbar, or paste or drop a picture)</li>
        <li>Table layout: insert rows above or below and columns left or right (tracked while tracking); delete rows, columns or the table; merge and split cells</li>
        <li>Pictures: click one to resize or rotate it with its handles; Layout Options set its text wrapping (Tight and Through follow its outline), position, size, crop, rotation, effects, picture style, shape, 3-D and artistic effects, and alt text; drag a floating picture to move it</li>
        <li>Tables whose rows or cells are fields, including repeating sections</li>
        <li>Legacy form fields (check boxes, drop-down lists and text fields), also in forms protected for filling in, with Tab between fields</li>
        <li>Draw signature: sign with your finger, a pen or the mouse (or type your name), and it goes in at the caret as a picture</li>
        <li>Comments: add (Ctrl+Alt+M), reply, edit, resolve and delete, including comments in headers, footers and notes</li>
        <li>Bold, italic, underline, strikethrough, superscript, subscript and highlight</li>
        <li>Font, size, colour and character style, from those the document already uses (its theme's fonts and colours too); Grow and Shrink Font (Ctrl+Shift+&gt;, Ctrl+Shift+&lt;); Clear formatting (Ctrl+Space)</li>
        <li>Zoom from 50% to 200% in the status bar, or with Ctrl+mouse wheel on the page</li>
        <li>AutoFormat as you type: smart quotes and dashes in the text's language, symbols such as (c) ©, (tm) ™ and --&gt; →, lists started by typing "1. " or "* ", and web and e-mail addresses made links when a space follows them (turn it off in the status bar)</li>
        <li>Line and paragraph spacing (Ctrl+1, Ctrl+5, Ctrl+2), space before and after, and indents (Ctrl+M, Ctrl+Shift+M), with Word's Paragraph dialog</li>
        <li>Print (Ctrl+P): just the document, on its paper size and in its margins, a new page at each page break, with the header and footer, and page numbers, on every page</li>
        <li>Save as PDF from Download's menu: the browser makes it from its print window, on this device</li>
        <li>About how many pages, in the status bar, with a dashed line where each page begins</li>
        <li>Several documents at once, each in its own tab: Open (Ctrl+O) opens another beside this one, and a document dropped on the page opens in a new tab</li>
        <li>Show formatting marks (¶): spaces, tabs, line breaks and paragraph ends</li>
        <li>Paragraph style (Normal, Title, Heading 1–3 and the document's own styles, such as Heading 4, Quote and Caption) and alignment</li>
        <li>Find and replace (Ctrl+F, Ctrl+H)</li>
        <li>The toolbar from the keyboard: Alt+F10 goes to it, the arrow keys along it, Esc back to the text</li>
        <li>Track Changes for your own edits, formatting included, with your name, as Word records them</li>
        <li>Review your changes before downloading, with a reminder to update the table of contents, cross-references and bibliography in Word</li>
        <li>Accept and reject tracked changes, one at a time or all at once</li>
        <li>Paste with bold, italic, underline, strikethrough, superscript, subscript, lists and links</li>
        <li>Links (Ctrl+K) to web pages, e-mail addresses and DOIs: insert, change or remove them</li>
        <li>Word and character count of the document and of the selection; spelling checked in the document's language where the browser can</li>
        <li>Template and form fields: text, check boxes, drop-down lists and dates</li>
        <li>Bulleted and numbered lists: start and end them, and change an item's level with Tab and Shift+Tab</li>
      </ul>
      <h2>What you can see but not edit</h2>
      <p>
        Comments (with replies and whether they are resolved), equations (drawn as in Word), fields, charts,
        SmartArt and groups of shapes are shown as they are, labelled in plain language, and saved unchanged.
        Sections in columns, paragraph borders and shading, drop caps, character spacing and theme colours
        are shown as in Word, and fonts embedded in the document are used; a font the document asks for that
        isn't on your device is named in the status bar.
        Click anything read-only, or try to type in it, and the status bar says what it is and where to edit it.
      </p>
      <h2>Protected, signed and encrypted documents</h2>
      <p>
        A document with Word's Restrict Editing on opens as Word allows: read-only, comments only, filling in
        its fields only, or with every edit tracked. A digitally signed document, or one marked read-only
        recommended, opens read-only. The bar above the toolbar says what applies; Edit anyway lifts it for
        the session, after saying what that means. The restriction stays in the file. A signed document
        downloaded unchanged keeps its signature; edited, its signature is removed, as in Word. Files
        encrypted with a password or by a sensitivity label can't be opened here, and say so.
      </p>
      <h2>Before you release a document</h2>
      <p>
        A document's sensitivity label (or classification) is shown beside its name and never changed.
        Before downloading, the review lists what the file carries beyond the page: comments, tracked
        changes (their deleted text can still be read), hidden text, personal information such as the
        author's name, other properties, custom XML data and embedded files. Comments, tracked changes,
        hidden text and personal information can be removed there, as with Word's Document Inspector, and
        removing personal information can be made the default for files without comments or tracked changes.
      </p>
      <h2>Sending a CV</h2>
      <p>
        Many employers' screening software reads a CV's text before anyone does. For a document that reads
        as a CV, the review says what such software is known to miss: contact details only in a header or
        footer, text boxes, columns, charts and pictures, and text laid out side by side in tables.
      </p>
      <h2>Accessibility and properties</h2>
      <p>
        The review checks accessibility as Word's checker does: pictures without alt text, tables without a
        header row, skipped heading levels, a missing language or title, faint text, unclear links and blank
        paragraphs used for spacing. Show goes to each one. Saving an edit records it in the document's
        properties (modified, by whom, revision), as Word does; the review can switch that off.
      </p>
      <h2>What is preserved</h2>
      <p>
        Everything the editor does not understand is kept as-is when you save: fonts, colours, sizes,
        spacing, list numbering, images, tables layout, the shapes of text boxes, footnotes, fields,
        equations, comment anchors, section settings, embedded fonts and tracked changes. Such content is shown greyed
        out (tracked changes with their author) and cannot be edited here, and so are locked form
        fields and fields linked to the document's data. In a table, such a field stops only the
        changes that would alter it: rows and columns elsewhere still change.
      </p>
      <h2>Limitations</h2>
      <ul>
        <li>3-D lighting is worked out for a bevel, not for a picture's depth; artistic effects are Intact's own versions of Word's, drawn into the picture (in Word they can't be taken off).</li>
        <li>Tracked changes inside locked content are accepted or rejected together, not one by one; text holding someone's tracked change can't be deleted together with it.</li>
        <li>Edits inside form fields, text moved by dragging and changes to a picture's layout aren't tracked.</li>
        <li>Deleting the text a comment is attached to deletes the comment, as in Word (Undo brings it back).</li>
        <li>Pictures, tables, page breaks and new comments can't go in a note, a header, a footer or a text box. A page number shows 1 until Word numbers the pages, and Update fields leaves page numbers to Word.</li>
        <li>A text box's size, position and look are kept, not changed here; a text box grows to show all its text where Word would cut it off. Groups of shapes, and shapes without text, can't be edited.</li>
        <li>The page isn't split into pages: columns are balanced, and each header and footer is shown once. The page count is worked out from the paper and margins, so Word may break a line or two differently.</li>
        <li>Headers and footers print once, at the start and the end, where they can't go on the right pages: sections with their own, or a page number beside a picture or table, or in a browser without page-margin boxes (Firefox, Safari).</li>
        <li>Equations need a math font to look as in Word (Windows and macOS have one).</li>
        <li>Legacy <code>.doc</code> files are not supported.</li>
      </ul>
      <h2>Keyboard shortcuts</h2>
      <p>The ones from Word for what you can do here (Cmd for Ctrl on a Mac):</p>
      <table class="shortcuts">
        <tr v-for="[keys, what] in SHORTCUTS" :key="keys"><th scope="row"><kbd>{{ keys }}</kbd></th><td>{{ what }}</td></tr>
      </table>
      <h2>Languages, templates and line numbers</h2>
      <p>
        Right-to-left text (Arabic, Hebrew, Persian, Urdu) is shown and edited right to left, with Word's
        alignment. Word templates (.dotx) and files with macros (.docm, .dotm) open and are saved as what they
        are; macros are never run, and kept. Line numbers are shown in the margin as Word prints them.
      </p>
      <h2>Privacy and hosting</h2>
      <p>
        The browser enforces the privacy promise: a security policy lets the page load only Intact's own
        files and connect nowhere, so a document opened here can't be sent off. Once opened online, Intact
        also works offline, and can be installed as an app from the browser's menu; for that it keeps only
        its own files. While a document has changes you haven't downloaded, a copy of it is kept in this
        browser, on this device only, so a reload or a crash doesn't lose them: the start page offers it
        back. The copy goes when you download or close the document, when you discard it, and after 7 days.
        Organisations can host it themselves: each release is a zip of static files, with checksums, a
        software bill of materials and a build attestation.
      </p>
      <p class="version">Version {{ version }}</p>
      <p><RouterLink to="/">← Back to start</RouterLink></p>
    </main>
  </div>
</template>

<script setup>
import AppHeader from '../components/AppHeader.vue';
import { SHORTCUTS } from '../editor/shortcuts.js';

const version = __APP_VERSION__; // eslint-disable-line no-undef -- set by vite.config.js
</script>

<style scoped>
.about { max-width: 720px; margin: 0 auto; padding: 32px 16px; line-height: 1.6; }
.shortcuts { border-collapse: collapse; }
.shortcuts th, .shortcuts td { text-align: left; vertical-align: top; padding: 4px 16px 4px 0; border-bottom: 1px solid var(--border); font-weight: normal; }
kbd { font-family: inherit; font-size: .92em; white-space: nowrap; }
</style>

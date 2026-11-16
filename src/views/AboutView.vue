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
        <li>Footnotes and endnotes, below the text: click a note's number in the text to go to it</li>
        <li>Headers and footers, in the page's margins: the first page's, other pages', even pages' and each section's, chosen with the tabs above and below the page</li>
        <li>The text in text boxes and shapes, shown where Word puts them with their fill and outline</li>
        <li>Insert tables and pictures (from the toolbar, or paste or drop a picture)</li>
        <li>Table layout: insert rows above or below and columns left or right (tracked while tracking); delete rows, columns or the table; merge and split cells</li>
        <li>Pictures: click one to resize or rotate it with its handles; Layout Options set its text wrapping (Tight and Through follow its outline), position, size, crop, rotation, effects, picture style, shape, 3-D and artistic effects, and alt text; drag a floating picture to move it</li>
        <li>Tables whose rows or cells are fields, including repeating sections</li>
        <li>Comments: add (Ctrl+Alt+M), reply, edit, resolve and delete, including comments in headers, footers and notes</li>
        <li>Bold, italic, underline, strikethrough, superscript, subscript and highlight</li>
        <li>Show formatting marks (¶): spaces, tabs, line breaks and paragraph ends</li>
        <li>Paragraph style (Normal, Title, Heading 1–3 and the document's own styles, such as Heading 4, Quote and Caption) and alignment</li>
        <li>Find and replace (Ctrl+F, Ctrl+H)</li>
        <li>Track Changes for your own edits, with your name, as Word records them</li>
        <li>Review your changes before downloading, with a reminder to update the table of contents, cross-references and bibliography in Word</li>
        <li>Accept and reject tracked changes, one at a time or all at once</li>
        <li>Paste with bold, italic, underline, strikethrough, superscript, subscript, lists and links</li>
        <li>Links (Ctrl+K) to web pages, e-mail addresses and DOIs: insert, change or remove them</li>
        <li>Word count of the document and of the selection; spelling checked in the document's language where the browser can</li>
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
      <h2>What is preserved</h2>
      <p>
        Everything the editor does not understand is kept as-is when you save: fonts, colours, sizes,
        spacing, list numbering, images, tables layout, the shapes of text boxes, footnotes, fields,
        equations, comment anchors, section settings, embedded fonts and tracked changes. Such content is shown greyed
        out (tracked changes with their author) and cannot be edited here, and so are locked form
        fields and fields linked to the document's data.
      </p>
      <h2>Limitations</h2>
      <ul>
        <li>3-D lighting is worked out for a bevel, not for a picture's depth; artistic effects are Intact's own versions of Word's, drawn into the picture (in Word they can't be taken off).</li>
        <li>Tracked changes inside locked content are accepted or rejected together, not one by one; text holding someone's tracked change can't be deleted together with it.</li>
        <li>Formatting changes, edits inside form fields and text moved by dragging aren't tracked.</li>
        <li>Deleting the text a comment is attached to deletes the comment, as in Word (Undo brings it back).</li>
        <li>New footnotes can't be added; pictures, tables, page breaks and new comments can't go in a note, a header, a footer or a text box.</li>
        <li>A text box's size, position and look are kept, not changed here; a text box grows to show all its text where Word would cut it off. Groups of shapes, and shapes without text, can't be edited.</li>
        <li>The page isn't split into pages: columns are balanced, and each header and footer is shown once.</li>
        <li>Equations need a math font to look as in Word (Windows and macOS have one).</li>
        <li>Legacy <code>.doc</code> files are not supported.</li>
      </ul>
      <h2>Keyboard shortcuts</h2>
      <p>The ones from Word for what you can do here (Cmd for Ctrl on a Mac):</p>
      <table class="shortcuts">
        <tr v-for="[keys, what] in SHORTCUTS" :key="keys"><th scope="row"><kbd>{{ keys }}</kbd></th><td>{{ what }}</td></tr>
      </table>
      <p><RouterLink to="/">← Back to start</RouterLink></p>
    </main>
  </div>
</template>

<script setup>
import AppHeader from '../components/AppHeader.vue';
import { SHORTCUTS } from '../editor/shortcuts.js';
</script>

<style scoped>
.about { max-width: 720px; margin: 0 auto; padding: 32px 16px; line-height: 1.6; }
.shortcuts { border-collapse: collapse; }
.shortcuts th, .shortcuts td { text-align: left; vertical-align: top; padding: 4px 16px 4px 0; border-bottom: 1px solid var(--border); font-weight: normal; }
kbd { font-family: inherit; font-size: .92em; white-space: nowrap; }
</style>

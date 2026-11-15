<template>
  <div>
    <AppHeader />
    <main class="about">
      <h1>About</h1>
      <p>
        A small, client-only editor for Microsoft Word <code>.docx</code> files. The file is unzipped and
        parsed in your browser; no data is sent to any server.
      </p>
      <h2>What you can edit</h2>
      <ul>
        <li>Text of paragraphs, headings and table cells; Tab in the last cell adds a row</li>
        <li>New comments on the text you select (Ctrl+Alt+M)</li>
        <li>Bold, italic, underline, strikethrough, superscript, subscript and highlight</li>
        <li>Show formatting marks (¶): spaces, tabs, line breaks and paragraph ends</li>
        <li>Paragraph style (Normal, Title, Heading 1–3) and alignment</li>
        <li>Find and replace (Ctrl+F, Ctrl+H)</li>
        <li>Track Changes for your own edits, with your name, as Word records them</li>
        <li>Review your changes before downloading</li>
        <li>Accept and reject tracked changes, one at a time or all at once</li>
        <li>Paste with bold, italic, underline, strikethrough, superscript, subscript and lists</li>
        <li>Template and form fields: text, check boxes, drop-down lists and dates</li>
        <li>Bulleted and numbered lists: start and end them, and change an item's level with Tab and Shift+Tab</li>
      </ul>
      <h2>What you can see but not edit</h2>
      <p>
        Headers and footers (all of them, one at a time), comments (with replies and whether they are resolved),
        footnotes and fields are shown as they are, labelled in plain language, and saved unchanged.
        Click anything read-only, or try to type in it, and the status bar says what it is and where to edit it.
      </p>
      <h2>What is preserved</h2>
      <p>
        Everything the editor does not understand is kept as-is when you save: fonts, colours, sizes,
        spacing, list numbering, images, tables layout, headers and footers, footnotes, fields,
        equations, comment anchors, section settings and tracked changes. Such content is shown greyed
        out (tracked changes with their author) and cannot be edited here, and so are locked form
        fields and fields linked to the document's data.
      </p>
      <h2>Limitations</h2>
      <ul>
        <li>No adding images or tables; tables only grow by a row at the end.</li>
        <li>Existing comments can't be replied to or resolved here.</li>
        <li>Tracked formatting changes, table rows and changes inside locked content can only be accepted or rejected in Word; text holding someone's tracked change can't be deleted together with it.</li>
        <li>Formatting changes, edits inside form fields and text moved by dragging aren't tracked.</li>
        <li>Deleting the text a comment is attached to deletes the comment, as in Word (Undo brings it back).</li>
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

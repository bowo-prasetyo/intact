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
        <li>Text of paragraphs, headings and table cells</li>
        <li>Bold, italic, underline, strikethrough</li>
        <li>Paragraph style (Normal, Title, Heading 1–3) and alignment</li>
        <li>Find and replace (Ctrl+F, Ctrl+H)</li>
        <li>Paste with bold, italic, underline, strikethrough and lists</li>
        <li>Template and form fields: text, check boxes, drop-down lists and dates</li>
        <li>Bulleted and numbered lists: start and end them, and change an item's level with Tab and Shift+Tab</li>
      </ul>
      <h2>What you can see but not edit</h2>
      <p>
        Headers and footers (all of them, one at a time), comments (with replies and whether they are resolved),
        footnotes and fields are shown as they are, labelled in plain language, and saved unchanged.
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
        <li>No adding images or tables; tables can't change shape.</li>
        <li>Tracked changes are kept but can't be accepted or rejected here.</li>
        <li>Legacy <code>.doc</code> files are not supported.</li>
      </ul>
      <p><RouterLink to="/">← Back to start</RouterLink></p>
    </main>
  </div>
</template>

<script setup>
import AppHeader from '../components/AppHeader.vue';
</script>

<style scoped>
.about { max-width: 720px; margin: 0 auto; padding: 32px 16px; line-height: 1.6; }
</style>

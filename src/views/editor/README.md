# Splitting the editor

`src/views/EditorView.vue` was one file of 11,305 lines: the template (about 1,270 lines), the
script (about 9,360, in some 78 sections marked `// ------ name`) and the styles (about 680). Every
change to the editor touched it, and two people working on two features met in it. It is being split
here, a section at a time, without changing what the editor does: 10,372 lines are left in it.

## Done

| Part | Now in | How it was checked |
|---|---|---|
| The styles | `EditorView.scoped.css`, `EditorView.css` (`<style src>`, scoped as before) | The built CSS is byte for byte the same size |
| Sorting a table | `tableSort.js` (`useTableSort`) | `tests/table-sort-formula.spec.js`, the full suite |
| Comparing documents | `compare.js` (`useCompare`) | `tests/compare.spec.js`, `compare-notes.spec.js`, `redline.spec.js`, the full suite |
| Redaction | `redaction.js` (`useRedaction`) | `tests/journalism.spec.js`, `release.spec.js`, the full suite |

## The pattern

A section becomes a module here that exports one `use…(ctx)` function:

- **`ctx` is the editor's shared state, read when it is used, never copied at setup.** Some of it is
  set after the module's function runs (the undo stack, made when the page mounts), so the editor
  passes such things as getters (`get undoStack() { return undoStack; }`) and the module reads
  `ctx.undoStack` inside its functions. Refs (`doc`, `page`, `message`, `tracking`) and hoisted
  functions (`refocus`, `markDirty`, `onInput`) can be passed as they are.
- **It returns what the template and the rest of the script use**, and the editor destructures it
  where the section was (`const { sortDialog, sortDialogEl, … } = useTableSort({ … })`), so template
  refs (`ref="sortDialogEl"`) and handlers keep their names.
- **It imports what it needs itself** (`t`, the docx and editor helpers), not through `ctx`.
- **Nothing else changes in the same pull request**: a section is moved as it is, with only
  `ctx.` added where it reads the editor's state, so a review can compare the two line by line.

Moving a section is mechanical: prefix each name it reads from the editor with `ctx.` (not a
property such as `x.page`, nor an object's key such as `{ author: … }`, but a name after a spread,
`[...page.value]`, or in a ternary, `a ? page : b`), pass as a getter what is declared after the
section or set later, and take out the editor's imports only the section used. Never inside a
comment or a string: "tracked changes" in a message is words, not the editor's `changes` (a move
once wrote "tracked ctx.changes" to the status bar; the tests caught it).

To see what a section needs before moving it, list the names it uses that are declared elsewhere
in the script (its `ctx`) and the names of its own that the template and the rest of the script use
(what it returns). A section whose own names the rest of the script uses a lot (read-only content,
events, selection and toolbar state) is the editor's core: move it last, if at all.

## Next, in this order

Sections used mostly by the template, with a small `ctx`, first:

1. AutoCorrect entries (about 100 lines) → `autoCorrect.js`
2. Change case, Format Painter (about 240) → `caseAndPainter.js`
3. Picture layout (about 440) → `pictureLayout.js`
4. Before you release (about 380) → `release.js`
5. References: citations, captions, indexes, tables of contents (about 500 in four sections) → `references.js`
6. Page numbers, notes and page setup (about 560, sharing much of the page's state: after the others)
7. The dialogs' template blocks, each with its section, as components in `src/components/`

Each is its own pull request, with the full suite green in Chromium, Firefox and WebKit.

<template>
  <div class="review-backdrop" @mousedown.self="$emit('close')">
    <section ref="dialog" class="review" role="dialog" aria-modal="true" aria-labelledby="review-title" @keydown.esc.stop="$emit('close')">
      <header>
        <h2 id="review-title">Review your changes</h2>
        <p class="summary">{{ summary }}</p>
      </header>

      <div class="body">
      <ul v-if="notes.length" class="notes">
        <li v-for="(n, i) in notes" :key="i" :class="n.kind">
          {{ n.text }}
          <label v-if="n.properties" class="update-fields">
            <input type="checkbox" :checked="review.properties.on" @change="$emit('update-properties', $event.target.checked)" />
            Record this edit in the document’s properties
          </label>
          <label v-if="n.fields" class="update-fields">
            <input type="checkbox" :checked="updateOnOpen" @change="setUpdateOnOpen($event.target.checked)" />
            Ask Word to update them when the file is opened
          </label>
        </li>
      </ul>

      <section v-if="review.inspection" class="release" aria-labelledby="release-title">
        <h3 id="release-title">Before you release</h3>
        <p v-if="!release.length" class="clean">Nothing hidden: no comments, tracked changes, hidden text or personal information.</p>
        <ul v-else>
          <li v-for="item in release" :key="item.id" :class="item.kind">
            <span>{{ item.text }}</span>
            <button v-if="item.action" type="button" :disabled="!review.inspection.canChange" @click="$emit('release', item.action)">{{ item.button }}</button>
            <label v-if="item.personal" class="personal">
              <input type="checkbox" :checked="review.inspection.removePersonal" :disabled="!review.inspection.canChange" @change="$emit('personal', $event.target.checked)" />
              Remove it from the downloaded file: names on comments and tracked changes become “Author”, and the author, last saved by, company and manager go, as with Word’s Document Inspector
            </label>
          </li>
        </ul>
        <p v-if="!review.inspection.canChange && release.some((i) => i.action || i.personal)" class="locked">The document is restricted: choose Edit anyway to remove these here.</p>
      </section>

      <section v-if="review.accessibility" class="release accessibility" aria-labelledby="a11y-title">
        <h3 id="a11y-title">Accessibility</h3>
        <p v-if="!review.accessibility.length" class="clean">No accessibility issues found: pictures have alt text, tables have header rows, headings are in order.</p>
        <ul v-else>
          <li v-for="a in review.accessibility" :key="a.id" :class="a.severity === 'tip' ? 'info' : 'warn'">
            <span><strong>{{ SEVERITY[a.severity] }}: {{ a.title }}.</strong> {{ a.text }}</span>
            <button v-if="a.count" type="button" @click="$emit('show-issue', a.id)">{{ a.count > 1 ? `Show (${a.count})` : 'Show' }}</button>
            <form v-if="a.form" class="title-form" @submit.prevent="title.trim() && $emit('title', title)">
              <input v-model="title" type="text" aria-label="Document title" placeholder="Title" :disabled="!canChangeTitle" />
              <button type="submit" :disabled="!title.trim() || !canChangeTitle">Add title</button>
            </form>
          </li>
        </ul>
      </section>

      <ol v-if="shown.length" class="changes">
        <li v-for="(c, i) in shown" :key="i" class="change" :class="c.kind">
          <div class="where">
            <span>{{ c.where ? `${c.where}, paragraph ${c.number}` : `Paragraph ${c.number}` }}{{ c.inTable ? ' (in a table)' : '' }}</span>
            <span class="kind">{{ KIND[c.kind] }}</span>
            <span v-if="c.tracked" class="tag">Tracked</span>
            <span v-else-if="c.resolved" class="tag resolved">Change accepted or rejected</span>
          </div>
          <p class="words">
            <template v-for="(w, k) in excerpt(c.words)" :key="k"><del v-if="w.type === 'del'">{{ w.text }}</del><ins v-else-if="w.type === 'ins'">{{ w.text }}</ins><span v-else-if="w.type === 'gap'" class="gap">…</span><span v-else>{{ w.text }}</span></template>
            <span v-if="isEmpty(c.words)" class="empty">(empty paragraph)</span>
          </p>
          <ul v-if="c.formatting" class="formatting">
            <li v-for="f in c.formatting" :key="f">{{ f }}</li>
          </ul>
        </li>
      </ol>
      <p v-if="review.paragraphs.length > shown.length" class="more">…and {{ review.paragraphs.length - shown.length }} more</p>
      </div>

      <footer>
        <label class="skip"><input type="checkbox" :checked="skip" @change="$emit('skip', $event.target.checked)" /> Download without this review next time</label>
        <div class="buttons">
          <button type="button" @click="$emit('close')">Keep editing</button>
          <button ref="downloadButton" type="button" class="primary" @click="$emit('download')">Download</button>
        </div>
      </footer>
    </section>
  </div>
</template>

<script setup>
// "Review my changes": shown before Download (or from the toolbar), so what will go back to the
// recipient is clear: every paragraph that changed, with its words, and anything else that changes.
import { computed, onMounted, ref } from 'vue';
import { PERSONAL } from '../docx/properties.js';

const props = defineProps({
  // From DocxDocument.review(): { paragraphs, commentsRemoved, insertions, deletions, tracking, trackingSwitched, stylesAdded, listsAdded },
  // with signature ({ count, removed }), protection ({ mode, password }) and overridden from the editor.
  review: { type: Object, required: true },
  author: { type: String, default: '' },
  skip: { type: Boolean, default: false },
});
const emit = defineEmits(['close', 'download', 'skip', 'update-fields', 'release', 'personal', 'show-issue', 'title', 'update-properties']);

const SEVERITY = { error: 'Error', warning: 'Warning', tip: 'Tip' };
const title = ref('');
const canChangeTitle = computed(() => props.review.inspection?.canChange !== false);

// Word's "update fields on open" (w:updateFields), set from the fields note.
const updateOnOpen = ref(!!props.review.fields?.updateOnOpen);
function setUpdateOnOpen(on) {
  updateOnOpen.value = on;
  emit('update-fields', on);
}

const KIND = { changed: 'Changed', added: 'Added', removed: 'Removed' };
const LIMIT = 100;
const CONTEXT = 8; // words of unchanged text kept around a change

const dialog = ref(null);
const downloadButton = ref(null);
onMounted(() => downloadButton.value?.focus());

const shown = computed(() => props.review.paragraphs.slice(0, LIMIT));

const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;

const summary = computed(() => {
  const count = { changed: 0, added: 0, removed: 0 };
  for (const p of props.review.paragraphs) count[p.kind]++;
  const parts = [];
  if (count.changed) parts.push(`${plural(count.changed, 'paragraph')} changed`);
  if (count.added) parts.push(`${count.added} added`);
  if (count.removed) parts.push(`${count.removed} removed`);
  const r = props.review;
  const l = r.layout;
  const layoutChanged = l && Object.keys(l.before).some((k) => l.before[k] !== l.after[k]);
  if (!parts.length) return layoutChanged || r.trackingSwitched || r.commentsRemoved.length || r.commentsAdded?.length || r.commentsEdited?.length || r.commentsResolved?.length || r.commentsReopened?.length
    ? 'The text is unchanged.'
    : 'No changes: the file is the same as the one you opened.';
  return `${parts.join(', ')}.`;
});

const notes = computed(() => {
  const r = props.review;
  const out = [];
  const label = r.label?.name || (r.label ? 'set in Office' : r.marking?.value);
  if (label) out.push({ kind: 'info', text: `${r.label ? 'Sensitivity label' : 'Classification'} “${label}” stays on the file, unchanged.` });
  // Restrictions and signatures first: they say what the recipient gets back.
  const sig = r.signature;
  if (sig?.removed) {
    out.push({ kind: 'warn', text: `The ${sig.count === 1 ? 'digital signature is' : `${sig.count} digital signatures are`} removed from the downloaded file: ${sig.count === 1 ? 'it' : 'they'} would no longer match the changed document, and Word removes ${sig.count === 1 ? 'it' : 'them'} too when a signed document is edited. If the document must be signed, ask for it to be signed again.` });
  } else if (sig) {
    out.push({ kind: 'good', text: `Nothing has changed, so the file downloads exactly as it was opened and its ${sig.count === 1 ? 'digital signature stays' : 'digital signatures stay'} valid.` });
  }
  const p = r.protection;
  if (p) {
    const label = { readOnly: 'no changes (read only)', comments: 'comments only', forms: 'filling in forms only', trackedChanges: 'tracked changes only' }[p.mode];
    out.push(r.overridden
      ? { kind: 'warn', text: `Restrict Editing in Word allows ${label}, and you chose to edit anyway. The restriction stays in the downloaded file${p.password ? ', with its password' : ''}, so Word still applies it.` }
      : { kind: 'info', text: `Restrict Editing (${label}) stays in the downloaded file, so Word still applies it.` });
  }
  // Saving records the edit in the properties, as Word does (unless switched off).
  const pr = r.properties;
  if (pr?.available && pr.changed) {
    const by = pr.by ? `last saved by ${pr.by}` : `last saved by kept as “${pr.lastBy || 'nobody'}” (give your name with a comment or Track changes to record it)`;
    const what = [`modified now`, r.inspection?.removePersonal ? null : by, pr.revision && `revision ${pr.revision}`].filter(Boolean).join(', ');
    out.push({ kind: 'info', properties: true, text: pr.on ? `The document’s properties record this edit: ${what}.` : 'The document’s properties stay as they were: they don’t record this edit.' });
  }
  // Fields Word works out (tables of contents, cross-references, caption numbers) show what they
  // showed when the file was last saved in Word: after editing the text they may be out of date.
  const f = r.fields;
  const l = r.layout;
  const edited = r.paragraphs.length > 0 || (l && Object.keys(l.before).some((k) => l.before[k] !== l.after[k]));
  if (f?.updatable.length && edited) {
    const what = f.updatable.join(', ').replace(/, ([^,]*)$/, ' and $1');
    out.push(f.alreadyUpdating
      ? { kind: 'info', text: `This document has ${what}. The file asks Word to update them when it is opened.` }
      : { kind: 'warn', fields: true, text: `This document has ${what}, which Word updates, not Intact: they may be out of date after your changes. In Word, select all (Ctrl+A) and press F9 to update them before you send the file.` });
  }
  if (f?.citationsChanged) {
    out.push({ kind: 'warn', text: 'Citations were added or removed, but the bibliography doesn’t change until it is refreshed in Word, with your reference manager (Zotero, EndNote, Mendeley) or References › Update Citations and Bibliography.' });
  }
  if (r.insertions || r.deletions || r.rowsTracked || r.cellsTracked || r.formatChanges) {
    const what = [r.insertions && plural(r.insertions, 'insertion'), r.deletions && plural(r.deletions, 'deletion'), r.formatChanges && plural(r.formatChanges, 'formatting change'), r.rowsTracked && plural(r.rowsTracked, 'table row'), r.cellsTracked && plural(r.cellsTracked, 'table cell')]
      .filter(Boolean).join(', ').replace(/, ([^,]*)$/, ' and $1');
    out.push({ kind: 'good', text: `Recorded as tracked changes (${what}), so the recipient can accept or reject them in Word.` });
  }
  if (r.resolved?.accepted || r.resolved?.rejected) {
    const parts = [r.resolved.accepted && `accepted ${plural(r.resolved.accepted, 'tracked change')}`, r.resolved.rejected && `rejected ${plural(r.resolved.rejected, r.resolved.accepted ? 'change' : 'tracked change')}`].filter(Boolean);
    out.push({ kind: 'info', text: `You ${parts.join(' and ')}. As in Word, accepting and rejecting aren’t themselves tracked.` });
  }
  if (r.paragraphs.some((p) => !p.tracked && !p.resolved && (!p.formatting || /^Picture/.test(p.formatting[0])))) {
    out.push({ kind: 'warn', text: 'Some text changes aren’t tracked: the recipient won’t see them marked. Undo them and turn on Track changes to mark them.' });
  }
  if (r.paragraphs.some((p) => p.formatting && !p.tracked && !/^(Paragraph break|Comment|Tracked change|Picture)/.test(p.formatting[0]))) {
    out.push({ kind: 'info', text: 'Formatting changed while Track changes was off is applied directly, not tracked.' });
  }
  out.push(...layoutNotes(r.layout));
  if (r.trackingSwitched) {
    out.push({ kind: 'info', text: r.tracking ? 'Track Changes will be on when the file is opened in Word.' : 'Track Changes will be off when the file is opened in Word.' });
  }
  const quote = (text) => `“${text.length > 80 ? `${text.slice(0, 80)}…` : text}”`;
  for (const c of r.commentsAdded || []) {
    out.push({ kind: 'good', text: c.parentId ? `Reply by ${c.author} to ${c.parentAuthor || 'a comment'}: ${quote(c.text)}` : `Comment added by ${c.author}: ${quote(c.text)}` });
  }
  for (const c of r.commentsEdited || []) out.push({ kind: 'info', text: `Comment by ${c.author} edited.` });
  for (const c of r.commentsResolved || []) out.push({ kind: 'info', text: `Comment by ${c.author} resolved.` });
  for (const c of r.commentsReopened || []) out.push({ kind: 'info', text: `Comment by ${c.author} reopened.` });
  for (const c of r.commentsRemoved) {
    out.push({ kind: 'info', text: `Comment by ${c.author} removed, with the text it was attached to.` });
  }
  if (r.stylesAdded) out.push({ kind: 'info', text: 'Styles you used are added to the file’s style list.' });
  if (r.listsAdded) out.push({ kind: 'info', text: 'Bulleted or numbered lists you started are added to the file’s list definitions.' });
  return out;
});

const list = (names) => names.map((n) => `“${n}”`).join(', ').replace(/, ([^,]*)$/, ' and $1');
const quoteValue = (v) => `“${v.length > 60 ? `${v.slice(0, 60)}…` : v}”`;

/**
 * "Before you release": what the file carries that a reader of the page may not see, as Word's
 * Document Inspector finds it, with a way to remove what can be removed here.
 */
const release = computed(() => {
  const i = props.review.inspection;
  if (!i) return [];
  const out = [];
  if (i.comments.count) {
    out.push({ id: 'comments', kind: 'warn', action: 'comments', button: 'Remove all comments', text: `${plural(i.comments.count, 'comment')}${i.comments.authors.length ? ` by ${list(i.comments.authors)}` : ''}. Comments go with the file: anyone who opens it can read them.` });
  }
  if (i.revisions) {
    out.push({ id: 'changes', kind: 'warn', action: 'changes', button: 'Accept all changes', text: `${plural(i.revisions, 'tracked change')}${i.revisionAuthors.length ? ` by ${list(i.revisionAuthors)}` : ''}. Until they are accepted or rejected, the file keeps the earlier wording, deleted text included, and anyone can see it in Word.` });
  }
  if (i.hidden.count) {
    const stuck = i.hidden.count - i.hidden.removable;
    out.push({
      id: 'hidden', kind: 'warn', action: i.hidden.removable ? 'hidden' : null, button: 'Remove hidden text',
      text: `Hidden text: ${plural(i.hidden.count, 'passage')}, such as ${quoteValue(i.hidden.sample)} – Word doesn’t show or print it, but anyone can show it again.${stuck ? ` ${stuck === i.hidden.count ? 'It is' : `${plural(stuck, 'passage')} ${stuck === 1 ? 'is' : 'are'}`} in content Intact can’t change: remove ${stuck === 1 ? 'it' : 'them'} in Word.` : ''}` });
  }
  const p = i.properties || { core: [], app: [], custom: [] };
  const props_ = [...p.core, ...p.app];
  const personal = props_.filter((x) => PERSONAL.has(x.label));
  const names = [...new Set([...i.comments.authors, ...i.revisionAuthors])].filter((n) => n !== 'Author');
  if (i.removePersonal) {
    out.push({ id: 'personal', kind: 'good', personal: true, text: 'Personal information is removed from the downloaded file.' });
  } else if (personal.length || names.length) {
    const parts = [personal.length && personal.map((x) => `${x.label} ${quoteValue(x.value)}`).join(', '), names.length && `names on comments and tracked changes (${list(names)})`].filter(Boolean);
    out.push({ id: 'personal', kind: 'warn', personal: true, text: `Personal information: ${parts.join('; ')}.` });
  }
  const other = props_.filter((x) => !PERSONAL.has(x.label));
  const custom = p.custom.filter((x) => !/^MSIP_Label_/.test(x.name) && x.value);
  if (other.length || custom.length) {
    const parts = [other.map((x) => `${x.label} ${quoteValue(x.value)}`).join(', '), custom.length && `custom: ${custom.map((x) => `${x.name} ${quoteValue(x.value)}`).join(', ')}`].filter(Boolean);
    out.push({ id: 'properties', kind: 'info', text: `Document properties: ${parts.join('; ')}. They go with the file: change them in Word (File › Info) if they shouldn’t.` });
  }
  if (i.customXml) out.push({ id: 'customXml', kind: 'info', text: `${plural(i.customXml, 'custom XML data part')}, such as document-library properties or data linked to fields. Kept as it is: check ${i.customXml === 1 ? 'it' : 'them'} with Word’s Document Inspector if ${i.customXml === 1 ? 'it' : 'they'} may hold information you don’t want to release.` });
  if (i.embedded) out.push({ id: 'embedded', kind: 'info', text: `${plural(i.embedded, 'embedded file')} (such as a spreadsheet behind a chart): ${i.embedded === 1 ? 'it' : 'they'} can hold more than the page shows. Check ${i.embedded === 1 ? 'it' : 'them'} in Word.` });
  if (i.macros) out.push({ id: 'macros', kind: 'warn', text: 'The file contains macros (Visual Basic code). Kept as they are; check them in Word before you release it.' });
  return out;
});

/** Tables, rows, columns and pictures added or removed (in all, as counted before and after). */
function layoutNotes(layout) {
  if (!layout) return [];
  const { before, after } = layout;
  const parts = [];
  const change = (key, one) => {
    const d = after[key] - before[key];
    if (d > 0) parts.push(`${plural(d, one)} added`);
    if (d < 0) parts.push(`${plural(-d, one)} removed`);
  };
  change('tables', 'table');
  const tables = after.tables - before.tables;
  // Rows and columns of tables added or removed whole aren't counted again.
  if (!tables) {
    change('rows', 'table row');
    change('columns', 'table column');
  }
  change('pictures', 'picture');
  return parts.length ? [{ kind: 'info', text: `${parts.join(', ').replace(/^./, (c) => c.toUpperCase())}.` }] : [];
}

const isEmpty = (words) => !words.some((w) => w.text.trim());

/** The words of a change, with long unchanged stretches shortened to a few words around it. */
function excerpt(words) {
  if (words.length === 1) return words;
  return words.flatMap((w, i) => {
    if (w.type !== 'same') return [w];
    const tokens = w.text.match(/\s+|[^\s]+/g) || [];
    const keep = CONTEXT * 2; // words and spaces
    if (tokens.length <= keep * 2) return [w];
    const first = i === 0;
    const last = i === words.length - 1;
    if (first) return [{ type: 'gap' }, { type: 'same', text: tokens.slice(-keep).join('') }];
    if (last) return [{ type: 'same', text: tokens.slice(0, keep).join('') }, { type: 'gap' }];
    return [{ type: 'same', text: tokens.slice(0, keep).join('') }, { type: 'gap' }, { type: 'same', text: tokens.slice(-keep).join('') }];
  });
}
</script>

<style scoped>
.review-backdrop {
  position: fixed; inset: 0; z-index: 20; display: flex; align-items: center; justify-content: center;
  padding: 16px; background: rgb(0 0 0 / 35%);
}
.review {
  width: min(720px, 100%); max-height: min(86vh, 900px); display: flex; flex-direction: column;
  background: var(--surface); color: var(--text); border-radius: 8px; box-shadow: 0 12px 40px rgb(0 0 0 / 30%);
}
header { padding: 18px 20px 6px; }
/* Everything between the title and the buttons scrolls, so Download stays in view. */
.body { flex: 1; min-height: 0; overflow: auto; padding-bottom: 12px; }
h2 { margin: 0; font-size: 18px; }
.summary { margin: 6px 0 0; color: var(--muted); }
.notes { margin: 8px 20px 0; padding: 0; list-style: none; display: flex; flex-direction: column; gap: 6px; font-size: 13px; }
.notes li { padding: 6px 10px; border-radius: 4px; background: var(--hover); }
.notes .good { border-left: 3px solid #2e7d32; }
.notes .warn { border-left: 3px solid #c77700; }
.notes .info { border-left: 3px solid var(--border); }
.release { margin: 14px 20px 0; font-size: 13px; }
.release h3 { margin: 0 0 6px; font-size: 13px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }
.release ul { margin: 0; padding: 0; list-style: none; display: flex; flex-direction: column; gap: 6px; }
.release li { padding: 6px 10px; border-radius: 4px; background: var(--hover); display: flex; flex-wrap: wrap; align-items: center; gap: 6px 12px; }
.release li > span { flex: 1 1 280px; }
.release .title-form { flex: 1 1 100%; display: flex; gap: 6px; }
.release .title-form input { flex: 1; min-width: 0; border: 1px solid var(--border); border-radius: 4px; padding: 3px 8px; font: inherit; background: var(--surface); color: inherit; }
.release .good { border-left: 3px solid #2e7d32; }
.release .warn { border-left: 3px solid #c77700; }
.release .info { border-left: 3px solid var(--border); }
.release button { border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; padding: 3px 10px; cursor: pointer; }
.release button:hover:not(:disabled) { background: var(--hover); }
.release button:disabled, .release input:disabled + * { opacity: .5; cursor: default; }
.release .personal { flex: 1 1 100%; display: flex; gap: 6px; align-items: flex-start; color: var(--muted); }
.release .clean, .release .locked { margin: 0; color: var(--muted); }
.update-fields { display: flex; gap: 6px; align-items: center; margin-top: 6px; font-weight: 600; }
.changes { margin: 12px 0 0; padding: 0 20px; list-style: none; }
.change { padding: 10px 0; border-top: 1px solid var(--border); }
.where { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; font-size: 12px; color: var(--muted); }
.kind { font-weight: 600; }
.tag { padding: 0 6px; border-radius: 8px; background: #efe6fb; color: #6f2dbd; }
.tag.resolved { background: #e6f2ea; color: #1e6b34; }
.words { margin: 4px 0 0; white-space: pre-wrap; overflow-wrap: anywhere; line-height: 1.5; }
del { color: #b3261e; background: rgb(179 38 30 / 8%); }
ins { color: #1e6b34; background: rgb(30 107 52 / 10%); text-decoration: underline; }
.gap, .empty { color: var(--muted); }
.formatting { margin: 4px 0 0; padding-left: 18px; font-size: 13px; color: var(--muted); }
.more { margin: 8px 20px 0; color: var(--muted); font-size: 13px; }
footer {
  display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between;
  padding: 14px 20px; border-top: 1px solid var(--border); margin-top: 12px;
}
.skip { font-size: 13px; color: var(--muted); display: flex; gap: 6px; align-items: center; }
.buttons { display: flex; gap: 8px; margin-left: auto; }
.buttons button { padding: 7px 14px; border-radius: 4px; border: 1px solid var(--border); background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.buttons .primary { background: var(--brand); border-color: var(--brand); color: #fff; font-weight: 600; }
@media (prefers-color-scheme: dark) {
  .tag { background: #3a2a52; color: #d1b3ff; }
  del { color: #ff8a80; }
  ins { color: #8fd19e; }
}
</style>

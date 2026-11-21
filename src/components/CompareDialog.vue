<template>
  <div class="cmp-backdrop" @mousedown.self="$emit('close')">
    <section class="cmp-dialog" role="dialog" aria-modal="true" :aria-label="t('Compare documents')" @keydown.esc.stop="$emit('close')">
      <header>
        <h2>{{ t('Compare documents') }}</h2>
        <button ref="closeButton" type="button" class="close" :aria-label="t('Close')" @click="$emit('close')">×</button>
      </header>
      <p class="what">{{ t('From {other} to this document: text taken out is struck through, text put in is underlined. Only the text is compared, not its formatting.', { other }) }}</p>
      <p class="summary" role="status">{{ summary }}</p>
      <ol v-if="result.changes.length" class="changes">
        <li v-for="(c, i) in result.changes" :key="i" :class="c.kind">
          <button type="button" class="go" :title="t('Go to it in this document')" @click="$emit('go', c)">
            <span class="kind">{{ t(KINDS[c.kind]) }}</span>
            <span class="text"><template v-for="(p, k) in c.parts" :key="k"><del v-if="p.kind === 'del'">{{ p.text }}</del><ins v-else-if="p.kind === 'ins'">{{ p.text }}</ins><span v-else>{{ shorten(p.text, k, c.parts.length) }}</span></template></span>
          </button>
        </li>
      </ol>
      <!-- A redline, as Word's Compare makes: a new document with these changes as tracked changes. -->
      <p v-if="result.changes.length" id="cmp-redline" class="redline-note">{{ t('A redline is a new Word document: this one, with these changes marked as tracked changes by {author}, to send back or to accept and reject in Word. This document stays as it is.', { author }) }}{{ protected ? ` ${t('It is encrypted with this document’s password.')}` : '' }}</p>
      <div class="actions">
        <template v-if="result.changes.length">
          <button type="button" class="confirm" aria-describedby="cmp-redline" @click="$emit('redline', 'open')">{{ t('Open redline') }}</button>
          <button type="button" class="secondary" aria-describedby="cmp-redline" @click="$emit('redline', 'download')">{{ t('Download redline') }}</button>
        </template>
        <button type="button" :class="result.changes.length ? 'secondary' : 'confirm'" @click="$emit('close')">{{ t('Close') }}</button>
      </div>
    </section>
  </div>
</template>

<script setup>
// The result of Compare documents (see editor/compare.js): each paragraph changed, added or taken out
// since the other version, its words struck through or underlined as Word's Compare marks them; a
// click goes to it in the document. The changes can be made into a redline (see docx/redline.js).
import { computed, ref, onMounted } from 'vue';
import { t } from '../i18n.js';

const props = defineProps({
  result: { type: Object, required: true },
  other: { type: String, required: true },
  author: { type: String, default: 'Author' },
  protected: { type: Boolean, default: false },
});
defineEmits(['close', 'go', 'redline']);
const closeButton = ref(null);
onMounted(() => closeButton.value?.focus());

const KINDS = { changed: 'Changed', added: 'Added', removed: 'Taken out' };
const summary = computed(() => {
  const { changed, added, removed } = props.result;
  if (!changed && !added && !removed) return t('The text is the same in both.');
  return t('{changed} changed, {added} added, {removed} taken out (paragraphs)', { changed, added, removed });
});

/** Unchanged text around a change, shortened to the words near it. */
function shorten(text, k, n) {
  const max = 60;
  if (text.length <= max * 2) return text;
  if (k === 0) return `…${text.slice(-max)}`;
  if (k === n - 1) return `${text.slice(0, max)}…`;
  return `${text.slice(0, max)} … ${text.slice(-max)}`;
}
</script>

<style scoped>
.cmp-backdrop {
  position: fixed; inset: 0; z-index: 20; display: flex; align-items: center; justify-content: center;
  padding: 16px; background: rgb(0 0 0 / 35%);
}
.cmp-dialog {
  width: min(680px, 100%); max-height: calc(100vh - 32px); display: flex; flex-direction: column; padding: 0 20px 18px; box-sizing: border-box;
  background: var(--surface); color: var(--text); border-radius: 8px; box-shadow: 0 12px 40px rgb(0 0 0 / 30%); font-size: 13px;
}
header { display: flex; align-items: center; justify-content: space-between; padding: 16px 0 4px; }
h2 { margin: 0; font-size: 18px; }
.close { border: 0; background: none; color: inherit; font-size: 22px; line-height: 1; width: 32px; height: 32px; border-radius: 4px; cursor: pointer; }
.close:hover { background: var(--hover); }
.what { margin: 4px 0; color: var(--muted); }
.summary { margin: 4px 0 8px; font-weight: 600; }
.changes { list-style: none; margin: 0; padding: 0; overflow: auto; border-top: 1px solid var(--border); }
.changes li { border-bottom: 1px solid var(--border); }
.go { display: flex; gap: 10px; width: 100%; padding: 8px 4px; border: 0; background: none; color: inherit; font: inherit; text-align: left; cursor: pointer; }
.go:hover { background: var(--hover); }
.go:focus-visible { outline: 2px solid var(--accent); outline-offset: -2px; }
.kind { flex: none; width: 76px; color: var(--muted); font-size: 12px; }
.text { white-space: pre-wrap; overflow-wrap: anywhere; }
del { color: var(--danger); text-decoration: line-through; }
ins { color: var(--text); text-decoration: underline; text-decoration-thickness: 2px; background: var(--active); }
.redline-note { margin: 12px 0 0; color: var(--muted); }
.actions { display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end; margin-top: 12px; }
.actions button.secondary { background: var(--surface); color: var(--text); border-color: var(--border); }
.actions button { font: inherit; padding: 6px 16px; border-radius: 6px; border: 1px solid var(--accent); background: var(--accent); color: var(--surface); cursor: pointer; }
</style>

<template>
  <ModalDialog title="Bookmark" @close="$emit('close')" @submit="add">
    <label>Bookmark name
      <input ref="nameEl" v-model.trim="name" aria-label="Bookmark name" maxlength="40" autocomplete="off" spellcheck="false" @input="error = ''" />
    </label>
    <p v-if="error" class="error" role="alert">{{ error }}</p>
    <p v-else class="hint">{{ selection ? 'Marks the selected text' : 'Marks the place at the caret' }}: a name of letters, numbers and underscores, starting with a letter, as Word names them. A cross-reference can then point to it.</p>
    <div v-if="bookmarks.length" class="items" role="listbox" aria-label="Bookmarks" tabindex="0" @keydown="onKeys">
      <div v-for="(b, i) in bookmarks" :key="b.name" role="option" :aria-selected="chosen === i" :class="{ on: chosen === i }" @click="choose(i)" @dblclick="choose(i); $emit('goto', b.name)">
        <span class="name">{{ b.name }}</span><span class="text">{{ b.text }}</span>
      </div>
    </div>
    <p v-else class="hint">The document has no bookmarks yet.</p>
    <template #actions>
      <button type="button" :disabled="chosen < 0" @click="$emit('remove', bookmarks[chosen].name)">Delete</button>
      <button type="button" :disabled="chosen < 0" @click="$emit('goto', bookmarks[chosen].name)">Go to</button>
      <button type="submit" class="confirm">{{ exists ? 'Move here' : 'Add' }}</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Word's Bookmark dialog (Insert > Bookmark): name the selected text or the place at the caret, to
// refer to it (Cross-reference > Bookmark), go to a bookmark, or delete one. A name already used
// moves that bookmark here, as in Word.
import { ref, computed, onMounted } from 'vue';
import ModalDialog from './ModalDialog.vue';

// bookmarks: [{ name, text }] in the document's order (Word's hidden ones, _Toc and _Ref, left out);
// selection: whether text is selected.
const props = defineProps({ bookmarks: { type: Array, required: true }, selection: { type: Boolean, default: false } });
const emit = defineEmits(['close', 'add', 'goto', 'remove']);
const name = ref('');
const error = ref('');
const chosen = ref(-1);
const nameEl = ref(null);
const exists = computed(() => props.bookmarks.some((b) => b.name.toLowerCase() === name.value.toLowerCase()));

// Word's rules for a bookmark's name: up to 40 letters, digits and underscores, a letter first.
const VALID = /^\p{L}[\p{L}\p{N}_]{0,39}$/u;

function choose(i) {
  chosen.value = i;
  name.value = props.bookmarks[i].name;
  error.value = '';
}
function onKeys(e) {
  const step = { ArrowDown: 1, ArrowUp: -1 }[e.key];
  if (!step || !props.bookmarks.length) return;
  e.preventDefault();
  choose(Math.max(0, Math.min(props.bookmarks.length - 1, chosen.value + step)));
}
function add() {
  if (!name.value) {
    error.value = 'Type a name for the bookmark';
    return;
  }
  if (!VALID.test(name.value)) {
    error.value = /^\p{N}|^_/u.test(name.value) ? 'A bookmark’s name starts with a letter' : 'A bookmark’s name has only letters, numbers and underscores (no spaces)';
    return;
  }
  emit('add', name.value);
}
onMounted(() => nameEl.value?.focus());
</script>

<style scoped>
.items { max-height: 200px; overflow: auto; border: 1px solid var(--border); border-radius: 4px; margin: 8px 0; }
.items div { display: flex; gap: 10px; padding: 4px 8px; cursor: pointer; white-space: nowrap; overflow: hidden; }
.items div:hover { background: var(--hover); }
.items div.on { background: var(--active); color: var(--accent); }
.items .name { font-weight: 600; flex: none; }
.items .text { color: var(--muted); overflow: hidden; text-overflow: ellipsis; }
.items:focus-visible { outline: 2px solid var(--accent); }
</style>

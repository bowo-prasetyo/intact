<template>
  <aside class="comments" aria-label="Comments">
    <header>
      <h2>Comments <span class="count">{{ threads.length }}</span></h2>
      <button type="button" class="close" title="Hide comments" aria-label="Hide comments" @click="$emit('close')">✕</button>
    </header>
    <p class="note">Add, reply to, edit, resolve and delete comments here.</p>
    <form v-if="composing" class="composer" aria-label="New comment" @submit.prevent="post" @keydown.esc.stop="$emit('cancel')">
      <!-- No browser autocomplete: its list of earlier entries would cover the buttons below. -->
      <label v-if="askName" class="field">Your name
        <input ref="nameInput" v-model="nameDraft" type="text" maxlength="100" autocomplete="off" />
      </label>
      <label class="field">New comment
        <textarea ref="textInput" v-model="text" rows="3" @keydown.enter.ctrl.prevent="post" @keydown.enter.meta.prevent="post" />
      </label>
      <p v-if="error" class="error" role="alert">{{ error }}</p>
      <div class="actions">
        <button type="button" @click="$emit('cancel')">Cancel</button>
        <button type="submit" class="primary" title="Comment (Ctrl+Enter)">Comment</button>
      </div>
    </form>
    <ol>
      <li v-for="t in threads" :key="t.id">
        <article
          class="thread"
          :class="{ active: t.id === active, done: t.done }"
          tabindex="0"
          :aria-current="t.id === active"
          @click="$emit('select', t.id)"
          @keydown.enter.prevent="$emit('select', t.id)"
        >
          <div v-for="c in [t, ...t.replies]" :key="c.id" class="comment" :class="{ reply: c !== t }">
            <div class="meta">
              <span class="avatar" aria-hidden="true">{{ initials(c) }}</span>
              <span class="author">{{ c.author }}</span>
              <time v-if="c.date" :datetime="c.date">{{ formatDate(c.date) }}</time>
            </div>
            <form v-if="editing?.id === c.id" class="inline" aria-label="Edit comment" @submit.prevent="saveEdit" @click.stop @keydown.stop @keydown.esc="editing = null">
              <textarea ref="editInput" v-model="editing.text" rows="3" aria-label="Comment text" @keydown.enter.ctrl.prevent="saveEdit" @keydown.enter.meta.prevent="saveEdit" />
              <div class="actions">
                <button type="button" @click="editing = null">Cancel</button>
                <button type="submit" class="primary">Save</button>
              </div>
            </form>
            <template v-else>
              <p v-for="(text, i) in c.paragraphs" :key="i" class="text">{{ text }}</p>
              <p v-if="!c.paragraphs.length" class="text empty">(no text)</p>
              <div v-if="t.anchored" class="comment-actions" @click.stop>
                <button type="button" class="link" :title="`Edit this comment by ${c.author}`" @click="startEdit(c)">Edit</button>
                <button v-if="c !== t" type="button" class="link" title="Delete this reply (Undo brings it back)" @click="$emit('delete', c.id)">Delete</button>
              </div>
            </template>
          </div>
          <form v-if="replying?.id === t.id" class="inline" aria-label="Reply" @submit.prevent="saveReply" @click.stop @keydown.stop @keydown.esc="replying = null">
            <label v-if="askName" class="field">Your name
              <input ref="replyName" v-model="replying.name" type="text" maxlength="100" autocomplete="off" />
            </label>
            <textarea ref="replyInput" v-model="replying.text" rows="2" aria-label="Reply text" @keydown.enter.ctrl.prevent="saveReply" @keydown.enter.meta.prevent="saveReply" />
            <p v-if="replying.error" class="error" role="alert">{{ replying.error }}</p>
            <div class="actions">
              <button type="button" @click="replying = null">Cancel</button>
              <button type="submit" class="primary">Reply</button>
            </div>
          </form>
          <div class="tags">
            <span v-if="t.n" class="tag">{{ t.n }}</span>
            <span v-else class="tag muted">{{ t.place || 'Not anchored in the text' }}</span>
            <span v-if="t.done" class="tag resolved">Resolved</span>
            <span v-if="t.isNew" class="tag new">New</span>
          </div>
          <div v-if="t.anchored" class="thread-actions" @click.stop>
            <button type="button" @click="startReply(t)">Reply</button>
            <button type="button" @click="$emit('resolve', { id: t.id, done: !t.done })">{{ t.done ? 'Reopen' : 'Resolve' }}</button>
            <button type="button" title="Delete this comment and its replies (Undo brings them back)" @click="$emit('delete', t.id)">Delete</button>
          </div>
          <p v-else class="hint">Not in the document’s text: reply to, edit or resolve it in Word.</p>
        </article>
      </li>
    </ol>
  </aside>
</template>

<script setup>
import { nextTick, ref, watch } from 'vue';

const props = defineProps({
  threads: { type: Array, required: true }, // [{ id, n, author, date, paragraphs, done, isNew, replies: [...] }]
  active: { type: String, default: null },
  composing: { type: Boolean, default: false }, // show the box for a new comment
  askName: { type: Boolean, default: false }, // the author's name isn't known yet
});
const emit = defineEmits(['select', 'close', 'post', 'cancel', 'delete', 'reply', 'edit', 'resolve']);

const text = ref('');
const nameDraft = ref('');
const error = ref('');
const textInput = ref(null);
const nameInput = ref(null);

watch(() => props.composing, async (on) => {
  if (!on) return;
  text.value = '';
  error.value = '';
  await nextTick();
  (props.askName ? nameInput.value : textInput.value)?.focus();
}, { immediate: true });

// Editing a comment, or replying to a thread: one at a time.
const editing = ref(null); // { id, text }
const replying = ref(null); // { id, text, name, error }
const editInput = ref(null);
const replyInput = ref(null);
const replyName = ref(null);
const first = (r) => (Array.isArray(r.value) ? r.value[0] : r.value);

async function startEdit(c) {
  replying.value = null;
  editing.value = { id: c.id, text: c.paragraphs.join('\n') };
  await nextTick();
  first(editInput)?.focus();
}

function saveEdit() {
  const { id, text } = editing.value;
  editing.value = null;
  emit('edit', { id, text: text.trim() });
}

async function startReply(t) {
  editing.value = null;
  replying.value = { id: t.id, text: '', name: '', error: '' };
  await nextTick();
  (props.askName ? first(replyName) : first(replyInput))?.focus();
}

function saveReply() {
  const r = replying.value;
  if (props.askName && !r.name.trim()) {
    r.error = 'Your name goes on the reply';
    return;
  }
  if (!r.text.trim()) {
    r.error = 'Write the reply first';
    return;
  }
  replying.value = null;
  emit('reply', { id: r.id, text: r.text.trim(), name: r.name.trim() });
}

function post() {
  if (props.askName && !nameDraft.value.trim()) {
    error.value = 'Your name goes on the comment';
    nameInput.value?.focus();
    return;
  }
  if (!text.value.trim()) {
    error.value = 'Write the comment first';
    textInput.value?.focus();
    return;
  }
  emit('post', { text: text.value.trim(), name: nameDraft.value.trim() });
}

function initials(c) {
  if (c.initials) return c.initials.slice(0, 3);
  return c.author.split(/\s+/).filter(Boolean).map((w) => w[0]).join('').slice(0, 2).toUpperCase() || '?';
}

function formatDate(iso) {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}
</script>

<style scoped>
.comments {
  width: 300px; flex: none; overflow: auto; background: var(--surface); border-left: 1px solid var(--border);
  display: flex; flex-direction: column; font-size: 13px;
}
header { display: flex; align-items: center; justify-content: space-between; padding: 10px 12px 0; }
h2 { font-size: 14px; margin: 0; }
.count { color: var(--muted); font-weight: normal; margin-left: 4px; }
.close { border: 0; background: transparent; color: inherit; cursor: pointer; font-size: 14px; padding: 4px 6px; border-radius: 4px; }
.close:hover { background: var(--hover); }
.note { margin: 4px 12px 8px; color: var(--muted); font-size: 12px; }
ol { list-style: none; margin: 0; padding: 0 8px 12px; display: flex; flex-direction: column; gap: 8px; }
.thread {
  border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; cursor: pointer;
  background: var(--surface); outline: none;
}
.thread:hover { border-color: #c9a29a; }
.thread:focus-visible, .thread.active { border-color: #d0705c; box-shadow: 0 0 0 2px #f3c3b8; }
.thread.done { opacity: .75; }
.comment + .comment { margin-top: 8px; padding-top: 8px; border-top: 1px dashed var(--border); }
.reply { margin-left: 14px; }
.meta { display: flex; align-items: center; gap: 6px; }
.avatar {
  width: 22px; height: 22px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center;
  background: #d0705c; color: #fff; font-size: 10px; font-weight: 600; flex: none;
}
.author { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
time { color: var(--muted); margin-left: auto; white-space: nowrap; }
.text { margin: 4px 0 0; white-space: pre-wrap; overflow-wrap: anywhere; }
.text.empty { color: var(--muted); font-style: italic; }
.tags { display: flex; gap: 6px; margin-top: 6px; }
.tag { font-size: 11px; padding: 1px 6px; border-radius: 9px; background: #f9ddd6; color: #7a2e1f; }
.tag.muted { background: var(--hover); color: var(--muted); }
.tag.resolved { background: #dcefe0; color: #1e6b34; }
.tag.new { background: var(--active); color: var(--accent); }
.comment-actions { display: flex; gap: 8px; margin-top: 2px; }
.link { border: 0; background: transparent; color: var(--muted); font: inherit; font-size: 12px; cursor: pointer; padding: 0; text-decoration: underline dotted; }
.link:hover { color: inherit; }
.thread-actions { display: flex; gap: 6px; margin-top: 8px; }
.thread-actions button { border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; font-size: 12px; padding: 2px 8px; cursor: pointer; }
.thread-actions button:hover { background: var(--hover); }
.inline { display: flex; flex-direction: column; gap: 6px; margin-top: 6px; }
.inline textarea { font: inherit; font-size: 13px; color: var(--text); background: var(--surface); border: 1px solid var(--border); border-radius: 4px; padding: 5px 6px; resize: vertical; }
.hint { margin: 6px 0 0; color: var(--muted); font-size: 12px; }
.composer { margin: 0 8px 8px; padding: 8px 10px; border: 1px solid #d0705c; border-radius: 8px; display: flex; flex-direction: column; gap: 6px; }
.field { display: flex; flex-direction: column; gap: 3px; font-size: 12px; color: var(--muted); }
.field input, .field textarea { font: inherit; font-size: 13px; color: var(--text); background: var(--surface); border: 1px solid var(--border); border-radius: 4px; padding: 5px 6px; resize: vertical; }
.error { margin: 0; color: var(--danger); font-size: 12px; }
.actions { display: flex; justify-content: flex-end; gap: 6px; }
.actions button { border: 1px solid var(--border); border-radius: 4px; padding: 4px 10px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.actions .primary { background: #d0705c; border-color: #d0705c; color: #fff; }

@media (max-width: 860px) {
  .comments {
    position: fixed; left: 0; right: 0; bottom: 0; width: auto; max-height: 45vh; z-index: 5;
    border-left: 0; border-top: 1px solid var(--border); box-shadow: 0 -4px 16px rgb(0 0 0 / 12%);
  }
}
</style>

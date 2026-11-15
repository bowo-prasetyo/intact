<template>
  <aside class="comments" aria-label="Comments">
    <header>
      <h2>Comments <span class="count">{{ threads.length }}</span></h2>
      <button type="button" class="close" title="Hide comments" aria-label="Hide comments" @click="$emit('close')">✕</button>
    </header>
    <p class="note">Shown read-only – reply to or resolve comments in Word.</p>
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
            <p v-for="(text, i) in c.paragraphs" :key="i" class="text">{{ text }}</p>
            <p v-if="!c.paragraphs.length" class="text empty">(no text)</p>
          </div>
          <div class="tags">
            <span v-if="t.n" class="tag">{{ t.n }}</span>
            <span v-else class="tag muted">Not anchored in the text</span>
            <span v-if="t.done" class="tag resolved">Resolved</span>
          </div>
        </article>
      </li>
    </ol>
  </aside>
</template>

<script setup>
defineProps({
  threads: { type: Array, required: true }, // [{ id, n, author, date, paragraphs, done, replies: [...] }]
  active: { type: String, default: null },
});
defineEmits(['select', 'close']);

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

@media (max-width: 860px) {
  .comments {
    position: fixed; left: 0; right: 0; bottom: 0; width: auto; max-height: 45vh; z-index: 5;
    border-left: 0; border-top: 1px solid var(--border); box-shadow: 0 -4px 16px rgb(0 0 0 / 12%);
  }
}
</style>

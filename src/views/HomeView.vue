<template>
  <div class="home">
    <AppHeader />
    <main class="home-main">
      <h1>Change what you need.<br /> Keep the rest <span class="brand-word">intact</span>.</h1>
      <p class="lead">Open a <code>.docx</code>, fix it and download it again, right here in your browser.</p>
      <ul class="promises" aria-label="What Intact promises">
        <li v-for="p in promises" :key="p.title">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3.5 8.5l3 3 6-7" /></svg>
          <span><strong>{{ p.title }}</strong> – {{ p.text }}</span>
        </li>
      </ul>

      <div class="actions">
        <button class="tile" type="button" :disabled="busy" @click="createNew">
          <span class="tile-icon" aria-hidden="true">＋</span>
          <span class="tile-title">Blank document</span>
          <span class="tile-sub">Start from scratch</span>
        </button>

        <label
          class="tile drop"
          :class="{ over: dragOver }"
          @dragover.prevent="dragOver = true"
          @dragleave="dragOver = false"
          @drop.prevent="onDrop"
        >
          <span class="tile-icon" aria-hidden="true">📂</span>
          <span class="tile-title">Open a .docx file</span>
          <span class="tile-sub">Click to browse or drop a file here</span>
          <input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" hidden :disabled="busy" @change="onPick" />
        </label>
      </div>

      <p v-if="busy" class="status">Opening…</p>
      <p v-if="error" class="error" role="alert">{{ error }}</p>
    </main>
  </div>
</template>

<script setup>
import { ref } from 'vue';
import { useRouter } from 'vue-router';
import AppHeader from '../components/AppHeader.vue';
import { openFile, newDocument } from '../store.js';

const router = useRouter();
const busy = ref(false);
const error = ref('');
const dragOver = ref(false);

// The promises from VISION.md. Keep them true: a claim here that a release breaks is a bug.
const promises = [
  { title: 'Private', text: 'never uploaded' },
  { title: 'Faithful', text: 'only your edits change' },
  { title: 'Instant', text: 'nothing to install' },
];

async function run(task) {
  busy.value = true;
  error.value = '';
  try {
    await task();
    router.push({ name: 'edit' });
  } catch (e) {
    console.error(e);
    error.value = `Could not open the document: ${e.message}`;
  } finally {
    busy.value = false;
  }
}

function load(file) {
  if (!file) return;
  if (!/\.docx$/i.test(file.name)) {
    error.value = 'Please choose a .docx file (older .doc files are not supported).';
    return;
  }
  run(() => openFile(file));
}

const onPick = (e) => { load(e.target.files[0]); e.target.value = ''; };
const onDrop = (e) => { dragOver.value = false; load(e.dataTransfer.files[0]); };
const createNew = () => run(newDocument);
</script>

<style scoped>
.home { min-height: 100vh; display: flex; flex-direction: column; }
.home-main { max-width: 760px; margin: 0 auto; padding: 48px 16px; text-align: center; }
h1 { font-size: clamp(1.6rem, 4.5vw, 2.5rem); line-height: 1.2; letter-spacing: -0.01em; margin: 0 0 14px; text-wrap: balance; }
.brand-word { color: var(--accent); }
.lead { color: var(--muted); font-size: 1.05rem; margin: 0 0 16px; text-wrap: balance; }
.promises {
  display: flex; flex-wrap: wrap; justify-content: center; gap: 8px 22px;
  list-style: none; padding: 0; margin: 0 0 36px; font-size: .95rem;
}
.promises li { display: flex; align-items: center; gap: 6px; color: var(--muted); }
.promises strong { color: var(--text); font-weight: 600; }
.promises svg { flex: none; fill: none; stroke: var(--accent); stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }
.actions { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; }
.tile {
  display: flex; flex-direction: column; align-items: center; gap: 6px; padding: 32px 20px;
  border: 1px solid var(--border); border-radius: 12px; background: var(--surface);
  cursor: pointer; font: inherit; color: inherit; transition: border-color .15s, box-shadow .15s;
}
.tile:hover, .drop.over { border-color: var(--accent); box-shadow: 0 4px 16px rgb(0 0 0 / 8%); }
.drop { border-style: dashed; }
.tile-icon { font-size: 2rem; line-height: 1; }
.tile-title { font-weight: 600; font-size: 1.1rem; }
.tile-sub { color: var(--muted); font-size: .9rem; }
.status { margin-top: 24px; color: var(--muted); }
.error { margin-top: 24px; color: var(--danger); }
</style>

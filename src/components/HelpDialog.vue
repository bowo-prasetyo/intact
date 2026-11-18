<template>
  <div class="help-backdrop" @mousedown.self="$emit('close')">
    <section class="help" role="dialog" aria-modal="true" aria-labelledby="help-title" @keydown.esc.stop="$emit('close')">
      <header>
        <h2 id="help-title">Help</h2>
        <button type="button" class="close" aria-label="Close help" @click="$emit('close')">×</button>
      </header>
      <div class="body">
        <section aria-labelledby="help-checks">
          <h3 id="help-checks">Check the document</h3>
          <p>Any time, not only when you download.</p>
          <div class="checks">
            <button ref="first" type="button" @click="$emit('check', '')">
              <strong>Review your changes</strong><span>What changed, and what goes back with the file</span>
            </button>
            <button type="button" @click="$emit('check', 'release')">
              <strong>Before you release</strong><span>Comments, tracked changes, hidden text and personal information in the file</span>
            </button>
            <button type="button" @click="$emit('check', 'accessibility')">
              <strong>Check accessibility</strong><span>Alt text, headings, tables and more, for people using screen readers</span>
            </button>
            <button type="button" @click="$emit('check', 'screening')">
              <strong>Check a CV for screening software</strong><span>Contact details, text boxes, columns and pictures that employers’ software may miss</span>
            </button>
          </div>
        </section>
        <section aria-labelledby="help-keys">
          <h3 id="help-keys">Keyboard shortcuts</h3>
          <p v-if="mac">On a Mac, use Cmd where it says Ctrl.</p>
          <table>
            <tbody>
              <tr v-for="[keys, what] in SHORTCUTS" :key="keys">
                <th scope="row"><kbd>{{ keys }}</kbd></th>
                <td>{{ what }}</td>
              </tr>
            </tbody>
          </table>
        </section>
      </div>
    </section>
  </div>
</template>

<script setup>
// Help: the checks the review offers, reachable without downloading, and Word's keyboard shortcuts
// as Intact supports them (the list in editor/shortcuts.js, also on the About page).
import { onMounted, ref } from 'vue';
import { SHORTCUTS } from '../editor/shortcuts.js';

defineEmits(['close', 'check']);
const first = ref(null);
const mac = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
onMounted(() => first.value?.focus());
</script>

<style scoped>
.help-backdrop {
  position: fixed; inset: 0; z-index: 20; display: flex; align-items: center; justify-content: center;
  padding: 16px; background: rgb(0 0 0 / 35%);
}
.help {
  width: min(640px, 100%); max-height: min(86vh, 900px); display: flex; flex-direction: column;
  background: var(--surface); color: var(--text); border-radius: 8px; box-shadow: 0 12px 40px rgb(0 0 0 / 30%);
}
header { display: flex; align-items: center; justify-content: space-between; padding: 16px 20px 4px; }
h2 { margin: 0; font-size: 18px; }
.close { border: 0; background: none; color: inherit; font-size: 22px; line-height: 1; width: 32px; height: 32px; border-radius: 4px; cursor: pointer; }
.close:hover { background: var(--hover); }
.body { overflow: auto; padding: 0 20px 18px; font-size: 13px; }
h3 { margin: 14px 0 4px; font-size: 13px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }
p { margin: 0 0 8px; color: var(--muted); }
.checks { display: grid; gap: 6px; }
.checks button {
  display: flex; flex-direction: column; align-items: flex-start; gap: 2px; text-align: left; padding: 8px 12px;
  border: 1px solid var(--border); border-radius: 6px; background: var(--surface); color: inherit; font: inherit; cursor: pointer;
}
.checks button:hover { background: var(--hover); border-color: var(--accent); }
.checks span { color: var(--muted); }
button:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
table { width: 100%; border-collapse: collapse; }
th, td { padding: 5px 0; border-top: 1px solid var(--border); text-align: left; vertical-align: top; font-weight: normal; }
th { padding-right: 16px; width: 42%; }
kbd { font: 12px ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
</style>

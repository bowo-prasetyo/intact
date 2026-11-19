<template>
  <div class="sym-backdrop" @mousedown.self="$emit('close')">
    <section class="sym-dialog" role="dialog" aria-modal="true" aria-labelledby="sym-title" @keydown.esc.stop="$emit('close')">
      <header>
        <h2 id="sym-title">Symbol</h2>
        <button type="button" class="close" aria-label="Close" @click="$emit('close')">×</button>
      </header>
      <input
        ref="searchInput"
        v-model="query"
        class="search"
        type="search"
        autocomplete="off"
        spellcheck="false"
        placeholder="Search: euro, tick, arrow… or a code such as 20AC"
        aria-label="Search symbols by name, or type a character code"
        @keydown.enter.prevent="insertFirst"
        @keydown.down.prevent="focusFirst"
      />
      <template v-if="!query.trim()">
        <template v-if="recent.length">
          <p class="label">Recently used</p>
          <div class="grid" role="group" aria-label="Recently used symbols" @keydown="onGridKeydown">
            <button v-for="c in recent" :key="`r${c}`" type="button" class="sym" :title="describe(c)" :aria-label="describe(c)" @focus="shown = c" @mouseenter="shown = c" @click="pick(c)">{{ visible(c) }}</button>
          </div>
        </template>
        <div class="tabs" role="tablist" aria-label="Groups of symbols">
          <button v-for="(g, i) in SYMBOL_GROUPS" :key="g.label" type="button" role="tab" :aria-selected="tab === i" :class="{ on: tab === i }" @click="tab = i">{{ g.label }}</button>
        </div>
      </template>
      <p v-else class="label" role="status">{{ results.length ? `${results.length} found` : 'No symbol by that name: try another word, or a character code such as 00A9' }}</p>
      <div ref="grid" class="grid main" role="group" :aria-label="query.trim() ? 'Symbols found' : SYMBOL_GROUPS[tab].label" @keydown="onGridKeydown">
        <button v-for="s in list" :key="s.char" type="button" class="sym" :title="describe(s.char)" :aria-label="describe(s.char)" @focus="shown = s.char" @mouseenter="shown = s.char" @click="pick(s.char)">{{ visible(s.char) }}</button>
      </div>
      <p class="preview" aria-hidden="true">
        <template v-if="shown"><span class="big">{{ visible(shown) }}</span>{{ nameOf(shown) }} · {{ codeOf(shown) }}</template>
        <template v-else>Click a symbol to put it at the caret.</template>
      </p>
      <div class="actions">
        <button type="button" @click="$emit('close')">Close</button>
      </div>
    </section>
  </div>
</template>

<script setup>
// Insert Symbol, as Word's Ω Symbol: groups of the symbols people look for, a search by name (or a
// character's code, as Word's "Character code" box takes it) and the symbols used last. A click puts
// the symbol at the caret, as a key press would (tracked while tracking), and closes the dialog.
import { ref, computed, onMounted, nextTick } from 'vue';
import { SYMBOL_GROUPS, searchSymbols, charFromCode, codeOf, nameOf, recentSymbols, rememberSymbol } from '../editor/symbols.js';

// insert(char): the symbol chosen.
const emit = defineEmits(['close', 'insert']);

const query = ref('');
const tab = ref(0);
const recent = ref(recentSymbols());
const shown = ref(null);
const searchInput = ref(null);
const grid = ref(null);

const results = computed(() => {
  const found = searchSymbols(query.value);
  const byCode = charFromCode(query.value);
  return byCode && !found.some((s) => s.char === byCode) ? [{ char: byCode, name: nameOf(byCode) }, ...found] : found;
});
const list = computed(() => (query.value.trim() ? results.value : SYMBOL_GROUPS[tab.value].symbols));

/** Spaces and invisible characters shown as a sign of what they are, in the grid. */
const visible = (c) => ({ ' ': '°', ' ': '⁝', ' ': '°', ' ': '⎵', ' ': '⎵', '­': '¬', '​': '⁞' }[c] || c);
const describe = (c) => `${nameOf(c)} (${codeOf(c)})`;

function pick(char) {
  recent.value = rememberSymbol(char);
  emit('insert', char);
}

function insertFirst() {
  if (list.value.length) pick(list.value[0].char);
}

function focusFirst() {
  grid.value?.querySelector('.sym')?.focus();
}

/** Arrow keys move around a grid of symbols (Enter and Space choose one, as buttons do). */
function onGridKeydown(e) {
  const buttons = [...e.currentTarget.querySelectorAll('.sym')];
  const at = buttons.indexOf(document.activeElement);
  if (at < 0) return;
  const cols = Math.max(1, Math.round(e.currentTarget.clientWidth / buttons[0].offsetWidth));
  const step = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: cols, ArrowUp: -cols, Home: -at, End: buttons.length - 1 - at }[e.key];
  if (step == null) return;
  e.preventDefault();
  if (e.key === 'ArrowUp' && at - cols < 0 && e.currentTarget === grid.value) {
    searchInput.value?.focus();
    return;
  }
  buttons[Math.max(0, Math.min(buttons.length - 1, at + step))]?.focus();
}

onMounted(() => nextTick(() => searchInput.value?.focus()));
</script>

<style scoped>
.sym-backdrop {
  position: fixed; inset: 0; z-index: 20; display: flex; align-items: center; justify-content: center;
  padding: 16px; background: rgb(0 0 0 / 35%);
}
.sym-dialog {
  width: min(520px, 100%); max-height: calc(100vh - 32px); overflow: auto; padding: 0 20px 18px; box-sizing: border-box;
  background: var(--surface); color: var(--text); border-radius: 8px; box-shadow: 0 12px 40px rgb(0 0 0 / 30%);
  font-size: 13px;
}
header { display: flex; align-items: center; justify-content: space-between; padding: 16px 0 8px; }
h2 { margin: 0; font-size: 18px; }
.close { border: 0; background: none; color: inherit; font-size: 22px; line-height: 1; width: 32px; height: 32px; border-radius: 4px; cursor: pointer; }
.close:hover { background: var(--hover); }
.search { width: 100%; box-sizing: border-box; padding: 6px 8px; font: inherit; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; }
.label { margin: 10px 0 4px; color: var(--muted); font-size: 12px; }
.tabs { display: flex; flex-wrap: wrap; gap: 4px; margin: 12px 0 6px; }
.tabs button { padding: 3px 9px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface); color: inherit; font: inherit; font-size: 12px; cursor: pointer; }
.tabs button.on { background: var(--active); border-color: var(--accent); color: var(--accent); }
.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(36px, 1fr)); gap: 4px; }
.grid.main { min-height: 120px; align-content: start; }
.sym {
  height: 36px; padding: 0; border: 1px solid var(--border); border-radius: 4px; background: #fff; color: #1a1a1a;
  font-size: 18px; line-height: 1; cursor: pointer; font-family: Calibri, Carlito, 'Segoe UI Symbol', 'Noto Sans Symbols', 'Noto Sans', sans-serif;
}
.sym:hover { border-color: var(--accent); }
.preview { display: flex; align-items: center; gap: 10px; min-height: 40px; margin: 10px 0 0; color: var(--muted); }
.big { font-size: 28px; color: var(--text); min-width: 32px; text-align: center; }
.actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 10px; }
.actions button { padding: 6px 12px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer; }
.actions button:hover { background: var(--hover); }
button:focus-visible, input:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
</style>

<template>
  <ModalDialog :title="editing ? 'Edit equation' : 'Equation'" @close="$emit('close')" @submit="insert">
    <label>Type the equation, as in LaTeX
      <textarea ref="input" v-model="src" rows="3" spellcheck="false" autocomplete="off" aria-label="Equation, as in LaTeX" placeholder="x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}" @keydown.enter.exact.prevent="insert" />
    </label>
    <div class="examples" role="group" aria-label="Examples">
      <button v-for="e in EXAMPLES" :key="e.src" type="button" :title="e.src" @click="src = e.src">{{ e.label }}</button>
    </div>
    <div class="preview" aria-live="polite">
      <p v-if="result.error" class="error">{{ result.error }}</p>
      <div v-else ref="previewEl" class="maths" aria-label="Preview" />
    </div>
    <label v-if="!editing" class="check"><input v-model="display" type="checkbox" /> On its own line (when the paragraph is empty)</label>
    <p class="hint">In a matrix, cases or lines of working, &amp; separates cells (in working, the point the lines are lined up at) and \\ starts a new row. Written as Word’s own equations (Office Math): Word edits them as its own.<template v-if="editing"> The equation is written again only if it is changed.</template></p>
    <template #actions>
      <button type="submit" class="confirm" :disabled="!!result.error">{{ editing ? 'Update' : 'Insert' }}</button>
    </template>
  </ModalDialog>
</template>

<script setup>
// Insert Equation: the equation typed as LaTeX (see editor/mathInput.js), shown as it will look,
// and put at the caret as Word's Office Math.
import { ref, computed, watch, onMounted, nextTick } from 'vue';
import ModalDialog from './ModalDialog.vue';
import { latexToOmml } from '../editor/mathInput.js';
import { ommlToMathml, MATH_NS } from '../docx/omml.js';
import { NS, parseXml } from '../docx/xml.js';

// To edit an equation already in the document: its LaTeX (`latex`) and whether it is on its own line.
const props = defineProps({ latex: { type: String, default: '' }, displayed: { type: Boolean, default: false } });
// insert({ xml, display, src, changed }): the equation's OMML, and the LaTeX it was made from.
const emit = defineEmits(['close', 'insert']);
const editing = !!props.latex;
const EXAMPLES = [
  { label: 'x²', src: 'x^2' },
  { label: 'a/b', src: '\\frac{a}{b}' },
  { label: '√x', src: '\\sqrt{x}' },
  { label: 'Σ', src: '\\sum_{i=1}^{n} x_i' },
  { label: '∫', src: '\\int_{0}^{1} f(x)\\,dx' },
  { label: 'Quadratic', src: 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}' },
  { label: 'α β π', src: '\\alpha + \\beta = \\pi' },
  // Matrices, cases and working lined up at its = signs: & between cells, \\ between rows.
  { label: 'Matrix', src: '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}' },
  { label: 'Cases', src: '|x| = \\begin{cases} x & x \\ge 0 \\\\ -x & x < 0 \\end{cases}' },
  { label: 'Working', src: '\\begin{aligned} 2x + 3 &= 11 \\\\ 2x &= 8 \\\\ x &= 4 \\end{aligned}' },
  { label: 'Vector', src: '\\vec{v} = \\overline{AB}' },
];
const src = ref(props.latex);
const display = ref(props.displayed);
const input = ref(null);
const previewEl = ref(null);
const result = computed(() => latexToOmml(src.value, { display: display.value }));

function drawPreview() {
  const el = previewEl.value;
  if (!el || result.value.error) return;
  try {
    const node = parseXml(`<w:p xmlns:w="${NS.w}" xmlns:m="${MATH_NS}">${result.value.xml}</w:p>`).documentElement.firstElementChild;
    el.replaceChildren(...ommlToMathml(node));
  } catch {
    el.textContent = '';
  }
}
watch(result, () => nextTick(drawPreview));

function insert() {
  if (result.value.error) return;
  emit('insert', { xml: result.value.xml, display: display.value, changed: src.value.trim() !== props.latex.trim() });
}
onMounted(() => nextTick(() => {
  drawPreview();
  input.value?.focus();
}));
</script>

<style scoped>
textarea { font-family: Consolas, 'Courier New', monospace !important; resize: vertical; }
.examples { display: flex; flex-wrap: wrap; gap: 4px; }
.examples button { padding: 3px 9px; border: 1px solid var(--border); border-radius: 12px; background: var(--surface); color: inherit; font: inherit; font-size: 12px; cursor: pointer; }
.examples button:hover { border-color: var(--accent); }
.preview { min-height: 56px; margin: 10px 0; padding: 8px; border: 1px solid var(--border); border-radius: 4px; background: #fff; color: #1a1a1a; display: flex; align-items: center; justify-content: center; }
.maths { font-size: 20px; font-family: 'Cambria Math', 'STIX Two Math', 'Latin Modern Math', serif; }
</style>

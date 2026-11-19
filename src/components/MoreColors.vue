<template>
  <div class="more-colors">
    <template v-if="remembered.length">
      <p class="palette-label">Your colours</p>
      <div class="swatches">
        <button v-for="hex in remembered" :key="hex" type="button" class="swatch" :class="{ current: current === hex }" :style="{ background: `#${hex}` }" :title="`#${hex}`" :aria-label="`Your colour #${hex}`" @click="emit('pick', hex)" />
      </div>
    </template>
    <p class="palette-label">More colours</p>
    <form class="hex-row" @submit.prevent="apply">
      <input class="picker" type="color" :value="`#${valid || current || '000000'}`" aria-label="Pick a colour" title="Pick a colour" @input="text = $event.target.value.toUpperCase()" @change="apply" />
      <input v-model="text" class="code" type="text" inputmode="text" maxlength="7" spellcheck="false" autocomplete="off" placeholder="#0E6467" aria-label="Colour code, such as #0E6467" :aria-invalid="text && !valid ? 'true' : null" />
      <button type="submit" :disabled="!valid">Apply</button>
    </form>
  </div>
</template>

<script setup>
// "More colours…" for a palette: a colour picker and a box for a hex code (a brand colour, say), and
// the colours chosen that way before, remembered in this browser (see editor/colors.js).
import { ref, computed } from 'vue';
import { parseHex, rememberedColors, rememberColor } from '../editor/colors.js';

defineProps({
  /** The colour applied now (RRGGBB), to mark its swatch. */
  current: { type: String, default: null },
});
// pick(hex): a colour chosen, as Word's RRGGBB.
const emit = defineEmits(['pick']);

const remembered = ref(rememberedColors());
const text = ref('');
const valid = computed(() => parseHex(text.value));

function apply() {
  const hex = valid.value;
  if (!hex) return;
  remembered.value = rememberColor(hex);
  emit('pick', hex);
}
</script>

<style scoped>
.more-colors { display: flex; flex-direction: column; gap: 6px; }
.palette-label { margin: 6px 0 2px; font-size: 12px; color: var(--muted); }
.swatches { display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; }
.swatch { height: 24px; border: 1px solid rgb(0 0 0 / 25%); border-radius: 3px; cursor: pointer; padding: 0; }
.swatch.current, .swatch:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.hex-row { display: flex; gap: 6px; align-items: center; }
.picker { flex: none; width: 32px; height: 28px; padding: 0 2px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); cursor: pointer; }
.code { flex: 1; min-width: 0; font: inherit; font-size: 13px; padding: 4px 6px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font-variant-numeric: tabular-nums; }
.code[aria-invalid="true"] { border-color: #b3261e; }
.hex-row button { flex: none; font: inherit; font-size: 13px; padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; cursor: pointer; }
.hex-row button:disabled { opacity: .5; cursor: default; }
</style>

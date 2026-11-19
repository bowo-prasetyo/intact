<template>
  <div class="md-backdrop" @mousedown.self="$emit('close')">
    <form class="md-dialog" role="dialog" aria-modal="true" :aria-labelledby="titleId" novalidate @submit.prevent="$emit('submit')" @keydown.esc.stop="$emit('close')">
      <header>
        <h2 :id="titleId">{{ title }}</h2>
        <button type="button" class="close" aria-label="Close" @click="$emit('close')">×</button>
      </header>
      <slot />
      <div class="actions">
        <slot name="actions" />
        <button type="button" @click="$emit('close')">Cancel</button>
      </div>
    </form>
  </div>
</template>

<script setup>
// A dialog over the page, as the References dialogs use (Caption, Cross-reference, Equation,
// Citation): a title, its fields, and its buttons with Cancel; Esc or a click outside closes it.
let next = 0;
defineProps({ title: { type: String, required: true } });
defineEmits(['close', 'submit']);
const titleId = `md-title-${++next}`;
</script>

<style scoped>
.md-backdrop {
  position: fixed; inset: 0; z-index: 20; display: flex; align-items: center; justify-content: center;
  padding: 16px; background: rgb(0 0 0 / 35%);
}
.md-dialog {
  width: min(520px, 100%); max-height: calc(100vh - 32px); overflow: auto; padding: 0 20px 18px; box-sizing: border-box;
  background: var(--surface); color: var(--text); border-radius: 8px; box-shadow: 0 12px 40px rgb(0 0 0 / 30%);
  font-size: 13px;
}
header { display: flex; align-items: center; justify-content: space-between; padding: 16px 0 8px; }
h2 { margin: 0; font-size: 18px; }
.close { border: 0; background: none; color: inherit; font-size: 22px; line-height: 1; width: 32px; height: 32px; border-radius: 4px; cursor: pointer; }
.close:hover { background: var(--hover); }
.actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
.actions :deep(button), .actions button {
  padding: 6px 14px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer;
}
.actions :deep(button.confirm) { background: var(--accent); border-color: var(--accent); color: #fff; }
.actions :deep(button:disabled) { opacity: .5; cursor: default; }
</style>

<style>
/* The fields in a dialog's slot: labels above their inputs. */
.md-dialog label { display: flex; flex-direction: column; gap: 3px; margin: 8px 0; color: var(--muted); font-size: 12px; }
.md-dialog label.check { flex-direction: row; align-items: center; gap: 6px; color: var(--text); font-size: 13px; }
.md-dialog input:not([type=checkbox]):not([type=radio]), .md-dialog select, .md-dialog textarea {
  padding: 6px 8px; font: inherit; font-size: 13px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: var(--text);
}
.md-dialog .hint { margin: 4px 0; color: var(--muted); font-size: 12px; }
.md-dialog .error { margin: 4px 0; color: #b3261e; font-size: 12px; }
.md-dialog .two { display: grid; grid-template-columns: 1fr 1fr; gap: 0 10px; }
.md-dialog fieldset { border: 1px solid var(--border); border-radius: 6px; margin: 10px 0; padding: 4px 10px 8px; }
.md-dialog legend { color: var(--muted); font-size: 12px; padding: 0 4px; }
</style>

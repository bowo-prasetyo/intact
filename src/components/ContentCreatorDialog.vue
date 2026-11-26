<template>
  <div class="creator-backdrop" role="presentation" @mousedown.self="$emit('close')">
    <section class="creator-dialog" role="dialog" aria-modal="true" aria-labelledby="creator-title">
      <header class="creator-header">
        <div><p class="eyebrow">PRIVATE · ON THIS DEVICE</p><h2 id="creator-title">Content Creator Studio</h2><p class="muted">Plan, review and prepare content without uploading your document.</p></div>
        <button class="close" type="button" aria-label="Close Content Creator Studio" @click="$emit('close')">×</button>
      </header>
      <nav class="creator-tabs" aria-label="Creator tools">
        <button v-for="tab in tabs" :key="tab.id" type="button" :aria-pressed="active === tab.id" :class="{ active: active === tab.id }" @click="active = tab.id">{{ tab.label }}</button>
      </nav>

      <div v-if="active === 'goals'" class="creator-content">
        <h3>Writing goals</h3>
        <p class="muted">Set a document target and keep a simple daily writing log in this browser.</p>
        <div class="creator-grid">
          <label>Document word target<input v-model.number="documentGoal" type="number" min="1" max="1000000" /></label>
          <label>Daily word target<input v-model.number="dailyGoal" type="number" min="1" max="100000" /></label>
        </div>
        <div class="progress-card">
          <div class="progress-heading"><strong>Current document</strong><span>{{ number(wordCount) }} / {{ number(documentGoal) }} words</span></div>
          <div class="progress-track"><span :style="{ width: progress(wordCount, documentGoal) + '%' }"></span></div>
          <p class="muted">{{ Math.max(0, documentGoal - wordCount).toLocaleString() }} words remaining · {{ progress(wordCount, documentGoal) }}% complete</p>
        </div>
        <div class="creator-grid">
          <label>Words written today<input v-model.number="todayWords" type="number" min="0" max="100000" /></label>
          <label>Today's date<input :value="today" type="text" readonly /></label>
        </div>
        <div class="progress-card">
          <div class="progress-heading"><strong>Daily progress</strong><span>{{ number(todayWords) }} / {{ number(dailyGoal) }} words</span></div>
          <div class="progress-track"><span :style="{ width: progress(todayWords, dailyGoal) + '%' }"></span></div>
          <p class="muted">Log your actual words written today; this is a manual log, not an estimate from document edits.</p>
        </div>
        <div class="actions"><button type="button" class="secondary" @click="saveGoals">Save goals and today's log</button><span class="feedback" role="status">{{ goalFeedback }}</span></div>
      </div>

      <div v-else-if="active === 'plan'" class="creator-content">
        <h3>Content brief and outline</h3>
        <p class="muted">Shape the idea before drafting. Your brief is saved locally and can be copied into your document.</p>
        <label>Working title<input v-model="plan.title" placeholder="e.g. A practical guide to…" /></label>
        <div class="creator-grid">
          <label>Target audience<input v-model="plan.audience" placeholder="Who is this for?" /></label>
          <label>Content format<select v-model="plan.format"><option>Blog article</option><option>Feature article</option><option>Newsletter</option><option>Guide</option><option>Opinion piece</option><option>Social post</option><option>Other</option></select></label>
        </div>
        <label>Purpose<textarea v-model="plan.purpose" rows="2" placeholder="What should readers understand or do?" /></label>
        <label>Key message<textarea v-model="plan.message" rows="2" placeholder="The one thing readers should remember…" /></label>
        <label>Outline<textarea v-model="plan.outline" rows="7" placeholder="Introduction&#10;1. First main point&#10;2. Second main point&#10;Conclusion" /></label>
        <div class="actions"><button type="button" class="secondary" @click="savePlan">Save brief</button><button type="button" @click="copy(planText)">Copy brief and outline</button><span class="feedback" role="status">{{ planFeedback || copyFeedback }}</span></div>
      </div>

      <div v-else-if="active === 'seo'" class="creator-content">
        <h3>SEO preparation checklist</h3>
        <p class="muted">Editorial checks for discoverability. They do not predict rankings or replace search research.</p>
        <div class="creator-grid">
          <label>SEO title<input v-model="seo.title" maxlength="100" placeholder="A clear, descriptive title" /><small>{{ seo.title.length }} characters</small></label>
          <label>Focus keyword<input v-model="seo.keyword" placeholder="Main search phrase" /></label>
        </div>
        <label>Meta description<textarea v-model="seo.description" rows="3" maxlength="300" placeholder="Summarize the page for a potential reader…" /><small>{{ seo.description.length }} characters</small></label>
        <div class="seo-results">
          <div class="progress-heading"><strong>Document checks</strong><span>{{ number(wordsIn(text)) }} words · {{ seoResult.occurrences }} keyword mentions</span></div>
          <ul><li v-for="check in seoResult.checks" :key="check.label"><span :class="check.ok ? 'check-ok' : 'check-warn'">{{ check.ok ? '✓' : '○' }}</span><div><strong>{{ check.label }}</strong><small>{{ check.detail }}</small></div></li></ul>
        </div>
        <p class="muted">{{ seoResult.note }}</p>
      </div>

      <div v-else-if="active === 'quality'" class="creator-content">
        <h3>Editorial quality check</h3>
        <p class="muted">A deterministic first-pass review of the open document. Suggestions never rewrite your text.</p>
        <div class="quality-summary"><strong>{{ editorial.words.toLocaleString() }} words</strong><span>{{ editorial.paragraphs }} paragraphs · {{ editorial.sentences }} sentences</span><span>{{ editorial.issues.length }} suggestions</span></div>
        <ul v-if="editorial.issues.length" class="issue-list"><li v-for="(issue, index) in editorial.issues" :key="index"><strong>{{ issue.type }}</strong><p>{{ issue.detail }}</p><small v-if="issue.excerpt">{{ issue.excerpt }}</small></li></ul>
        <p v-else class="empty-state">No issues found by these checks. A human editorial review is still recommended.</p>
        <p class="muted">Checks include repeated adjacent words, long sentences, long paragraphs and frequently repeated words. Names, deliberate repetition and specialist terminology may be flagged.</p>
      </div>

      <div v-else class="creator-content">
        <h3>Repurpose this content</h3>
        <p class="muted">Create editable starting drafts from the current document. Review every output before publishing.</p>
        <label>Post title<input v-model="repurposeTitle" :placeholder="title || 'Give this content a title'" /></label>
        <div class="repurpose-list">
          <article v-for="item in repurposeItems" :key="item.id">
            <div class="progress-heading"><strong>{{ item.label }}</strong><button type="button" class="text-button" @click="copy(item.text)">Copy</button></div>
            <textarea :value="item.text" rows="4" readonly :aria-label="item.label + ' draft'" />
            <small v-if="item.id === 'x'">{{ item.text.length }} / 280 characters</small>
          </article>
        </div>
        <p class="muted">These are draft transformations, not platform integrations. Verify facts, links, tone and current platform limits before posting.</p>
      </div>

      <footer class="creator-footer"><span>Intact keeps this workspace and its saved notes in this browser.</span><button type="button" @click="$emit('close')">Close</button></footer>
    </section>
  </div>
</template>

<script setup>
import { computed, onMounted, onBeforeUnmount, reactive, ref } from 'vue';
import { editorialChecks, repurpose, seoChecks, wordsIn } from '../editor/creator.js';

const props = defineProps({
  text: { type: String, default: '' },
  wordCount: { type: Number, default: 0 },
  title: { type: String, default: '' },
});
const emit = defineEmits(['close']);

const tabs = [
  { id: 'goals', label: 'Writing goals' },
  { id: 'plan', label: 'Brief & outline' },
  { id: 'seo', label: 'SEO' },
  { id: 'quality', label: 'Quality check' },
  { id: 'repurpose', label: 'Repurpose' },
];
const active = ref('goals');
const today = new Date().toLocaleDateString('en-CA');
const number = (value) => (Number(value) || 0).toLocaleString();
const progress = (value, goal) => Math.min(100, Math.round((Math.max(0, Number(value) || 0) / Math.max(1, Number(goal) || 1)) * 100));
const read = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } };
const documentGoal = ref(read('intact.creator.documentGoal', 1200));
const dailyGoal = ref(read('intact.creator.dailyGoal', 500));
const dailyLog = read('intact.creator.dailyLog', {});
const todayWords = ref(Number(dailyLog[today] || 0));
const goalFeedback = ref('');
const plan = reactive({ title: '', audience: '', format: 'Blog article', purpose: '', message: '', outline: '' });
Object.assign(plan, read('intact.creator.plan', {}));
const planFeedback = ref('');
const copyFeedback = ref('');
const seo = reactive({ title: '', keyword: '', description: '' });
const editorial = computed(() => editorialChecks(props.text));
const seoResult = computed(() => seoChecks({ ...seo, text: props.text }));
const repurposeTitle = ref(props.title);
const repurposeItems = computed(() => {
  const drafts = repurpose(props.text, repurposeTitle.value);
  return [
    { id: 'linkedin', label: 'LinkedIn post', text: drafts.linkedin },
    { id: 'x', label: 'X post', text: drafts.x },
    { id: 'instagram', label: 'Instagram caption', text: drafts.instagram },
    { id: 'newsletter', label: 'Newsletter introduction', text: drafts.newsletter },
  ];
});
const planText = computed(() => [
  plan.title && `# ${plan.title}`, plan.audience && `Audience: ${plan.audience}`,
  plan.format && `Format: ${plan.format}`, plan.purpose && `Purpose: ${plan.purpose}`,
  plan.message && `Key message: ${plan.message}`, plan.outline && `## Outline\\n${plan.outline}`,
].filter(Boolean).join('\\n\\n'));
function saveGoals() {
  try {
    const log = read('intact.creator.dailyLog', {});
    log[today] = Math.max(0, Number(todayWords.value) || 0);
    localStorage.setItem('intact.creator.documentGoal', String(Math.max(1, Number(documentGoal.value) || 1)));
    localStorage.setItem('intact.creator.dailyGoal', String(Math.max(1, Number(dailyGoal.value) || 1)));
    localStorage.setItem('intact.creator.dailyLog', JSON.stringify(log));
    goalFeedback.value = 'Saved on this device.';
  } catch { goalFeedback.value = 'Could not save: browser storage is unavailable.'; }
}
function savePlan() {
  try { localStorage.setItem('intact.creator.plan', JSON.stringify(plan)); planFeedback.value = 'Brief saved on this device.'; }
  catch { planFeedback.value = 'Could not save: browser storage is unavailable.'; }
}
async function copy(value) {
  try { await navigator.clipboard.writeText(String(value)); copyFeedback.value = 'Copied to clipboard.'; }
  catch { copyFeedback.value = 'Clipboard access unavailable; select and copy the text manually.'; }
}
function onKeydown(event) { if (event.key === 'Escape') emit('close'); }
onMounted(() => window.addEventListener('keydown', onKeydown));
onBeforeUnmount(() => window.removeEventListener('keydown', onKeydown));
</script>

<style scoped>
.creator-backdrop { position: fixed; inset: 0; z-index: 1200; display: grid; place-items: center; padding: 18px; background: rgb(15 22 28 / 52%); }
.creator-dialog { width: min(820px, 100%); max-height: min(90vh, 900px); overflow: auto; background: var(--surface); color: var(--text); border: 1px solid var(--border); border-radius: 12px; box-shadow: 0 18px 60px rgb(0 0 0 / 28%); }
.creator-header { display: flex; justify-content: space-between; gap: 16px; padding: 22px 24px 14px; }
.creator-header h2 { margin: 2px 0 5px; font-size: 1.35rem; }
.eyebrow { margin: 0; color: var(--accent); font-size: .68rem; font-weight: 700; letter-spacing: .12em; }
.muted, small { color: var(--muted); line-height: 1.45; }
.creator-header .muted { margin: 0; }
.close { width: 34px; height: 34px; border: 0; border-radius: 6px; background: transparent; color: inherit; font-size: 25px; cursor: pointer; }
.close:hover, .text-button:hover { background: var(--hover); }
.creator-tabs { display: flex; gap: 4px; overflow-x: auto; padding: 0 20px 12px; border-bottom: 1px solid var(--border); }
.creator-tabs button { flex: none; padding: 8px 11px; border: 1px solid transparent; border-radius: 6px; background: transparent; color: inherit; font: inherit; cursor: pointer; }
.creator-tabs button.active { color: var(--accent); border-color: var(--border); background: var(--active); font-weight: 650; }
.creator-content { padding: 18px 24px 22px; }
.creator-content h3 { margin: 0 0 4px; font-size: 1.05rem; }
.creator-content > .muted:first-of-type { margin: 0 0 18px; }
.creator-content label { display: grid; gap: 5px; margin: 12px 0; font-size: .88rem; font-weight: 600; }
.creator-content input, .creator-content textarea, .creator-content select { width: 100%; min-width: 0; padding: 9px 10px; border: 1px solid var(--border); border-radius: 6px; background: var(--surface); color: inherit; font: inherit; font-weight: 400; }
.creator-content textarea { resize: vertical; line-height: 1.5; }
.creator-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
.progress-card, .seo-results { margin: 14px 0; padding: 14px; border: 1px solid var(--border); border-radius: 8px; }
.progress-heading { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
.progress-heading span { color: var(--muted); font-size: .85rem; font-variant-numeric: tabular-nums; }
.progress-track { height: 8px; overflow: hidden; margin: 10px 0 5px; border-radius: 8px; background: var(--hover); }
.progress-track span { display: block; height: 100%; border-radius: inherit; background: var(--accent); }
.progress-card p { margin: 5px 0 0; font-size: .82rem; }
.actions { display: flex; align-items: center; flex-wrap: wrap; gap: 9px; margin-top: 14px; }
.actions button, .creator-footer button { padding: 8px 13px; border: 1px solid var(--border); border-radius: 6px; background: var(--active); color: var(--text); font: inherit; cursor: pointer; }
.actions .secondary { background: var(--brand); color: white; border-color: var(--brand); }
.feedback { color: var(--muted); font-size: .82rem; }
.seo-results ul, .issue-list { display: grid; gap: 10px; margin: 12px 0 0; padding: 0; list-style: none; }
.seo-results li { display: flex; align-items: flex-start; gap: 10px; }
.seo-results li div { display: grid; gap: 2px; }
.seo-results li small { font-size: .8rem; }
.check-ok { color: #237a43; font-weight: 800; }
.check-warn { color: var(--muted); font-weight: 800; }
.quality-summary { display: flex; flex-wrap: wrap; gap: 12px; margin: 14px 0; padding: 12px; background: var(--hover); border-radius: 7px; }
.quality-summary span { color: var(--muted); }
.issue-list li { padding: 11px 12px; border: 1px solid var(--border); border-radius: 7px; }
.issue-list p { margin: 4px 0; }
.issue-list small { display: block; overflow-wrap: anywhere; }
.empty-state { padding: 16px; border: 1px solid var(--border); border-radius: 7px; }
.repurpose-list { display: grid; gap: 14px; margin-top: 14px; }
.repurpose-list article { display: grid; gap: 7px; }
.text-button { padding: 5px 9px; border: 0; border-radius: 5px; background: transparent; color: var(--accent); font: inherit; cursor: pointer; }
.creator-footer { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 12px 24px; border-top: 1px solid var(--border); color: var(--muted); font-size: .8rem; }
@media (max-width: 600px) { .creator-backdrop { padding: 0; place-items: stretch; } .creator-dialog { max-height: 100vh; border-radius: 0; } .creator-header, .creator-content { padding-left: 16px; padding-right: 16px; } .creator-tabs { padding-left: 12px; padding-right: 12px; } .creator-grid { grid-template-columns: 1fr; gap: 0; } .creator-footer { padding: 12px 16px; } }
</style>

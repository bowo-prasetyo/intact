<!--
VISION.md is the only door for what Intact does. A pull request either serves an entry already in
its Direction, or changes VISION.md itself, in this pull request, where the change can be reviewed.
A feature that does neither isn't ready for review, however good it is: start with the VISION entry.
Delete the comments as you fill this in; leave a box unticked, with a reason, rather than delete it.
-->

## What this changes

<!-- What someone using Intact sees or gets, in a sentence or two. -->

## Why: the VISION entry it serves

<!--
Quote the line in VISION.md's Direction this delivers, or the job it fixes ("Fix" or "Write").
For a new kind of user or a new kind of feature, add the entry to VISION.md in this pull request
and quote it here. A bug fix or test-only change can say so and skip to the checklist's last part.
-->

> 

## Checklist

**Vision**
- [ ] It serves the VISION entry quoted above, or adds or changes that entry in VISION.md in this pull request
- [ ] It goes into the document, or works on it: not a workspace beside it (a brief, goals, a log, drafts for other platforms)
- [ ] It doesn't touch a non-goal; or it says here which one it comes near, and the very good reason VISION asks for
- [ ] It doesn't repeat a tool Intact already has (Word count, the accessibility check, Copy for the web, Fill in…); it extends that tool instead
- [ ] `docs/product/`, if used, describes a feature VISION already accepts; it isn't where a feature is proposed

**Promises**
- [ ] Private: no new network requests or third-party code; nothing kept in this browser beyond a convenience the user can see and clear, or that expires (as drafts do)
- [ ] Faithful: what it can change has round-trip cases (`tests/roundtrip.spec.js`, or a fixture in `tests/fixtures/`)
- [ ] Instant: the start page loads no more than before, or the reason is given here
- [ ] Honest: what it can't edit or keep is in the README's Limitations; nothing in a document it makes names Intact or its menus

**A document made here is a Word document** (only for something that writes to the file)
- [ ] Written as Word writes it: properties in schema order, parts related and typed, checked by a test (as `tests/starters.spec.js` does)
- [ ] Paper, units and language follow where the browser is, as the other starters and dialogs do
- [ ] Opened in Microsoft Word, by hand, before release: by <!-- who -->, on <!-- date -->, or still to do

**Where it lives**
- [ ] A new tool goes where Word users look for it (the ribbon's tab, a dialog Word has, the right-click menu); the header is kept for the file and its buttons
- [ ] It works on a phone, or the reason it doesn't is given here

**Words**
- [ ] Every new interface string goes through `t()` and has its Indonesian in `src/locales/id.js`
- [ ] README's Features, the About page (`src/views/AboutView.vue`) and VISION's Direction (✔ where it's done) say what changed, in the same words

## How it was tested

<!--
The specs added or changed, and what ran: `npm test` (Chromium and Firefox), or which project and
specs, and what couldn't run where you were. What was checked by hand, and in which browsers.
-->

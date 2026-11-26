# Content Creator Essentials

## Purpose

Content Creator Studio adds a small, private editorial workspace to Intact. It supports the work surrounding a DOCX draft without changing the product into a cloud publishing platform.

All checks and transformations run in the browser. The open document text is read for the current session only; saved goals, the daily word log and the content brief are stored in this browser's local storage. Nothing is uploaded and none of the tools silently rewrites the document.

## P0 scope

### 1. Writing goals

- Show current document word count against a configurable document target.
- Set a daily word target and record the writer's actual daily output.
- Persist goals and the daily log locally.
- Label the daily entry as manual: document word-count differences cannot reliably represent words written that day.

### 2. Content brief and outline

- Capture a working title, audience, format, purpose, key message and outline.
- Save the brief locally and copy it into the clipboard for use in a document.
- Do not automatically modify the open DOCX.

### 3. SEO preparation

- Check editorial title and meta-description length ranges.
- Check whether the focus keyword appears in the title and opening 100 words.
- Count keyword occurrences and inspect actual document headings, hyperlinks and image alt text.
- Clearly describe these as heuristics, not ranking predictions. The current text-only interface does not claim to audit image alt text, live links, canonical tags, structured data or search-engine indexing.

### 4. Editorial quality checker

- Flag adjacent repeated words, repeated three-word phrases, frequently repeated words, sentences over 30 words and paragraphs over 150 words.
- Present suggestions only; never rewrite text or imply that a heuristic is a grammar authority.
- Explain that deliberate repetition, names and specialist terms can be false positives.

### 5. Content repurposing

- Produce starting drafts for LinkedIn, X, Instagram and newsletter introductions.
- Keep outputs editable and copyable.
- Bound the X draft to 280 characters as a conservative editorial check; users must verify current platform rules before publishing.
- Do not publish, connect to social accounts or invent facts. Authors review every draft.

## Acceptance criteria

- Creator Studio opens from the editor and can be dismissed with Escape, the close button or clicking the backdrop.
- The UI works on narrow screens and exposes named controls and status feedback.
- Goals and briefs survive a reload in the same browser.
- Daily output is an explicit user-entered value and is never inferred from a single document's word count.
- SEO checks disclose their heuristic nature.
- Editorial findings do not mutate the DOCX.
- Repurposed drafts are editable before copying.
- No network requests, accounts, server-side processing or third-party AI services are introduced.

## Deliberate limitations

P0 is a local editorial companion, not an SEO crawler, grammar engine, editorial calendar, social publishing integration or multi-document project manager. Future work should preserve Intact's privacy and DOCX-fidelity promises.

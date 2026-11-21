// Saved comments, to reuse when marking: feedback written often ("Cite a primary source here")
// saved once and put into a new comment with a click or Alt+1 to Alt+9. Kept in this browser, as
// AutoCorrect entries are, never in a document or sent anywhere; taken to another browser or device
// as a CSV file (one comment a row, under a "Comment" header), which a spreadsheet opens too.

const KEY = 'intact.commentBank';
export const MAX_SAVED = 500;
export const MAX_LENGTH = 2000;
export const MAX_FILE = 1024 * 1024;

const clean = (s) => String(s ?? '').replace(/\r\n?/g, '\n').trim();

/** The saved comments, as kept in this browser (none where it keeps nothing). */
export function loadCommentBank() {
  try {
    const v = JSON.parse(localStorage.getItem(KEY) || '[]');
    return Array.isArray(v) ? v.filter((s) => typeof s === 'string' && s.trim()).slice(0, MAX_SAVED) : [];
  } catch {
    return [];
  }
}

/** Keeps the saved comments in this browser; false where it refuses storage (this session only, then). */
export function saveCommentBank(list) {
  try {
    localStorage.setItem(KEY, JSON.stringify(list));
    return true;
  } catch {
    return false;
  }
}

/**
 * `incoming` comments added to `list`, after it, leaving out empty ones, those already there (the
 * same text) and any past the limit: { list, added, skipped }.
 */
export function addToBank(list, incoming) {
  const out = [...list];
  const seen = new Set(out.map((s) => s.toLowerCase()));
  let added = 0;
  let skipped = 0;
  for (const raw of incoming) {
    const text = clean(raw).slice(0, MAX_LENGTH);
    if (!text || seen.has(text.toLowerCase()) || out.length >= MAX_SAVED) {
      skipped++;
      continue;
    }
    seen.add(text.toLowerCase());
    out.push(text);
    added++;
  }
  return { list: out, added, skipped };
}

/**
 * The saved comments as a CSV file's text: a header row "Comment", then a comment a row (quoted
 * where it has a comma, quote or line break), in UTF-8 with a byte order mark, as Excel's "CSV
 * UTF-8" saves it.
 */
export function bankToCsv(list) {
  const field = (s) => (/[",;\t\r\n]|^\s|\s$/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  return `﻿${['Comment', ...list.map((s) => field(s.replace(/\n/g, '\r\n')))].join('\r\n')}\r\n`;
}

/**
 * The comments in a CSV file's text (see bankToCsv; also a column saved from a spreadsheet, with or
 * without its header row): the first field of each row. `parseCsv` is mail merge's (docx/merge.js).
 */
export function bankFromCsv(text, parseCsv) {
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r\n|\n|\r/, 1)[0];
  if (!/^\s*"?\s*(comments?|komentar|feedback)\s*"?\s*([,;\t]|$)/i.test(first)) text = `Comment\n${text}`;
  return parseCsv(text).rows.map((r) => clean(r[0])).filter(Boolean);
}

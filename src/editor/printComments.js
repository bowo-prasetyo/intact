// Comments in a printout or PDF, for work marked here and returned as a PDF (as a learning platform
// often asks for): the comments listed after the document, each with its number as the text shows
// it, its author and date, the words it is on and its replies; those in a header, footer or note
// say where they are. Made just before printing, beside the page (never in it, so neither the
// document nor Undo sees it), and taken away after.

const MAX_QUOTE = 90;

/** `text` shortened to about `max` characters, at a word, with … after. */
function shorten(text, max = MAX_QUOTE) {
  const t = text.replace(/\s+/g, ' ').trim();
  if (t.length <= max) return t;
  const cut = t.slice(0, max);
  return `${cut.slice(0, Math.max(cut.lastIndexOf(' '), max * 0.6))}…`;
}

const formatDate = (iso) => {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
};

function el(tag, className, text) {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text != null) e.textContent = text;
  return e;
}

/** A comment's author and date, and its paragraphs, in `into`. */
function addComment(into, c, lead) {
  const who = el('p', 'who');
  if (lead) who.append(lead);
  who.append(el('strong', null, c.author || 'Author'));
  const date = c.date ? formatDate(c.date) : '';
  if (date) who.append(`, ${date}`);
  into.append(who);
  for (const text of c.paragraphs?.length ? c.paragraphs : ['(no text)']) into.append(el('p', 'text', text));
}

/**
 * The list of comments to print after the document, as a section: `threads` as the comments panel
 * has them ([{ id, n, place, author, date, paragraphs, done, replies }]), `quoted(id)` the text a
 * comment is on (or ''), `title` its heading. Comments in the text come first, by their numbers.
 */
export function printCommentsSection(threads, quoted, title = 'Comments') {
  const section = el('section', 'print-comments');
  section.setAttribute('aria-hidden', 'true'); // printed only: not part of the page on screen
  section.append(el('h2', null, title));
  const list = el('ol');
  const ordered = [...threads].sort((a, b) => (a.n ?? Infinity) - (b.n ?? Infinity));
  for (const t of ordered) {
    const item = el('li');
    const head = el('div', 'comment');
    const lead = el('span', 'n', t.n ? `${t.n}` : '–');
    addComment(head, t, lead);
    const on = quoted(t.id);
    const where = [on ? `On “${shorten(on)}”` : t.place ? `In ${t.place.replace(/^./, (c) => c.toLowerCase())}` : '', t.done ? 'resolved' : ''].filter(Boolean).join(' · ');
    if (where) head.querySelector('.who').after(el('p', 'on', where));
    item.append(head);
    for (const r of t.replies || []) {
      const reply = el('div', 'reply');
      addComment(reply, r, 'Reply from ');
      item.append(reply);
    }
    list.append(item);
  }
  section.append(list);
  return section;
}

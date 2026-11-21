// Copy for the web: the document (or the selection) as a website's editor (a CMS), an e-mail or a
// Markdown editor takes it, without Word's formatting: headings, paragraphs, bold, italic,
// strikethrough, superscript and subscript, links, bulleted and numbered lists, line breaks and
// tables. The text is the text as it will read once its tracked changes are accepted: deleted text
// and hidden text (a test's answers) are left out. Pictures, text boxes, equations and notes are left
// out too, and counted, so the editor can say so. Read from the page as shown; nothing is sent anywhere.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';
// What isn't copied: deleted text, hidden text, the editor's marks, and kept content out of view.
const SKIP = '[data-tc="del"], .raw-inline.deleted, [data-hidden], .raw-hidden, .comment-ref, .note-ref, .picture, .textbox, .docx-hf-band, .docx-notes';
const BULLET = /^[•◦▪·•◦▪\-–—*o§]$/u;

/**
 * The content of `page` (within `range`, if given): { blocks, left: { pictures, textBoxes, notes,
 * equations } }. A block is { type: 'heading', level, inlines } | { type: 'paragraph', inlines } |
 * { type: 'item', ordered, level, marker, inlines } | { type: 'table', rows: [[cell inlines]] };
 * an inline is { text, b, i, s, sup, sub, href } or { br: true }. `levelOf(el)` gives a list
 * paragraph's level (0 for the first).
 */
export function webContent(page, { range = null, levelOf = () => 0 } = {}) {
  const left = { pictures: 0, textBoxes: 0, notes: 0, equations: 0 };
  const within = (node) => !range || range.intersectsNode(node);
  const count = (root) => {
    for (const el of root.querySelectorAll('.picture, .textbox, .note-ref, .raw-inline')) {
      if (!within(el) || el.closest('.docx-hf-band, .docx-notes')) continue;
      if (el.matches('.picture')) left.pictures++;
      else if (el.matches('.textbox')) left.textBoxes++;
      else if (el.matches('.note-ref')) left.notes++;
      else if (el.querySelector('math')) left.equations++;
    }
  };
  count(page);
  // Pictures inside text boxes are the text box's.
  left.pictures -= [...page.querySelectorAll('.textbox .picture')].filter(within).length;

  const blocks = [];
  // In the order of the text: paragraphs, and tables (a table's paragraphs are copied with it, and a
  // table in a cell is copied as its text).
  for (const el of page.querySelectorAll(`${BLOCKS}, table`)) {
    if (!within(el) || el.closest(SKIP) || el.parentElement.closest('table')) continue;
    if (el.matches('table')) {
      const rows = [...el.rows].filter((tr) => tr.closest('table') === el && within(tr)).map((tr) => [...tr.cells].map((td) => {
        const out = [];
        for (const b of td.querySelectorAll(BLOCKS)) {
          if (b.closest(SKIP)) continue;
          const inl = inlines(b, range);
          if (!inl.length) continue;
          if (out.length) out.push({ br: true });
          out.push(...inl);
        }
        return out;
      }));
      if (rows.some((r) => r.some((c) => c.length))) blocks.push({ type: 'table', rows });
      continue;
    }
    const inl = inlines(el, range);
    if (!inl.some((x) => x.text?.trim())) continue;
    if (/^H\d$/.test(el.tagName)) blocks.push({ type: 'heading', level: Number(el.tagName[1]), inlines: inl });
    else if (el.classList.contains('list') && el.dataset.marker) {
      const marker = el.dataset.marker.trim();
      blocks.push({ type: 'item', ordered: !BULLET.test(marker), level: levelOf(el), marker, inlines: inl });
    } else blocks.push({ type: 'paragraph', inlines: inl });
  }
  return { blocks, left };
}

/** The inlines of block `el`: its text with the formatting a web page keeps, within `range`. */
function inlines(el, range) {
  const out = [];
  const heading = /^H\d$/.test(el.tagName);
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => {
      if (n.nodeType === 1) {
        if (n.matches(SKIP) || (n !== el && n.matches(`${BLOCKS}, table`))) return NodeFilter.FILTER_REJECT;
        return n.tagName === 'BR' ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
      }
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 1) {
      if (!range || range.intersectsNode(n)) out.push({ br: true });
      continue;
    }
    if (range && !range.intersectsNode(n)) continue;
    let text = n.data;
    if (range) {
      const end = n === range.endContainer ? range.endOffset : text.length;
      const start = n === range.startContainer ? range.startOffset : 0;
      text = text.slice(start, end);
    }
    text = text.replace(/[\t ]/g, (c) => (c === '\t' ? ' ' : ' '));
    if (!text) continue;
    const parent = n.parentElement;
    const style = getComputedStyle(parent);
    const link = parent.closest('a[href]');
    const href = link && !link.getAttribute('href').startsWith('#') ? link.href : null;
    const deco = style.textDecorationLine || '';
    const run = {
      text,
      b: !heading && Number(style.fontWeight) >= 600,
      i: style.fontStyle === 'italic' || style.fontStyle.startsWith('oblique'),
      s: deco.includes('line-through') && !parent.closest('[data-tc]'),
      sup: !!parent.closest('sup') || style.verticalAlign === 'super',
      sub: !!parent.closest('sub') || style.verticalAlign === 'sub',
      href,
    };
    const last = out[out.length - 1];
    if (last && !last.br && ['b', 'i', 's', 'sup', 'sub', 'href'].every((k) => last[k] === run[k])) last.text += text;
    else out.push(run);
  }
  // Spaces at the ends of a block aren't the web's.
  while (out.length && out[0].text != null && !out[0].text.trim()) out.shift();
  if (out[0]?.text) out[0].text = out[0].text.replace(/^\s+/, '');
  const last = out[out.length - 1];
  if (last?.text) last.text = last.text.replace(/\s+$/, '');
  return out.filter((x) => x.br || x.text);
}

const escapeHtml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The blocks as clean HTML: no classes, styles or Word markup. */
export function toHtml(blocks) {
  const inline = (list) => list.map((x) => {
    if (x.br) return '<br>';
    let h = escapeHtml(x.text);
    if (x.sup) h = `<sup>${h}</sup>`;
    if (x.sub) h = `<sub>${h}</sub>`;
    if (x.s) h = `<s>${h}</s>`;
    if (x.i) h = `<em>${h}</em>`;
    if (x.b) h = `<strong>${h}</strong>`;
    if (x.href) h = `<a href="${escapeHtml(x.href)}">${h}</a>`;
    return h;
  }).join('');
  const out = [];
  const open = []; // the lists open: [{ tag, level }]
  const closeTo = (level) => {
    while (open.length && open[open.length - 1].level >= level) out.push(`</li></${open.pop().tag}>`);
  };
  for (const b of blocks) {
    if (b.type !== 'item') {
      closeTo(0);
      if (b.type === 'heading') out.push(`<h${b.level}>${inline(b.inlines)}</h${b.level}>`);
      else if (b.type === 'paragraph') out.push(`<p>${inline(b.inlines)}</p>`);
      else out.push(`<table>${b.rows.map((r) => `<tr>${r.map((c) => `<td>${inline(c)}</td>`).join('')}</tr>`).join('')}</table>`);
      continue;
    }
    const tag = b.ordered ? 'ol' : 'ul';
    const top = open[open.length - 1];
    if (top && top.level > b.level) closeTo(b.level + 1);
    const cur = open[open.length - 1];
    if (cur && cur.level === b.level && cur.tag !== tag) closeTo(b.level);
    const now = open[open.length - 1];
    if (!now || now.level < b.level) {
      out.push(`<${tag}><li>`);
      open.push({ tag, level: b.level });
    } else out.push('</li><li>');
    out.push(inline(b.inlines));
  }
  closeTo(0);
  return out.join('');
}

/** The blocks as plain text: paragraphs apart by a blank line, list items with their markers, cells by tabs. */
export function toText(blocks) {
  const inline = (list) => list.map((x) => (x.br ? '\n' : x.text)).join('');
  const out = [];
  blocks.forEach((b, k) => {
    const prev = blocks[k - 1];
    if (k && !(b.type === 'item' && prev?.type === 'item')) out.push('');
    if (b.type === 'item') out.push(`${'  '.repeat(b.level)}${b.ordered ? b.marker : '•'} ${inline(b.inlines)}`);
    else if (b.type === 'table') out.push(b.rows.map((r) => r.map((c) => inline(c).replace(/\n/g, ' ')).join('\t')).join('\n'));
    else out.push(inline(b.inlines));
  });
  return `${out.join('\n')}\n`;
}

const escapeMd = (s) => s.replace(/([\\`*_[\]<>|])/g, '\\$1');

/** The blocks as Markdown (CommonMark, with GitHub's tables and ~~strikethrough~~). */
export function toMarkdown(blocks) {
  const inline = (list) => list.map((x) => {
    if (x.br) return '  \n';
    const lead = x.text.match(/^\s*/)[0];
    const trail = x.text.match(/\s*$/)[0];
    let m = escapeMd(x.text.trim());
    if (!m) return x.text;
    if (x.sup) m = `<sup>${m}</sup>`;
    if (x.sub) m = `<sub>${m}</sub>`;
    if (x.s) m = `~~${m}~~`;
    if (x.i) m = `*${m}*`;
    if (x.b) m = `**${m}**`;
    if (x.href) m = `[${m}](${x.href.replace(/[()]/g, (c) => encodeURIComponent(c))})`;
    return lead + m + trail;
  }).join('');
  const out = [];
  const numbers = []; // per level, the next number of an ordered list
  blocks.forEach((b, k) => {
    const prev = blocks[k - 1];
    if (k && !(b.type === 'item' && prev?.type === 'item')) out.push('');
    if (b.type !== 'item') numbers.length = 0;
    if (b.type === 'heading') out.push(`${'#'.repeat(b.level)} ${inline(b.inlines)}`);
    else if (b.type === 'paragraph') out.push(inline(b.inlines).replace(/^(\s*)(#|>|[-+]\s|\d+[.)]\s)/, '$1\\$2'));
    else if (b.type === 'item') {
      numbers.length = b.level + 1;
      const n = (numbers[b.level] = (numbers[b.level] || 0) + 1);
      out.push(`${'   '.repeat(b.level)}${b.ordered ? `${n}.` : '-'} ${inline(b.inlines)}`);
    } else {
      const width = Math.max(...b.rows.map((r) => r.length));
      const row = (r) => `| ${Array.from({ length: width }, (_, i) => inline(r[i] || []).replace(/ {2}\n/g, '<br>')).join(' | ')} |`;
      out.push(row(b.rows[0]), `|${' --- |'.repeat(width)}`, ...b.rows.slice(1).map(row));
    }
  });
  return `${out.join('\n')}\n`;
}

/**
 * Puts `formats` ({ 'text/html': …, 'text/plain': … }) on the clipboard, through a copy event the
 * page answers itself (so it works in every browser, from a click). True if the browser copied.
 */
export function copyFormats(formats) {
  let copied = false;
  const onCopy = (e) => {
    for (const [type, data] of Object.entries(formats)) e.clipboardData.setData(type, data);
    e.preventDefault();
    e.stopImmediatePropagation();
    copied = true;
  };
  document.addEventListener('copy', onCopy, { capture: true });
  try {
    document.execCommand('copy');
  } finally {
    document.removeEventListener('copy', onCopy, { capture: true });
  }
  return copied;
}

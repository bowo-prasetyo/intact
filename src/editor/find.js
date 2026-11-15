// Find and replace over the editor DOM. Text is searched per "segment": a run of text nodes in the
// same paragraph and the same editing host, so matches cross formatting ("**John** Smith") but
// never paragraphs, line breaks, fields or preserved content.

const STRUCTURAL = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV', 'TD', 'TR', 'TABLE', 'LI', 'BR', 'IMG']);
const BLOCK = 'p, h1, h2, h3, h4, h5, h6, td, div';

/** Text segments of `root` in document order: { text, nodes: [{ node, start }], editable }. */
export function textSegments(root) {
  const segments = [];
  let current = null;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.nodeType === 1 && n.classList.contains('raw-hidden') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 1) {
      if (STRUCTURAL.has(n.tagName) || n.matches('.sdt, .sdt-content, .raw-inline')) current = null;
      continue;
    }
    if (!n.data) continue;
    const parent = n.parentElement;
    const block = parent.closest(BLOCK);
    const host = parent.closest('[contenteditable]');
    if (!current || current.block !== block || current.host !== host) {
      current = { block, host, editable: host?.getAttribute('contenteditable') === 'true', text: '', nodes: [] };
      segments.push(current);
    }
    current.nodes.push({ node: n, start: current.text.length });
    current.text += n.data;
  }
  return segments;
}

function pattern(query, { matchCase, wholeWord }) {
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const source = wholeWord ? `(?<![\\p{L}\\p{N}_])${escaped}(?![\\p{L}\\p{N}_])` : escaped;
  return new RegExp(source, matchCase ? 'gu' : 'giu');
}

/** All matches of `query` in `roots` (an element or a list, in order): { segment, start, end, editable }. */
export function findMatches(roots, query, options = {}) {
  if (!query) return [];
  const re = pattern(query, options);
  const out = [];
  for (const segment of [roots].flat().flatMap(textSegments)) {
    for (const m of segment.text.matchAll(re)) {
      if (m[0]) out.push({ segment, start: m.index, end: m.index + m[0].length, editable: segment.editable });
    }
  }
  return out;
}

/** The DOM point at text offset `offset` of a segment (`atEnd` prefers the end of a node). */
function point(segment, offset, atEnd) {
  const { nodes } = segment;
  for (let i = 0; i < nodes.length; i++) {
    const { node, start } = nodes[i];
    const end = start + node.data.length;
    if (atEnd ? offset <= end : offset < end || i === nodes.length - 1) return [node, Math.min(offset - start, node.data.length)];
  }
  return [nodes[0].node, 0];
}

export function matchRange(match) {
  const range = document.createRange();
  range.setStart(...point(match.segment, match.start, false));
  range.setEnd(...point(match.segment, match.end, true));
  return range;
}

/**
 * Replaces a match's text. The replacement goes into the first text node of the match, so it takes
 * the formatting of the match's first character, as in Word. Returns the point just after it.
 */
export function replaceMatch(match, text) {
  const { segment, start, end } = match;
  let after = null;
  for (const { node, start: nodeStart } of segment.nodes) {
    const nodeEnd = nodeStart + node.data.length;
    if (nodeEnd <= start || nodeStart >= end) continue;
    const from = Math.max(start, nodeStart) - nodeStart;
    const to = Math.min(end, nodeEnd) - nodeStart;
    if (!after) {
      node.data = node.data.slice(0, from) + text + node.data.slice(to);
      after = { node, offset: from + text.length };
    } else {
      node.data = node.data.slice(0, from) + node.data.slice(to);
    }
  }
  // A paragraph left with no text still needs something to put the caret in.
  const block = segment.block;
  if (block && !block.textContent && !block.querySelector('br, img, .raw-inline:not(.raw-hidden), .sdt')) block.append(document.createElement('br'));
  return after;
}

/** Compares a match's start with a DOM point: negative if the match starts before it. */
export function compareToPoint(match, node, offset) {
  const range = document.createRange();
  range.setStart(...point(match.segment, match.start, false));
  const at = document.createRange();
  at.setStart(node, offset);
  return range.compareBoundaryPoints(Range.START_TO_START, at);
}

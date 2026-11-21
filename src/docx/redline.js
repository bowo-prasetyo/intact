// A redline, as Word's Review › Compare makes one: a new document that is this version, with what
// changed since the other version written into it as tracked changes, so whoever gets it sees the
// changes marked in Word (and can accept or reject them). Made from the two documents' XML, on this
// device. Only the body's text is compared (see editor/compare.js): words put in are marked inserted
// where they are, with their formatting; words taken out are put back, marked deleted, in the
// formatting of the text beside them; paragraphs added and taken out are marked whole, with their
// paragraph marks. Tracked changes already in this version are accepted first, as Word's Compare
// compares the documents as if they were.

import { NS, createW, wChild, ensureOrdered, PPR_ORDER, parseXml, serializeXml } from './xml.js';
import { resolveRevisionsIn, countRevisions } from './resolveAll.js';
import { bodyParagraphs, compareParagraphs, readParagraph } from '../editor/compare.js';

const W = NS.w;
const isW = (n, local) => n?.nodeType === 1 && n.namespaceURI === W && (!local || n.localName === local);

/** The largest w:id used in `xml` (annotations' ids: comments, bookmarks, tracked changes), or 0. */
export function largestId(xml) {
  let max = 0;
  for (const m of xml.matchAll(/\bw:id="(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return max;
}

/**
 * The redline of `currentXml` (this version's document.xml) against `otherXml` (the other's):
 * { xml (the new document.xml), changes (compareParagraphs' result), accepted (how many tracked
 * changes this version had, accepted first) }. `author` and `date` mark the changes; their ids start
 * after `firstId`; `styleIds` (the styles this document defines) lets a paragraph taken out keep
 * its style where this document has it.
 */
export function redline(currentXml, otherXml, { author = 'Author', date = new Date(), firstId = 0, styleIds = new Set() } = {}) {
  const dom = parseXml(currentXml);
  const other = parseXml(otherXml);
  const body = dom.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('This document has no body to compare');
  const accepted = countRevisions(body);
  if (accepted) resolveRevisionsIn(body, true);

  const before = bodyParagraphs(other);
  const after = bodyParagraphs(dom);
  const changes = compareParagraphs(before.map((p) => p.text), after.map((p) => p.text));
  let id = firstId;
  const stamp = date.toISOString().replace(/\.\d+Z$/, 'Z');
  const mark = (local) => createW(dom, local, { id: ++id, author, date: stamp });

  // Paragraphs taken out go before the one now where they were (after the last, at the end), in order.
  let lastAtEnd = null;
  // `before` paragraphs taken out, in order, to find their styles: compareParagraphs lists them in turn.
  const removedTexts = new Map();
  for (const b of before) {
    if (!removedTexts.has(b.text)) removedTexts.set(b.text, []);
    removedTexts.get(b.text).push(b.p);
  }

  for (const change of changes.changes) {
    if (change.kind === 'changed') markChanged(after[change.at].p, change.parts, mark, dom);
    else if (change.kind === 'added') markAdded(after[change.at].p, mark);
    else {
      const text = change.parts[0].text;
      const source = removedTexts.get(text)?.shift();
      const p = deletedParagraph(dom, text, source, styleIds, mark);
      if (change.at < after.length) after[change.at].p.before(p);
      else {
        const anchor = lastAtEnd || after.at(-1)?.p;
        if (anchor) anchor.after(p);
        else body.insertBefore(p, wChild(body, 'sectPr'));
        lastAtEnd = p;
      }
    }
  }
  return { xml: serializeXml(dom), changes, accepted };
}

/** Wraps `runs` in w:ins, runs side by side in one. */
function wrapInserted(runs, mark) {
  let last = null;
  for (const run of runs) {
    if (last && last.nextSibling === run) {
      last.append(run);
      continue;
    }
    const ins = mark('ins');
    run.before(ins);
    ins.append(run);
    last = ins;
  }
}

/** The runs of paragraph `p` itself (not of a text box's paragraphs inside it). */
function ownRuns(p) {
  return [...p.getElementsByTagNameNS(W, 'r')].filter((r) => {
    for (let n = r.parentNode; n; n = n.parentNode) {
      if (isW(n, 'p')) return n === p;
      if (isW(n, 'r')) return false; // a run inside a run's drawing
    }
    return false;
  });
}

/** Marks the paragraph mark of `p` as inserted or deleted (`local`), as Word does in pPr's rPr. */
function markParagraphMark(p, local, mark) {
  let pPr = wChild(p, 'pPr');
  if (!pPr) {
    pPr = createW(p.ownerDocument, 'pPr');
    p.prepend(pPr);
  }
  const rPr = ensureOrdered(pPr, 'rPr', PPR_ORDER);
  rPr.prepend(mark(local));
}

/** A paragraph's properties without its mark's (pPr's rPr), to tell paragraphs formatted alike. */
function paragraphFormat(p) {
  const pPr = wChild(p, 'pPr');
  if (!pPr) return '';
  const copy = pPr.cloneNode(true);
  for (const c of [...copy.childNodes]) if (isW(c, 'rPr')) c.remove();
  return new XMLSerializer().serializeToString(copy);
}

/**
 * Marks an added paragraph: its runs inserted, and a paragraph break inserted. Which break: where
 * the paragraph before is formatted the same, that one's, as Word records Enter pressed at its end
 * and the new paragraph typed (and as LibreOffice reads best); else the new paragraph's own, as Word
 * records a paragraph typed before the next. Rejected, either leaves the text as it was.
 */
function markAdded(p, mark) {
  wrapInserted(ownRuns(p), mark);
  const prev = p.previousElementSibling;
  const prevPPr = isW(prev, 'p') ? wChild(prev, 'pPr') : null;
  const prevRPr = wChild(prevPPr, 'rPr');
  const alike = isW(prev, 'p') && !wChild(prevRPr, 'ins') && !wChild(prevRPr, 'del') && !wChild(prevPPr, 'sectPr') && paragraphFormat(prev) === paragraphFormat(p);
  if (alike) markParagraphMark(prev, 'ins', mark);
  else markParagraphMark(p, 'ins', mark);
}

/** A run of deleted text, in `rPr`'s formatting: w:delText, and w:tab for tabs. */
function deletedRun(dom, text, rPr) {
  const run = createW(dom, 'r');
  if (rPr) {
    const copy = rPr.cloneNode(true);
    for (const c of [...copy.childNodes]) if (isW(c, 'rPrChange')) c.remove();
    run.append(copy);
  }
  text.split('\t').forEach((piece, i) => {
    if (i) run.append(createW(dom, 'tab'));
    if (!piece) return;
    const t = createW(dom, 'delText');
    t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
    t.textContent = piece;
    run.append(t);
  });
  return run;
}

/** A paragraph taken out: its text deleted, and its mark, in its style if this document has it. */
function deletedParagraph(dom, text, source, styleIds, mark) {
  const p = createW(dom, 'p');
  const pPr = createW(dom, 'pPr');
  const style = source && wChild(wChild(source, 'pPr'), 'pStyle')?.getAttributeNS(W, 'val');
  if (style && styleIds.has(style)) pPr.append(createW(dom, 'pStyle', { val: style }));
  p.append(pPr);
  const del = mark('del');
  del.append(deletedRun(dom, text, null));
  p.append(del);
  markParagraphMark(p, 'del', mark);
  return p;
}

/**
 * Splits the runs of `p` so that a run starts or ends at `pos` in its text (as readParagraph counts
 * it): a w:t split in two, and its run in two, each with the run's formatting.
 */
function splitAt(p, pos) {
  let { atoms } = readParagraph(p);
  const inside = atoms.find((a) => a.start < pos && pos < a.end && isW(a.node, 't'));
  if (inside) {
    const k = pos - inside.start;
    const t = inside.node;
    const rest = t.cloneNode(false);
    rest.textContent = t.textContent.slice(k);
    t.textContent = t.textContent.slice(0, k);
    for (const el of [t, rest]) el.setAttributeNS(NS.xml, 'xml:space', 'preserve');
    t.after(rest);
    ({ atoms } = readParagraph(p));
  }
  const a = atoms.find((x) => x.end === pos && x.start < pos);
  const b = atoms.find((x) => x.start === pos && x.end > pos);
  if (!a || !b || !a.run || a.run !== b.run || a.node.parentNode !== a.run) return;
  const run = a.run;
  const second = run.cloneNode(false);
  const rPr = wChild(run, 'rPr');
  if (rPr) second.append(rPr.cloneNode(true));
  while (a.node.nextSibling) second.append(a.node.nextSibling);
  run.after(second);
}

/** Marks a changed paragraph: its words put in inserted, those taken out put back deleted. */
function markChanged(p, parts, mark, dom) {
  const inserted = [];
  const deleted = [];
  let pos = 0;
  for (const part of parts) {
    if (part.kind === 'del') deleted.push({ at: pos, text: part.text });
    else {
      if (part.kind === 'ins') inserted.push([pos, pos + part.text.length]);
      pos += part.text.length;
    }
  }
  const cuts = new Set([...inserted.flat(), ...deleted.map((d) => d.at)]);
  for (const at of [...cuts].sort((x, y) => y - x)) splitAt(p, at);
  let { atoms } = readParagraph(p);
  for (const [s, e] of inserted) {
    const runs = [...new Set(atoms.filter((a) => a.run && a.start >= s && a.end <= e && a.end > a.start).map((a) => a.run))]
      // Only runs all of whose text is in the insertion (a run that couldn't be split stays as it is).
      .filter((run) => atoms.every((a) => a.run !== run || (a.start >= s && a.end <= e)));
    wrapInserted(runs, mark);
  }
  ({ atoms } = readParagraph(p));
  for (const { at, text } of deleted) {
    const before = [...atoms].reverse().find((a) => a.end === at && a.run);
    const after = atoms.find((a) => a.start === at && a.run && a.end > a.start);
    const near = before || after;
    const del = mark('del');
    del.append(deletedRun(dom, text, near ? wChild(near.run, 'rPr') : null));
    const outer = (run) => (isW(run.parentNode, 'ins') ? run.parentNode : run);
    if (before) outer(before.run).after(del);
    else if (after) outer(after.run).before(del);
    else {
      const pPr = wChild(p, 'pPr');
      if (pPr) pPr.after(del);
      else p.prepend(del);
    }
  }
}

/** The redline's package: `zip` (this document's, as downloaded, not encrypted) with `xml` as its main part. */
export async function redlinePackage(JSZip, zipBytes, mainPath, otherXml, options) {
  const zip = await JSZip.loadAsync(zipBytes);
  const current = await zip.file(mainPath).async('string');
  // Ids after every one the package uses, so the new tracked changes clash with none.
  let firstId = 0;
  for (const f of Object.values(zip.files)) {
    if (!f.dir && /\.xml$/i.test(f.name)) firstId = Math.max(firstId, largestId(await f.async('string')));
  }
  const styles = zip.file('word/styles.xml') ? await zip.file('word/styles.xml').async('string') : '';
  const styleIds = new Set([...styles.matchAll(/w:styleId="([^"]+)"/g)].map((m) => m[1]));
  const result = redline(current, otherXml, { ...options, firstId, styleIds });
  zip.file(mainPath, result.xml);
  return { zip, ...result };
}

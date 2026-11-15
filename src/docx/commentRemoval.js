// Comments whose text was deleted. In Word, deleting the text a comment is attached to deletes the
// comment with it. Intact does the same, but completely: the comment, its replies and its entries in
// the companion parts go, and no half of a comment range is left behind. Comment parts are only
// rewritten when a comment was actually removed, so an untouched file keeps them byte for byte.
import { NS, parseXml, serializeXml, wAttr } from './xml.js';

const W = NS.w;
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';

/** The comment parts, by relationship type suffix (as in DocxDocument's relOf). */
export const COMMENT_PART_RELS = {
  comments: '/comments',
  extended: '/commentsExtended',
  ids: '/commentsIds',
  extensible: '/commentsExtensible',
};

const ids = (root, local) => new Set([...(root?.getElementsByTagNameNS(W, local) || [])].map((e) => wAttr(e, 'id')));

/** Ids of the comments referenced (by a w:commentReference) under `root`. */
export function commentIdsIn(root) {
  return ids(root, 'commentReference');
}

/** Ids of the comments whose range has both a start and an end marker under `root`. */
export function pairedCommentRanges(root) {
  const ends = ids(root, 'commentRangeEnd');
  return new Set([...ids(root, 'commentRangeStart')].filter((id) => ends.has(id)));
}

/**
 * The comments to remove, as a Set of ids: those that were referenced in the body when the document
 * was opened and no longer are (unless also referenced from a header, footer or note), plus the
 * replies of any comment removed, as Word removes a whole thread with its first comment.
 */
export function removedComments(comments, { wasInBody, present, elsewhere }) {
  const gone = new Set([...wasInBody].filter((id) => !present.has(id) && !elsewhere.has(id)));
  for (let grew = gone.size > 0; grew;) {
    grew = false;
    for (const c of comments) {
      if (c.parentId && gone.has(c.parentId) && !gone.has(c.id)) {
        gone.add(c.id);
        grew = true;
      }
    }
  }
  return gone;
}

/**
 * Removes what is left in the body of the comments in `gone` (range markers and references), and the
 * leftover half of any range in `paired` whose other marker was deleted with the text.
 */
export function pruneCommentMarkers(body, gone, paired) {
  const all = (local) => [...body.getElementsByTagNameNS(W, local)];
  for (const local of ['commentRangeStart', 'commentRangeEnd']) {
    for (const el of all(local)) if (gone.has(wAttr(el, 'id'))) el.remove();
  }
  for (const ref of all('commentReference')) {
    if (!gone.has(wAttr(ref, 'id'))) continue;
    const run = ref.parentNode;
    ref.remove();
    const rest = [...run.childNodes].filter((n) => n.nodeType === 1 && !(n.namespaceURI === W && n.localName === 'rPr'));
    if (run.namespaceURI === W && run.localName === 'r' && !rest.length) run.remove();
  }
  const starts = ids(body, 'commentRangeStart');
  const ends = ids(body, 'commentRangeEnd');
  for (const id of paired) {
    if (gone.has(id) || starts.has(id) === ends.has(id)) continue;
    for (const el of all(starts.has(id) ? 'commentRangeStart' : 'commentRangeEnd')) if (wAttr(el, 'id') === id) el.remove();
  }
}

/**
 * The comment parts with the comments in `gone` taken out: { path: xml }. `parts` holds the original
 * text of each part ({ comments: { path, text }, extended, ids, extensible }); with nothing to remove,
 * the original text is returned, so a later save after an undo puts the comments back.
 */
export function pruneCommentParts(parts, comments, gone) {
  const out = {};
  if (!gone.size) {
    for (const p of Object.values(parts)) out[p.path] = p.text;
    return out;
  }
  const paraIds = new Set(comments.filter((c) => gone.has(c.id) && c.paraId).map((c) => c.paraId));
  const durableIds = new Set();
  const edit = (key, change) => {
    const part = parts[key];
    if (!part) return;
    const doc = parseXml(part.text);
    change(doc);
    out[part.path] = serializeXml(doc);
  };
  const removeWhere = (doc, ns, local, test) => {
    for (const el of [...doc.getElementsByTagNameNS(ns, local)]) if (test(el)) el.remove();
  };
  edit('comments', (doc) => removeWhere(doc, W, 'comment', (el) => gone.has(wAttr(el, 'id'))));
  edit('extended', (doc) => removeWhere(doc, W15, 'commentEx', (el) => paraIds.has(el.getAttributeNS(W15, 'paraId'))));
  edit('ids', (doc) => removeWhere(doc, W16CID, 'commentId', (el) => {
    if (!paraIds.has(el.getAttributeNS(W16CID, 'paraId'))) return false;
    durableIds.add(el.getAttributeNS(W16CID, 'durableId'));
    return true;
  }));
  edit('extensible', (doc) => removeWhere(doc, W16CEX, 'commentExtensible', (el) => durableIds.has(el.getAttributeNS(W16CEX, 'durableId'))));
  return out;
}

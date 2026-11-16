// Links: what a typed address means, and which addresses may become links. A link is written as
// Word writes one: <w:hyperlink r:id> with an external relationship to the address, in the part
// (document, footnotes, endnotes) the link is in (see DocxDocument.linkRel and writeLinks).

const DOI = /^(?:doi:\s*|https?:\/\/(?:dx\.)?doi\.org\/)?(10\.\d{4,9}\/\S+)$/i;
const EMAIL = /^[^\s@<>()]+@[^\s@<>()]+\.[^\s@<>()]+$/;

/**
 * The address for what was typed in the Insert link dialog, or null if it isn't one: a web
 * address (https://…, or www.… made https://), an e-mail address (made mailto:) or a DOI
 * (10.1000/xyz or doi:10.1000/xyz, made https://doi.org/10.1000/xyz).
 */
export function normalizeLink(input) {
  const text = (input || '').trim();
  if (!text || /\s/.test(text)) return null;
  const doi = DOI.exec(text);
  if (doi) return `https://doi.org/${doi[1]}`;
  if (/^https?:\/\/[^\s/]+\.[^\s]+/i.test(text) || /^https?:\/\/localhost\b/i.test(text)) return text;
  if (/^mailto:[^\s]+@/i.test(text)) return text;
  if (/^www\.[^\s]+\.[^\s]+/i.test(text)) return `https://${text}`;
  if (EMAIL.test(text)) return `mailto:${text}`;
  return null;
}

/** True for an address a pasted link may keep: web and e-mail addresses only. */
export function isLinkable(href) {
  return /^(https?:\/\/|mailto:)\S+$/i.test((href || '').trim());
}

/** The relationships part of a package part: word/footnotes.xml → word/_rels/footnotes.xml.rels. */
export function relsPathOf(partPath) {
  const slash = partPath.lastIndexOf('/');
  return `${slash >= 0 ? partPath.slice(0, slash + 1) : ''}_rels/${partPath.slice(slash + 1)}.rels`;
}

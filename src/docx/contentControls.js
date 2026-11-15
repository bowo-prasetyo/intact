// Content controls (<w:sdt>): the fill-in fields of templates and forms. Text, checkbox, drop-down
// and date controls are editable; controls Word itself would not let the user change (locked ones),
// or whose content Word would overwrite on open (bound to the document's XML data), stay read-only.
import { NS, wChild, wChildren, wAttr, onOff } from './xml.js';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const w14Child = (el, local) => (el ? [...el.children].find((c) => c.namespaceURI === W14 && c.localName === local) : null);
const w14Attr = (el, name) => el?.getAttributeNS(W14, name) ?? null;

/**
 * What kind of control `sdt` is and whether it can be edited here.
 * kind: 'text' | 'checkbox' | 'dropdown' | 'combo' | 'date'; readOnly: a reason to show, or null.
 */
export function controlInfo(sdt) {
  const sdtPr = wChild(sdt, 'sdtPr');
  const has = (local) => !!wChild(sdtPr, local);
  const checkbox = w14Child(sdtPr, 'checkbox');
  let kind = 'text';
  if (checkbox) kind = 'checkbox';
  else if (has('dropDownList')) kind = 'dropdown';
  else if (has('comboBox')) kind = 'combo';
  else if (has('date')) kind = 'date';

  const gallery = wAttr(wChild(wChild(sdtPr, 'docPartObj'), 'docPartGallery'), 'val') || '';
  const label = wAttr(wChild(sdtPr, 'alias'), 'val') || wAttr(wChild(sdtPr, 'tag'), 'val') || gallery;
  const lock = wAttr(wChild(sdtPr, 'lock'), 'val');
  let readOnly = null;
  if (lock === 'contentLocked' || lock === 'sdtContentLocked') readOnly = 'locked in Word';
  else if (has('dataBinding')) readOnly = 'linked to document data, edit it in Word';
  else if (/table of contents|bibliograph/i.test(gallery)) readOnly = 'updated by Word';
  else if (['picture', 'citation', 'bibliography', 'equation', 'group', 'docPartList'].some(has)) readOnly = 'not editable here';

  return {
    kind,
    label,
    readOnly,
    placeholder: onOff(wChild(sdtPr, 'showingPlcHdr')) === true,
    plain: has('text'),
    multiLine: ['1', 'true', 'on'].includes(wAttr(wChild(sdtPr, 'text'), 'multiLine') || ''),
    checked: checkbox ? ['1', 'true'].includes(w14Attr(w14Child(checkbox, 'checked'), 'val')) : false,
    symbols: checkbox ? {
      on: symbol(w14Child(checkbox, 'checkedState'), '2612'),
      off: symbol(w14Child(checkbox, 'uncheckedState'), '2610'),
    } : null,
  };
}

function symbol(state, fallback) {
  return String.fromCodePoint(parseInt(w14Attr(state, 'val') || fallback, 16));
}

const KIND_NAMES = { text: 'Text field', checkbox: 'Check box', dropdown: 'Drop-down list', combo: 'Combo box', date: 'Date' };

export function controlTitle(info) {
  const name = info.label ? `“${info.label}”` : KIND_NAMES[info.kind].toLowerCase();
  if (info.readOnly) return `Content control ${name} – ${info.readOnly}`;
  const how = {
    checkbox: 'click or press Alt+↓ to tick or clear',
    dropdown: 'click or press Alt+↓ to choose',
    date: 'click or press Alt+↓ to pick a date',
    combo: 'type, or click ▾ / press Alt+↓ to choose',
  }[info.kind];
  return `${KIND_NAMES[info.kind]}${info.label ? ` “${info.label}”` : ''}${how ? ` – ${how}` : ''}`;
}

/** The choices of a drop-down or combo box control: [{ text, value }]. */
export function controlChoices(sdt) {
  const sdtPr = wChild(sdt, 'sdtPr');
  const list = wChild(sdtPr, 'dropDownList') || wChild(sdtPr, 'comboBox');
  return wChildren(list, 'listItem').map((li) => {
    const value = wAttr(li, 'value') ?? '';
    return { text: wAttr(li, 'displayText') ?? value, value };
  });
}

/** The date control's settings: { iso: 'YYYY-MM-DD' | '', format, locale }. */
export function controlDate(sdt) {
  const date = wChild(wChild(sdt, 'sdtPr'), 'date');
  return {
    iso: (wAttr(date, 'fullDate') || '').slice(0, 10),
    format: wAttr(wChild(date, 'dateFormat'), 'val') || 'M/d/yyyy',
    locale: wAttr(wChild(date, 'lid'), 'val') || undefined,
  };
}

/** Formats an ISO date ('2026-03-05') with a Word date picture such as "d MMMM yyyy". */
export function formatWordDate(iso, format, locale) {
  const [y, m, d] = iso.split('-').map(Number);
  const date = new Date(Date.UTC(y, m - 1, d));
  const name = (opts) => {
    try { return new Intl.DateTimeFormat(locale, { ...opts, timeZone: 'UTC' }).format(date); } catch { return new Intl.DateTimeFormat('en-US', { ...opts, timeZone: 'UTC' }).format(date); }
  };
  const tokens = {
    dddd: () => name({ weekday: 'long' }), ddd: () => name({ weekday: 'short' }),
    dd: () => String(d).padStart(2, '0'), d: () => String(d),
    MMMM: () => name({ month: 'long' }), MMM: () => name({ month: 'short' }),
    MM: () => String(m).padStart(2, '0'), M: () => String(m),
    yyyy: () => String(y), yy: () => String(y % 100).padStart(2, '0'),
  };
  return format.replace(/'([^']*)'|dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy/g, (t, lit) => (lit !== undefined ? lit : tokens[t]()));
}

/**
 * Applies what the user did to a cloned <w:sdt>, after its content has been rewritten:
 * replaced placeholder text, a ticked box, a chosen item or date.
 */
export function updateControlProperties(sdt, el, { placeholderReplaced, isPlaceholderStyle }) {
  const sdtPr = wChild(sdt, 'sdtPr');
  const content = wChild(sdt, 'sdtContent');
  if (placeholderReplaced) {
    for (const c of wChildren(sdtPr, 'showingPlcHdr')) sdtPr.removeChild(c);
    const defaults = wChild(sdtPr, 'rPr');
    for (const r of content.getElementsByTagNameNS(NS.w, 'r')) {
      let rPr = wChild(r, 'rPr');
      const rStyle = wChild(rPr, 'rStyle');
      if (rStyle && isPlaceholderStyle(wAttr(rStyle, 'val'))) rPr.removeChild(rStyle);
      if (rPr && !rPr.firstElementChild) { r.removeChild(rPr); rPr = null; }
      // New text takes the control's own formatting, as in Word.
      if (!rPr && defaults?.firstElementChild) r.insertBefore(defaults.cloneNode(true), r.firstChild);
    }
  }
  if (el.dataset.checked != null) {
    const checked = w14Child(w14Child(sdtPr, 'checkbox'), 'checked');
    checked?.setAttributeNS(W14, 'w14:val', el.dataset.checked === '1' ? '1' : '0');
  }
  if (el.dataset.value != null) {
    const list = wChild(sdtPr, 'dropDownList') || wChild(sdtPr, 'comboBox');
    list?.setAttributeNS(NS.w, 'w:lastValue', el.dataset.value);
  }
  if (el.dataset.fullDate != null) {
    wChild(sdtPr, 'date')?.setAttributeNS(NS.w, 'w:fullDate', `${el.dataset.fullDate}T00:00:00Z`);
  }
}

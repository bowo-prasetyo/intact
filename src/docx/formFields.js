// Legacy form fields: Word's older fill-in fields (Developer → Legacy Tools), still common in school,
// medical and club forms. Each is a complex field (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN) whose
// settings are in a <w:ffData> on its begin marker:
//
//   <w:r><w:fldChar w:fldCharType="begin"><w:ffData>
//     <w:name/> <w:enabled/> … then one of <w:textInput>, <w:checkBox> or <w:ddList>
//   </w:ffData></w:fldChar></w:r>
//
// A text field's value is its result, edited as text. A check box and a drop-down list have no
// result Word reads: Word draws the box, or the chosen entry, from <w:ffData> itself
// (ECMA-376 Part 1, 17.16.17). So ticking a box or choosing an entry changes the begin marker's
// <w:ffData> (w:checked, w:result), and nothing else.
import { NS, wChild, wChildren, wAttr, onOff, createW } from './xml.js';

const W = NS.w;

/** The <w:ffData> of field begin marker run `r`, or null. */
function ffDataOf(r) {
  return wChild(wChild(r, 'fldChar'), 'ffData');
}

/**
 * What legacy form field begin marker run `r` starts, or null if it starts none.
 * kind: 'text' | 'checkbox' | 'dropdown'; enabled: false where Word doesn't let it be filled in.
 */
export function formFieldInfo(r) {
  const ff = ffDataOf(r);
  if (!ff) return null;
  const check = wChild(ff, 'checkBox');
  const list = wChild(ff, 'ddList');
  const input = wChild(ff, 'textInput');
  const kind = check ? 'checkbox' : list ? 'dropdown' : input ? 'text' : null;
  if (!kind) return null;
  const help = wChild(ff, 'statusText') || wChild(ff, 'helpText');
  const info = {
    kind,
    name: wAttr(wChild(ff, 'name'), 'val') || '',
    enabled: onOff(wChild(ff, 'enabled')) !== false,
    // Help text Word shows in its status bar (w:type="text"; "autoText" names an entry, not text).
    help: help && wAttr(help, 'type') !== 'autoText' ? wAttr(help, 'val') || '' : '',
  };
  if (check) {
    const own = onOff(wChild(check, 'checked'));
    info.checked = own ?? onOff(wChild(check, 'default')) ?? false;
  }
  if (list) {
    info.entries = wChildren(list, 'listEntry').map((e) => wAttr(e, 'val') || '');
    const at = (local) => {
      const v = wAttr(wChild(list, local), 'val');
      return v != null && /^\d+$/.test(v) ? Number(v) : null;
    };
    const chosen = at('result') ?? at('default') ?? 0;
    info.selected = chosen < info.entries.length ? chosen : 0;
  }
  if (input) info.maxLength = Number(wAttr(wChild(input, 'maxLength'), 'val')) || 0;
  return info;
}

/** Ticks (or clears) the check box that begin marker run `r` starts, in place. */
export function setFormCheck(r, on) {
  const check = wChild(ffDataOf(r), 'checkBox');
  if (!check) return;
  const byDefault = onOff(wChild(check, 'default')) ?? false;
  let checked = wChild(check, 'checked');
  // CT_FFCheckBox: (size | sizeAuto), default?, checked?. Absent, w:checked is the default.
  if (on === byDefault) {
    checked?.remove();
    return;
  }
  if (!checked) {
    checked = createW(r.ownerDocument, 'checked');
    check.append(checked);
  }
  if (on) checked.removeAttributeNS(W, 'val');
  else checked.setAttributeNS(W, 'w:val', '0');
}

/** Chooses entry `index` of the drop-down list that begin marker run `r` starts, in place. */
export function setFormDropdown(r, index) {
  const list = wChild(ffDataOf(r), 'ddList');
  if (!list) return;
  let result = wChild(list, 'result');
  // CT_FFDDList: result?, default?, listEntry*.
  if (!result) {
    result = createW(r.ownerDocument, 'result');
    list.insertBefore(result, list.firstChild);
  }
  result.setAttributeNS(W, 'w:val', String(index));
}

/** What a legacy form field is, for its tooltip. */
export function formFieldTitle(info) {
  const what = { text: 'Text form field', checkbox: 'Check box', dropdown: 'Drop-down form field' }[info.kind];
  const named = info.help || info.name;
  if (!info.enabled) return `${what}${named ? ` (${named})` : ''} – turned off in Word, so it can’t be filled in`;
  const how = { text: 'type in it', checkbox: 'click to tick or clear it', dropdown: 'click to choose' }[info.kind];
  return `${what}${named ? ` (${named})` : ''} – ${how}`;
}

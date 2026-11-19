// Lengths in the dialogs (margins, indents, tab stops, a picture's size and position), in
// centimetres or inches, as Word shows them in the unit of the place it is used. Here the unit is
// the one chosen in a dialog (remembered in this browser), or else the document's paper: inches for
// the North American sizes (Letter, Legal, Executive, Tabloid), centimetres for the others.

export const UNITS = { cm: 'cm', in: 'in' };
export const UNIT_NAMES = { cm: 'centimetres', in: 'inches' };
export const TWIPS_PER = { cm: 1440 / 2.54, in: 1440 };
export const EMU_PER = { cm: 360000, in: 914400 };

const KEY = 'intact.unit';
// Short sides of the North American paper sizes, in twips: Letter and Legal, Executive, Tabloid.
const US_SHORT_SIDES = [12240, 10440, 15840];

/** The unit chosen in a dialog before, or null. */
export function savedUnit() {
  try {
    const u = localStorage.getItem(KEY);
    return u === 'cm' || u === 'in' ? u : null;
  } catch {
    return null;
  }
}

/** Remembers the unit chosen, for the next dialog and document. */
export function saveUnit(unit) {
  try { localStorage.setItem(KEY, unit); } catch { /* this session only */ }
}

/** Inches for paper of a North American size (`w` × `h` in twips), else centimetres. */
export function paperUnit(w, h) {
  const short = Math.min(w || 0, h || 0);
  return US_SHORT_SIDES.some((s) => Math.abs(short - s) <= 20) ? 'in' : 'cm';
}

/** The unit to show lengths in, for a document on paper `w` × `h` (twips). */
export function unitFor(w, h) {
  return savedUnit() || paperUnit(w, h);
}

/** `twips` in `unit`, to two decimals. */
export const fromTwips = (twips, unit) => Math.round((twips / TWIPS_PER[unit]) * 100) / 100;

/** A length typed in `unit`, in twips. */
export const toTwips = (value, unit) => Math.round(value * TWIPS_PER[unit]);

/**
 * Switches a dialog's lengths to `unit`. `state` holds the lengths shown, `initial` those it opened
 * with (in the same unit, to tell what was changed) and `twips` their exact values. A length left
 * as it was is shown afresh from its exact value, so it still reads as unchanged; one changed is
 * converted as typed.
 */
export function switchUnit(state, initial, twips, keys, unit) {
  const from = state.unit;
  if (!from || from === unit) return;
  for (const k of keys) {
    const shown = state[k];
    const untouched = initial && Number(shown) === initial[k];
    if (initial && twips && k in twips) initial[k] = fromTwips(twips[k], unit);
    if (untouched && initial) state[k] = initial[k];
    else if (shown !== '' && Number.isFinite(Number(shown))) state[k] = fromTwips(toTwips(Number(shown), from), unit);
  }
  state.unit = unit;
}

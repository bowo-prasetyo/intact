// Today's date as Word writes it in a DATE field's picture (\@ "MMMM d, yyyy"), worked out from how
// the browser writes it in the text's language, so the date Word puts there each time it opens the
// file looks as the one inserted here.

// 3 February 2001: a day and a month of one digit, to see where the format pads them.
const PROBE = new Date(2001, 1, 3);

function partsOf(lang, options, date) {
  try { return new Intl.DateTimeFormat(lang, options).formatToParts(date); } catch { return new Intl.DateTimeFormat('en', options).formatToParts(date); }
}

/**
 * The Word date picture for a date written with `options` (Intl.DateTimeFormat's) in `lang`:
 * "MMMM d, yyyy" for en-US's long date, "dd/MM/yyyy" for en-GB's short one. Text that isn't a
 * part of the date (", ", " de ") is kept, letters quoted as Word quotes them.
 */
export function wordDatePicture(lang, options) {
  const probe = partsOf(lang, options, PROBE);
  const name = (type, width) => partsOf(lang, { [type]: width }, PROBE).find((p) => p.type === type)?.value;
  return probe.map(({ type, value }) => {
    if (type === 'day') return value.length > 1 ? 'dd' : 'd';
    if (type === 'year') return value.length === 2 ? 'yy' : 'yyyy';
    if (type === 'month') {
      if (/^\d+$/.test(value)) return value.length > 1 ? 'MM' : 'M';
      return value === name('month', 'long') ? 'MMMM' : 'MMM';
    }
    if (type === 'weekday') return value === name('weekday', 'long') ? 'dddd' : 'ddd';
    // Literal text: letters (and quotes) quoted, as Word's pictures write words in a date.
    return /[\p{L}'"]/u.test(value) ? `'${value.replace(/'/g, '')}'` : value;
  }).join('');
}

/** A DATE field's instruction, as Word's Date and Time dialog writes it with "Update automatically". */
export function dateFieldInstruction(picture) {
  return ` DATE \\@ "${picture}" `;
}

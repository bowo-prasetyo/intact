// Word's table formulas (Layout > Formula, the = field): =SUM(ABOVE), =AVERAGE(LEFT), =B2*C2,
// =SUM(B2:B5)*11%, worked out from the numbers in the table's cells as Word works them out, and
// shown in a number format (\# "#,##0.00") with the document's own separators (1.234,56 in
// Indonesian, 1,234.56 in English). Works on the page: a cell is a td of the table the field is in.
//
// ABOVE, LEFT, BELOW and RIGHT take the cells beside the formula's cell in that direction, up to the
// first one that is empty or doesn't hold a number (a heading), as Word does; cell references are
// Word's: columns A, B, C… and rows 1, 2, 3… of the table, as its grid lays them out.

const FUNCTIONS = {
  SUM: (xs) => xs.reduce((a, b) => a + b, 0),
  AVERAGE: (xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0),
  MIN: (xs) => (xs.length ? Math.min(...xs) : 0),
  MAX: (xs) => (xs.length ? Math.max(...xs) : 0),
  COUNT: (xs) => xs.length,
  PRODUCT: (xs) => xs.reduce((a, b) => a * b, 1),
  ABS: (xs) => Math.abs(xs[0] ?? 0),
  ROUND: (xs) => {
    const f = 10 ** (xs[1] ?? 0);
    return Math.round((xs[0] ?? 0) * f) / f;
  },
  INT: (xs) => Math.trunc(xs[0] ?? 0),
};
export const FORMULA_FUNCTIONS = Object.keys(FUNCTIONS);
const DIRECTIONS = new Set(['ABOVE', 'BELOW', 'LEFT', 'RIGHT']);

/** The separators numbers are written with in `lang`: { group: '.', decimal: ',' } for id-ID. */
export function separators(lang) {
  try {
    const parts = new Intl.NumberFormat(lang || 'en').formatToParts(12345.6);
    return { group: parts.find((p) => p.type === 'group')?.value || ',', decimal: parts.find((p) => p.type === 'decimal')?.value || '.' };
  } catch {
    return { group: ',', decimal: '.' };
  }
}

/**
 * The number a cell's text holds, written with `seps` ("Rp 1.250.000", "(2,5)", "12%"), or null when
 * it holds none: an empty cell, or words.
 */
export function parseCellNumber(text, seps = { group: ',', decimal: '.' }) {
  const t = (text || '').replace(/[\s  ]/g, '');
  if (!/\d/.test(t)) return null;
  // Words around the number (a currency, a unit) are allowed; words with it are not ("Room 12").
  const m = /^([^\d\-−(]{0,3})(\(|[-−])?([\d.,'’]*\d[\d.,'’]*)(\))?(%?)([^\d]{0,4})$/.exec(t);
  if (!m) return null;
  let digits = m[3].replace(/['’]/g, '');
  digits = digits.split(seps.group).join('');
  if (seps.decimal !== '.') digits = digits.replace(seps.decimal, '.');
  if (!/^\d*\.?\d+$|^\d+\.?$/.test(digits)) return null;
  let n = Number(digits);
  if (m[2] || (m[4] && m[2] === '(')) n = -n;
  if (m[5]) n /= 100;
  return Number.isFinite(n) ? n : null;
}

/** The table's grid: rows of cells by grid column (a merged cell in each column it spans). */
function grid(table) {
  const rows = [...table.rows].filter((r) => r.closest('table') === table);
  return rows.map((tr) => {
    const out = [];
    for (const td of tr.cells) for (let i = 0; i < (td.colSpan || 1); i++) out.push(td);
    return out;
  });
}

/** The text of a cell, without a formula field's own result where `skip` is that field. */
const cellText = (td) => td.textContent;

/** "B3" as { r: 2, c: 1 }. */
function cellRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]) - 1, c: c - 1 };
}

/**
 * Tokens of a formula (without its "="): numbers, names (SUM, ABOVE, B3), operators and brackets.
 * Numbers are written with the document's decimal sign; arguments are separated by ";" (and by ","
 * where the decimal sign isn't one), as Word's list separator is.
 */
function tokenize(src, seps) {
  const out = [];
  const dec = seps.decimal === ',' ? ',' : '\\.';
  const re = new RegExp(`\\s*(?:(\\d+(?:${dec}\\d+)?%?)|([A-Za-z]+\\d*(?::[A-Za-z]+\\d+)?)|([-+*/^(),;]))`, 'y');
  let at = 0;
  while (at < src.length) {
    if (/^\s*$/.test(src.slice(at))) break;
    re.lastIndex = at;
    const m = re.exec(src);
    if (!m) throw new Error('syntax');
    at = re.lastIndex;
    if (m[1]) {
      let s = m[1];
      const pct = s.endsWith('%');
      if (pct) s = s.slice(0, -1);
      const n = Number(s.replace(',', '.'));
      out.push({ num: pct ? n / 100 : n });
    } else if (m[2]) out.push({ name: m[2].toUpperCase() });
    else out.push({ op: m[3] === ';' ? ',' : m[3] });
  }
  return out;
}

/**
 * The value of formula `formula` ("=SUM(ABOVE)") in cell `td`, with the numbers written as `lang`
 * writes them. Throws an Error('syntax') for a formula Word would call a syntax error.
 */
export function evaluateFormula(formula, td, lang) {
  const seps = separators(lang);
  const src = formula.trim().replace(/^=/, '');
  const tokens = tokenize(src, seps);
  const table = td?.closest('table');
  const g = table ? grid(table) : [];
  const r0 = table ? g.findIndex((row) => row.includes(td)) : -1;
  const c0 = r0 >= 0 ? g[r0].indexOf(td) : -1;
  const num = (cell) => (cell ? parseCellNumber(cellText(cell), seps) : null);
  // ABOVE, LEFT…: the cells beside this one, up to the first that holds no number.
  const direction = (dir) => {
    if (r0 < 0) throw new Error('syntax');
    const [dr, dc] = { ABOVE: [-1, 0], BELOW: [1, 0], LEFT: [0, -1], RIGHT: [0, 1] }[dir];
    const out = [];
    for (let r = r0 + dr, c = c0 + dc; g[r] && g[r][c]; r += dr, c += dc) {
      if (g[r][c] === td) continue; // a cell spanning the formula's own column or row
      const v = num(g[r][c]);
      if (v == null) break;
      out.push(v);
    }
    return out;
  };
  const range = (name) => {
    const [a, b = a] = name.split(':').map(cellRef);
    if (!a || !b) throw new Error('syntax');
    const out = [];
    for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++) {
      for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) {
        const v = num(g[r]?.[c]);
        if (v != null && !out.includes(g[r][c])) out.push(v);
      }
    }
    return out;
  };

  let i = 0;
  const peek = () => tokens[i];
  const take = (op) => (tokens[i]?.op === op ? tokens[i++] : null);
  const expect = (op) => {
    if (!take(op)) throw new Error('syntax');
  };
  // expr := term (('+'|'-') term)*; term := power (('*'|'/') power)*; power := unary ('^' unary)*
  function expr() {
    let v = term();
    for (let t; (t = take('+') || take('-'));) v = t.op === '+' ? v + term() : v - term();
    return v;
  }
  function term() {
    let v = power();
    for (let t; (t = take('*') || take('/'));) {
      const w = power();
      if (t.op === '/' && w === 0) throw new Error('zero');
      v = t.op === '*' ? v * w : v / w;
    }
    return v;
  }
  function power() {
    let v = unary();
    while (take('^')) v **= unary();
    return v;
  }
  function unary() {
    if (take('-')) return -unary();
    if (take('+')) return unary();
    return primary();
  }
  // A function's arguments: numbers, expressions, directions and ranges, each a list of values.
  function args() {
    const vals = [];
    expect('(');
    if (take(')')) return vals;
    do {
      const t = peek();
      if (t?.name && DIRECTIONS.has(t.name)) {
        i++;
        vals.push(...direction(t.name));
      } else if (t?.name && /^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(t.name) && t.name.includes(':')) {
        i++;
        vals.push(...range(t.name));
      } else vals.push(expr());
    } while (take(','));
    expect(')');
    return vals;
  }
  function primary() {
    const t = tokens[i++];
    if (!t) throw new Error('syntax');
    if (t.num != null) return t.num;
    if (t.op === '(') {
      const v = expr();
      expect(')');
      return v;
    }
    if (t.name && FUNCTIONS[t.name]) return FUNCTIONS[t.name](args());
    if (t.name && /^[A-Z]+\d+$/.test(t.name)) {
      const v = range(t.name);
      return v[0] ?? 0;
    }
    throw new Error('syntax');
  }
  if (!tokens.length) throw new Error('syntax');
  const value = expr();
  if (i !== tokens.length) throw new Error('syntax');
  return value;
}

/**
 * `value` in Word's number picture `picture` ("#,##0.00", "Rp #.##0", "0%"), whose separators are
 * the document's, written with them; with no picture, as Word writes a plain result.
 */
export function formatResult(value, picture, lang) {
  const seps = separators(lang);
  if (!picture) {
    const s = String(Math.round(value * 1e10) / 1e10);
    return seps.decimal === '.' ? s : s.replace('.', seps.decimal);
  }
  const m = /[#0][#0.,]*/.exec(picture);
  if (!m) return picture;
  const body = m[0];
  const decAt = body.lastIndexOf(seps.decimal);
  const intPart = decAt >= 0 ? body.slice(0, decAt) : body;
  const decimals = decAt >= 0 ? body.slice(decAt + 1).replace(/[^#0]/g, '').length : 0;
  const grouped = intPart.includes(seps.group);
  const pct = /%/.test(picture.slice(m.index + body.length));
  const v = pct ? value * 100 : value;
  const fixed = Math.abs(v).toFixed(decimals);
  let [whole, frac] = fixed.split('.');
  if (grouped) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, seps.group);
  const minWhole = intPart.replace(/[^0]/g, '').length;
  whole = whole.padStart(minWhole, '0') || '0';
  const number = `${v < 0 ? '-' : ''}${whole}${frac ? seps.decimal + frac : ''}`;
  return picture.slice(0, m.index) + number + picture.slice(m.index + body.length);
}

/** A formula field's instruction (' =SUM(ABOVE) \# "#,##0" ') as { formula, picture }, or null if it isn't one. */
export function parseFormulaInstr(instr) {
  const m = /^\s*(=[^\\]*?)\s*(?:\\#\s*(?:"([^"]*)"|(\S+)))?\s*(?:\\\*\s*MERGEFORMAT\s*)?$/i.exec(instr || '');
  return m ? { formula: m[1].trim(), picture: m[2] ?? m[3] ?? '' } : null;
}

/** The instruction Word writes for a formula, as its Formula dialog does: ' =SUM(ABOVE) \# "#,##0" '. */
export function formulaInstr(formula, picture) {
  return ` ${formula.trim()} ${picture ? `\\# "${picture}" ` : ''}`;
}

/** The result of formula field `instr` in cell `td`, or Word's error text for one it can't work out. */
export function formulaResult(instr, td, lang) {
  const f = parseFormulaInstr(instr);
  if (!f) return null;
  try {
    return formatResult(evaluateFormula(f.formula, td, lang), f.picture, lang);
  } catch (e) {
    return e.message === 'zero' ? '!Zero Divide' : '!Syntax Error';
  }
}

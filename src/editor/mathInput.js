// Typing an equation, as Word's equation editor takes LaTeX: a small part of LaTeX's maths, written
// as Word's Office Math (OMML), which Word then edits as its own. Fractions (\frac{a}{b}),
// superscripts and subscripts (x^2, x_{i}), roots (\sqrt{x}, \sqrt[3]{x}), sums, products and
// integrals with limits (\sum_{i=1}^{n}), brackets that grow (\left( … \right)), Greek letters and
// the usual symbols (\alpha, \pi, \times, \le, \infty …), and functions (\sin, \log …) upright.

const SYMBOLS = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ',
  vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ',
  tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  times: '×', cdot: '⋅', div: '÷', pm: '±', mp: '∓', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠',
  approx: '≈', equiv: '≡', sim: '∼', propto: '∝', infty: '∞', partial: '∂', nabla: '∇', to: '→', rightarrow: '→',
  leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐', leftrightarrow: '↔', iff: '⇔', in: '∈', notin: '∉', subset: '⊂',
  subseteq: '⊆', cup: '∪', cap: '∩', emptyset: '∅', forall: '∀', exists: '∃', neg: '¬', land: '∧', lor: '∨',
  degree: '°', circ: '∘', ldots: '…', cdots: '⋯', angle: '∠', perp: '⊥', parallel: '∥', prime: '′',
  ',': ' ', ';': ' ', quad: ' ', qquad: '  ', '{': '{', '}': '}', '%': '%', $: '$', '#': '#', '&': '&', _: '_',
};
// Large operators, with where their limits go (under and over, or beside), as Word places them.
const NARY = { sum: ['∑', 'undOvr'], prod: ['∏', 'undOvr'], coprod: ['∐', 'undOvr'], bigcup: ['⋃', 'undOvr'], bigcap: ['⋂', 'undOvr'], int: ['∫', 'subSup'], iint: ['∬', 'subSup'], iiint: ['∭', 'subSup'], oint: ['∮', 'subSup'] };
const FUNCTIONS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh', 'log', 'ln', 'exp', 'lim', 'max', 'min', 'det', 'gcd']);
const DELIMS = { '(': '(', ')': ')', '[': '[', ']': ']', '|': '|', '.': '', '\\{': '{', '\\}': '}', '\\langle': '⟨', '\\rangle': '⟩', '\\lvert': '|', '\\rvert': '|' };

class Parser {
  constructor(src) {
    this.src = src;
    this.i = 0;
  }

  error(msg) {
    throw new Error(msg);
  }

  peek() {
    while (this.i < this.src.length && /\s/.test(this.src[this.i])) this.i++;
    return this.src[this.i];
  }

  command() {
    // After a backslash: a name of letters, or one other character.
    const m = /^[A-Za-z]+|^./.exec(this.src.slice(this.i));
    if (!m) this.error('A \\ needs a name after it, such as \\frac or \\alpha');
    this.i += m[0].length;
    return m[0];
  }

  /** Nodes up to the end, or to `stop` (a closing brace, \right or ]). */
  list(stop = null) {
    const out = [];
    for (;;) {
      const c = this.peek();
      if (c === undefined) {
        if (stop === '}') this.error('A { has no } to close it');
        if (stop === 'right') this.error('A \\left has no \\right to close it');
        if (stop === ']') this.error('A [ has no ] to close it');
        return out;
      }
      if (stop === '}' && c === '}') { this.i++; return out; }
      if (stop === ']' && c === ']') { this.i++; return out; }
      if (stop === 'right' && this.src.startsWith('\\right', this.i)) { this.i += 6; return out; }
      if (c === '}') this.error('A } has no { before it');
      if (c === '^' || c === '_') {
        this.i++;
        const base = out.pop() || { t: 'run', text: '' };
        // "ab^2": the script goes on the last letter only.
        let b = base;
        if (base.t === 'run' && [...base.text].length > 1 && !base.fn) {
          const chars = [...base.text];
          out.push({ t: 'run', text: chars.slice(0, -1).join('') });
          b = { t: 'run', text: chars[chars.length - 1] };
        }
        out.push(this.script(b, c));
        continue;
      }
      const node = this.atom();
      const last = out[out.length - 1];
      if (node.t === 'run' && !node.fn && !node.op && last?.t === 'run' && !last.fn && !last.op) last.text += node.text;
      else out.push(node);
    }
  }

  /** `base` with the script starting at `kind` (^ or _), and the other one if it follows. */
  script(base, kind) {
    let sub = null;
    let sup = null;
    if (kind === '^') sup = this.argument(); else sub = this.argument();
    const next = this.peek();
    if ((next === '^' && !sup) || (next === '_' && !sub)) {
      this.i++;
      if (next === '^') sup = this.argument(); else sub = this.argument();
    }
    if (base.t === 'nary') return { ...base, sub: sub ?? base.sub, sup: sup ?? base.sup };
    if (sub && sup) return { t: 'subsup', e: [base], sub, sup };
    return sup ? { t: 'sup', e: [base], sup } : { t: 'sub', e: [base], sub };
  }

  /** A braced group, or one character or command. */
  argument() {
    const c = this.peek();
    if (c === undefined) this.error('Something should follow ^, _, \\frac or \\sqrt');
    if (c === '{') {
      this.i++;
      return this.list('}');
    }
    return [this.atom()];
  }

  atom() {
    const c = this.peek();
    if (c === '{') {
      this.i++;
      return { t: 'group', e: this.list('}') };
    }
    if (c !== '\\') {
      this.i++;
      // Digits and letters run together; an operator on its own.
      return /[+\-=<>±×÷*/,]/.test(c) ? { t: 'run', text: c === '-' ? '−' : c === '*' ? '∗' : c, op: true } : { t: 'run', text: c };
    }
    this.i++;
    const name = this.command();
    if (name === 'frac' || name === 'dfrac' || name === 'tfrac') return { t: 'frac', num: this.argument(), den: this.argument() };
    if (name === 'sqrt') {
      let deg = null;
      if (this.peek() === '[') {
        this.i++;
        deg = this.list(']');
      }
      return { t: 'rad', deg, e: this.argument() };
    }
    if (name === 'left') {
      const open = this.delimiter();
      const e = this.list('right');
      const close = this.delimiter();
      return { t: 'delim', open, close, e };
    }
    if (name === 'right') this.error('A \\right has no \\left before it');
    if (NARY[name]) {
      const [chr, limLoc] = NARY[name];
      const node = { t: 'nary', chr, limLoc, sub: null, sup: null, e: [] };
      // Its limits, then what it applies to: the next term.
      let n = node;
      while (this.peek() === '_' || this.peek() === '^') {
        const k = this.src[this.i++];
        n = { ...n, [k === '_' ? 'sub' : 'sup']: this.argument() };
      }
      if (this.peek() !== undefined && this.peek() !== '}' && !this.src.startsWith('\\right', this.i)) {
        const term = this.atom();
        let e = [term];
        if (this.peek() === '^' || this.peek() === '_') {
          const k = this.src[this.i++];
          e = [this.script(term, k)];
        }
        n.e = e;
      }
      return n;
    }
    if (FUNCTIONS.has(name)) return { t: 'run', text: name, fn: true };
    if (name === 'text' || name === 'mathrm' || name === 'operatorname') {
      // Its text as typed, spaces too (elsewhere spaces are LaTeX's, and left out).
      if (this.peek() === '{') {
        const end = this.src.indexOf('}', this.i);
        if (end < 0) this.error('A { has no } to close it');
        const text = this.src.slice(this.i + 1, end);
        this.i = end + 1;
        return { t: 'run', text, fn: true };
      }
      return { t: 'run', text: this.argument().map(textOf).join(''), fn: true };
    }
    if (SYMBOLS[name] !== undefined) return { t: 'run', text: SYMBOLS[name], op: /[^\p{L}\p{N}]/u.test(SYMBOLS[name]) };
    this.error(`\\${name} isn’t one this editor knows: try \\frac, \\sqrt, \\sum, \\alpha or ^ and _`);
    return null;
  }

  delimiter() {
    const c = this.peek();
    if (c === undefined) this.error('\\left and \\right need a bracket after them, such as ( or [');
    if (c === '\\') {
      const start = this.i;
      this.i++;
      const name = `\\${this.command()}`;
      if (DELIMS[name] === undefined) {
        this.i = start;
        this.error(`${name} isn’t a bracket \\left and \\right take`);
      }
      return DELIMS[name];
    }
    this.i++;
    if (DELIMS[c] === undefined) this.error(`${c} isn’t a bracket \\left and \\right take`);
    return DELIMS[c];
  }
}

const textOf = (n) => (n.t === 'run' ? n.text : n.e ? n.e.map(textOf).join('') : '');
const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const FONT = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';

function omml(nodes) {
  return nodes.map((n) => {
    switch (n.t) {
      case 'run':
        if (!n.text) return '';
        return `<m:r>${n.fn ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}${FONT}<m:t xml:space="preserve">${esc(n.text)}</m:t></m:r>`;
      case 'group': return omml(n.e);
      case 'frac': return `<m:f><m:num>${omml(n.num)}</m:num><m:den>${omml(n.den)}</m:den></m:f>`;
      case 'sup': return `<m:sSup><m:e>${omml(n.e)}</m:e><m:sup>${omml(n.sup)}</m:sup></m:sSup>`;
      case 'sub': return `<m:sSub><m:e>${omml(n.e)}</m:e><m:sub>${omml(n.sub)}</m:sub></m:sSub>`;
      case 'subsup': return `<m:sSubSup><m:e>${omml(n.e)}</m:e><m:sub>${omml(n.sub)}</m:sub><m:sup>${omml(n.sup)}</m:sup></m:sSubSup>`;
      case 'rad': return n.deg
        ? `<m:rad><m:deg>${omml(n.deg)}</m:deg><m:e>${omml(n.e)}</m:e></m:rad>`
        : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${omml(n.e)}</m:e></m:rad>`;
      case 'nary': return `<m:nary><m:naryPr><m:chr m:val="${n.chr}"/><m:limLoc m:val="${n.limLoc}"/>${n.sub ? '' : '<m:subHide m:val="1"/>'}${n.sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr>`
        + `<m:sub>${omml(n.sub || [])}</m:sub><m:sup>${omml(n.sup || [])}</m:sup><m:e>${omml(n.e)}</m:e></m:nary>`;
      case 'delim': {
        const pr = n.open === '(' && n.close === ')' ? '' : `<m:dPr><m:begChr m:val="${esc(n.open)}"/><m:endChr m:val="${esc(n.close)}"/></m:dPr>`;
        return `<m:d>${pr}<m:e>${omml(n.e)}</m:e></m:d>`;
      }
      default: return '';
    }
  }).join('');
}

/**
 * The equation typed as `src` (LaTeX), as OMML: { xml } (an m:oMath, or with `display` an
 * m:oMathPara holding one, for an equation on its own line), or { error } saying what is wrong.
 */
export function latexToOmml(src, { display = false } = {}) {
  const text = (src || '').trim().replace(/^\$+|\$+$/g, '').trim();
  if (!text) return { error: 'Type an equation, such as x^2 + \\frac{1}{2}' };
  try {
    const body = omml(new Parser(text).list());
    const math = `<m:oMath>${body}</m:oMath>`;
    return { xml: display ? `<m:oMathPara>${math}</m:oMathPara>` : math };
  } catch (e) {
    return { error: e.message };
  }
}

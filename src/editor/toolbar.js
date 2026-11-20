// Keyboard use of the editor's toolbar, as WAI-ARIA's toolbar pattern describes it: the toolbar is a
// single Tab stop (the control last used in it), ← and → move along its controls (wrapping around),
// Home and End go to the first and last; Esc goes back to the text. Controls that are disabled or
// not shown are skipped. A <select> keeps ↑ and ↓ (and Alt+↓) to choose its value.
// On a phone the toolbar is one row that scrolls sideways: data-more ("start", "end" or "both")
// says at which edge there are more controls than fit, for the page to fade that edge.

const CONTROLS = 'button, select, input';

/**
 * Makes `toolbar` keyboard-navigable. `onEscape()` is called on Esc (back to the text). Returns
 * { focus(), dispose() }: focus() puts the focus on the toolbar's current control (Alt+F10).
 */
export function rovingToolbar(toolbar, { onEscape } = {}) {
  let current = null;

  const all = () => [...toolbar.querySelectorAll(CONTROLS)];
  const usable = (el) => !el.matches(':disabled') && !el.hidden && el.getClientRects().length > 0;
  // On a phone (one row) the most used controls come first (CSS order): the arrows follow what is shown.
  const items = () => {
    const list = all().filter(usable);
    if (getComputedStyle(toolbar).flexWrap !== 'nowrap') return list;
    const at = new Map(list.map((el) => [el, el.getBoundingClientRect().left]));
    return list.sort((a, b) => at.get(a) - at.get(b));
  };

  /** One control takes Tab (the current one, or the first usable); the rest are reached by arrows. */
  function sync() {
    const list = items();
    if (!list.includes(current)) current = list[0] || null;
    for (const el of all()) {
      const index = el === current ? '0' : '-1';
      if (el.getAttribute('tabindex') !== index) el.setAttribute('tabindex', index);
    }
  }

  function move(to) {
    if (!to) return;
    current = to;
    sync();
    to.focus();
  }

  function onKeydown(e) {
    if (e.altKey || e.ctrlKey || e.metaKey) return;
    const list = items();
    const at = list.indexOf(e.target);
    if (at < 0) return;
    let to = null;
    if (e.key === 'ArrowRight') to = list[(at + 1) % list.length];
    else if (e.key === 'ArrowLeft') to = list[(at - 1 + list.length) % list.length];
    else if (e.key === 'Home') to = list[0];
    else if (e.key === 'End') to = list[list.length - 1];
    else if (e.key === 'Escape' && onEscape) {
      e.preventDefault();
      onEscape();
      return;
    } else return;
    e.preventDefault(); // a <select> would change its value on ← and → in some browsers
    move(to);
  }

  function onFocusin(e) {
    if (e.target.matches?.(CONTROLS) && e.target !== current) {
      current = e.target;
      sync();
    }
  }

  // Controls come and go and are enabled and disabled as the selection moves: keep one Tab stop.
  let frame = 0;
  const observer = new MutationObserver(() => {
    if (frame) return;
    frame = requestAnimationFrame(() => { frame = 0; sync(); });
  });
  observer.observe(toolbar, { subtree: true, childList: true, attributes: true, attributeFilter: ['disabled', 'hidden', 'class', 'style', 'data-active'] });
  toolbar.addEventListener('keydown', onKeydown);
  toolbar.addEventListener('focusin', onFocusin);
  sync();

  function edges() {
    const { scrollLeft: at, scrollWidth, clientWidth } = toolbar;
    const start = at > 1;
    const end = at + clientWidth < scrollWidth - 1;
    const more = start && end ? 'both' : start ? 'start' : end ? 'end' : null;
    if (more) toolbar.dataset.more = more;
    else delete toolbar.dataset.more;
  }
  const resized = new ResizeObserver(edges);
  resized.observe(toolbar);
  toolbar.addEventListener('scroll', edges, { passive: true });
  edges();

  return {
    focus() {
      sync();
      current?.focus();
    },
    dispose() {
      observer.disconnect();
      resized.disconnect();
      toolbar.removeEventListener('scroll', edges);
      cancelAnimationFrame(frame);
      toolbar.removeEventListener('keydown', onKeydown);
      toolbar.removeEventListener('focusin', onFocusin);
    },
  };
}

/**
 * Text labels for the toolbar's icon buttons (their `.short-label`), where there is room for them:
 * each in turn, along the toolbar, as long as it fits without the toolbar taking another row, so
 * the toolbar never grows. On a wide screen that is all of them, on a laptop some; on a phone (one
 * row that scrolls) none. Worked out again when the toolbar's width changes. Returns { refit, dispose }.
 */
export function fitLabels(toolbar) {
  const LABELLED = 'with-label';
  const candidates = () => [...toolbar.querySelectorAll('button')].filter((b) => b.querySelector(':scope > .short-label') && b.getClientRects().length);
  const rowCount = () => {
    const middles = [];
    for (const el of toolbar.querySelectorAll('button, select')) {
      const r = el.getBoundingClientRect();
      if (r.width) middles.push(r.top + r.height / 2);
    }
    // Controls of one row are centred alike, within a few pixels.
    return middles.sort((a, b) => a - b).filter((m, i, all) => i === 0 || m - all[i - 1] > 8).length;
  };
  let width = 0;
  let busy = false;

  function refit() {
    if (busy) return;
    busy = true;
    try {
      for (const b of toolbar.querySelectorAll(`.${LABELLED}`)) b.classList.remove(LABELLED);
      if (getComputedStyle(toolbar).flexWrap === 'nowrap') return; // a phone: one row, scrolling
      const rows = rowCount();
      for (const b of candidates()) {
        b.classList.add(LABELLED);
        if (rowCount() > rows) b.classList.remove(LABELLED); // no room for it: its icon alone
      }
    } finally {
      busy = false;
    }
  }

  const resized = new ResizeObserver(() => {
    if (toolbar.clientWidth === width) return; // only its width matters (its height follows the fit)
    width = toolbar.clientWidth;
    refit();
  });
  resized.observe(toolbar);
  return { refit, dispose: () => resized.disconnect() };
}

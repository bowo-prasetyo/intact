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
  const items = () => all().filter(usable);

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
  observer.observe(toolbar, { subtree: true, childList: true, attributes: true, attributeFilter: ['disabled', 'hidden', 'class', 'style'] });
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

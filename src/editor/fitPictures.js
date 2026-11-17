// Pictures in line with the text that are wider than their paragraph, on a narrow screen where the
// page is laid out to the screen's width rather than the paper's (see the editor's styles at
// 860px): shown smaller (CSS zoom) so the whole picture can be seen without scrolling sideways.
// Only how it shows changes: its size in the document stays as it is, and a picture resized here is
// resized in the document's terms (see pictureZoom).

const NARROW = '(max-width: 860px)';
const PICTURES = '.raw-inline.picture[data-wrap="inline"]';

/** How much a picture is shown smaller than its size (1: not at all). */
export function pictureZoom(span) {
  return Number(span?.style.zoom) || 1;
}

/** Fits the inline pictures in `root` to their paragraphs while `NARROW` applies. Returns a disposer. */
export function fitPictures(root) {
  const narrow = window.matchMedia(NARROW);
  let frame = 0;

  function fit() {
    frame = 0;
    for (const span of root.querySelectorAll(PICTURES)) {
      const width = parseFloat(span.style.width); // its own size, as laid out at 100%
      const block = span.parentElement?.closest('p, h1, h2, h3, h4, h5, h6, li, td, th') || span.parentElement;
      let zoom = 1;
      if (narrow.matches && width && block) {
        const cs = getComputedStyle(block);
        const room = block.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
        if (room > 0 && width > room) zoom = Math.floor((room / width) * 1000) / 1000;
      }
      const value = zoom < 1 ? String(zoom) : '';
      if (span.style.zoom !== value) span.style.zoom = value;
    }
  }
  const schedule = () => { frame ||= requestAnimationFrame(fit); };

  // Pictures are restyled (resized, cropped, laid out again) by replacing their style: fit again.
  const observer = new MutationObserver(schedule);
  observer.observe(root, { subtree: true, childList: true, attributes: true, attributeFilter: ['style', 'data-wrap'] });
  const resized = new ResizeObserver(schedule);
  resized.observe(root);
  narrow.addEventListener('change', schedule);
  schedule();

  return () => {
    observer.disconnect();
    resized.disconnect();
    narrow.removeEventListener('change', schedule);
    cancelAnimationFrame(frame);
  };
}

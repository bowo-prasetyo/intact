// Word's keyboard shortcuts, for what Intact can do. Cmd works for Ctrl on a Mac. Keys are matched
// by their position where the layout matters (digits, Space), so other keyboard layouts work too.

/** The command a key press asks for, or null. See SHORTCUTS for the list. */
export function shortcutFor(e) {
  if (!(e.ctrlKey || e.metaKey) || e.isComposing) return null;
  const key = e.key.toLowerCase();
  const space = e.code === 'Space' || key === ' ';
  const altGr = e.getModifierState?.('AltGraph');
  // Ctrl+= subscript, Ctrl+Shift+= (Ctrl++) superscript; Ctrl+Shift+8 (Ctrl+*, Cmd+8) formatting marks.
  if (!e.altKey && e.code === 'Equal') return e.shiftKey ? 'superscript' : 'subscript';
  if (!e.altKey && e.code === 'Digit8') return 'formattingMarks';
  // Ctrl+1, Ctrl+2, Ctrl+5: single, double and 1.5 line spacing; Ctrl+0: space before on or off.
  if (!e.altKey && !e.shiftKey) {
    const spacing = { Digit1: 'lineSingle', Digit2: 'lineDouble', Digit5: 'lineOneHalf', Digit0: 'spaceBefore' }[e.code];
    if (spacing) return spacing;
    // Ctrl+] and Ctrl+[: the font 1 pt bigger or smaller.
    if (e.code === 'BracketRight') return 'growFontPoint';
    if (e.code === 'BracketLeft') return 'shrinkFontPoint';
  }
  // Ctrl+Shift+> and Ctrl+Shift+<: Word's Grow Font and Shrink Font (the keys of > and <).
  if (e.shiftKey && !e.altKey) {
    if (e.code === 'Period') return 'growFont';
    if (e.code === 'Comma') return 'shrinkFont';
  }
  if (e.altKey && !e.shiftKey && !altGr && e.code === 'KeyH') return 'highlight';
  if (e.altKey && !e.shiftKey && !altGr && e.code === 'KeyM') return 'comment';
  if (e.altKey && !e.shiftKey && !altGr && e.code === 'KeyF') return 'footnote';
  if (!e.shiftKey && !e.altKey) {
    // Bold, italic and underline too: not every browser does them in editable text by itself.
    return {
      b: 'bold', i: 'italic', u: 'underline', e: 'center', l: 'left', r: 'right', j: 'justify', m: 'indent', enter: 'pageBreak', k: 'link',
    }[key] || (space ? 'clearFormatting' : null);
  }
  if (e.shiftKey && !e.altKey) {
    return { n: 'normal', l: 'bullets', m: 'outdent' }[key] || (space ? 'nonBreakingSpace' : null);
  }
  // Ctrl+Alt+digit; not AltGr (Ctrl+Alt on some layouts), which types characters.
  if (e.altKey && !e.shiftKey && /^Digit[123]$/.test(e.code) && !e.getModifierState?.('AltGraph')) return `heading${e.code.slice(-1)}`;
  return null;
}

/** The shortcuts, as shown on the About page. */
export const SHORTCUTS = [
  ['Ctrl+B, Ctrl+I, Ctrl+U', 'Bold, italic, underline'],
  ['Ctrl+Space', 'Clear character formatting (bold, italic, font, size, colour…); highlight stays'],
  ['Ctrl+Shift+=, Ctrl+=', 'Superscript, subscript'],
  ['Ctrl+Shift+>, Ctrl+Shift+<', 'Grow or shrink the font (Word\'s steps: 10, 11, 12, 14…)'],
  ['Ctrl+], Ctrl+[', 'Font 1 pt bigger or smaller'],
  ['Ctrl+Alt+H', 'Highlight (in the last colour chosen)'],
  ['Ctrl+Shift+8', 'Show or hide formatting marks (¶)'],
  ['Ctrl+Alt+M', 'New comment'],
  ['Ctrl+K', 'Insert a link, or change the one at the caret'],
  ['Tab in the last cell of a table', 'New row'],
  ['Ctrl+L, Ctrl+E, Ctrl+R, Ctrl+J', 'Align left, centre, right, justify'],
  ['Ctrl+Shift+N', 'Normal style'],
  ['Ctrl+Alt+1, 2, 3', 'Heading 1, 2, 3'],
  ['Ctrl+Shift+L', 'Bulleted list'],
  ['1. or * then a space, at the start of a paragraph', 'Numbered or bulleted list (Ctrl+Z types it as it is)'],
  ['Tab, Shift+Tab', 'List level down or up'],
  ['Ctrl+M, Ctrl+Shift+M', 'Indent or outdent (in a list: list level down or up)'],
  ['Ctrl+1, Ctrl+5, Ctrl+2', 'Single, 1.5 and double line spacing (where the browser leaves these keys to the page)'],
  ['Ctrl+0', 'Space before the paragraph on or off'],
  ['Ctrl+Enter', 'Page break'],
  ['Ctrl+Alt+F', 'Insert a footnote'],
  ['Shift+Enter', 'Line break'],
  ['Ctrl+Shift+Space', 'Non-breaking space'],
  ['Ctrl+Shift+E', 'Track Changes on or off'],
  ['Ctrl+F, Ctrl+H', 'Find, replace'],
  ['Ctrl+Z, Ctrl+Y', 'Undo, redo'],
  ['Ctrl+mouse wheel (or pinch) on the page', 'Zoom in or out (the buttons at the right of the status bar too)'],
  ['Ctrl+S', 'Review changes and download'],
  ['Alt+F10, then Esc', 'Go to the toolbar (arrow keys move along it), and back to the text'],
  ['Ctrl+/', 'Help: these shortcuts, and checks before you send the file'],
];

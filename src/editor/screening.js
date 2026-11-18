// A check of a CV before it goes to an employer, many of whom put applications through screening
// software (an applicant tracking system, ATS) that reads the file's text to fill in a profile and
// match keywords. Such software reads plain paragraphs well and often misses or mixes up the rest:
// text in headers and footers, text boxes and columns, cells of layout tables, pictures and charts.
// Like the accessibility check it reads the page as rendered, and each issue lists the elements it
// was found in, so the editor can show them one by one. It doesn't change anything.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';
const EMAIL = /[^\s@<>()"',;:]+@[^\s@<>()"',;:]+\.[a-z]{2,}/i;
const PHONE = /(?:\+|\b)\d[\d ().-]{7,}\d\b/;
const WEB = /\b(?:https?:\/\/|www\.)\S+|\blinkedin\.com\/\S+/i;
// Section headings of a CV, in a few languages: two of them (or the file's name) make it a CV.
const SECTIONS = new RegExp(`^(${[
  'experience', 'work experience', 'professional experience', 'employment', 'employment history', 'work history', 'career history',
  'education', 'qualifications', 'skills', 'key skills', 'core skills', 'technical skills', 'summary', 'professional summary',
  'profile', 'personal profile', 'objective', 'career objective', 'certifications', 'references', 'languages', 'achievements',
  'pengalaman', 'pengalaman kerja', 'pendidikan', 'keahlian', 'keterampilan', 'ringkasan',
  'berufserfahrung', 'ausbildung', 'kenntnisse', 'expérience', 'expérience professionnelle', 'formation', 'compétences',
  'experiencia', 'experiencia laboral', 'educación', 'formación', 'habilidades', 'experiência', 'educação', 'competências',
].join('|')})\\s*:?$`, 'i');
const CV_NAME = /\b(cv|resume|résumé|curriculum|lebenslauf)\b/i;
const GENERIC_NAME = /^(cv|resume|résumé|curriculum[ _-]?vitae|my[ _-]?cv|document\d*|untitled|new[ _-]?document|template)([ _-]*(final|new|v?\d+))*$/i;

const inBody = (el) => !el.closest('.docx-notes, .docx-hf-band');
const textOf = (el) => (el.innerText ?? el.textContent ?? '').trim();
const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;

/** Whether `page` (with the file's `name`, no extension) reads as a CV. */
export function looksLikeCv(page, name = '') {
  if (CV_NAME.test(name.replace(/[_-]+/g, ' '))) return true;
  const headings = new Set();
  for (const el of page.querySelectorAll(BLOCKS)) {
    if (!inBody(el) || el.closest('table')) continue;
    const t = textOf(el).replace(/\s+/g, ' ');
    if (t.length <= 40 && SECTIONS.test(t)) headings.add(t.toLowerCase().replace(/\s*:$/, ''));
    if (headings.size >= 2) return true;
  }
  return false;
}

/**
 * The screening issues in `page` (the editor's page): { cv, issues: [{ id, severity, title, text,
 * targets }] }, severity 'warning' or 'tip', `cv` whether the document reads as a CV. `name` is the
 * file's name without its extension.
 */
export function checkScreening(page, name = '') {
  const issues = [];
  const add = (id, severity, title, text, targets = []) => issues.push({ id: `ats-${id}`, severity, title, text, targets });
  const cv = looksLikeCv(page, name);
  const bodyText = [...page.children].filter((el) => !el.matches('.docx-notes, .docx-hf-band')).map(textOf).join('\n');
  const contact = (t) => EMAIL.test(t) || PHONE.test(t) || WEB.test(t);

  // Contact details only in a header or footer: some screening software skips them.
  const bands = [...page.querySelectorAll(':scope > .docx-hf-band .docx-hf-part')].filter((el) => contact(textOf(el)));
  const emailInBody = EMAIL.test(bodyText);
  if (bands.length) {
    const what = bands.map(textOf).join(' ');
    const missing = [EMAIL.test(what) && !emailInBody && 'e-mail address', PHONE.test(what) && !PHONE.test(bodyText) && 'phone number', WEB.test(what) && !WEB.test(bodyText) && 'web address'].filter(Boolean);
    if (missing.length) {
      add('contact', 'warning', 'Contact details in the header or footer', `Your ${missing.join(' and ')} ${missing.length === 1 ? 'is' : 'are'} only in the ${bands.some((el) => el.dataset.which === 'header') ? 'header' : 'footer'}: some screening software skips headers and footers, so the employer’s system may not find how to reach you. Put ${missing.length === 1 ? 'it' : 'them'} in the text at the top of the first page too.`, bands);
    }
  }
  if (cv && !emailInBody && !bands.some((el) => EMAIL.test(textOf(el)))) {
    add('email', 'warning', 'No e-mail address', 'No e-mail address was found in the text: screening software looks for one to fill in your contact details.');
  }

  // Text boxes and shapes: often read out of order, or not at all.
  const boxes = [...page.querySelectorAll('.textbox')].filter((el) => inBody(el) && textOf(el.querySelector(':scope > .textbox-content') || el));
  if (boxes.length) {
    add('textbox', 'warning', 'Text in text boxes', `${plural(boxes.length, 'text box', 'text boxes')} or ${boxes.length === 1 ? 'shape holds' : 'shapes hold'} text: screening software often reads ${boxes.length === 1 ? 'it' : 'them'} out of order, or not at all. Put what matters in the page’s own paragraphs.`, boxes);
  }

  // Columns: read across, two lists become one.
  const columns = [...page.querySelectorAll(':scope > .docx-columns')].filter((el) => Number(el.dataset.cols) > 1 && textOf(el));
  if (columns.length) {
    add('columns', 'warning', 'Text in columns', `${plural(columns.length, 'part')} of the document ${columns.length === 1 ? 'is' : 'are'} set in columns: screening software may read across them, mixing the lines of one column with the next. A single column reads reliably.`, columns);
  }

  // Tables laid out in columns: most software reads them, some read across the cells.
  const tables = [...page.querySelectorAll('table')].filter((t) => inBody(t) && !t.parentElement.closest('table')
    && [...t.rows].some((r) => [...r.cells].filter((c) => textOf(c)).length > 1));
  if (tables.length) {
    add('tables', 'tip', 'Text laid out in a table', `${plural(tables.length, 'table')} ${tables.length === 1 ? 'sets' : 'set'} text side by side: much screening software reads tables, but some read across the cells, joining a job title to the dates beside it. Check that each line still makes sense read row by row, or use tabs instead.`, tables);
  }

  // Charts, diagrams and other objects kept as they are, and pictures: no text to read.
  const objects = [...page.querySelectorAll('.raw-block')].filter(inBody);
  if (objects.length) {
    add('objects', 'warning', 'Charts or diagrams', `${plural(objects.length, 'chart, diagram or object', 'charts, diagrams or objects')}: screening software can’t read the text in ${objects.length === 1 ? 'it' : 'them'}. Say the same in words in the text.`, objects);
  }
  const pictures = [...page.querySelectorAll('.picture')].filter((el) => !el.closest('.docx-notes'));
  if (pictures.length) {
    add('pictures', 'tip', 'Pictures', `${plural(pictures.length, 'picture')} (a photo, logo, icons or skill bars): screening software reads no text in pictures. Make sure nothing important, such as a skill level or contact detail, is only in a picture.`, pictures);
  }

  // A file called "CV": the recruiter has many.
  if (cv && GENERIC_NAME.test(name.trim())) {
    add('name', 'tip', 'File name', `The file is called “${name}”: recruiters receive many of those. A name such as Firstname_Lastname_CV is easy to find again.`);
  }
  return { cv, issues };
}

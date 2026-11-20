// Starter documents for the start page: a letter, a memo, meeting minutes, a quote, an invoice, a CV,
// a report and a manuscript, each a Word document written as Word writes one, in Word's own styles (Title,
// Heading 1 and 2, Table Grid, List Paragraph, Footer), with placeholders such as [Your name] that
// Fill in fills in once for every place they appear. Nothing here is fetched: each is built in the
// browser from the blank document's parts (see blank.js).
import { blankDocxFiles, textWidth } from './blank.js';

const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A run: `text` with { b, i, color, sz (half-points), caps, u (underline) }; a tab for '\t'. */
function run(text, f = {}) {
  const props = `${f.b ? '<w:b/>' : ''}${f.i ? '<w:i/>' : ''}${f.caps ? '<w:caps/>' : ''}${f.color ? `<w:color w:val="${f.color}"/>` : ''}${f.sz ? `<w:sz w:val="${f.sz}"/><w:szCs w:val="${f.sz}"/>` : ''}${f.u ? '<w:u w:val="single"/>' : ''}`;
  const rPr = props ? `<w:rPr>${props}</w:rPr>` : '';
  return text.split('\t').map((part, i) => `${i ? `<w:r>${rPr}<w:tab/></w:r>` : ''}${part ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(part)}</w:t></w:r>` : ''}`).join('');
}

/**
 * A paragraph of `content` (text, or runs made with run()): { style, jc, before, after, line
 * (240ths of a line), tabs ([twips]), keepNext, border ('bottom'), bullet, ind (left twips), first
 * (first-line indent, twips: 0 takes a style's off) }.
 */
function p(content = '', o = {}) {
  const runs = typeof content === 'string' ? run(content, o.run) : content.join('');
  const pPr = [
    o.style ? `<w:pStyle w:val="${o.style}"/>` : '',
    o.keepNext ? '<w:keepNext/>' : '',
    o.bullet ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : '',
    o.border === 'thickThin' ? '<w:pBdr><w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="1" w:color="auto"/></w:pBdr>' : o.border ? '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="auto"/></w:pBdr>' : '',
    o.tabs ? `<w:tabs>${o.tabs.map((t) => (typeof t === 'number' ? `<w:tab w:val="left" w:pos="${t}"/>` : `<w:tab w:val="${t.align}" w:pos="${t.pos}"/>`)).join('')}</w:tabs>` : '',
    o.before != null || o.after != null || o.line != null
      ? `<w:spacing${o.before != null ? ` w:before="${o.before}"` : ''}${o.after != null ? ` w:after="${o.after}"` : ''}${o.line != null ? ` w:line="${o.line}" w:lineRule="auto"` : ''}/>` : '',
    o.ind != null || o.first != null || o.hanging != null
      ? `<w:ind${o.ind != null ? ` w:left="${o.ind}"` : ''}${o.hanging != null ? ` w:hanging="${o.hanging}"` : o.first != null ? ` w:firstLine="${o.first}"` : ''}/>` : '',
    o.jc ? `<w:jc w:val="${o.jc}"/>` : '',
  ].join('');
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
}

// A cell without borders (a layout table: signatures side by side), as Word's No Border writes it.
const NO_BORDERS = `<w:tcBorders>${['top', 'left', 'bottom', 'right'].map((side) => `<w:${side} w:val="nil"/>`).join('')}</w:tcBorders>`;

/**
 * A table across the page in Table Grid: `widths` (shares of the text width), `rows` (arrays of
 * cell texts), the first row a heading row (bold, shaded, repeated on each page) unless `heading`
 * is false, and cells in columns `right` aligned right (amounts).
 */
function table(paper, widths, rows, { heading = true, right = [], shade = 'D9E2F3', borders = true, center = false, width = null } = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const tw = widths.map((x) => Math.floor(((width ?? textWidth(paper)) * x) / total));
  const grid = tw.map((x) => `<w:gridCol w:w="${x}"/>`).join('');
  const body = rows.map((cells, r) => {
    const head = heading && r === 0;
    const trPr = head ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
    const tcs = cells.map((text, c) => {
      const shd = head ? `<w:shd w:val="clear" w:color="auto" w:fill="${shade}"/>` : '';
      const jc = right.includes(c) ? 'right' : center ? 'center' : null;
      const bold = head || /^\*\*.*\*\*$/.test(text);
      const t = text.replace(/^\*\*(.*)\*\*$/, '$1');
      // A cell of several lines (a signature: the role, room to sign, the name) is a paragraph each.
      const lines = t.split('\n').map((line) => p(line.replace(/^_(.*)_$/, '$1'), { jc, after: 0, run: { b: bold, u: /^_.*_$/.test(line) } }));
      return `<w:tc><w:tcPr><w:tcW w:w="${tw[c]}" w:type="dxa"/>${borders ? '' : NO_BORDERS}${shd}</w:tcPr>${lines.join('')}</w:tc>`;
    }).join('');
    return `<w:tr>${trPr}${tcs}</w:tr>`;
  }).join('');
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

const LABEL = 1701; // 3 cm: where a "Label:" line's value starts (a left tab stop)
const label = (name, value) => p([run(`${name}:`, { b: true }), run(`\t${value}`)], { tabs: [LABEL], after: 40 });

// Indonesian government documents (naskah dinas) as the rules for them (tata naskah dinas) lay them
// out: on A4, in Arial 12 pt, single-spaced, with the margins official documents commonly use (3 cm
// on the left, 2 cm elsewhere). Each starts with the office's letterhead (kop: its name in capitals,
// its unit, its address and contacts, and a thick-and-thin line under them), and ends with the
// signatory's position, name and NIP, at the right. A logo can go beside the letterhead with Insert
// picture. Labels such as "Nomor" line up at a tab stop, with the colon after it, as offices write
// them; the placeholders are in Indonesian, so Fill in asks for them in Indonesian.
const ID_MARGINS = { top: 1134, right: 1134, bottom: 1134, left: 1701 };
const ID_WIDTH = 11906 - ID_MARGINS.left - ID_MARGINS.right; // A4 between those margins
const ID_LABEL = 1701; // 3 cm: where a label's colon and value start
const ID_SIGN = 4820; // 8.5 cm: where the signature block starts, on the right half
const ID_SINGLE = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>';

/** The office's letterhead (kop surat): name, unit, address and contacts, a thick-and-thin line. */
const kop = () => [
  p('[NAMA INSTANSI]', { jc: 'center', run: { b: true, sz: 32 } }),
  p('[NAMA UNIT KERJA]', { jc: 'center', run: { b: true, sz: 28 } }),
  p('[Alamat kantor] Telepon [Nomor telepon]', { jc: 'center', run: { sz: 20 } }),
  p('Laman [Alamat laman] Pos-el [Alamat pos-el]', { jc: 'center', border: 'thickThin', after: 360, run: { sz: 20 } }),
];
/** "Nomor<tab>: [Nomor surat]": a label, its colon at the tab stop, the value after it. */
const idLabel = (name, value, o = {}) => p(`${name}\t: ${value}`, { tabs: [o.at || ID_LABEL], ...o });
/** The signature block, on the right: the position, room to sign, the name (underlined) and NIP. */
const signature = (before = 240) => [
  p('[Nama jabatan],', { ind: ID_SIGN, before }),
  p('', { ind: ID_SIGN, after: 1080 }),
  p('[Nama lengkap]', { ind: ID_SIGN, run: { b: true, u: true } }),
  p('NIP [NIP]', { ind: ID_SIGN }),
];
/** Tembusan (copies to), numbered as offices type them. */
const tembusan = () => [
  p('Tembusan:', { before: 360 }),
  p('1. [Tembusan pertama]'),
  p('2. [Tembusan kedua]'),
];

function naskahDinas() {
  const common = { lang: 'id', paper: 'A4', margins: ID_MARGINS, language: 'id-ID', font: { name: 'Arial', sz: 24 }, normal: ID_SINGLE };
  return [
    {
      id: 'surat-dinas',
      title: 'Surat dinas',
      description: 'Surat resmi instansi, dengan kop, nomor, hal, dan tanda tangan',
      file: 'Surat dinas',
      ...common,
      build: () => ({
        ...common,
        body: [
          ...kop(),
          p('[Kota], [Tanggal surat]', { jc: 'right', after: 240 }),
          idLabel('Nomor', '[Nomor surat]'),
          idLabel('Sifat', '[Biasa / Segera / Rahasia]'),
          idLabel('Lampiran', '[Jumlah lampiran]'),
          idLabel('Hal', '[Hal surat]', { after: 240 }),
          p('Yth. [Jabatan penerima]'),
          p('[Nama instansi penerima]'),
          p('di [Tempat]', { after: 240 }),
          p('[Paragraf pembuka: rujukan atau latar belakang surat.]', { jc: 'both', first: 709, after: 120 }),
          p('[Paragraf isi: hal yang disampaikan, diminta, atau diberitahukan.]', { jc: 'both', first: 709, after: 120 }),
          p('Demikian kami sampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.', { jc: 'both', first: 709 }),
          ...signature(),
          ...tembusan(),
        ].join(''),
      }),
    },
    {
      id: 'nota-dinas',
      title: 'Nota dinas',
      description: 'Naskah internal antarpejabat: yth., dari, hal, dan tanggal',
      file: 'Nota dinas',
      ...common,
      build: () => ({
        ...common,
        body: [
          ...kop(),
          p('NOTA DINAS', { jc: 'center', run: { b: true } }),
          p('NOMOR [Nomor nota dinas]', { jc: 'center', after: 240 }),
          idLabel('Yth.', '[Jabatan penerima]'),
          idLabel('Dari', '[Jabatan pengirim]'),
          idLabel('Tembusan', '[Jabatan tembusan]'),
          idLabel('Sifat', '[Biasa / Segera / Rahasia]'),
          idLabel('Lampiran', '[Jumlah lampiran]'),
          idLabel('Hal', '[Hal nota dinas]'),
          idLabel('Tanggal', '[Tanggal]'),
          p('', { border: 'bottom', after: 240 }),
          p('[Isi nota dinas: pokok persoalan, uraian, dan hal yang diminta atau diusulkan.]', { jc: 'both', first: 709, after: 120 }),
          p('Demikian disampaikan, mohon arahan lebih lanjut.', { jc: 'both', first: 709 }),
          ...signature(),
        ].join(''),
      }),
    },
    {
      id: 'surat-undangan',
      title: 'Surat undangan',
      description: 'Undangan rapat atau acara: hari, waktu, tempat, dan acara',
      file: 'Surat undangan',
      ...common,
      build: () => ({
        ...common,
        body: [
          ...kop(),
          p('[Kota], [Tanggal surat]', { jc: 'right', after: 240 }),
          idLabel('Nomor', '[Nomor surat]'),
          idLabel('Sifat', '[Biasa / Segera]'),
          idLabel('Lampiran', '[Jumlah lampiran]'),
          idLabel('Hal', 'Undangan [Nama kegiatan]', { after: 240 }),
          p('Yth. [Daftar undangan]'),
          p('di [Tempat]', { after: 240 }),
          p('Dalam rangka [maksud kegiatan], dengan hormat kami mengundang Bapak/Ibu untuk hadir pada:', { jc: 'both', first: 709, after: 120 }),
          idLabel('hari, tanggal', '[Hari, tanggal]', { ind: 709, at: 3119 }),
          idLabel('waktu', '[Pukul mulai] s.d. [Pukul selesai]', { ind: 709, at: 3119 }),
          idLabel('tempat', '[Tempat kegiatan]', { ind: 709, at: 3119 }),
          idLabel('acara', '[Acara]', { ind: 709, at: 3119, after: 120 }),
          p('Mengingat pentingnya acara tersebut, kami mohon Bapak/Ibu hadir tepat waktu. Atas perhatian Bapak/Ibu, kami ucapkan terima kasih.', { jc: 'both', first: 709 }),
          ...signature(),
          ...tembusan(),
        ].join(''),
      }),
    },
    {
      id: 'berita-acara',
      title: 'Berita acara',
      description: 'Catatan resmi suatu kejadian atau serah terima, ditandatangani kedua pihak',
      file: 'Berita acara',
      ...common,
      build: () => ({
        ...common,
        styles: ['tableGrid'],
        body: [
          ...kop(),
          p('BERITA ACARA [JUDUL BERITA ACARA]', { jc: 'center', run: { b: true } }),
          p('NOMOR [Nomor berita acara]', { jc: 'center', after: 240 }),
          p('Pada hari ini [Hari], tanggal [Tanggal dalam huruf], bulan [Bulan], tahun [Tahun dalam huruf], bertempat di [Tempat], kami yang bertanda tangan di bawah ini:', { jc: 'both', after: 120 }),
          idLabel('1. Nama', '[Nama pihak pertama]', { at: 2268 }),
          idLabel('NIP', '[NIP pihak pertama]', { ind: 284, at: 2268 }),
          idLabel('Jabatan', '[Jabatan pihak pertama]', { ind: 284, at: 2268 }),
          p('selanjutnya disebut PIHAK PERTAMA.', { ind: 284, after: 120 }),
          idLabel('2. Nama', '[Nama pihak kedua]', { at: 2268 }),
          idLabel('NIP', '[NIP pihak kedua]', { ind: 284, at: 2268 }),
          idLabel('Jabatan', '[Jabatan pihak kedua]', { ind: 284, at: 2268 }),
          p('selanjutnya disebut PIHAK KEDUA.', { ind: 284, after: 120 }),
          p('[Uraian: apa yang terjadi, diperiksa, atau diserahterimakan, dan dalam keadaan apa.]', { jc: 'both', first: 709, after: 120 }),
          p('Demikian berita acara ini dibuat dengan sebenarnya untuk dipergunakan sebagaimana mestinya.', { jc: 'both', first: 709, after: 360 }),
          idTable([1, 1], [[
            'PIHAK KEDUA,\n\n\n\n_[Nama pihak kedua]_\nNIP [NIP pihak kedua]',
            'PIHAK PERTAMA,\n\n\n\n_[Nama pihak pertama]_\nNIP [NIP pihak pertama]',
          ]]),
        ].join(''),
      }),
    },
  ];
}
/** A table across the text between the naskah dinas' margins, with no borders: signatures side by side. */
const idTable = (widths, rows) => table(null, widths, rows, { heading: false, borders: false, center: true, width: ID_WIDTH });

/**
 * The starters: { id, title, description, file, build(paper) → blankDocxFiles options, and for those
 * in another language than English its `lang` ('id') and `paper` }. Their placeholders name what to
 * type, in brackets, so Fill in lists them.
 */
export const STARTERS = [
  {
    id: 'letter',
    title: 'Letter',
    description: 'A formal letter, with your address and theirs',
    file: 'Letter',
    build: (paper) => ({
      body: [
        p('[Your name]', { after: 0, run: { b: true } }),
        p('[Your street address]', { after: 0 }),
        p('[Your city and postal code]', { after: 0 }),
        p('[Your phone] · [Your email]', { after: 360 }),
        p('[Date]', { after: 360 }),
        p('[Recipient name]', { after: 0 }),
        p('[Recipient title]', { after: 0 }),
        p('[Company name]', { after: 0 }),
        p('[Company address]', { after: 360 }),
        p([run('Subject: ', { b: true }), run('[Subject]', { b: true })], { after: 240 }),
        p('Dear [Recipient name],'),
        p('[Say why you are writing, in a sentence or two.]'),
        p('[Give the details: what happened, what you need, or what you offer.]'),
        p('[Say what happens next, or what you would like them to do.]'),
        p('Yours sincerely,', { after: 960 }),
        p('[Your name]', { after: 0 }),
        p('[Your job title]'),
      ].join(''),
    }),
  },
  {
    id: 'memo',
    title: 'Memo',
    description: 'A short internal note: to, from, date and subject',
    file: 'Memo',
    build: () => ({
      body: [
        p('Memo', { style: 'Title', after: 240 }),
        label('To', '[Recipient names]'),
        label('From', '[Your name]'),
        label('Date', '[Date]'),
        label('Subject', '[Subject]'),
        p('', { border: 'bottom', after: 240 }),
        p('[Start with the point of the memo: what the reader needs to know or do.]'),
        p('[Add the background and details the reader needs.]'),
        p('[End with what happens next, and by when.]'),
      ].join(''),
    }),
  },
  {
    id: 'minutes',
    title: 'Meeting minutes',
    description: 'Who came, what was decided, and who does what',
    file: 'Minutes',
    build: (paper) => ({
      styles: ['tableGrid'],
      body: [
        p('[Meeting name] minutes', { style: 'Title', after: 240 }),
        label('Date', '[Date and time]'),
        label('Place', '[Place]'),
        label('Chair', '[Chair name]'),
        label('Minutes', '[Note taker name]'),
        p('Attendance', { style: 'Heading1' }),
        p([run('Present: ', { b: true }), run('[Names of those present]')]),
        p([run('Apologies: ', { b: true }), run('[Names of those absent]')]),
        p('Agenda and decisions', { style: 'Heading1' }),
        p('1. [First agenda item]', { style: 'Heading2' }),
        p('[What was said about the first item.]'),
        p([run('Decision: ', { b: true }), run('[What was decided on the first item.]')]),
        p('2. [Second agenda item]', { style: 'Heading2' }),
        p('[What was said about the second item.]'),
        p([run('Decision: ', { b: true }), run('[What was decided on the second item.]')]),
        p('Actions', { style: 'Heading1', keepNext: true }),
        table(paper, [6, 3, 2], [['Action', 'Owner', 'Due'], ['', '', ''], ['', '', ''], ['', '', '']]),
        p('Next meeting', { style: 'Heading1', before: 360 }),
        p('[Date, time and place of the next meeting]'),
      ].join(''),
    }),
  },
  {
    id: 'quote',
    title: 'Quote',
    description: 'Prices for a client, with a total and terms',
    file: 'Quote',
    build: (paper) => ({
      styles: ['tableGrid'],
      body: [
        p('[Your business name]', { after: 0, run: { b: true, sz: 32 } }),
        p('[Your address] · [Your phone] · [Your email]', { after: 360, run: { color: '595959' } }),
        p('Quote', { style: 'Title', after: 240 }),
        label('Quote no.', '[Quote number]'),
        label('Date', '[Date]'),
        label('Valid until', '[Valid until]'),
        p([run('For: ', { b: true }), run('[Client name], [Client address]')], { before: 240, after: 240 }),
        table(paper, [6, 1.4, 2.2, 2.2], [
          ['Description', 'Qty', 'Unit price', 'Amount'],
          ['', '', '', ''], ['', '', '', ''], ['', '', '', ''],
          ['**Subtotal**', '', '', '[Subtotal]'],
          ['**Tax**', '', '', '[Tax]'],
          ['**Total**', '', '', '**[Total]**'],
        ], { right: [1, 2, 3] }),
        p('Terms', { style: 'Heading2', before: 360 }),
        p('[Payment terms, delivery time and what the price includes.]'),
        p('Thank you for considering us. To accept this quote, sign below and send it back to [Your email].', { before: 240 }),
        p('Accepted by: ______________________\tDate: ____________', { before: 480, tabs: [{ align: 'right', pos: textWidth(paper) }] }),
      ].join(''),
    }),
  },
  {
    id: 'invoice',
    title: 'Invoice',
    description: 'A bill to a client, with how to pay',
    file: 'Invoice',
    build: (paper) => ({
      styles: ['tableGrid'],
      body: [
        p('[Your business name]', { after: 0, run: { b: true, sz: 32 } }),
        p('[Your address] · [Your phone] · [Your email]', { after: 360, run: { color: '595959' } }),
        p('Invoice', { style: 'Title', after: 240 }),
        label('Invoice no.', '[Invoice number]'),
        label('Date', '[Date]'),
        label('Due date', '[Due date]'),
        p([run('Bill to: ', { b: true }), run('[Client name], [Client address]')], { before: 240, after: 240 }),
        table(paper, [6, 1.4, 2.2, 2.2], [
          ['Description', 'Qty', 'Unit price', 'Amount'],
          ['', '', '', ''], ['', '', '', ''], ['', '', '', ''],
          ['**Subtotal**', '', '', '[Subtotal]'],
          ['**Tax**', '', '', '[Tax]'],
          ['**Total due**', '', '', '**[Total]**'],
        ], { right: [1, 2, 3] }),
        p('How to pay', { style: 'Heading2', before: 360 }),
        label('Bank', '[Bank name]'),
        label('Account', '[Account name], [Account number]'),
        label('Reference', '[Invoice number]'),
        p('Thank you for your business.', { before: 240 }),
      ].join(''),
    }),
  },
  {
    id: 'cv',
    title: 'CV',
    description: 'Experience, education and skills, easy for screening software to read',
    file: 'CV',
    build: () => ({
      bullets: true,
      body: [
        p('[Your name]', { style: 'Title', after: 80 }),
        // Contact details in the text, not in a header: screening software reads them there.
        p('[Your email] · [Your phone] · [Your city] · [Your LinkedIn or website]', { after: 240, run: { color: '595959' } }),
        p('Profile', { style: 'Heading1' }),
        p('[Two or three sentences: what you do, your strengths, what you are looking for.]'),
        p('Experience', { style: 'Heading1' }),
        p('[Job title], [Company]', { style: 'Heading2' }),
        p('[Start month and year] – [End month and year, or Present]', { run: { i: true, color: '595959' } }),
        p('[An achievement, with a number if you can: what you did and what came of it]', { style: 'ListParagraph', bullet: true }),
        p('[Another achievement]', { style: 'ListParagraph', bullet: true }),
        p('[Previous job title], [Previous company]', { style: 'Heading2' }),
        p('[Earlier start month and year] – [Earlier end month and year]', { run: { i: true, color: '595959' } }),
        p('[An achievement in that job]', { style: 'ListParagraph', bullet: true }),
        p('Education', { style: 'Heading1' }),
        p('[Degree or qualification], [School or university]', { style: 'Heading2' }),
        p('[Year completed]', { run: { i: true, color: '595959' } }),
        p('Skills', { style: 'Heading1' }),
        p('[A skill, and the tools you use for it]', { style: 'ListParagraph', bullet: true }),
        p('[Languages you speak, and how well]', { style: 'ListParagraph', bullet: true }),
      ].join(''),
    }),
  },
  {
    id: 'report',
    title: 'Report',
    description: 'Headings for a summary, findings and recommendations, numbered pages',
    file: 'Report',
    build: () => ({
      bullets: true,
      footer: true,
      body: [
        p('[Report title]', { style: 'Title', after: 80 }),
        p('[Your name] · [Organisation] · [Date]', { after: 360, run: { color: '595959' } }),
        p('Summary', { style: 'Heading1' }),
        p('[In a paragraph: what the report is about, what you found, what you advise.]'),
        p('Background', { style: 'Heading1' }),
        p('[Why this report was written, and what it covers.]'),
        p('Findings', { style: 'Heading1' }),
        p('[First finding]', { style: 'Heading2' }),
        p('[The evidence for the first finding.]'),
        p('[Second finding]', { style: 'Heading2' }),
        p('[The evidence for the second finding.]'),
        p('Recommendations', { style: 'Heading1' }),
        p('[A recommendation: what should be done, by whom and by when]', { style: 'ListParagraph', bullet: true }),
        p('[Another recommendation]', { style: 'ListParagraph', bullet: true }),
        p('Conclusion', { style: 'Heading1' }),
        p('[What the reader should take away.]'),
      ].join(''),
    }),
  },
  {
    id: 'manuscript',
    title: 'Manuscript',
    description: 'A novel in standard manuscript format: title page, chapters, double-spaced',
    file: 'Manuscript',
    build: () => manuscript(), // defined below, with what it is made of
  },
  // Indonesian government documents (naskah dinas), in Indonesian, as the start page offers them
  // when the interface is in Indonesian: defined below.
  ...naskahDinas(),
];

// A novel's manuscript as agents and publishers ask for it (standard manuscript format): Times New
// Roman 12 pt, double-spaced, every paragraph's first line indented half an inch (Normal's, so
// paragraphs typed or pasted follow), no space between paragraphs, 2.54 cm margins. A title page
// with the author's contact details and the word count, the title halfway down; a header on the
// other pages with the author's surname, a short title and the page number; each chapter a Heading 1
// (in the Headings pane) that starts a new page a third of the way down, and a scene break as a
// centred #.
const SINGLE = { line: 240, first: 0, after: 0 };
const manuscript = () => {
  const pageField = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const chapter = (title) => p(title, { style: 'Heading1' });
  const text = (t) => p(t);
  return {
    font: { name: 'Times New Roman', sz: 24 },
    normal: '<w:spacing w:after="0" w:line="480" w:lineRule="auto"/><w:ind w:firstLine="720"/>',
    // A chapter starts a new page, a third of the way down, its title centred in the text's own size.
    heading1: {
      pPr: '<w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:spacing w:before="3600" w:after="480"/><w:ind w:firstLine="0"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/>',
      rPr: '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/>',
    },
    header: `<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr>${run('[Surname] / [Short title] / ')}${pageField}</w:p>`,
    titlePg: true,
    body: [
      // The word count at the top right, on a line of its own (Intact shows a right tab stop's text at
      // Word's default stops, so a line of its own, aligned right, looks the same in both).
      p('about [Word count, rounded] words', { ...SINGLE, jc: 'right' }),
      p('[Your legal name]', SINGLE),
      p('[Street address]', SINGLE),
      p('[City, region and postcode]', SINGLE),
      p('[Phone number]', SINGLE),
      p('[Email address]', SINGLE),
      p('[Novel title]', { ...SINGLE, before: 4320, after: 240, jc: 'center', run: { caps: true } }),
      p('by [Pen name]', { ...SINGLE, jc: 'center' }),
      chapter('Chapter One'),
      text('[The first paragraph of your story. Every paragraph starts half an inch in and is double-spaced, as the style Normal has it: just type, or paste your text.]'),
      text('[The next paragraph.]'),
      p('#', { first: 0, jc: 'center' }),
      text('[A new scene, after a scene break: a # on a line of its own, centred.]'),
      chapter('Chapter Two'),
      text('[Each chapter is a Heading 1: it starts a new page, and the Headings pane lists it.]'),
      p('THE END', { first: 0, jc: 'center', before: 480 }),
    ].join(''),
  };
};

// Countries that use Letter paper rather than A4 (the United States, Canada, Mexico, the
// Philippines and others in the Americas), by the region of the browser's language.
const LETTER_REGIONS = new Set(['US', 'CA', 'MX', 'PH', 'CL', 'CO', 'VE', 'GT', 'CR', 'PA', 'DO', 'PR', 'SV', 'NI', 'HN', 'BZ']);

/** The paper a new document is on where the browser is: Letter in the Americas that use it, else A4. */
export function localPaper(language = (typeof navigator !== 'undefined' && navigator.language) || 'en') {
  const region = /[-_]([A-Za-z]{2})\b/.exec(language)?.[1]?.toUpperCase();
  return region && LETTER_REGIONS.has(region) ? 'Letter' : 'A4';
}

/**
 * The files of starter `id` on `paper`: its parts as blankDocxFiles makes them, titled as the starter.
 * A starter with paper of its own (the naskah dinas, on A4 as their rules ask) keeps it.
 */
export function starterFiles(id, paper = localPaper()) {
  const starter = STARTERS.find((s) => s.id === id);
  if (!starter) throw new Error(`No starter called ${id}`);
  const on = starter.paper || paper;
  return blankDocxFiles({ ...starter.build(on), paper: on });
}

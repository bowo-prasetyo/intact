// Starter documents for the start page: a letter, a memo, meeting minutes, a quote, an invoice, a CV,
// a report, a board paper and a manuscript, each a Word document written as Word writes one, in Word's own styles (Title,
// Heading 1 and 2, Table Grid, List Paragraph, Footer), with placeholders such as [Your name] that
// Fill in fills in once for every place they appear. Nothing here is fetched: each is built in the
// browser from the blank document's parts (see blank.js).
import { blankDocxFiles, textWidth, sectionXml, listId, LIST_INDENT, HEADER_REF } from './blank.js';

const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A run: `text` with { b, i, color, sz (half-points), caps, u (underline) }; a tab for '\t'. */
function run(text, f = {}) {
  const props = `${f.b ? '<w:b/>' : ''}${f.i ? '<w:i/>' : ''}${f.caps ? '<w:caps/>' : ''}${f.color ? `<w:color w:val="${f.color}"/>` : ''}${f.sz ? `<w:sz w:val="${f.sz}"/><w:szCs w:val="${f.sz}"/>` : ''}${f.u ? '<w:u w:val="single"/>' : ''}`;
  const rPr = props ? `<w:rPr>${props}</w:rPr>` : '';
  return text.split('\t').map((part, i) => `${i ? `<w:r>${rPr}<w:tab/></w:r>` : ''}${part ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(part)}</w:t></w:r>` : ''}`).join('');
}

/**
 * A paragraph of `content` (text, or runs made with run()): { style, jc, before, after, line
 * (240ths of a line), tabs ([twips]), keepNext, border ('bottom'), bullet, num ([numId, level]: a
 * numbered list's item), ind (left twips), first (first-line indent, twips: 0 takes a style's off),
 * section (sectionXml's options: the paragraph ends a section, as Word keeps a section break) }.
 */
function p(content = '', o = {}) {
  const runs = typeof content === 'string' ? run(content, o.run) : content.join('');
  const pPr = [
    o.style ? `<w:pStyle w:val="${o.style}"/>` : '',
    o.keepNext ? '<w:keepNext/>' : '',
    o.bullet ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : '',
    o.num ? `<w:numPr><w:ilvl w:val="${o.num[1] || 0}"/><w:numId w:val="${o.num[0]}"/></w:numPr>` : '',
    o.border === 'thickThin' ? '<w:pBdr><w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="1" w:color="auto"/></w:pBdr>' : o.border ? '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="auto"/></w:pBdr>' : '',
    o.tabs ? `<w:tabs>${o.tabs.map((t) => (typeof t === 'number' ? `<w:tab w:val="left" w:pos="${t}"/>` : `<w:tab w:val="${t.align}" w:pos="${t.pos}"/>`)).join('')}</w:tabs>` : '',
    o.before != null || o.after != null || o.line != null
      ? `<w:spacing${o.before != null ? ` w:before="${o.before}"` : ''}${o.after != null ? ` w:after="${o.after}"` : ''}${o.line != null ? ` w:line="${o.line}" w:lineRule="auto"` : ''}/>` : '',
    o.ind != null || o.first != null || o.hanging != null
      ? `<w:ind${o.ind != null ? ` w:left="${o.ind}"` : ''}${o.hanging != null ? ` w:hanging="${o.hanging}"` : o.first != null ? ` w:firstLine="${o.first}"` : ''}/>` : '',
    o.jc ? `<w:jc w:val="${o.jc}"/>` : '',
    o.section ? `<w:sectPr>${sectionXml(o.section)}</w:sectPr>` : '',
  ].join('');
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
}

// A cell without borders (a layout table: signatures side by side), as Word's No Border writes it.
const NO_BORDERS = `<w:tcBorders>${['top', 'left', 'bottom', 'right'].map((side) => `<w:${side} w:val="nil"/>`).join('')}</w:tcBorders>`;

/**
 * A table across the page in Table Grid: `widths` (shares of the text width), `rows` (arrays of
 * cell texts), the first row a heading row (bold, shaded, repeated on each page) unless `heading`
 * is false, and cells in columns `right` aligned right (amounts); rows other than the heading at
 * least `rowHeight` twips high (room to write in a rubric's cells), but for a `totalRow` at the end.
 */
function table(paper, widths, rows, { heading = true, right = [], shade = 'D9E2F3', borders = true, center = false, width = null, rowHeight = 0, totalRow = false } = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const tw = widths.map((x) => Math.floor(((width ?? textWidth(paper)) * x) / total));
  const grid = tw.map((x) => `<w:gridCol w:w="${x}"/>`).join('');
  const body = rows.map((cells, r) => {
    const head = heading && r === 0;
    const tall = rowHeight && !(totalRow && r === rows.length - 1);
    const trPr = head ? '<w:trPr><w:tblHeader/></w:trPr>' : tall ? `<w:trPr><w:trHeight w:val="${rowHeight}"/></w:trPr>` : '';
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

// The other naskah dinas, each of more than a page now and then: their pages numbered from the
// second on, at the top in the middle between dashes (- 2 -), as the rules for official documents
// number them; the first page has none (a different first page, its header empty).
const PAGE_FIELD = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>- 2 -</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
const ID_PAGED = {
  header: `<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>${PAGE_FIELD}</w:p>`,
  titlePg: true,
  section: { numFmt: 'numberInDash' },
};
const ID_LABEL_COL = 2268; // 4 cm: a decree's labels (Menimbang, Mengingat, KESATU)
const ID_COLON_COL = 284; // 0.5 cm: their colons

/** A numbered point: an item of point list `n` (the nth of the starter's `lists`, see blank.js), at `level`. */
const point = (n, text, o = {}) => p(text, { num: [listId(n), o.level || 0], jc: 'both', after: 0, ...o });
/** Text under a point (at `level`), lined up with the point's text. */
const under = (text, level = 0, o = {}) => p(text, { ind: LIST_INDENT * (level + 1), jc: 'both', after: 120, ...o });

/**
 * A table without borders, as decrees lay out their considerations and articles: a label (Menimbang,
 * KESATU), a colon, and the text, its paragraphs (made with p() or point()) in `rows` as [label, [paragraphs]].
 */
function labelled(rows) {
  const widths = [ID_LABEL_COL, ID_COLON_COL, ID_WIDTH - ID_LABEL_COL - ID_COLON_COL];
  const cell = (w, inner) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${NO_BORDERS}</w:tcPr>${inner}</w:tc>`;
  const body = rows.map(([labelText, paras]) => `<w:tr>${cell(widths[0], p(labelText, { after: 0 }))}${cell(widths[1], p(':', { after: 0 }))}${cell(widths[2], paras.join(''))}</w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`
    + `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${body}</w:tbl>`;
}

/** Runs of `parts`: text, or [text, run options] (a placeholder in capitals: { caps: true }). */
const runs = (...parts) => parts.map((x) => (Array.isArray(x) ? run(x[0], x[1]) : run(x)));
/** Where and when an official document is set (Ditetapkan di / pada tanggal), above the signature. */
const setAt = (where = 'Ditetapkan di', when = 'pada tanggal') => [
  p(`${where} [Kota]`, { ind: ID_SIGN, before: 360 }),
  p(`${when} [Tanggal penetapan]`, { ind: ID_SIGN }),
];
/** The signature block of a decree: the deciding official's position (in capitals, as written), room to sign, the name and NIP. */
const decreeSignature = (before = 0) => [
  p(runs(['[Jabatan penetap]', { caps: true }], ','), { ind: ID_SIGN, before }),
  p('', { ind: ID_SIGN, after: 1080 }),
  p('[Nama lengkap]', { ind: ID_SIGN, run: { b: true, u: true } }),
  p('NIP [NIP]', { ind: ID_SIGN }),
];

/** A label and its value lined up at a tab stop `at`, indented by `ind`: Nama : […]. */
const party = (name, value, ind = LIST_INDENT) => idLabel(name, value, { ind, at: 2835, after: 0 });

function naskahDinas() {
  const common = { lang: 'id', paper: 'A4', margins: ID_MARGINS, language: 'id-ID', font: { name: 'Arial', sz: 24 }, normal: ID_SINGLE };
  const paged = { ...common, ...ID_PAGED, section: { ...ID_PAGED.section } };
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
    {
      // A decree (keputusan) and its attachment (lampiran) in a section of its own, on a new page:
      // the attachment's pages numbered from 1 again, and its paper turned in Page setup (This
      // section) for a wide table. The deciding official and the decree's title are typed once, in
      // ordinary case, and shown in capitals where a decree writes them so (w:caps).
      id: 'keputusan',
      title: 'Keputusan',
      description: 'Keputusan pejabat: menimbang, mengingat, memutuskan, diktum, dan lampiran',
      file: 'Keputusan',
      ...paged,
      build: () => ({
        ...paged,
        styles: ['tableGrid'],
        lists: ['lowerLetter', 'decimal', 'decimal'],
        section: { ...paged.section, numStart: 1 },
        body: [
          p(runs(['KEPUTUSAN ', { b: true }], ['[Jabatan penetap]', { b: true, caps: true }]), { jc: 'center' }),
          p('NOMOR [Nomor keputusan]', { jc: 'center' }),
          p('TENTANG', { jc: 'center' }),
          p(runs(['[Judul keputusan]', { b: true, caps: true }]), { jc: 'center', after: 240 }),
          p('DENGAN RAHMAT TUHAN YANG MAHA ESA', { jc: 'center', after: 240 }),
          p(runs(['[Jabatan penetap]', { caps: true }], ','), { jc: 'center', after: 240 }),
          labelled([
            ['Menimbang', [
              point(0, 'bahwa [pertimbangan yang menjadi alasan keputusan];'),
              point(0, 'bahwa berdasarkan pertimbangan sebagaimana dimaksud dalam huruf a, perlu menetapkan Keputusan [Jabatan penetap] tentang [Judul keputusan];', { after: 120 }),
            ]],
            ['Mengingat', [
              point(1, '[Undang-undang yang menjadi dasar];'),
              point(1, '[Peraturan yang menjadi dasar];'),
            ]],
          ]),
          p('MEMUTUSKAN:', { jc: 'center', before: 240, after: 240 }),
          labelled([
            ['Menetapkan', [p(runs(['Keputusan [Jabatan penetap] tentang [Judul keputusan]', { caps: true }], '.'), { jc: 'both', after: 120 })]],
            ['KESATU', [p('[Isi diktum kesatu].', { jc: 'both', after: 120 })]],
            ['KEDUA', [p('[Isi diktum kedua].', { jc: 'both', after: 120 })]],
            ['KETIGA', [p('Keputusan ini mulai berlaku pada tanggal ditetapkan.', { jc: 'both' })]],
          ]),
          ...setAt(),
          ...decreeSignature(240),
          p('Salinan Keputusan ini disampaikan kepada Yth.:', { before: 360 }),
          point(2, '[Penerima salinan pertama];', { jc: 'left' }),
          point(2, '[Penerima salinan kedua].', { jc: 'left' }),
          // The decree's section ends here, its pages numbered from its second; the attachment's
          // begins on a new page (the final section, numbered from 1 again).
          p('', { section: { paper: 'A4', margins: ID_MARGINS, numFmt: 'numberInDash', refs: HEADER_REF, titlePg: true } }),
          p('LAMPIRAN', { ind: ID_SIGN }),
          p(runs('KEPUTUSAN ', ['[Jabatan penetap]', { caps: true }]), { ind: ID_SIGN }),
          p('NOMOR [Nomor keputusan]', { ind: ID_SIGN }),
          p(runs('TENTANG ', ['[Judul keputusan]', { caps: true }]), { ind: ID_SIGN, after: 480 }),
          p(runs(['[Judul lampiran]', { b: true, caps: true }]), { jc: 'center', after: 240 }),
          table(null, [1, 5, 4, 3], [
            ['NO.', 'NAMA/NIP', 'JABATAN', 'KEDUDUKAN DALAM TIM'],
            ['1.', '[Nama anggota pertama]\nNIP [NIP anggota pertama]', '[Jabatan anggota pertama]', 'Ketua'],
            ['2.', '[Nama anggota kedua]\nNIP [NIP anggota kedua]', '[Jabatan anggota kedua]', 'Anggota'],
          ], { shade: 'F2F2F2', width: ID_WIDTH }),
          ...decreeSignature(480),
        ].join(''),
      }),
    },
    {
      id: 'surat-tugas',
      title: 'Surat tugas',
      description: 'Penugasan pegawai: dasar, kepada siapa, dan untuk apa',
      file: 'Surat tugas',
      ...paged,
      build: () => ({
        ...paged,
        styles: ['tableGrid'],
        lists: ['decimal', 'decimal'],
        body: [
          ...kop(),
          p('SURAT TUGAS', { jc: 'center', run: { b: true } }),
          p('NOMOR [Nomor surat tugas]', { jc: 'center', after: 240 }),
          labelled([
            ['Dasar', [
              point(0, '[Dasar penugasan: peraturan, surat, atau perintah];'),
              point(0, '[Dasar penugasan lain, bila ada].'),
            ]],
          ]),
          p('MEMBERI TUGAS', { jc: 'center', before: 240, after: 240, run: { b: true } }),
          labelled([
            ['Kepada', [
              idLabel('Nama', '[Nama pegawai]', { at: 2268, after: 0 }),
              idLabel('NIP', '[NIP pegawai]', { at: 2268, after: 0 }),
              idLabel('Pangkat/golongan', '[Pangkat/golongan]', { at: 2268, after: 0 }),
              idLabel('Jabatan', '[Jabatan pegawai]', { at: 2268, after: 120 }),
            ]],
            ['Untuk', [
              point(1, '[Tugas yang dilaksanakan, di mana, dan kapan];'),
              point(1, 'melaporkan hasil pelaksanaan tugas kepada [Jabatan pemberi tugas].'),
            ]],
          ]),
          p('[Kota], [Tanggal surat]', { ind: ID_SIGN, before: 480 }),
          ...signature(0),
        ].join(''),
      }),
    },
    {
      id: 'surat-edaran',
      title: 'Surat edaran',
      description: 'Pemberitahuan tentang hal yang perlu segera dilaksanakan',
      file: 'Surat edaran',
      ...paged,
      build: () => ({
        ...paged,
        lists: ['decimal'],
        body: [
          ...kop(),
          p(runs(['SURAT EDARAN ', { b: true }], ['[Jabatan penanda tangan]', { b: true, caps: true }]), { jc: 'center' }),
          p('NOMOR [Nomor surat edaran]', { jc: 'center' }),
          p('TENTANG', { jc: 'center' }),
          p(runs(['[Hal yang diedarkan]', { b: true, caps: true }]), { jc: 'center', after: 240 }),
          p('Yth. [Pejabat atau pihak yang dituju]', { after: 240 }),
          point(0, 'Latar Belakang', { run: { b: true } }),
          under('[Mengapa surat edaran ini diterbitkan.]'),
          point(0, 'Maksud dan Tujuan', { run: { b: true } }),
          under('[Maksud dan tujuan surat edaran ini.]'),
          point(0, 'Ruang Lingkup', { run: { b: true } }),
          under('[Kepada siapa dan untuk hal apa surat edaran ini berlaku.]'),
          point(0, 'Dasar', { run: { b: true } }),
          under('[Peraturan perundang-undangan yang menjadi dasar.]'),
          point(0, 'Isi Edaran', { run: { b: true } }),
          under('[Hal yang diminta atau diberitahukan, rinci dan jelas.]'),
          point(0, 'Penutup', { run: { b: true } }),
          under('Demikian surat edaran ini disampaikan untuk dilaksanakan dengan penuh tanggung jawab.'),
          ...setAt('Dikeluarkan di', 'pada tanggal'),
          ...signature(240),
        ].join(''),
      }),
    },
    {
      id: 'notula',
      title: 'Notula rapat',
      description: 'Catatan jalannya rapat: peserta, pembahasan, keputusan, dan tindak lanjut',
      file: 'Notula rapat',
      ...paged,
      build: () => ({
        ...paged,
        styles: ['tableGrid'],
        lists: ['decimal'],
        body: [
          p('NOTULA RAPAT', { jc: 'center', run: { b: true } }),
          p(runs(['[Nama rapat]', { b: true, caps: true }]), { jc: 'center', after: 240 }),
          idLabel('Hari, tanggal', '[Hari, tanggal]', { at: 2835 }),
          idLabel('Waktu', '[Pukul mulai] s.d. [Pukul selesai]', { at: 2835 }),
          idLabel('Tempat', '[Tempat rapat]', { at: 2835 }),
          idLabel('Pimpinan rapat', '[Nama pimpinan rapat]', { at: 2835 }),
          idLabel('Notulis', '[Nama notulis]', { at: 2835 }),
          idLabel('Peserta', '[Peserta rapat, atau lihat daftar hadir terlampir]', { at: 2835 }),
          idLabel('Acara', '[Acara rapat]', { at: 2835, after: 240 }),
          p('Jalannya rapat:', { run: { b: true }, after: 120 }),
          point(0, 'Pembukaan', { run: { b: true } }),
          under('[Siapa membuka rapat, dan pokok pengantarnya.]'),
          point(0, 'Pembahasan', { run: { b: true } }),
          under('[Pokok bahasan, pendapat yang disampaikan, dan oleh siapa.]'),
          point(0, 'Kesimpulan atau keputusan', { run: { b: true } }),
          under('[Apa yang disepakati atau diputuskan.]'),
          point(0, 'Tindak lanjut', { run: { b: true } }),
          under('[Siapa mengerjakan apa, dan kapan.]'),
          p('Rapat ditutup pada pukul [Pukul selesai].', { before: 120, after: 360 }),
          idTable([1, 1], [[
            'Pimpinan Rapat,\n\n\n\n_[Nama pimpinan rapat]_\nNIP [NIP pimpinan rapat]',
            'Notulis,\n\n\n\n_[Nama notulis]_\nNIP [NIP notulis]',
          ]]),
        ].join(''),
      }),
    },
    {
      id: 'laporan',
      title: 'Laporan kegiatan',
      description: 'Pendahuluan, kegiatan yang dilaksanakan, hasil, kesimpulan dan saran',
      file: 'Laporan kegiatan',
      ...paged,
      build: () => ({
        ...paged,
        lists: ['upperLetter'],
        body: [
          p('LAPORAN', { jc: 'center', run: { b: true } }),
          p(runs(['[Judul laporan]', { b: true, caps: true }]), { jc: 'center', after: 360 }),
          point(0, 'Pendahuluan', { run: { b: true } }),
          point(0, 'Umum', { level: 1, run: { b: true } }),
          under('[Gambaran umum tentang hal yang dilaporkan.]', 1),
          point(0, 'Maksud dan Tujuan', { level: 1, run: { b: true } }),
          under('[Maksud dan tujuan laporan ini.]', 1),
          point(0, 'Ruang Lingkup', { level: 1, run: { b: true } }),
          under('[Apa saja yang dilaporkan, dan batasnya.]', 1),
          point(0, 'Dasar', { level: 1, run: { b: true } }),
          under('[Peraturan, surat tugas, atau perintah yang menjadi dasar kegiatan.]', 1),
          point(0, 'Kegiatan yang Dilaksanakan', { run: { b: true } }),
          under('[Apa yang dilaksanakan, di mana, kapan, dan oleh siapa.]'),
          point(0, 'Hasil yang Dicapai', { run: { b: true } }),
          under('[Hasil kegiatan, dengan data pendukung.]'),
          point(0, 'Kesimpulan dan Saran', { run: { b: true } }),
          under('[Kesimpulan dari hasil kegiatan, dan saran tindak lanjut.]'),
          point(0, 'Penutup', { run: { b: true } }),
          under('Demikian laporan ini dibuat untuk dapat dipergunakan sebagaimana mestinya.'),
          p('[Kota], [Tanggal laporan]', { ind: ID_SIGN, before: 360 }),
          ...signature(0),
        ].join(''),
      }),
    },
    {
      id: 'telaahan-staf',
      title: 'Telaahan staf',
      description: 'Kajian untuk pimpinan: persoalan, fakta, analisis, dan saran',
      file: 'Telaahan staf',
      ...paged,
      build: () => ({
        ...paged,
        lists: ['upperLetter'],
        body: [
          p('TELAAHAN STAF', { jc: 'center', run: { b: true }, after: 240 }),
          idLabel('Kepada', 'Yth. [Jabatan penerima]'),
          idLabel('Dari', '[Jabatan pengirim]'),
          idLabel('Tanggal', '[Tanggal]'),
          idLabel('Hal', '[Hal yang ditelaah]'),
          p('', { border: 'bottom', after: 240 }),
          point(0, 'Persoalan', { run: { b: true } }),
          under('[Persoalan yang dihadapi, dirumuskan singkat dan jelas.]'),
          point(0, 'Praanggapan', { run: { b: true } }),
          under('[Dugaan yang beralasan tentang keadaan yang mungkin terjadi.]'),
          point(0, 'Fakta yang Mempengaruhi', { run: { b: true } }),
          under('[Fakta dan data yang berkaitan dengan persoalan.]'),
          point(0, 'Analisis', { run: { b: true } }),
          under('[Hubungan antara persoalan, fakta, dan praanggapan, dengan pilihan penyelesaiannya.]'),
          point(0, 'Kesimpulan', { run: { b: true } }),
          under('[Pilihan penyelesaian yang terbaik, dari analisis.]'),
          point(0, 'Saran', { run: { b: true } }),
          under('[Tindakan yang disarankan kepada pimpinan.]'),
          ...signature(360),
        ].join(''),
      }),
    },
    {
      id: 'surat-kuasa',
      title: 'Surat kuasa',
      description: 'Pelimpahan kewenangan dari pemberi kepada penerima kuasa',
      file: 'Surat kuasa',
      ...common,
      build: () => ({
        ...common,
        styles: ['tableGrid'],
        body: [
          ...kop(),
          p('SURAT KUASA', { jc: 'center', run: { b: true } }),
          p('NOMOR [Nomor surat kuasa]', { jc: 'center', after: 240 }),
          p('Yang bertanda tangan di bawah ini:', { after: 120 }),
          party('Nama', '[Nama pemberi kuasa]'),
          party('NIP', '[NIP pemberi kuasa]'),
          party('Jabatan', '[Jabatan pemberi kuasa]'),
          p('selanjutnya disebut Pemberi Kuasa.', { ind: LIST_INDENT, after: 120 }),
          p('Dengan ini memberi kuasa kepada:', { after: 120 }),
          party('Nama', '[Nama penerima kuasa]'),
          party('NIP', '[NIP penerima kuasa]'),
          party('Jabatan', '[Jabatan penerima kuasa]'),
          p('selanjutnya disebut Penerima Kuasa.', { ind: LIST_INDENT, after: 120 }),
          p('untuk [tindakan yang dikuasakan, dan batasnya].', { jc: 'both', after: 120 }),
          p('Demikian surat kuasa ini dibuat untuk dipergunakan sebagaimana mestinya.', { jc: 'both', after: 360 }),
          idTable([1, 1], [[
            '\nPenerima Kuasa,\n\n\n\n_[Nama penerima kuasa]_\nNIP [NIP penerima kuasa]',
            '[Kota], [Tanggal surat]\nPemberi Kuasa,\n\n\n\n_[Nama pemberi kuasa]_\nNIP [NIP pemberi kuasa]',
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
 * in another language than English its `lang` ('id') and `paper`; for teachers' ones, `group`
 * 'teach' (the start page shows them as a group of their own) }. Their placeholders name what to
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
    // A board paper (a decision paper), as boards and executive committees ask for one: what is
    // asked of them first, in a box, then the recommendation, a summary, the options weighed, the
    // money, the risks, who was consulted and what happens next; numbered pages, for the board pack.
    id: 'board-paper',
    title: 'Board paper',
    description: 'A decision paper: what is asked first, then options, costs, risks and next steps',
    file: 'Board paper',
    build: (paper) => ({
      styles: ['tableGrid'],
      bullets: true,
      footer: true,
      body: [
        p('[Paper title]', { style: 'Title', after: 80 }),
        p('[Organisation] · [Board or committee] · [Meeting date]', { after: 240, run: { color: '595959' } }),
        label('Paper no.', '[Paper number]'),
        label('From', '[Author name], [Author role]'),
        label('Sponsor', '[Executive sponsor]'),
        label('Purpose', '[For decision, for discussion or for information]'),
        label('Status', '[Confidential, or Public]'),
        p('', { border: 'bottom', after: 240 }),
        table(paper, [1], [
          ['Decision sought'],
          ['The [Board or committee] is asked to [approve, endorse or note] [what], [by when].'],
        ]),
        p('Recommendation', { style: 'Heading1', before: 360 }),
        p('That the [Board or committee] [approves the recommended option].', { style: 'ListParagraph', bullet: true }),
        p('That the [Board or committee] [notes the risks and how they are managed].', { style: 'ListParagraph', bullet: true }),
        p('Summary', { style: 'Heading1' }),
        p('[In three or four sentences: the issue, what you propose and why, what it costs, and what happens if nothing is done.]'),
        p('Background', { style: 'Heading1' }),
        p('[What the reader needs to know: how this came about, and any earlier decisions.]'),
        p('Options considered', { style: 'Heading1', keepNext: true }),
        table(paper, [3, 4, 4, 2], [
          ['Option', 'Benefits', 'Risks', 'Cost'],
          ['1. [Do nothing]', '[Benefits of doing nothing]', '[Risks of doing nothing]', '[Cost]'],
          ['2. [Second option]', '[Benefits]', '[Risks]', '[Cost]'],
          ['3. [Recommended option]', '[Benefits]', '[Risks]', '[Cost]'],
        ], { right: [3] }),
        p('[Why the recommended option is the best of these.]', { before: 120 }),
        p('Financial implications', { style: 'Heading1' }),
        p('[The cost, the budget it comes from, any savings or return, and over what period.]'),
        p('Risks', { style: 'Heading1', keepNext: true }),
        table(paper, [4, 2.2, 2.2, 4.6], [
          ['Risk', 'Likelihood', 'Impact', 'How it is managed'],
          ['[A risk of the recommended option]', '[Low/Med/High]', '[Low/Med/High]', '[What reduces it, and who watches it]'],
          ['', '', '', ''],
        ]),
        p('Consultation', { style: 'Heading1', before: 360 }),
        p('[Who was consulted (finance, legal, people, those affected) and what they said.]'),
        p('Next steps', { style: 'Heading1', keepNext: true }),
        table(paper, [6, 3, 2], [['Action', 'Owner', 'Due'], ['[First step once approved]', '[Owner]', '[Date]'], ['', '', '']]),
        p('Appendices', { style: 'Heading1', before: 360 }),
        p('[Appendix A: title]', { style: 'ListParagraph', bullet: true }),
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

// ------------------------------------------------------------------ for teachers

// A test paper's questions and answer choices, numbered by Word (1., 2. and A., B., C., D. under
// each: see blank.js's question lists), so a question or a choice added with Enter, or moved with
// Tab and Shift+Tab, is numbered as Word numbers it. `list` is the list's numId (2, 3…), each
// part of a test that numbers its questions from 1 having its own.
const question = (text, list = 2, o = {}) => p(text, { num: [list, 0], before: 160, after: 60, keepNext: true, ...o });
const choice = (text, list = 2, o = {}) => p(text, { num: [list, 1], after: 0, ...o });
/** A question with four answer choices; the last can end a section (`section`). */
const multipleChoice = (text, choices, list, section) => [
  question(text, list),
  ...choices.map((c, i) => choice(c, list, { keepNext: i < choices.length - 1, ...(i === choices.length - 1 && section ? { section } : {}) })),
];

/**
 * Ruled lines to write an answer on: a table of `n` rows, `width` twips wide, indented `indent`,
 * each cell with a line under it and no other borders (cell borders, as Word's Borders button
 * writes them), so Word and Intact draw the same lines.
 */
function answerLines(width, n = 3, indent = 397) {
  const none = (side) => `<w:${side} w:val="nil"/>`;
  const borders = `<w:tcBorders>${none('top')}${none('left')}<w:bottom w:val="single" w:sz="4" w:space="0" w:color="808080"/>${none('right')}</w:tcBorders>`;
  const w = width - indent;
  const row = `<w:tr><w:trPr><w:trHeight w:val="454"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${borders}</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:tc></w:tr>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="${w}" w:type="dxa"/><w:tblInd w:w="${indent}" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>`
    + `<w:tblGrid><w:gridCol w:w="${w}"/></w:tblGrid>${row.repeat(n)}</w:tbl>`;
}

/** A student's details in a table (Name, Class…), the labels bold: `rows` of [label, value] pairs. */
const details = (paper, widths, rows, o = {}) => table(paper, widths, rows.map((r) => r.map((c, i) => (i % 2 ? c : `**${c}**`))), { heading: false, ...o });

/** The English teachers' templates: a test, a worksheet, a rubric, a lesson plan and a letter to parents. */
function forTeachers() {
  return [
    {
      id: 'test',
      title: 'Test',
      description: 'Multiple choice in two columns, then questions with lines to answer on',
      file: 'Test',
      group: 'teach',
      build: (paper) => ({
        questions: 1,
        styles: ['tableGrid'],
        // The questions after Part A are in one column again, on the same page.
        section: { type: 'continuous' },
        body: [
          p('[School name]', { jc: 'center', after: 0, run: { b: true, sz: 28 } }),
          p('[Subject] test: [Topic]', { jc: 'center', after: 240, run: { b: true, sz: 32 } }),
          details(paper, [1.2, 4, 1.1, 2], [['Name', '', 'Class', ''], ['Date', '', 'Score', '      / [Total marks]']]),
          p([run('Instructions: ', { b: true }), run('Answer all the questions. You have [Time allowed]. Write clearly, in pen.')], { before: 240 }),
          p('Part A: Multiple choice', { style: 'Heading2' }),
          // The end of the first section (one column): Part A's questions are in two, on the same page.
          p('Circle the letter of the best answer.', { run: { i: true }, section: { paper } }),
          ...[1, 2, 3, 4].flatMap((n) => multipleChoice(`[Question ${n}]`, ['Answer A', 'Answer B', 'Answer C', 'Answer D'], 2,
            n === 4 ? { paper, type: 'continuous', cols: 2, sep: true } : null)),
          p('Part B: Short answers', { style: 'Heading2' }),
          p('Write your answers on the lines.', { run: { i: true } }),
          question('[Question 5]'),
          answerLines(textWidth(paper), 3),
          question('[Question 6]'),
          answerLines(textWidth(paper), 4),
          p('End of the test. Check your answers.', { jc: 'center', before: 360, run: { i: true } }),
        ].join(''),
      }),
    },
    {
      id: 'worksheet',
      title: 'Worksheet',
      description: 'A goal, activities to match, fill in and write, and a self-check',
      file: 'Worksheet',
      group: 'teach',
      build: (paper) => ({
        questions: 1,
        styles: ['tableGrid'],
        body: [
          p('[School name]', { after: 0, run: { color: '595959' } }),
          p('[Worksheet title]', { style: 'Title', after: 160 }),
          details(paper, [1, 3, 0.9, 1.6, 0.8, 1.6], [['Name', '', 'Class', '', 'Date', '']]),
          p([run('Learning goal: ', { b: true }), run('[What students will be able to do]')], { before: 240 }),
          p('1. Match each word to its meaning', { style: 'Heading2' }),
          p('Write the letter of the meaning next to each word.', { run: { i: true } }),
          table(paper, [3, 1, 6], [['Word', 'Letter', 'Meaning'], ['', '', 'A. '], ['', '', 'B. '], ['', '', 'C. '], ['', '', 'D. ']]),
          p('2. Fill in the blanks', { style: 'Heading2', before: 240 }),
          p([run('Use these words: ', { i: true }), run('[Words for the blanks]')]),
          question('[First sentence, with a blank] __________'),
          question('[Second sentence, with a blank] __________'),
          question('[Third sentence, with a blank] __________'),
          p('3. Think and write', { style: 'Heading2', before: 240 }),
          p('[An open question for students to answer in their own words]'),
          answerLines(textWidth(paper), 5, 0),
          p('How did I do?', { style: 'Heading2', before: 240 }),
          p('☐ I can do this on my own     ☐ I need a little help     ☐ I need more practice'),
        ].join(''),
      }),
    },
    {
      id: 'rubric',
      title: 'Rubric',
      description: 'Criteria against four levels, with a score for each, on a landscape page',
      file: 'Rubric',
      group: 'teach',
      build: (paper) => ({
        styles: ['tableGrid'],
        section: { landscape: true },
        body: [
          p('[Assignment name] rubric', { style: 'Title', after: 160 }),
          label('Class', '[Class]'),
          label('Teacher', '[Teacher name]'),
          label('Student', ''),
          p('', { after: 120 }),
          table(null, [2.4, 2.2, 2.2, 2.2, 2.2, 1], [
            ['Criteria', 'Excellent (4)', 'Good (3)', 'Fair (2)', 'Needs work (1)', 'Score'],
            ['**[First criterion]**', '', '', '', '', ''],
            ['**[Second criterion]**', '', '', '', '', ''],
            ['**[Third criterion]**', '', '', '', '', ''],
            ['**[Fourth criterion]**', '', '', '', '', ''],
            ['**Total**', '', '', '', '', '/ 16'],
          ], { width: textWidth(paper, true), right: [5], rowHeight: 1134, totalRow: true }),
          p('Comments', { style: 'Heading2', before: 240 }),
          answerLines(textWidth(paper, true), 3, 0),
        ].join(''),
      }),
    },
    {
      id: 'lesson-plan',
      title: 'Lesson plan',
      description: 'Objectives, materials, the lesson’s stages, support and assessment',
      file: 'Lesson plan',
      group: 'teach',
      build: (paper) => ({
        bullets: true,
        styles: ['tableGrid'],
        body: [
          p('Lesson plan', { style: 'Title', after: 160 }),
          details(paper, [1, 2.4, 1, 2.4], [['Subject', '[Subject]', 'Class', '[Class]'], ['Topic', '[Topic]', 'Date', '[Date]'], ['Teacher', '[Teacher name]', 'Duration', '[Duration]']]),
          p('Learning objectives', { style: 'Heading1' }),
          p('By the end of the lesson, students will be able to:'),
          p('[First objective]', { style: 'ListParagraph', bullet: true }),
          p('[Second objective]', { style: 'ListParagraph', bullet: true }),
          p('Materials', { style: 'Heading1' }),
          p('[Books, equipment, worksheets and links the lesson needs]', { style: 'ListParagraph', bullet: true }),
          p('Lesson', { style: 'Heading1', keepNext: true }),
          table(paper, [1.6, 5.4, 1.1], [
            ['Stage', 'What the teacher and students do', 'Time'],
            ['**Starter**', '[How the lesson begins: a question, a recap or a hook]', ''],
            ['**Main activity**', '[What students do, alone, in pairs or in groups]', ''],
            ['**Plenary**', '[How the lesson ends: what was learned, and what comes next]', ''],
          ]),
          p('Differentiation', { style: 'Heading1', before: 360 }),
          p([run('Support: ', { b: true }), run('[How you will help students who need it]')], { style: 'ListParagraph', bullet: true }),
          p([run('Challenge: ', { b: true }), run('[How you will stretch students who are ready]')], { style: 'ListParagraph', bullet: true }),
          p('Assessment', { style: 'Heading1' }),
          p('[How you will check what students have learned]'),
          p('Reflection', { style: 'Heading1' }),
          p('[After the lesson: what went well, and what to change next time]'),
        ].join(''),
      }),
    },
    {
      id: 'parent-letter',
      title: 'Letter to parents',
      description: 'To each student’s family, ready for mail merge from a class list',
      file: 'Letter to parents',
      group: 'teach',
      build: () => ({
        body: [
          p('[School name]', { after: 0, run: { b: true, sz: 32 } }),
          p('[School address] · [School phone] · [School email]', { after: 360, run: { color: '595959' } }),
          p('[Date]', { after: 360 }),
          // [Parent name], [Student name] and [Class] are the columns of a class list, for mail merge.
          p('Dear [Parent name],'),
          p('I am writing to you about [Student name], in [Class].'),
          p('[What you are writing about: an event, a trip, progress or a concern, and what it means for your child.]'),
          p('[What you would like the family to do, and by when.]'),
          p('If you have any questions, please contact me at [Teacher email].'),
          p('Yours sincerely,', { after: 960 }),
          p('[Teacher name]', { after: 0 }),
          p('[Class] teacher', { after: 480 }),
          p('', { border: 'bottom', after: 240 }),
          p('Reply slip', { style: 'Heading2' }),
          p('Please return this slip to school by [Reply date].'),
          p('Student: [Student name], [Class]'),
          p('Parent’s name: ______________________     Signature: ______________________'),
        ].join(''),
      }),
    },
  ];
}

// Teachers in Indonesia (untuk guru): a test paper (lembar soal), a test blueprint (kisi-kisi), a
// lesson module (modul ajar, as Kurikulum Merdeka sets it out), a rubric and a letter to parents,
// in Indonesian, on A4 in Times New Roman 12 pt with 2 cm margins (the letter as an official
// letter: Arial, 3 cm on the left, see naskahDinas), each headed by the school's letterhead (kop).
const SCHOOL_MARGINS = { top: 1134, right: 1134, bottom: 1134, left: 1134 };
const SCHOOL_WIDTH = 11906 - 2 * 1134; // A4 between those margins
const SCHOOL_WIDE = 16838 - 2 * 1134; // A4 landscape
const SCHOOL = { paper: 'A4', margins: SCHOOL_MARGINS };

/** The school's letterhead (kop sekolah): its name, address and NPSN, over a thick-and-thin line. */
const kopSekolah = () => [
  p('[NAMA SEKOLAH]', { jc: 'center', after: 0, run: { b: true, sz: 32 } }),
  p('[Alamat sekolah] Telepon [Nomor telepon sekolah]', { jc: 'center', after: 0, run: { sz: 20 } }),
  p('NPSN [NPSN] · Pos-el [Alamat pos-el sekolah]', { jc: 'center', border: 'thickThin', after: 240, run: { sz: 20 } }),
];
/** "Mata pelajaran<tab>: [Mata pelajaran]", lined up at `at`. */
const schoolLabel = (name, value, at = 2835, o = {}) => p(`${name}\t: ${value}`, { tabs: [at], after: 0, ...o });
/** The head of school's and the teacher's signatures, side by side, without borders. */
// Each side's lines in step, so the names line up.
const schoolSignatures = (width, right = '[Kota], [Tanggal]\nGuru mata pelajaran,\n\n\n\n_[Nama guru]_\nNIP [NIP guru]') => table(null, [1, 1], [[
  'Mengetahui,\nKepala [Nama sekolah],\n\n\n\n_[Nama kepala sekolah]_\nNIP [NIP kepala sekolah]',
  right,
]], { heading: false, borders: false, center: true, width });

function untukGuru() {
  const common = { lang: 'id', group: 'teach', paper: 'A4', language: 'id-ID' };
  const school = { ...common, margins: SCHOOL_MARGINS, font: { name: 'Times New Roman', sz: 24 }, normal: '<w:spacing w:after="0" w:line="264" w:lineRule="auto"/>' };
  return [
    {
      id: 'soal',
      title: 'Lembar soal',
      description: 'Pilihan ganda A–D dalam dua kolom, isian, dan uraian dengan garis jawaban',
      file: 'Lembar soal',
      ...school,
      build: () => ({
        ...school,
        questions: 3, // pilihan ganda, isian, uraian: each numbered from 1
        styles: ['tableGrid'],
        section: { type: 'continuous' },
        body: [
          ...kopSekolah(),
          p('PENILAIAN SUMATIF [AKHIR SEMESTER GANJIL/GENAP]', { jc: 'center', run: { b: true } }),
          p('TAHUN PELAJARAN [TAHUN PELAJARAN]', { jc: 'center', after: 240, run: { b: true } }),
          table(null, [1.4, 2.6, 1.1, 2], [
            ['Mata pelajaran', ': [Mata pelajaran]', 'Nama', ': '],
            ['Kelas/Fase', ': [Kelas/Fase]', 'No. absen', ': '],
            ['Hari, tanggal', ': [Hari, tanggal ujian]', 'Nilai', ': '],
            ['Waktu', ': [Waktu pengerjaan]', 'Paraf guru', ': '],
          ], { heading: false, borders: false, width: SCHOOL_WIDTH }),
          p('Petunjuk umum', { before: 240, run: { b: true } }),
          ...['Berdoalah sebelum mengerjakan soal.', 'Tulislah nama dan nomor absen pada tempat yang tersedia.', 'Bacalah setiap soal dengan teliti sebelum menjawab.', 'Periksalah kembali jawaban sebelum diserahkan kepada pengawas.']
            .map((t, i) => p(`${i + 1}.\t${t}`, { ind: 397, hanging: 397, tabs: [397] })),
          p('I. Pilihan ganda', { before: 240, keepNext: true, run: { b: true } }),
          // The end of the first section (one column): the multiple choice is in two, on the same page.
          p('Berilah tanda silang (X) pada huruf A, B, C, atau D di depan jawaban yang paling tepat!', { run: { i: true }, section: SCHOOL }),
          ...[1, 2, 3, 4].flatMap((n) => multipleChoice(`[Soal pilihan ganda ${n}]`, ['Jawaban A', 'Jawaban B', 'Jawaban C', 'Jawaban D'], 2,
            n === 4 ? { ...SCHOOL, type: 'continuous', cols: 2, sep: true } : null)),
          p('II. Isian', { before: 240, keepNext: true, run: { b: true } }),
          p('Isilah titik-titik di bawah ini dengan jawaban yang tepat!', { run: { i: true } }),
          question('[Soal isian 1] ....................', 3),
          question('[Soal isian 2] ....................', 3),
          p('III. Uraian', { before: 240, keepNext: true, run: { b: true } }),
          p('Jawablah pertanyaan-pertanyaan berikut dengan jelas dan tepat!', { run: { i: true } }),
          question('[Soal uraian 1]', 4),
          answerLines(SCHOOL_WIDTH, 3),
          question('[Soal uraian 2]', 4),
          answerLines(SCHOOL_WIDTH, 3),
          p('Selamat mengerjakan', { jc: 'center', before: 360, run: { i: true } }),
        ].join(''),
      }),
    },
    {
      id: 'kisi-kisi',
      title: 'Kisi-kisi soal',
      description: 'Tujuan pembelajaran, indikator, level kognitif, dan nomor soal, pada halaman mendatar',
      file: 'Kisi-kisi soal',
      ...school,
      build: () => ({
        ...school,
        styles: ['tableGrid'],
        section: { landscape: true },
        body: [
          p('KISI-KISI PENILAIAN SUMATIF [AKHIR SEMESTER GANJIL/GENAP]', { jc: 'center', run: { b: true } }),
          p('TAHUN PELAJARAN [TAHUN PELAJARAN]', { jc: 'center', after: 240, run: { b: true } }),
          schoolLabel('Satuan pendidikan', '[Nama sekolah]', 3402),
          schoolLabel('Mata pelajaran', '[Mata pelajaran]', 3402),
          schoolLabel('Kelas/Fase/Semester', '[Kelas/Fase] / [Semester]', 3402),
          schoolLabel('Kurikulum', 'Kurikulum Merdeka', 3402),
          schoolLabel('Jumlah dan bentuk soal', '[Jumlah soal] soal; [Bentuk soal]', 3402, { after: 240 }),
          table(null, [0.6, 3, 3, 2, 3.4, 1.3, 1.3, 1.1], [
            ['No.', 'Capaian pembelajaran', 'Tujuan pembelajaran', 'Materi', 'Indikator soal', 'Level kognitif', 'Bentuk soal', 'Nomor soal'],
            ['1.', '', '', '', '', '', '', ''],
            ['2.', '', '', '', '', '', '', ''],
            ['3.', '', '', '', '', '', '', ''],
          ], { width: SCHOOL_WIDE, rowHeight: 1134 }),
          p('', { after: 240 }),
          schoolSignatures(SCHOOL_WIDE),
        ].join(''),
      }),
    },
    {
      id: 'modul-ajar',
      title: 'Modul ajar',
      description: 'Informasi umum, komponen inti, dan lampiran, sesuai Kurikulum Merdeka',
      file: 'Modul ajar',
      ...school,
      build: () => ({
        ...school,
        bullets: true,
        styles: ['tableGrid'],
        body: [
          p('MODUL AJAR', { jc: 'center', after: 0, run: { b: true, sz: 28 } }),
          p('[MATA PELAJARAN]: [TOPIK]', { jc: 'center', after: 240, run: { b: true } }),
          p('A. Informasi umum', { style: 'Heading1' }),
          p('1. Identitas modul', { style: 'Heading2' }),
          details(null, [1, 2], [['Penyusun', '[Nama guru]'], ['Satuan pendidikan', '[Nama sekolah]'], ['Fase/Kelas', '[Fase/Kelas]'], ['Mata pelajaran', '[Mata pelajaran]'], ['Alokasi waktu', '[Jumlah JP] JP'], ['Tahun pelajaran', '[Tahun pelajaran]']], { width: SCHOOL_WIDTH }),
          p('2. Kompetensi awal', { style: 'Heading2' }),
          p('[Pengetahuan dan keterampilan yang perlu dimiliki peserta didik sebelum mempelajari topik ini.]'),
          p('3. Dimensi profil lulusan', { style: 'Heading2' }),
          p('[Dimensi yang dikembangkan, misalnya penalaran kritis, kreativitas, kolaborasi, atau komunikasi.]'),
          p('4. Sarana dan prasarana', { style: 'Heading2' }),
          p('[Media, alat, bahan, dan sumber belajar.]'),
          p('5. Target peserta didik', { style: 'Heading2' }),
          p('[Peserta didik reguler, yang mengalami kesulitan belajar, atau yang berprestasi tinggi.]'),
          p('6. Model pembelajaran', { style: 'Heading2' }),
          p('[Misalnya pembelajaran berbasis masalah, berbasis proyek, atau inkuiri.]'),
          p('B. Komponen inti', { style: 'Heading1' }),
          p('1. Tujuan pembelajaran', { style: 'Heading2' }),
          p('[Tujuan pembelajaran pertama]', { style: 'ListParagraph', bullet: true }),
          p('[Tujuan pembelajaran kedua]', { style: 'ListParagraph', bullet: true }),
          p('2. Pemahaman bermakna', { style: 'Heading2' }),
          p('[Manfaat pembelajaran ini bagi peserta didik dalam kehidupan sehari-hari.]'),
          p('3. Pertanyaan pemantik', { style: 'Heading2' }),
          p('[Pertanyaan yang memantik rasa ingin tahu peserta didik]', { style: 'ListParagraph', bullet: true }),
          p('4. Kegiatan pembelajaran', { style: 'Heading2', keepNext: true }),
          table(null, [1.4, 5, 1.1], [
            ['Tahap', 'Kegiatan', 'Waktu'],
            ['**Pendahuluan**', '[Salam, doa, apersepsi, dan penyampaian tujuan pembelajaran]', ''],
            ['**Inti**', '[Kegiatan peserta didik sesuai model pembelajaran]', ''],
            ['**Penutup**', '[Refleksi, kesimpulan, dan rencana pertemuan berikutnya]', ''],
          ], { width: SCHOOL_WIDTH }),
          p('5. Asesmen', { style: 'Heading2', before: 240 }),
          p([run('Diagnostik: ', { b: true }), run('[Asesmen diagnostik]')], { style: 'ListParagraph', bullet: true }),
          p([run('Formatif: ', { b: true }), run('[Asesmen formatif]')], { style: 'ListParagraph', bullet: true }),
          p([run('Sumatif: ', { b: true }), run('[Asesmen sumatif]')], { style: 'ListParagraph', bullet: true }),
          p('6. Pengayaan dan remedial', { style: 'Heading2' }),
          p([run('Pengayaan: ', { b: true }), run('[Kegiatan pengayaan]')], { style: 'ListParagraph', bullet: true }),
          p([run('Remedial: ', { b: true }), run('[Kegiatan remedial]')], { style: 'ListParagraph', bullet: true }),
          p('7. Refleksi guru dan peserta didik', { style: 'Heading2' }),
          p('[Pertanyaan refleksi untuk guru dan peserta didik.]'),
          p('C. Lampiran', { style: 'Heading1' }),
          ...['Lembar kerja peserta didik (LKPD)', 'Bahan bacaan guru dan peserta didik', 'Glosarium', 'Daftar pustaka'].map((t) => p(t, { style: 'ListParagraph', bullet: true })),
          p('', { after: 240 }),
          schoolSignatures(SCHOOL_WIDTH),
        ].join(''),
      }),
    },
    {
      id: 'rubrik',
      title: 'Rubrik penilaian',
      description: 'Kriteria dengan empat tingkat capaian dan skor, pada halaman mendatar',
      file: 'Rubrik penilaian',
      ...school,
      build: () => ({
        ...school,
        styles: ['tableGrid'],
        section: { landscape: true },
        body: [
          p('RUBRIK PENILAIAN', { jc: 'center', after: 0, run: { b: true, sz: 28 } }),
          p('[Nama tugas atau proyek]', { jc: 'center', after: 240, run: { b: true } }),
          schoolLabel('Mata pelajaran', '[Mata pelajaran]'),
          schoolLabel('Kelas', '[Kelas]'),
          schoolLabel('Nama peserta didik', '', 2835, { after: 240 }),
          table(null, [2.4, 2.2, 2.2, 2.2, 2.2, 1], [
            ['Kriteria', 'Sangat baik (4)', 'Baik (3)', 'Cukup (2)', 'Perlu bimbingan (1)', 'Skor'],
            ['**[Kriteria 1]**', '', '', '', '', ''],
            ['**[Kriteria 2]**', '', '', '', '', ''],
            ['**[Kriteria 3]**', '', '', '', '', ''],
            ['**[Kriteria 4]**', '', '', '', '', ''],
            ['**Jumlah skor**', '', '', '', '', ''],
          ], { width: SCHOOL_WIDE, right: [5], rowHeight: 1134, totalRow: true }),
          p('Nilai = (jumlah skor ÷ 16) × 100', { before: 240 }),
          p('Catatan guru', { before: 240, run: { b: true } }),
          answerLines(SCHOOL_WIDE, 3, 0),
        ].join(''),
      }),
    },
    {
      id: 'surat-orang-tua',
      title: 'Surat untuk orang tua',
      description: 'Pemberitahuan kepada orang tua/wali tiap siswa, siap untuk gabung surat dari daftar kelas',
      file: 'Surat untuk orang tua',
      ...common,
      margins: ID_MARGINS,
      font: { name: 'Arial', sz: 24 },
      normal: ID_SINGLE,
      build: () => ({
        ...common,
        margins: ID_MARGINS,
        font: { name: 'Arial', sz: 24 },
        normal: ID_SINGLE,
        styles: ['tableGrid'], // the signatures' table
        body: [
          ...kopSekolah(),
          p('[Kota], [Tanggal surat]', { jc: 'right', after: 240 }),
          idLabel('Nomor', '[Nomor surat]'),
          idLabel('Lampiran', '[Jumlah lampiran]'),
          idLabel('Hal', '[Hal surat]', { after: 240 }),
          // [Nama orang tua], [Nama siswa] and [Kelas] are the columns of a class list, for mail merge.
          p('Yth. Bapak/Ibu [Nama orang tua]'),
          p('Orang tua/wali dari [Nama siswa], kelas [Kelas]'),
          p('di tempat', { after: 240 }),
          p('Dengan hormat,', { after: 120 }),
          p('[Isi surat: kegiatan, informasi, atau hal yang perlu diketahui orang tua/wali.]', { jc: 'both', first: 709, after: 120 }),
          p('[Hal yang diminta dari orang tua/wali, dan batas waktunya.]', { jc: 'both', first: 709, after: 120 }),
          p('Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.', { jc: 'both', first: 709, after: 360 }),
          schoolSignatures(ID_WIDTH, '\nWali kelas [Kelas],\n\n\n\n_[Nama wali kelas]_\nNIP [NIP wali kelas]'),
        ].join(''),
      }),
    },
  ];
}
// Added here, after what they are made of (constants above), as the start page's teachers' group.
STARTERS.push(...forTeachers(), ...untukGuru());

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

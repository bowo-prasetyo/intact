# A memo written by python-docx: sections (one landscape), headers and footers, a different first
# page, styles, a table with merged cells, a page break, core properties. Fictional content only.
import sys
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Pt, Cm, RGBColor

doc = Document()
doc.core_properties.author = 'Records Office'
doc.core_properties.last_modified_by = 'Records Office'
doc.core_properties.title = 'Internal memo'

sec = doc.sections[0]
sec.different_first_page_header_footer = True
sec.first_page_header.paragraphs[0].text = 'MEMO — first page'
sec.header.paragraphs[0].text = 'Internal memo'
sec.footer.paragraphs[0].text = 'Records Office · not for circulation'
sec.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_heading('Internal memo', 0)
p = doc.add_paragraph('To: ', style='Normal')
p.add_run('All departments').bold = True
p = doc.add_paragraph('Subject: ')
r = p.add_run('Moving to the new building')
r.italic = True
r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
doc.add_heading('Timeline', level=1)
for s in ['Pack by Friday.', 'Movers arrive on Monday.', 'Unpack by Wednesday.']:
    doc.add_paragraph(s, style='List Number')
doc.add_heading('Rooms', level=1)
t = doc.add_table(rows=4, cols=3)
t.style = 'Table Grid'
for i, row in enumerate([('Team', 'Old room', 'New room'), ('Finance', '101', '3.12'), ('Legal', '102', '3.14'), ('Archive', '—', 'Basement')]):
    for j, v in enumerate(row):
        t.cell(i, j).text = v
t.cell(3, 1).merge(t.cell(3, 2))
doc.add_paragraph('Notes', style='Intense Quote')
p = doc.add_paragraph('Please label every box with ')
r = p.add_run('your team and room')
r.underline = True
r.font.size = Pt(13)
p.add_run('.')
p.add_run().add_break(WD_BREAK.PAGE)
doc.add_paragraph('Floor plans follow on the next, landscape page.')
new = doc.add_section()
new.orientation = WD_ORIENT.LANDSCAPE
new.page_width, new.page_height = sec.page_height, sec.page_width
new.left_margin = new.right_margin = Cm(2)
doc.add_heading('Floor plans', level=1)
doc.add_paragraph('Third floor: Finance and Legal.', style='List Bullet')
doc.add_paragraph('Basement: Archive.', style='List Bullet')
doc.add_paragraph('Questions go to the move coordinators.')
doc.save(sys.argv[1])

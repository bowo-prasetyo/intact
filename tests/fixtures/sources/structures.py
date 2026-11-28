"""Writes libreoffice-structures.docx with LibreOffice, through its UNO API (fictional content only).

A document of the structures the round-trip promise most often meets in real files, each made by
LibreOffice itself and saved as Word 2007-365 .docx: an embedded chart, an embedded spreadsheet (an
OLE object), a table nested in a table's cell, numbering in an unusual style (upper Roman, then
letters in brackets, starting from 4), a drawing shape with text (written as mc:AlternateContent,
with a VML fallback) and three page styles with headers of their own (a first page, the body, a
landscape annex). Run: python3 structures.py OUT.docx (needs soffice and python3-uno).
"""
import os, subprocess, sys, time, uno
from com.sun.star.beans import PropertyValue
from com.sun.star.awt import Size, Point
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

def prop(name, value):
    p = PropertyValue(); p.Name = name; p.Value = value; return p

out = os.path.abspath(sys.argv[1])
profile = f"file:///tmp/lo-structures-{os.getpid()}"
office = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', f'-env:UserInstallation={profile}',
                           '--accept=socket,host=127.0.0.1,port=2083;urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', uno.getComponentContext())
for _ in range(60):
    try:
        ctx = resolver.resolve('uno:socket,host=127.0.0.1,port=2083;urp;StarOffice.ComponentContext'); break
    except Exception:
        time.sleep(0.5)
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desktop.loadComponentFromURL('private:factory/swriter', '_blank', 0, (prop('Hidden', True),))
text = doc.Text
cursor = text.createTextCursor()

def para(s, style='Standard'):
    cursor.ParaStyleName = style
    text.insertString(cursor, s, False)
    text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

# Page styles: a first page, the body (its own header), a landscape annex (its own header).
styles = doc.StyleFamilies.getByName('PageStyles')
for name, header in (('First Page', 'Cover header: Harbour Works quarterly'), ('Default Page Style', 'Body header: Harbour Works'), ('Landscape', 'Annex header: figures')):
    st = styles.getByName(name) if styles.hasByName(name) else None
    if st is None:
        st = doc.createInstance('com.sun.star.style.PageStyle'); styles.insertByName(name, st)
    st.HeaderIsOn = True
    st.HeaderText.setString(header)
landscape = styles.getByName('Landscape')
landscape.IsLandscape = True
w, h = landscape.Width, landscape.Height
if w < h: landscape.Width, landscape.Height = h, w

cursor.PageDescName = 'First Page'
para('Harbour Works: structures report', 'Title')
para('A fictional report holding the structures real documents carry.')

cursor.PageDescName = 'Default Page Style'
para('Numbering in an unusual style', 'Heading 1')
# A list style of its own: upper Roman from IV, then letters in brackets.
from com.sun.star.style.NumberingType import ROMAN_UPPER, CHARS_LOWER_LETTER
lists = doc.StyleFamilies.getByName('NumberingStyles')
listStyle = doc.createInstance('com.sun.star.style.NumberingStyle')
lists.insertByName('Inspection list', listStyle)
rules = listStyle.NumberingRules
def level(i, **values):
    props = list(rules.getByIndex(i))
    for key, value in values.items():
        props = [p for p in props if p.Name != key] + [prop(key, value)]
    uno.invoke(rules, 'replaceByIndex', (i, uno.Any('[]com.sun.star.beans.PropertyValue', tuple(props))))
level(0, NumberingType=ROMAN_UPPER, StartWith=4, ListFormat='%1%.')
level(1, NumberingType=CHARS_LOWER_LETTER, StartWith=1, ListFormat='(%2%)')
listStyle.NumberingRules = rules
for lvl, s_ in ((0, 'Quay walls inspected'), (1, 'North quay: sound'), (1, 'South quay: cracks noted'), (0, 'Cranes serviced')):
    cursor.ParaStyleName = 'Standard'
    cursor.NumberingStyleName = 'Inspection list'
    cursor.NumberingLevel = lvl
    text.insertString(cursor, s_, False)
    text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
cursor.NumberingStyleName = ''

para('A table in a table', 'Heading 1')
outer = doc.createInstance('com.sun.star.text.TextTable'); outer.initialize(2, 2)
text.insertTextContent(cursor, outer, False)
outer.getCellByName('A1').setString('Berth'); outer.getCellByName('B1').setString('Details')
outer.getCellByName('A2').setString('Berth 3')
cell = outer.getCellByName('B2')
inner = doc.createInstance('com.sun.star.text.TextTable'); inner.initialize(2, 2)
cell.insertTextContent(cell.createTextCursor(), inner, False)
for name, value in (('A1', 'Depth'), ('B1', '12 m'), ('A2', 'Length'), ('B2', '240 m')):
    inner.getCellByName(name).setString(value)
para('After the tables.')

para('An embedded chart', 'Heading 1')
chart = doc.createInstance('com.sun.star.text.TextEmbeddedObject')
chart.CLSID = '12dcae26-281f-416f-a234-c3086127382e'
chart.Width, chart.Height = 12000, 7000
text.insertTextContent(cursor, chart, False)
model = chart.getExtendedControlOverEmbeddedObject().getComponent()
data = model.getData()
data.setRowDescriptions(('Q1', 'Q2', 'Q3')); data.setColumnDescriptions(('Ships', 'Cargo (kt)'))
data.setData(((42.0, 310.0), (51.0, 365.0), (47.0, 342.0)))
text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

para('An embedded spreadsheet', 'Heading 1')
sheet = doc.createInstance('com.sun.star.text.TextEmbeddedObject')
sheet.CLSID = '47BBB4CB-CE4C-4E80-A591-42D9AE74950F'
sheet.Width, sheet.Height = 9000, 3000
text.insertTextContent(cursor, sheet, False)
calc = sheet.getExtendedControlOverEmbeddedObject().getComponent()
s0 = calc.Sheets.getByIndex(0)
for r, row in enumerate((('Berth', 'Calls'), ('1', '18'), ('2', '23'))):
    for c, v in enumerate(row):
        s0.getCellByPosition(c, r).setFormula(v)
text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

para('A shape with text', 'Heading 1')
shape = doc.createInstance('com.sun.star.drawing.RectangleShape')
shape.Size = Size(6000, 2000)
shape.AnchorType = uno.Enum('com.sun.star.text.TextContentAnchorType', 'AS_CHARACTER')
text.insertTextContent(cursor, shape, False)
shape.setString('Safety first')
text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

cursor.PageDescName = 'Landscape'
para('Annex: figures on a landscape page', 'Heading 1')
para('The annex has a header of its own.')

doc.DocumentProperties.Author = 'Fixture author'
doc.storeToURL(uno.systemPathToFileUrl(out), (prop('FilterName', 'MS Word 2007 XML'),))
doc.close(True)
try: desktop.terminate()
except Exception: pass
office.wait(timeout=30)
print('wrote', out)

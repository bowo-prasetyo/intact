"""Writes python-docx-customxml.docx (fictional content only): a short document with a custom XML
part, as SharePoint and records systems attach metadata to a document (customXml/item2.xml, with its
properties part, related to the main document part), and a content control in the text whose
w:dataBinding points into it. Run: python3 customxml.py OUT.docx (needs python-docx).
"""
import sys
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml

ITEM = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<record xmlns="urn:example:harbour-records"><reference>HW-2026-114</reference><classification>Internal</classification><owner>Records office</owner></record>'
PROPS = b'<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="{6C3E2A51-0F43-4B5E-9C2D-3A1F7E8B9D10}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="urn:example:harbour-records"/></ds:schemaRefs></ds:datastoreItem>'

doc = Document()
doc.add_heading('Records note', level=1)
doc.add_paragraph('This note carries its record metadata in a custom XML part, as a records system adds it.')
p = doc.add_paragraph('Reference: ')
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
sdt = parse_xml(f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:alias w:val="Reference"/><w:id w:val="4711"/>'
                '<w:dataBinding w:prefixMappings="xmlns:r=\'urn:example:harbour-records\'" w:xpath="/r:record[1]/r:reference[1]" w:storeItemID="{6C3E2A51-0F43-4B5E-9C2D-3A1F7E8B9D10}"/>'
                '<w:text/></w:sdtPr><w:sdtContent><w:r><w:t>HW-2026-114</w:t></w:r></w:sdtContent></w:sdt>')
p._p.append(sdt)
doc.add_paragraph('The reference above is bound to the custom XML part: Word fills it in from there.')

pkg = doc.part.package
item = Part(PackURI('/customXml/item2.xml'), 'application/xml', ITEM, pkg)
props = Part(PackURI('/customXml/itemProps2.xml'), 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml', PROPS, pkg)
item.relate_to(props, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps')
doc.part.relate_to(item, RT.CUSTOM_XML)
doc.core_properties.author = 'Fixture author'
doc.save(sys.argv[1])
print('wrote', sys.argv[1])

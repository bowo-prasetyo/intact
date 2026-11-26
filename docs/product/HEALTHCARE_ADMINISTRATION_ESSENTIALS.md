# Healthcare administration templates

## Scope

Intact remains a general-purpose private `.docx` editor. This adds a group of administrative
starter templates for healthcare organisations. It does not turn Intact into an electronic health
record, clinical decision-support system, records-management system or compliance-certified
product. The group comes last on the start page, after the everyday templates and the teachers'.

## Included templates

- Healthcare policy
- Standard operating procedure
- Patient administration form
- Discharge administration checklist
- Incident report
- Patient complaint record
- Healthcare quality audit
- Corrective action record (CAPA)

They are built like every other starter (`src/docx/starters.js`, group `health`): in the browser,
on the paper used where the browser is (Letter in the Americas that use it, else A4), with tables
sized to the text width. Each is a controlled document:

- the organisation, document ID and version in the header on every page, and numbered pages;
- a document-control table: organisation, department, owner, status, document ID, version,
  effective date and next review date;
- check boxes that are Word's own (`w14:checkbox` content controls, ticked with a click);
- a "Before approval or sharing" checklist and a sign-off table (prepared, reviewed and approved
  by, each with its own date).

## Placeholders

Fill in fills a placeholder once for every place it appears, so every placeholder names one thing:
`[Effective date]` and `[Next review date]`, `[Prepared by]` and `[Reviewed by]`, `[Containment
due date]` and `[Corrective due date]` are separate. Only `[Organisation]`, `[Document ID]` and
`[Version]` repeat, in the body and the header, where they mean the same thing. A table's first row
has placeholders; the rows after it are left empty to fill in.

## What the documents say, and what they don't

The documents contain only what an organisation would keep in them. They don't mention Intact or its
menus, so nothing about the tool reaches whoever receives the file. The patient administration form
and the discharge checklist say, as form text, that diagnoses and clinical readiness belong in the
clinical record.

The start page's note for the group says these are administrative templates, not clinical forms or a
compliance certification, and that **File > Inspect document** checks for comments, tracked
changes, hidden text and personal information before a document is shared. That check is a review
aid, not automated patient-data detection, de-identification, legal advice or a certification. It
cannot establish whether a disclosure is authorised or whether a deployment meets HIPAA or other
applicable requirements.

Approval names and dates are placeholders, not an electronic-signature or approval workflow.

## Acceptance checks (`tests/healthcare-starters.spec.js`)

- The group lists all eight templates, last on the start page, with its note.
- Each opens with Fill in, downloads under its own title, and is written as Word writes a document
  (properties in schema order, styles defined, parts related and typed), with the header, page
  numbers and check boxes above.
- No placeholder other than organisation, document ID and version is used twice in a document;
  filling in a CAPA's dates and names puts each in its own place.
- No table is wider than the text, on A4 or on Letter.
- The documents don't mention Intact or its menus.
- To do before release, by hand: open each in Microsoft Word, as the vision asks of each new kind
  of document.

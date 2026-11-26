// Healthcare administration starter documents. These are administrative templates, not clinical
// forms or clinical guidance. They are generated locally as ordinary Word-compatible DOCX files.
import { blankDocxFiles } from './blank.js';

const esc = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const p = (text, style = '') => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
const heading = (text) => p(text, 'Heading1');
const subheading = (text) => p(text, 'Heading2');
const field = (label, value = '[Complete this field]') => p(`${label}: ${value}`);
const checklist = (items) => items.map((item) => p(`☐  ${item}`)).join('');
const table = (headers, rows) => {
  const row = (cells, header = false) => `<w:tr>${cells.map((cell) => `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>${header ? '<w:shd w:fill="D9E2F3"/>' : ''}</w:tcPr>${p(cell)}</w:tc>`).join('')}</w:tr>`;
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${headers.map(() => '<w:gridCol w:w="2400"/>').join('')}</w:tblGrid>${row(headers, true)}${rows.map((r) => row(r)).join('')}</w:tbl>`;
};
const document = (title, description, sections) => ({
  id: `healthcare-${title.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`,
  title,
  description,
  file: title,
  group: 'healthcare',
  build: () => ({
    paper: 'A4',
    language: 'en-US',
    font: { name: 'Arial', sz: 22 },
    styles: ['tableGrid'],
    body: [
      p(title, 'Title'),
      p('HEALTHCARE ADMINISTRATION • TEMPLATE'),
      field('Organisation', '[Organisation / facility name]'),
      field('Department', '[Department / service]'),
      field('Document ID', '[Unique document ID]'),
      field('Version', '[Version number]'),
      field('Owner', '[Role responsible for this document]'),
      field('Effective date', '[YYYY-MM-DD]'),
      field('Next review date', '[YYYY-MM-DD]'),
      p('Use the organisation-approved policy, terminology and approval route. Replace every bracketed placeholder before use. Do not enter patient-identifiable or other sensitive information unless it is necessary, authorised and handled under local policy. This template is an administrative aid, not clinical guidance or a compliance certification.'),
      ...sections,
      heading('Before approval or sharing'),
      checklist([
        'Confirm the document owner, version, effective date and review date.',
        'Check that the recipient is authorised and the distribution list is correct.',
        'Remove unnecessary patient/staff identifiers and verify any necessary disclosure is authorised.',
        'Use Review → Inspect document / Before you release to check comments, tracked changes, hidden text and document properties.',
        'Resolve comments and tracked changes intentionally; do not assume deleting visible text removes earlier revisions.',
        'Confirm the final document follows the organisation’s records-retention and secure-storage procedures.',
      ]),
      field('Prepared by', '[Name / role]'),
      field('Reviewed by', '[Name / role]'),
      field('Approved by', '[Authorised approver / role]'),
      field('Approval date', '[YYYY-MM-DD]'),
    ].join(''),
  }),
});

export const HEALTHCARE_STARTERS = [
  document('Healthcare policy', 'A controlled policy with purpose, scope, responsibilities and review details', [
    heading('1. Purpose'), p('[State the administrative purpose and intended outcome.]'),
    heading('2. Scope'), p('[Define the departments, roles, sites and processes covered.]'),
    heading('3. Policy statement'), p('[State the organisation-approved policy requirements.]'),
    heading('4. Roles and responsibilities'), table(['Role', 'Responsibility'], [['[Role]', '[Responsibility]'], ['[Role]', '[Responsibility]']]),
    heading('5. Procedure and related documents'), p('[Describe the approved administrative process and link related controlled documents.]'),
    heading('6. Monitoring and review'), p('[Describe the owner, review cycle, evidence and escalation route.]'),
  ]),
  document('Standard operating procedure', 'A step-by-step administrative SOP with controls and escalation', [
    heading('1. Purpose and scope'), p('[What this procedure covers and when it applies.]'),
    heading('2. Preconditions and resources'), checklist(['[Required authorisation]', '[System access / equipment]', '[Approved forms or reference documents]']),
    heading('3. Procedure'), table(['Step', 'Action', 'Responsible role', 'Record / evidence'], [['1', '[Action]', '[Role]', '[Record]'], ['2', '[Action]', '[Role]', '[Record]'], ['3', '[Action]', '[Role]', '[Record]']]),
    heading('4. Exceptions and escalation'), p('[Describe administrative exceptions, who to contact and how to record escalation.]'),
    heading('5. Records'), p('[Identify records created, storage location, access controls and retention schedule.]'),
  ]),
  document('Patient administration form', 'An administrative intake and registration form; avoid unnecessary clinical details', [
    heading('Registration details'), field('Patient / record reference', '[Use the approved identifier]'), field('Date received', '[YYYY-MM-DD]'), field('Registration type', '[New / update / other]'),
    heading('Contact and representative details'), field('Preferred contact method', '[Approved option]'), field('Authorised representative', '[If applicable; verify authority]'), field('Interpreter / accessibility arrangement', '[If requested and relevant]'),
    heading('Administrative checks'), checklist(['Identity checked using the approved procedure', 'Authority to act / representative status checked where applicable', 'Required administrative documents received', 'Contact preferences recorded accurately', 'Outstanding items and follow-up owner recorded']),
    heading('Outstanding items'), table(['Item', 'Owner', 'Due date', 'Status'], [['[Item]', '[Role]', '[YYYY-MM-DD]', '[Open / complete]'], ['[Item]', '[Role]', '[YYYY-MM-DD]', '[Open / complete]']]),
    p('Do not use this form to record diagnoses, treatment decisions or other clinical assessments. Use the approved clinical record for clinical information.'),
  ]),
  document('Discharge administration checklist', 'A non-clinical administrative handover and records checklist', [
    heading('Administrative reference'), field('Patient / record reference', '[Use the approved identifier]'), field('Discharge date', '[YYYY-MM-DD]'), field('Responsible administrative role', '[Role]'),
    heading('Completion checklist'), checklist(['Required administrative discharge fields completed', 'Authorised recipient / destination confirmed where applicable', 'Approved documents prepared and checked', 'Appointments or follow-up administration routed to the responsible team', 'Equipment / property / belongings process recorded where applicable', 'Outstanding administrative tasks assigned with due dates', 'Record filing and retention steps completed']),
    heading('Outstanding actions'), table(['Action', 'Owner', 'Due date', 'Completion evidence'], [['[Action]', '[Role]', '[YYYY-MM-DD]', '[Evidence]'], ['[Action]', '[Role]', '[YYYY-MM-DD]', '[Evidence]']]),
    field('Administrative handover confirmed by', '[Name / role]'),
    p('This checklist does not determine clinical readiness for discharge. Clinical decisions and instructions must be made and documented by authorised clinical staff in the approved clinical record.'),
  ]),
  document('Incident report', 'A factual, non-judgemental administrative incident record', [
    heading('Incident details'), field('Incident reference', '[Unique reference]'), field('Date and time', '[YYYY-MM-DD HH:MM and time zone]'), field('Location / service', '[Location]'), field('Category', '[Organisation-approved category]'),
    heading('People and immediate response'), field('People involved', '[Use approved identifiers; minimum necessary detail]'), field('Reported by', '[Name / role]'), field('Immediate actions taken', '[Factual description; do not speculate]'), field('Manager / safety lead notified', '[Role and notification time]'),
    heading('Factual account'), p('[Record what was observed or reported, separating facts from attributed statements. Avoid blame, assumptions and unnecessary personal or clinical details.]'),
    heading('Follow-up'), table(['Action', 'Responsible role', 'Due date', 'Status'], [['[Action]', '[Role]', '[YYYY-MM-DD]', '[Open]'], ['[Action]', '[Role]', '[YYYY-MM-DD]', '[Open]']]),
    field('Related incident / risk reference', '[Reference, if applicable]'),
  ]),
  document('Patient complaint record', 'A respectful complaint log with acknowledgement and response tracking', [
    heading('Complaint reference'), field('Reference number', '[Unique reference]'), field('Date received', '[YYYY-MM-DD]'), field('Channel', '[In person / phone / letter / online / other]'), field('Preferred response method', '[As authorised and recorded]'),
    heading('Complaint summary'), p('[Summarise the concern in the complainant’s own terms where practical. Keep information relevant and factual.]'),
    heading('Acknowledgement and handling'), field('Acknowledgement sent', '[YYYY-MM-DD / pending]'), field('Case owner', '[Role]'), field('Target response date', '[YYYY-MM-DD]'), field('Escalation route', '[Organisation-approved route]'),
    heading('Actions and outcome'), table(['Action / response', 'Owner', 'Date', 'Status'], [['[Action]', '[Role]', '[YYYY-MM-DD]', '[Open]'], ['[Outcome]', '[Role]', '[YYYY-MM-DD]', '[Pending]']]),
    field('Outcome communicated on', '[YYYY-MM-DD]'), field('Re-open / appeal information provided', '[Yes / no / not applicable]'),
  ]),
  document('Healthcare quality audit', 'An administrative quality audit with evidence, findings and follow-up', [
    heading('Audit plan'), field('Audit scope', '[Process / department / site]'), field('Audit period', '[Start and end dates]'), field('Audit lead', '[Role]'), field('Criteria and source', '[Approved standard / policy / version]'),
    heading('Findings'), table(['Criterion', 'Evidence reviewed', 'Result', 'Finding / gap'], [['[Criterion]', '[Evidence reference]', '[Met / partial / not met]', '[Finding]'], ['[Criterion]', '[Evidence reference]', '[Met / partial / not met]', '[Finding]'], ['[Criterion]', '[Evidence reference]', '[Met / partial / not met]', '[Finding]']]),
    heading('Summary and limitations'), p('[Summarise the evidence and limitations. Do not include identifiable patient data in the report; cite approved, access-controlled evidence references.]'),
    heading('Recommendations'), p('[List proportionate recommendations, owners and target dates.]'),
  ]),
  document('Corrective action record (CAPA)', 'A corrective and preventive action record with effectiveness review', [
    heading('Issue and source'), field('CAPA reference', '[Unique reference]'), field('Source', '[Audit / incident / complaint / other]'), field('Date raised', '[YYYY-MM-DD]'), field('Process owner', '[Role]'),
    heading('Problem statement and containment'), p('[Describe the process issue using evidence. Record immediate containment and any escalation.]'),
    heading('Cause analysis'), p('[Record the approved analysis method, evidence and validated contributing factors. Avoid unsupported conclusions.]'),
    heading('Action plan'), table(['Action type', 'Action', 'Owner', 'Due date', 'Evidence'], [['Containment', '[Action]', '[Role]', '[YYYY-MM-DD]', '[Evidence]'], ['Corrective', '[Action]', '[Role]', '[YYYY-MM-DD]', '[Evidence]'], ['Preventive', '[Action]', '[Role]', '[YYYY-MM-DD]', '[Evidence]']]),
    heading('Effectiveness review'), field('Review date', '[YYYY-MM-DD]'), field('Measure / acceptance criteria', '[Defined measure]'), field('Result', '[Effective / further action required]'), field('Reviewed by', '[Role]'),
  ]),
];

export function healthcareStarterFiles(id) {
  const starter = HEALTHCARE_STARTERS.find((item) => item.id === id);
  if (!starter) return null;
  return blankDocxFiles({ ...starter.build(), paper: 'A4' });
}

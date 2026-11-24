import type { WorkspaceTab } from '../types/procurement';

export interface HelpProblem {
  symptom: string;
  cause: string;
  fix: string;
}

export interface HelpTopic {
  id: string;
  title: string;
  summary: string;
  /** Plain-language search words, used in addition to the title and summary */
  keywords: string[];
  steps: string[];
  checklist?: string[];
  problems: HelpProblem[];
  /** Filename in src/assets/help (screenshot of the real screen) */
  screenshot?: string;
  screenshotCaption?: string;
}

export const HELP_TOPICS: HelpTopic[] = [
  {
    id: 'overview',
    title: 'How the workflow fits together',
    summary: 'Work goes from the supplier offer, through pricing, to the customer documents and shipment.',
    keywords: ['start', 'workflow', 'stages', 'where to begin', 'getting started'],
    steps: [
      'Internal Quote: enter or import the supplier offer (items, unit costs, terms).',
      'Price Calc (Internal): choose the markup and VAT, and enter box sizes for freight. This is confidential and never sent to the customer.',
      'Quotation (Master): set the end user and job controls. These flow into every other document.',
      'Proforma Invoice, Supplier PO, Commercial Invoice, Packing List, Delivery Note: issue each one when its checks pass.',
      'Each issued document gets a number from the Master Registry. Issued documents are locked; see "Changing an issued document".',
    ],
    problems: [
      {
        symptom: 'I cannot find where to start.',
        cause: 'Nothing has been entered for this job yet.',
        fix: 'Start on Internal Quote, or import a supplier quote with "Open .xlsx / .pdf" in the top bar.',
      },
    ],
  },
  {
    id: 'internalquote',
    title: 'Internal Quote: the supplier offer',
    summary: 'Records what the supplier offers. Costs entered here feed Price Calc.',
    keywords: ['supplier', 'offer', 'internal quote', 'cost', 'supplier name', 'terms'],
    steps: [
      'Open Internal Quote from the sidebar.',
      'Enter the supplier name, quote reference, currency, incoterm and payment terms. Imported quotes fill these in for you.',
      'For each line, check the requested item, the offered specification, the UoM, the quantity and the unit cost.',
      'Select several rows with the checkboxes to set the UoM or clear offers for all of them at once.',
      'Go to Price Calc to set the markup and freight.',
    ],
    checklist: [
      'Supplier name is filled in.',
      'Every line has a quantity and a supplier unit cost.',
      'Currency matches the supplier quote.',
    ],
    problems: [
      {
        symptom: 'Costs do not appear in Price Calc.',
        cause: 'The row has no supplier unit cost, or the row is empty.',
        fix: 'Enter the unit cost on the row. Rows with no description or quantity are ignored.',
      },
      {
        symptom: 'The fields are greyed out.',
        cause: 'The job has an issued document, so supplier details are locked.',
        fix: 'Use "Unlock to amend" on the issued document, enter the admin password, and give a reason when you save the revision.',
      },
    ],
    screenshot: 'internal-quote.jpg',
    screenshotCaption: 'Internal Quote: supplier details and the item table.',
  },
  {
    id: 'importquote',
    title: 'Importing a supplier quote (PDF or Excel)',
    summary: 'Step-by-step: upload the quote, check what was read, fix anything missed, then apply it to a job.',
    keywords: ['import', 'supplier quote', 'pdf', 'xlsx', 'excel', 'column mapper', 'paste', 'scanned', 'upload'],
    steps: [
      'Upload: click "Open .xlsx / .pdf" in the top bar (or use the import button on Internal Quote or Import & Export).',
      'Review Items: check the detected supplier, quote reference, currency, incoterm and payment terms. Untick any line you do not want. Edit a row to correct it.',
      'Column Mapper Assistant: if lines were missed, open the Column Mapper. Pick the table, set the header row, then assign Description, UoM, Qty, Unit Price or Line Total (÷ Qty). Turn on "Merge wrapped lines" if a description spans two lines.',
      'Paste / Fix Text: for a scanned or image-only PDF, select and copy the quote table from the PDF, paste it here, choose the delimiter (Auto, Tab, Spaces, Pipe or Comma), then extract.',
      'Apply: choose "Active job" to load the items into this job, or "New job" to create a separate one. Then check the result in Internal Quote and Price Calc.',
    ],
    checklist: [
      'Every line you want has a description, quantity and unit price.',
      'Currency and incoterm are correct.',
      'Unwanted lines are unticked.',
    ],
    problems: [
      {
        symptom: 'A scanned PDF loads no items.',
        cause: 'The PDF is a photo, so there is no text to read.',
        fix: 'Copy the table from the PDF (Live Text works on iPad), paste it into Paste / Fix Text, and extract.',
      },
      {
        symptom: 'A price is 100 times too high (for example 1250 instead of 12.50).',
        cause: 'A European decimal comma was read as a thousands separator in an older import.',
        fix: 'Correct the unit price in Review Items. Current versions read "12,50" as 12.50, so check the source line.',
      },
      {
        symptom: 'The currency is wrong.',
        cause: 'The currency is taken from the price lines, but a quote may mention another currency elsewhere.',
        fix: 'Change the currency in the Review Items details before applying.',
      },
      {
        symptom: 'Some lines are missing.',
        cause: 'The table layout was not recognised automatically.',
        fix: 'Open the Column Mapper Assistant and map the columns by hand.',
      },
    ],
    screenshot: 'import-review.jpg',
    screenshotCaption: 'Review Items: check the supplier details and each line before loading it.',
  },
  {
    id: 'pricecalc',
    title: 'Price Calc: markup, VAT and freight',
    summary: 'Internal costing. Sets the selling price for each line and the freight charged to the customer.',
    keywords: ['markup', 'vat', 'freight', 'cbm', 'kg', 'price', 'margin', 'profit', 'box size', 'l w h'],
    steps: [
      'Pick the markup: 10%, 15%, 18% or 20%, or type a custom percentage.',
      'Choose Include VAT (the VAT is folded into the unit price) or Remove VAT (0%).',
      'Enter the freight rate per CBM and per KG, and the documentation charge.',
      'For each line with boxes, enter the box length, width and height (cm) and the kilograms per box. CBM is calculated from these.',
      'Column R is the active unit price. It feeds the customer documents. Select several rows to set a supplier cost or clear box sizes for all of them.',
    ],
    checklist: [
      'Each line with boxes has a full box size (L, W and H).',
      'Markup and VAT are the ones agreed for this job.',
    ],
    problems: [
      {
        symptom: 'Freight is zero or too low.',
        cause: 'Lines have boxes but no box size, so there is no CBM to charge on.',
        fix: 'Enter the box length, width and height for those lines. Issuing a quotation or invoice is blocked until they are complete.',
      },
      {
        symptom: 'The Price Calc fields are greyed out.',
        cause: 'A document has been issued for this job.',
        fix: 'Use "Unlock to amend" on the issued document and save a revision with a reason.',
      },
    ],
    screenshot: 'price-calc.jpg',
    screenshotCaption: 'Price Calc: markup, VAT, freight and the costing matrix.',
  },
  {
    id: 'documents',
    title: 'Issuing the trade documents',
    summary: 'Quotation, Proforma Invoice, Supplier PO, Commercial Invoice, Packing List and Delivery Note.',
    keywords: ['issue', 'issuing', 'quotation', 'proforma', 'invoice', 'packing list', 'delivery note', 'supplier po', 'blocked', 'pre-issue', 'checklist'],
    steps: [
      'Fill in the job controls on Quotation (Master): end user, RFQ reference, payment terms, lead time, and shipment details.',
      'Check the Price Calc figures for the markup and VAT you want.',
      'Open the document and review the figures in the A4 preview.',
      'Click "Issue" when the checks pass. The document gets the next number from the Master Registry, and its figures are frozen.',
      'Packing List: enter a carton number and gross weight for every line. Issuing it fills in the Delivery Note defaults.',
      'Delivery Note: enter the delivery date and transport reference, then issue it. Issuing the Delivery Note marks the job complete.',
    ],
    checklist: [
      'End user is selected (Supplier PO: supplier name is entered).',
      'Every line has a description, UoM and quantity greater than zero.',
      'Quotation, Proforma, Commercial Invoice and Supplier PO: every line has a supplier unit cost.',
      'Packing List and Delivery Note: every line has a carton number and gross weight. Delivery Note also needs a delivery date and transport reference.',
    ],
    problems: [
      {
        symptom: 'The Issue button is blocked.',
        cause: 'A pre-issue check is failing.',
        fix: 'Read the red messages above the button. Each one names the row and the field to fix.',
      },
      {
        symptom: 'A "Prerequisite Step Incomplete" warning appears.',
        cause: 'The suggested earlier step is not finished, for example the Quotation before the Proforma.',
        fix: 'Complete the step, or choose "Issue Anyway" if you accept the risk.',
      },
      {
        symptom: 'I need to issue a second document of the same type for this job.',
        cause: 'Each document type can be issued once per job.',
        fix: 'Use "New Job to Re-Issue" to copy the job into a new one.',
      },
    ],
    screenshot: 'quotation.jpg',
    screenshotCaption: 'Quotation (Master): job controls and the line items.',
  },
  {
    id: 'changing',
    title: 'Changing an issued document',
    summary: 'Issued figures are locked. Changes are made through a logged revision (Rev 2, Rev 3 ...).',
    keywords: ['change', 'amend', 'revision', 'unlock', 'locked', 'password', 'change log', 'audit', 'discard', 'rev'],
    steps: [
      'Click "Unlock to amend" on the issued document. Enter the admin password.',
      'Make the changes: quantity, supplier cost, supplier details, dates and so on. The job is open for amendment until you save or discard.',
      'Click "Save as Rev N" and enter a reason for the change (at least 5 characters).',
      'The previous figures are kept in the document history, and every changed field is written to the change log with the reason.',
      'To cancel instead, click "Discard changes". The job goes back to its saved state.',
    ],
    checklist: [
      'You have the admin password.',
      'You have a reason for every change you save.',
    ],
    problems: [
      {
        symptom: '"Incorrect admin password."',
        cause: 'The password does not match the one set in Clients & Settings.',
        fix: 'Enter the password the admin set. Ask the admin to reset it in Clients & Settings if it is lost.',
      },
      {
        symptom: 'Other documents in the job still show the old totals.',
        cause: 'Only the document you unlocked is revised.',
        fix: 'Amend each other issued document separately if it needs the same change.',
      },
      {
        symptom: 'Earlier revisions are gone.',
        cause: 'Once the Delivery Note is issued, the job is complete and earlier revisions are removed to keep the file small.',
        fix: 'Export the change log from the Master Registry before completing a job if you need the history.',
      },
    ],
  },
  {
    id: 'registry',
    title: 'Master Registry: document numbers and the change log',
    summary: 'Every issued document number is recorded here. Numbers are never reused.',
    keywords: ['registry', 'number', 'sequence', 'void', 'sync', 'export', 'change log', 'csv', 'excel', 'duplicate'],
    steps: [
      'The next available number is shown at the top. Numbers are issued in order, and only once.',
      'To retire a number, use "Mark VOID" and give a reason. A voided number is never issued again.',
      'To select records, tick the boxes and use "Export selected (.xlsx)".',
      'To export the change log, use Change log (.csv), (.xlsx) or (.json). The log records every change to an issued document, with its reason.',
      'Sync Master_Registry.xlsx merges a registry from another computer. A VOID number stays VOID.',
    ],
    problems: [
      {
        symptom: 'Two computers issued the same number.',
        cause: 'Each computer keeps its own registry. The app cannot see the other computer\'s numbers.',
        fix: 'Sync Master_Registry.xlsx between the computers and resolve the duplicate by voiding one of them.',
      },
      {
        symptom: 'The registry reset my numbering.',
        cause: '"Reset registry" starts again at #001.',
        fix: 'Only reset with a backup and a clear reason. Restoring a backup keeps every number already issued.',
      },
    ],
    screenshot: 'registry.jpg',
    screenshotCaption: 'Master Registry: the numbers issued so far.',
  },
  {
    id: 'importexport',
    title: 'Import & Export: files and backups',
    summary: 'Open files, export Excel and PDF documents, and make a full backup.',
    keywords: ['export', 'backup', 'restore', 'excel', 'xlsx', 'pdf', 'json', 'download', 'open file', 'current job workbook'],
    steps: [
      'Open a supplier quote or job file: "Open .xlsx / .pdf" in the top bar.',
      'Export an issued document as a values-only Excel file from its tab.',
      'Export the Current Job Workbook from Import & Export. It has live formulas linked across the sheets.',
      'Back up everything with the JSON backup. Keep the file somewhere safe.',
      'Restore only from a backup you trust. Restoring merges the registry, so numbers already issued are kept.',
    ],
    problems: [
      {
        symptom: 'The download does not start on iPad.',
        cause: 'iPad Safari uses the share sheet instead of a plain download.',
        fix: 'Choose "Save to Files" or "Open in Excel" from the share sheet that appears.',
      },
      {
        symptom: 'The storage warning appears.',
        cause: 'Browser storage is nearly full, usually because of large logos.',
        fix: 'Export a JSON backup, then remove old jobs.',
      },
    ],
    screenshot: 'import-export.jpg',
    screenshotCaption: 'Import & Export: opening files, exports and backups.',
  },
  {
    id: 'dashboard',
    title: 'Dashboard and reports',
    summary: 'Pipeline overview, job lifecycle matrix, and the PDF and Excel reports.',
    keywords: ['dashboard', 'report', 'pdf', 'landscape', 'portrait', 'margin', 'pending', 'lifecycle', 'matrix', 'scope'],
    steps: [
      'Use the scope and period controls to choose which jobs and weeks are shown.',
      'Read the lifecycle matrix. "Pending" means the document is not issued yet; "OK" means it is issued.',
      'Click "PDF Report", choose A4 Landscape (best for the full matrix) or A4 Portrait (stacked summary), then download.',
      'Click "Excel Report" for the same figures as a spreadsheet.',
    ],
    problems: [
      {
        symptom: 'A job is missing from the report.',
        cause: 'The report only includes jobs in the active scope, and completed jobs are hidden unless you include them.',
        fix: 'Change the scope, or turn on "include completed".',
      },
    ],
    screenshot: 'dashboard.jpg',
    screenshotCaption: 'Dashboard: the pipeline overview.',
  },
  {
    id: 'settings',
    title: 'Clients & Settings',
    summary: 'End-user directory, company branding, document styles and the admin password.',
    keywords: ['settings', 'clients', 'end user', 'logo', 'admin', 'password', 'style', 'theme', 'branding'],
    steps: [
      'Add or edit end users in the directory. Each one appears in the end-user list on documents.',
      'Upload the company logo (PNG or JPG).',
      'Choose the document style (v1 Executive Navy or v3 Charcoal and Amber).',
      'Change the admin password here. It is needed to unlock issued documents.',
    ],
    checklist: [
      'The admin password has been changed from the default.',
    ],
    problems: [
      {
        symptom: 'The logo looks wrong on printed documents.',
        cause: 'The image is too large or not a PNG or JPG.',
        fix: 'Use a PNG or JPG logo of a few hundred pixels across.',
      },
    ],
    screenshot: 'settings.jpg',
    screenshotCaption: 'Clients & Settings: the end-user directory and branding.',
  },
];

/** Which help topic opens for each workspace page */
const TAB_TOPIC: Record<WorkspaceTab, string> = {
  TRACKER: 'dashboard',
  INTERNAL_QUOTE: 'internalquote',
  PRICE_CALC: 'pricecalc',
  QT: 'documents',
  PI: 'documents',
  PO: 'documents',
  CI: 'documents',
  PL: 'documents',
  DN: 'documents',
  REGISTRY: 'registry',
  FILES: 'importexport',
  STUDIO: 'settings',
  HELP: 'overview',
};

export function helpTopicForTab(tab: WorkspaceTab): string {
  return TAB_TOPIC[tab] ?? 'overview';
}

/** Case-insensitive search across titles, summaries, keywords and symptoms */
export function searchHelpTopics(query: string): HelpTopic[] {
  const q = query.trim().toLowerCase();
  if (!q) return HELP_TOPICS;
  const words = q.split(/\s+/);
  return HELP_TOPICS.filter((t) => {
    const haystack = [
      t.title,
      t.summary,
      ...t.keywords,
      ...t.problems.map((p) => `${p.symptom} ${p.cause}`),
    ]
      .join(' ')
      .toLowerCase();
    return words.every((w) => haystack.includes(w));
  });
}

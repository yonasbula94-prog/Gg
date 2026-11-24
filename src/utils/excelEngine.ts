import ExcelJS from 'exceljs';
import {
  DOC_META,
  END_USER_OPTIONS,
  EXCEL_THEMES,
  resolveDocumentColumns,
} from '../constants/themes';
import {
  CompanyProfile,
  ComputedRow,
  CustomColumnDef,
  CustomColumnPosition,
  DocPrefix,
  JobItemRow,
  ProcurementJob,
  RegistryEntry,
  ThemePalette,
  ThemeVariantId,
} from '../types/procurement';
import { calculateJob, getResolvedEndUser } from './calculations';
import { createBlankJob } from '../store/storage';
import { shareOrDownloadBlob } from './platform';
import { ParsedSupplierQuote, parseGenericSupplierXlsx } from './supplierQuoteParser';

function resolveExcelTheme(themeOrId: ThemeVariantId | ThemePalette): ThemePalette {
  return typeof themeOrId === 'string' ? EXCEL_THEMES[themeOrId] : themeOrId;
}

function filterCustomCols(
  cols: CustomColumnDef[],
  pos: CustomColumnPosition
): CustomColumnDef[] {
  return cols.filter((c) => (c.position || 'after_no') === pos);
}

function colToLetter(colNum: number): string {
  let n = colNum;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

async function triggerDownload(buffer: ArrayBuffer, filename: string): Promise<void> {
  const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const blob = new Blob([buffer], { type: mime });
  await shareOrDownloadBlob(blob, filename, mime);
}

function thinBorder(argb: string): Partial<ExcelJS.Border> {
  return { style: 'thin', color: { argb } };
}

function mediumBorder(argb: string): Partial<ExcelJS.Border> {
  return { style: 'medium', color: { argb } };
}

function detectImageExtension(dataUrl: string): 'png' | 'jpeg' {
  if (dataUrl.startsWith('data:image/jpeg') || dataUrl.startsWith('data:image/jpg')) {
    return 'jpeg';
  }
  return 'png';
}

function styleBannerHeader(
  wb: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  theme: ThemePalette,
  company: CompanyProfile,
  docTitle: string,
  docNumber: string,
  colCount: number
): void {
  // Col A is clean 2-char margin spacer; content spans Col B .. Col (colCount + 1)
  const lastCol = colCount + 1;
  const fontName = theme.fontName || 'Calibri';

  ws.getRow(1).height = 8;

  // Row 2: Brand banner with embedded logo in Col B (white tile) + Company name in Col C..
  ws.getRow(2).height = 44;

  if (company.logoDataUrl && company.logoDataUrl.startsWith('data:image/')) {
    try {
      const imgId = wb.addImage({
        base64: company.logoDataUrl,
        extension: detectImageExtension(company.logoDataUrl),
      });
      ws.getCell(2, 2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFFFF' },
      };
      ws.addImage(imgId, {
        tl: { col: 1.08, row: 1.08 },
        ext: { width: 46, height: 48 },
      });
    } catch {
      // Fallback cleanly if custom logo is unreadable
    }
  }

  ws.mergeCells(2, 3, 2, lastCol - 2);
  const brandCell = ws.getCell(2, 3);
  brandCell.value = company.name;
  brandCell.font = { name: fontName, size: 16, bold: true, color: { argb: theme.excel.headerFont } };
  brandCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  ws.mergeCells(2, lastCol - 1, 2, lastCol);
  const titleCell = ws.getCell(2, lastCol - 1);
  titleCell.value = docTitle;
  titleCell.font = { name: fontName, size: 13, bold: true, color: { argb: theme.excel.headerFont } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };

  for (let c = 3; c <= lastCol; c++) {
    ws.getCell(2, c).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: theme.excel.headerFill },
    };
  }
  if (!company.logoDataUrl) {
    ws.getCell(2, 2).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: theme.excel.headerFill },
    };
  }

  // Row 3: Subtitle + document number (omitted cleanly when unissued — never prints "DRAFT")
  ws.mergeCells(3, 2, 3, lastCol - 2);
  const subCell = ws.getCell(3, 2);
  subCell.value = `${company.addressLine1}, ${company.addressLine2}  |  ${company.email}`;
  subCell.font = { name: fontName, size: 9, color: { argb: 'FF475569' } };
  subCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  ws.mergeCells(3, lastCol - 1, 3, lastCol);
  const docNoCell = ws.getCell(3, lastCol - 1);
  docNoCell.value = docNumber.trim() ? `DOC NO: ${docNumber.trim()}` : '';
  docNoCell.font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.strongBorderColor } };
  docNoCell.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };

  ws.getRow(3).height = 20;
  for (let c = 2; c <= lastCol; c++) {
    ws.getCell(3, c).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: theme.excel.subHeaderFill },
    };
    ws.getCell(3, c).border = {
      bottom: mediumBorder(theme.excel.accentFill),
    };
  }
}

/**
 * 1. VALUES-ONLY STYLED .XLSX PER ISSUED DOCUMENT
 * Strictly excludes PRICE CALC, controls block, and macros.
 * Hard-codes the claimed Document No. and freezes values.
 */
export async function exportIssuedDocumentXlsx(
  job: ProcurementJob,
  prefix: DocPrefix,
  themeOrId: ThemeVariantId | ThemePalette,
  company: CompanyProfile
): Promise<string> {
  const theme = resolveExcelTheme(themeOrId);
  const fontName = theme.fontName || 'Calibri';
  const meta = DOC_META[prefix];
  const calc = calculateJob(job);
  const issued = job.issuedDocs[prefix];
  const docNumber = issued && !issued.isVoided ? issued.docNumber : '';
  const endUser = getResolvedEndUser(job) || '—';
  const isLogistics = prefix === 'PL' || prefix === 'DN';
  const isPO = prefix === 'PO';
  const customCols = job.customColumns || [];
  const customMap = new Map(customCols.map((c) => [`custom:${c.id}`, c]));
  const orderedKeys = resolveDocumentColumns(job.columnOrder, customCols, isLogistics);
  const colCount = 1 + orderedKeys.length;
  const lastCol = colCount + 1;

  const wb = new ExcelJS.Workbook();
  wb.creator = company.name;
  wb.created = new Date();

  const ws = wb.addWorksheet(meta.sheetName, {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 11 }],
    pageSetup: {
      paperSize: 9, // A4
      orientation: isLogistics || customCols.length > 0 ? 'landscape' : 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.4, right: 0.4, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 },
    },
  });

  interface IssuedColSpec {
    key: string;
    header: string;
    width: number;
    align: 'left' | 'center' | 'right';
    wrapText?: boolean;
    numFmt?: string;
    getValue: (row: ComputedRow, idx: number) => string | number | null;
  }

  const buildSpec = (key: string): IssuedColSpec | null => {
    if (key.startsWith('custom:')) {
      const cc = customMap.get(key);
      if (!cc) return null;
      return {
        key,
        header: cc.label,
        width: 16,
        align: 'left',
        wrapText: true,
        getValue: (row) => row.customFields?.[cc.id] || '',
      };
    }
    switch (key) {
      case 'requested':
        return {
          key: 'requested',
          header: 'Requested Description',
          width: 26,
          align: 'left',
          wrapText: true,
          getValue: (row) => row.requestedDescription,
        };
      case 'offered':
        return {
          key: 'offered',
          header: 'Offered Description',
          width: 30,
          align: 'left',
          wrapText: true,
          getValue: (row) => row.offeredDescription,
        };
      case 'uom':
        return {
          key: 'uom',
          header: 'UoM',
          width: 9,
          align: 'center',
          getValue: (row) => row.uom,
        };
      case 'qty':
        return {
          key: 'qty',
          header: 'Qty',
          width: 9,
          align: 'right',
          numFmt: '#,##0.##',
          getValue: (row) => row.qty,
        };
      case 'carton':
        return {
          key: 'carton',
          header: 'Carton No.',
          width: 15,
          align: 'center',
          getValue: (row) => row.cartonNo,
        };
      case 'net_wt':
        return {
          key: 'net_wt',
          header: 'Net Wt (KG)',
          width: 14,
          align: 'right',
          numFmt: '#,##0.00',
          getValue: (row) => row.netWtKg,
        };
      case 'gross_wt':
        return {
          key: 'gross_wt',
          header: 'Gross Wt (KG)',
          width: 14,
          align: 'right',
          numFmt: '#,##0.00',
          getValue: (row) => row.grossWtKg,
        };
      case 'unit_val':
        return isPO
          ? {
              key: 'unit_val',
              header: `Unit Cost (${job.controls.currency})`,
              width: 16,
              align: 'right',
              numFmt: '#,##0.00',
              getValue: (row) => row.supplierUnitCost,
            }
          : {
              key: 'unit_val',
              header: `Unit Price (${job.controls.currency})`,
              width: 16,
              align: 'right',
              numFmt: '#,##0.00',
              getValue: (row) => row.activeUnitPrice,
            };
      case 'total_val':
        return isPO
          ? {
              key: 'total_val',
              header: `Total Cost (${job.controls.currency})`,
              width: 17,
              align: 'right',
              numFmt: '#,##0.00',
              getValue: (row) => row.supplierTotalCost,
            }
          : {
              key: 'total_val',
              header: `Total Price (${job.controls.currency})`,
              width: 17,
              align: 'right',
              numFmt: '#,##0.00',
              getValue: (row) => row.activeTotalPrice,
            };
      default:
        return null;
    }
  };

  const colSpecs: IssuedColSpec[] = [
    {
      key: 'no',
      header: 'No.',
      width: 6,
      align: 'center',
      getValue: (_row, idx) => idx + 1,
    },
    ...orderedKeys.map(buildSpec).filter((s): s is IssuedColSpec => s !== null),
  ];

  // Column widths (Col A is margin spacer; Col B..lastCol match colSpecs)
  ws.columns = [{ width: 2 }, ...colSpecs.map((c) => ({ width: c.width }))];

  styleBannerHeader(wb, ws, theme, company, meta.title, docNumber, colCount);

  // Metadata Block (Rows 5..9)
  const issueDate = issued
    ? new Date(issued.issuedAt).toISOString().slice(0, 10)
    : new Date().toISOString().slice(0, 10);

  const resolvedAddress =
    job.controls.endUserSelect !== 'OTHER'
      ? (company.endUserAddresses?.[job.controls.endUserSelect] ?? job.controls.endUserAddress)
      : job.controls.endUserAddress;

  const leftMeta: [string, string][] = isPO
    ? [
        ['SUPPLIER', job.controls.supplierName || '—'],
        ['ADDRESS', job.controls.supplierAddress || '—'],
        ['SUPPLIER REF', job.controls.supplierRef || '—'],
        ['INCOTERM', `${job.controls.incoterm} · ${job.controls.destination}`],
        ['CURRENCY', job.controls.currency],
      ]
    : [
        [meta.recipientLabel, endUser],
        ['ATTN / CONTACT', job.controls.endUserContact || '—'],
        ['ADDRESS', resolvedAddress || '—'],
        ['INCOTERM / DEST', `${job.controls.incoterm} · ${job.controls.destination}`],
        ['CURRENCY', job.controls.currency],
      ];

  const rightMeta: [string, string][] =
    prefix === 'PL'
      ? [
          ['DATE', issueDate],
          ['SHIP DATE', job.controls.shipDate || '—'],
          ['PORT OF LOADING', job.controls.portOfLoading || '—'],
          ['PI REFERENCE', job.issuedDocs.PI?.docNumber || '—'],
          ['TOTAL BOXES', String(calc.totalBoxes || '—')],
        ]
      : prefix === 'DN'
      ? [
          ['DATE', issueDate],
          ['DELIVERY DATE', job.controls.deliveryDate || '—'],
          ['TRANSPORT REF', job.controls.transportRef || '—'],
          ['RECEIVED BY', job.controls.receivedBy || '—'],
          ['PI REFERENCE', job.issuedDocs.PI?.docNumber || '—'],
        ]
      : [
          ['DATE', issueDate],
          ['RFQ / TENDER REF', job.controls.rfqRef || '—'],
          ['CUSTOMER PO REF', job.controls.customerPoRef || '—'],
          ['PAYMENT TERMS', job.controls.paymentTerms || '—'],
          ['LEAD TIME', job.controls.deliveryLeadTime || '—'],
        ];

  ws.getRow(4).height = 8;
  const rightLabelCol = Math.max(6, lastCol - 2);
  const rightValStart = rightLabelCol + 1;
  for (let i = 0; i < 5; i++) {
    const r = 5 + i;
    ws.getRow(r).height = 19;

    ws.getCell(r, 2).value = leftMeta[i][0];
    ws.getCell(r, 2).font = { name: fontName, size: 9, bold: true, color: { argb: 'FF64748B' } };
    ws.mergeCells(r, 3, r, rightLabelCol - 2);
    ws.getCell(r, 3).value = leftMeta[i][1];
    ws.getCell(r, 3).font = { name: fontName, size: 10, bold: i === 0, color: { argb: theme.excel.formulaFont } };

    ws.getCell(r, rightLabelCol).value = rightMeta[i][0];
    ws.getCell(r, rightLabelCol).font = { name: fontName, size: 9, bold: true, color: { argb: 'FF64748B' } };
    if (lastCol > rightValStart) {
      ws.mergeCells(r, rightValStart, r, lastCol);
    }
    ws.getCell(r, rightValStart).value = rightMeta[i][1];
    ws.getCell(r, rightValStart).font = { name: fontName, size: 10, bold: i === 0, color: { argb: theme.excel.formulaFont } };
    ws.getCell(r, rightValStart).alignment = { horizontal: 'right', vertical: 'middle' };
  }

  ws.getRow(10).height = 10;

  // Table Headers (Row 11)
  const headerRow = ws.getRow(11);
  headerRow.height = 26;
  colSpecs.forEach((spec, idx) => {
    const cell = headerRow.getCell(idx + 2);
    cell.value = spec.header;
    cell.font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
    cell.alignment = {
      vertical: 'middle',
      horizontal: spec.align,
      wrapText: true,
    };
    cell.border = {
      bottom: mediumBorder(theme.excel.strongBorderColor),
    };
  });

  // Populated Line Items (values only, no phantom zeros)
  const activeRows = calc.rows.filter((r) => !r.isEmpty);
  let cursor = 12;

  activeRows.forEach((row, idx) => {
    const r = ws.getRow(cursor);
    r.height = 24;
    const isZebra = idx % 2 === 1;

    colSpecs.forEach((spec, cIdx) => {
      const v = spec.getValue(row, idx);
      const cell = r.getCell(cIdx + 2);
      cell.value = v === null || v === undefined ? '' : v;
      cell.font = { name: fontName, size: 10, color: { argb: theme.excel.formulaFont } };
      if (isZebra) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.zebraFill } };
      }
      cell.border = { bottom: thinBorder(theme.excel.borderColor) };
      cell.alignment = {
        vertical: 'middle',
        horizontal: spec.align,
        wrapText: Boolean(spec.wrapText),
      };
      if (spec.numFmt && typeof v === 'number') {
        cell.numFmt = spec.numFmt;
      }
    });
    cursor += 1;
  });

  // Dynamic column lookup helper for totals alignment (resilient to any column reordering)
  const colNumOf = (key: string) => colSpecs.findIndex((c) => c.key === key) + 2;

  // Summary / Totals Block
  cursor += 1;
  if (isLogistics) {
    const qtyCol = colNumOf('qty');
    const cartonCol = colNumOf('carton');
    const netWtCol = colNumOf('net_wt');
    const grossWtCol = colNumOf('gross_wt');
    const totRow = ws.getRow(cursor);
    totRow.height = 24;

    totRow.getCell(2).value = 'TOTALS';
    totRow.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };

    if (qtyCol >= 2) {
      totRow.getCell(qtyCol).value = calc.totalQty;
      totRow.getCell(qtyCol).numFmt = '#,##0.##';
      totRow.getCell(qtyCol).alignment = { horizontal: 'right', vertical: 'middle' };
    }
    if (cartonCol >= 2) {
      totRow.getCell(cartonCol).value = `${calc.totalBoxes} CTNS`;
      totRow.getCell(cartonCol).alignment = { horizontal: 'center', vertical: 'middle' };
    }
    if (netWtCol >= 2) {
      totRow.getCell(netWtCol).value = calc.plTotalNetWt;
      totRow.getCell(netWtCol).numFmt = '#,##0.00';
      totRow.getCell(netWtCol).alignment = { horizontal: 'right', vertical: 'middle' };
    }
    if (grossWtCol >= 2) {
      totRow.getCell(grossWtCol).value = calc.plTotalGrossWt;
      totRow.getCell(grossWtCol).numFmt = '#,##0.00';
      totRow.getCell(grossWtCol).alignment = { horizontal: 'right', vertical: 'middle' };
    }

    for (let c = 2; c <= lastCol; c++) {
      const cell = totRow.getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.totalFill } };
      cell.font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.strongBorderColor } };
      cell.border = {
        top: thinBorder(theme.excel.strongBorderColor),
        bottom: mediumBorder(theme.excel.strongBorderColor),
      };
    }
    cursor += 2;
  } else if (isPO) {
    // Always place summary neatly flush-right on the sheet regardless of table column order
    const valCol = lastCol;
    const labelStartCol = Math.max(2, valCol - 2);
    const totRow = ws.getRow(cursor);
    totRow.height = 26;
    if (valCol - 1 > labelStartCol) {
      ws.mergeCells(cursor, labelStartCol, cursor, valCol - 1);
    }
    totRow.getCell(labelStartCol).value = `TOTAL ORDER COST (${job.controls.currency})`;
    totRow.getCell(labelStartCol).alignment = { horizontal: 'right', vertical: 'middle' };
    totRow.getCell(valCol).value = calc.totalSupplierCost;
    totRow.getCell(valCol).numFmt = '#,##0.00';
    totRow.getCell(valCol).alignment = { horizontal: 'right', vertical: 'middle' };
    for (let c = labelStartCol; c <= valCol; c++) {
      totRow.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.totalFill } };
      totRow.getCell(c).font = { name: fontName, size: 11, bold: true, color: { argb: theme.excel.strongBorderColor } };
      totRow.getCell(c).border = {
        top: thinBorder(theme.excel.strongBorderColor),
        bottom: mediumBorder(theme.excel.strongBorderColor),
      };
    }
    cursor += 2;
  } else {
    // QT / PI / CI: End-user documents show Goods Subtotal (VAT included in Unit Price) -> Freight -> Grand Total (ZERO VAT row)
    const valCol = lastCol;
    const summaryLines: [string, number, boolean][] = [
      ['GOODS SUBTOTAL', calc.goodsSubtotal, false],
      [`FREIGHT & DOCUMENTATION (${job.controls.incoterm})`, calc.customerFreight, false],
      [`GRAND TOTAL (${job.controls.currency})`, calc.grandTotal, true],
    ];

    const labelStartCol = Math.max(2, valCol - 3);
    summaryLines.forEach(([label, val, isGrand]) => {
      const r = ws.getRow(cursor);
      r.height = isGrand ? 26 : 21;
      if (valCol - 1 > labelStartCol) {
        ws.mergeCells(cursor, labelStartCol, cursor, valCol - 1);
      }
      const lCell = r.getCell(labelStartCol);
      lCell.value = label;
      lCell.alignment = { horizontal: 'right', vertical: 'middle' };

      const vCell = r.getCell(valCol);
      vCell.value = val;
      vCell.numFmt = '#,##0.00';
      vCell.alignment = { horizontal: 'right', vertical: 'middle' };

      for (let c = labelStartCol; c <= valCol; c++) {
        const cell = r.getCell(c);
        cell.font = {
          name: fontName,
          size: isGrand ? 11 : 10,
          bold: isGrand,
          color: { argb: isGrand ? theme.excel.strongBorderColor : theme.excel.formulaFont },
        };
        if (isGrand) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.totalFill } };
          cell.border = {
            top: thinBorder(theme.excel.strongBorderColor),
            bottom: mediumBorder(theme.excel.strongBorderColor),
          };
        } else {
          cell.border = { bottom: thinBorder(theme.excel.borderColor) };
        }
      }
      cursor += 1;
    });
    cursor += 1;
  }

  // Footer Remarks & Authorization (100% editable per job; empty omits cleanly)
  if (job.controls.remarks.trim()) {
    ws.mergeCells(cursor, 2, cursor, lastCol);
    const rem = ws.getCell(cursor, 2);
    rem.value = `Remarks: ${job.controls.remarks.trim()}`;
    rem.font = { name: fontName, size: 9, italic: true, color: { argb: 'FF475569' } };
    cursor += 1;
  }

  if (prefix === 'PI' || prefix === 'CI') {
    ws.mergeCells(cursor, 2, cursor, lastCol);
    const bank = ws.getCell(cursor, 2);
    bank.value = `${company.bankDetails}  |  ${company.trn}`;
    bank.font = { name: fontName, size: 9, color: { argb: 'FF475569' } };
    cursor += 2;
  } else {
    cursor += 1;
  }

  // Editable Signature Lines (left & right; leaving empty omits label cleanly)
  const leftSig = (job.controls.authorizedSignatory ?? '').trim();
  const rightSig = (job.controls.counterpartySignatory ?? '').trim();
  if (leftSig || rightSig) {
    if (leftSig) {
      ws.getCell(cursor, 2).value = `${leftSig}: ___________________________`;
      ws.getCell(cursor, 2).font = { name: fontName, size: 9, bold: true, color: { argb: 'FF334155' } };
    }
    if (rightSig) {
      const rightSigCol = Math.max(6, lastCol - 2);
      ws.getCell(cursor, rightSigCol).value = `${rightSig}: ___________________________`;
      ws.getCell(cursor, rightSigCol).font = { name: fontName, size: 9, bold: true, color: { argb: 'FF334155' } };
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  const safeBase = docNumber
    ? docNumber.replace(/[^a-zA-Z0-9_-]/g, '_')
    : job.jobName.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 28) || 'Job';
  const filename = docNumber
    ? `${safeBase}_${meta.sheetName.replace(/\s+/g, '_')}.xlsx`
    : `${meta.sheetName.replace(/\s+/g, '_')}_${safeBase}.xlsx`;
  await triggerDownload(buffer as ArrayBuffer, filename);
  return filename;
}

/**
 * 2. FULL LIVE-FORMULA GG_Job_Template.xlsx BUILDER
 * Includes HOW TO USE, PROFORMA INVOICE (rows 17-36 + K:L controls + L5 VAT),
 * PRICE CALC (E3 markup selector, 10/15/18/20% scenarios, Col R -> PI Unit Price, MAX freight + doc fee + markup),
 * QUOTATION, SUPPLIER PO, PACKING LIST, DELIVERY NOTE, COMMERCIAL INVOICE, and CONTROLS (with Pharmecore fixed).
 */
export async function exportJobTemplateXlsx(
  job: ProcurementJob,
  themeOrId: ThemeVariantId | ThemePalette,
  includeCurrentData: boolean,
  company: CompanyProfile
): Promise<string> {
  const theme = resolveExcelTheme(themeOrId);
  const fontName = theme.fontName || 'Calibri';
  const wb = new ExcelJS.Workbook();
  wb.creator = company.name;
  wb.created = new Date();

  // --- SHEET 1: HOW TO USE ---
  const wsHow = wb.addWorksheet('HOW TO USE', { views: [{ showGridLines: false }] });
  wsHow.columns = [{ width: 3 }, { width: 26 }, { width: 78 }];
  wsHow.mergeCells('B2:C2');
  wsHow.getCell('B2').value = `${company.name} — PROCUREMENT WORKBOOK GUIDE`;
  wsHow.getCell('B2').font = { name: fontName, size: 15, bold: true, color: { argb: theme.excel.headerFont } };
  wsHow.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
  wsHow.getRow(2).height = 30;

  const guideRows: [string, string][] = [
    ['1. Single Entry Point', 'Enter items ONCE on PROFORMA INVOICE rows 17–36 (Requested Description, Offered Description, UoM, Qty). Never insert or sort rows on a single tab.'],
    ['2. Internal Costing & VAT', 'Enter Supplier Unit Cost, Boxes, Box L/W/H (cm), and Kg/Box on PRICE CALC rows 17–36. Set active markup % in PRICE CALC!E3 and internal VAT % in PRICE CALC!E4 (set to 0% to remove VAT).'],
    ['3. Confidential Unit Pricing', 'Unit Price = ROUND(Supplier Cost * (1 + Markup %) * (1 + Internal VAT %), 2). PRICE CALC Column R feeds PI Unit Price directly so end-user documents NEVER show a separate VAT row.'],
    ['4. Freight Calculation', 'Base Freight = MAX(Total CBM * Rate/CBM, Total Gross KG * Rate/KG). Documentation charge (default 10%) is added first, then selected markup % is applied. Shown as a separate line before Grand Total.'],
    ['5. Sequential Numbering', 'Claim numbers exclusively from Master_Registry.xlsx (PREFIX-GG-SEQ-YEAR), starting from #001.'],
    ['6. Issuing & Confidentiality', 'Export values-only .xlsx per document tab. NEVER send PRICE CALC or the master job workbook to customers.'],
    ['7. End-User List (Updated)', 'CONTROLS!A2:A10 contains the confirmed end-user dropdown list including Pharmecore and OTHER.'],
  ];

  guideRows.forEach(([title, desc], idx) => {
    const r = 4 + idx;
    wsHow.getRow(r).height = 30;
    wsHow.getCell(r, 2).value = title;
    wsHow.getCell(r, 2).font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.strongBorderColor } };
    wsHow.getCell(r, 2).alignment = { vertical: 'middle' };
    wsHow.getCell(r, 3).value = desc;
    wsHow.getCell(r, 3).font = { name: fontName, size: 10, color: { argb: 'FF1E293B' } };
    wsHow.getCell(r, 3).alignment = { vertical: 'middle', wrapText: true };
    wsHow.getCell(r, 2).border = { bottom: thinBorder(theme.excel.borderColor) };
    wsHow.getCell(r, 3).border = { bottom: thinBorder(theme.excel.borderColor) };
  });

  // --- SHEET 2: PROFORMA INVOICE (Entry sheet, rows 17-36, Controls K:L, TO in C8, VAT in L5) ---
  const wsPI = wb.addWorksheet('PROFORMA INVOICE', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 16 }],
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  wsPI.columns = [
    { width: 2 },
    { width: 6 },  // B: No.
    { width: 28 }, // C: Requested Desc
    { width: 32 }, // D: Offered Desc
    { width: 9 },  // E: UoM
    { width: 9 },  // F: Qty
    { width: 15 }, // G: Unit Price
    { width: 16 }, // H: Total Price
    { width: 3 },  // I: Spacer
    { width: 3 },  // J: Spacer
    { width: 18 }, // K: Control Labels
    { width: 22 }, // L: Control Inputs
  ];

  styleBannerHeader(
    wb,
    wsPI,
    theme,
    company,
    'PROFORMA INVOICE',
    includeCurrentData && job.issuedDocs.PI && !job.issuedDocs.PI.isVoided ? job.issuedDocs.PI.docNumber : '',
    7
  );

  wsPI.getCell('B8').value = 'TO:';
  wsPI.getCell('B8').font = { name: fontName, size: 10, bold: true, color: { argb: 'FF475569' } };
  wsPI.getCell('C8').value = includeCurrentData ? job.controls.endUserSelect : company.endUsers[0] || 'AIAA (Asmara International Airport)';
  wsPI.getCell('C8').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.inputFont } };
  const lastEndUserRow = Math.max(2, company.endUsers.length + 1);
  wsPI.getCell('C8').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: [`CONTROLS!$A$2:$A$${lastEndUserRow}`],
  };

  wsPI.getCell('B9').value = 'CUSTOM NAME:';
  wsPI.getCell('B9').font = { name: fontName, size: 9, bold: true, color: { argb: 'FF64748B' } };
  wsPI.getCell('C9').value = includeCurrentData ? job.controls.customEndUser : '';
  wsPI.getCell('C9').font = { name: fontName, size: 10, color: { argb: theme.excel.inputFont } };

  wsPI.getCell('B10').value = 'RFQ REF:';
  wsPI.getCell('C10').value = includeCurrentData ? job.controls.rfqRef : '';
  wsPI.getCell('B11').value = 'CUSTOMER PO:';
  wsPI.getCell('C11').value = includeCurrentData ? job.controls.customerPoRef : '';

  // Controls block K2:L7 (End-user safe: NO VAT cell exposed on PI)
  wsPI.getCell('K2').value = 'JOB CONTROLS';
  wsPI.getCell('K2').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
  wsPI.getCell('K2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.accentFill } };
  wsPI.getCell('L2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.accentFill } };

  const piControls: [number, string, string | number][] = [
    [3, 'Currency', includeCurrentData ? job.controls.currency : 'USD'],
    [4, 'Incoterm', includeCurrentData ? job.controls.incoterm : 'CIF'],
    [5, 'Destination', includeCurrentData ? job.controls.destination : 'Massawa Seaport'],
    [6, 'Payment Terms', includeCurrentData ? job.controls.paymentTerms : '50% Advance with PO, 50% Prior to Shipment'],
    [7, 'Lead Time', includeCurrentData ? job.controls.deliveryLeadTime : '4–6 Weeks from PO Confirmation'],
  ];
  piControls.forEach(([rowIdx, lbl, val]) => {
    wsPI.getCell(rowIdx, 11).value = lbl;
    wsPI.getCell(rowIdx, 11).font = { name: fontName, size: 9, bold: true, color: { argb: 'FF475569' } };
    wsPI.getCell(rowIdx, 12).value = val;
    wsPI.getCell(rowIdx, 12).font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.inputFont } };
    wsPI.getCell(rowIdx, 11).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.subHeaderFill } };
    wsPI.getCell(rowIdx, 12).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.subHeaderFill } };
    wsPI.getCell(rowIdx, 11).border = { bottom: thinBorder(theme.excel.borderColor) };
    wsPI.getCell(rowIdx, 12).border = { bottom: thinBorder(theme.excel.borderColor) };
  });

  // Table header on Row 16 so items sit strictly on Rows 17-36
  const piHeaders = ['No.', 'Requested Description', 'Offered Description', 'UoM', 'Qty', 'Unit Price', 'Total Price'];
  wsPI.getRow(16).height = 25;
  piHeaders.forEach((h, i) => {
    const cell = wsPI.getCell(16, i + 2);
    cell.value = h;
    cell.font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
    cell.alignment = { vertical: 'middle', horizontal: i === 0 || i === 3 ? 'center' : i >= 4 ? 'right' : 'left' };
  });

  for (let i = 0; i < 20; i++) {
    const r = 17 + i;
    const item = includeCurrentData ? job.items[i] : null;
    wsPI.getRow(r).height = 22;

    wsPI.getCell(r, 2).value = { formula: `IF(AND(C${r}="",D${r}=""),"",${i + 1})` };
    wsPI.getCell(r, 3).value = item ? item.requestedDescription : '';
    wsPI.getCell(r, 4).value = item ? item.offeredDescription : '';
    wsPI.getCell(r, 5).value = item ? item.uom : '';
    wsPI.getCell(r, 6).value = item && item.qty !== null ? item.qty : null;
    // Unit Price linked strictly from PRICE CALC Col R
    wsPI.getCell(r, 7).value = { formula: `IF('PRICE CALC'!R${r}="","",'PRICE CALC'!R${r})` };
    wsPI.getCell(r, 8).value = { formula: `IF(OR(F${r}="",G${r}=""),"",F${r}*G${r})` };

    for (let c = 2; c <= 8; c++) {
      const cell = wsPI.getCell(r, c);
      const isInput = c >= 3 && c <= 6;
      cell.font = {
        name: fontName,
        size: 10,
        color: { argb: isInput ? theme.excel.inputFont : theme.excel.formulaFont },
      };
      if (i % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.zebraFill } };
      }
      cell.border = { bottom: thinBorder(theme.excel.borderColor) };
      cell.alignment = {
        vertical: 'middle',
        horizontal: c === 2 || c === 5 ? 'center' : c >= 6 ? 'right' : 'left',
        wrapText: c === 3 || c === 4,
      };
    }
    wsPI.getCell(r, 6).numFmt = '#,##0.##';
    wsPI.getCell(r, 7).numFmt = '#,##0.00';
    wsPI.getCell(r, 8).numFmt = '#,##0.00';
  }

  // PI Totals (Rows 38-40 — ZERO VAT row on customer-facing PI)
  const piSummary: [number, string, string, boolean][] = [
    [38, 'GOODS SUBTOTAL', 'SUM(H17:H36)', false],
    [39, 'FREIGHT & DOCUMENTATION', "'PRICE CALC'!R41", false],
    [40, 'GRAND TOTAL', 'H38+H39', true],
  ];
  piSummary.forEach(([r, label, formula, isGrand]) => {
    wsPI.getRow(r).height = isGrand ? 26 : 21;
    wsPI.mergeCells(r, 5, r, 7);
    wsPI.getCell(r, 5).value = label;
    wsPI.getCell(r, 5).alignment = { horizontal: 'right', vertical: 'middle' };
    wsPI.getCell(r, 8).value = { formula };
    wsPI.getCell(r, 8).numFmt = '#,##0.00';
    wsPI.getCell(r, 8).alignment = { horizontal: 'right', vertical: 'middle' };
    for (let c = 5; c <= 8; c++) {
      wsPI.getCell(r, c).font = { name: fontName, size: isGrand ? 11 : 10, bold: isGrand, color: { argb: theme.excel.strongBorderColor } };
      if (isGrand) {
        wsPI.getCell(r, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.totalFill } };
        wsPI.getCell(r, c).border = { top: thinBorder(theme.excel.strongBorderColor), bottom: mediumBorder(theme.excel.strongBorderColor) };
      } else {
        wsPI.getCell(r, c).border = { bottom: thinBorder(theme.excel.borderColor) };
      }
    }
  });

  // --- SHEET 3: PRICE CALC (INTERNAL ONLY, E3 = Markup %, Col R = Active Unit Price, R41 = Customer Freight) ---
  const wsPC = wb.addWorksheet('PRICE CALC', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 16 }],
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  wsPC.columns = [
    { width: 2 },  // A spacer
    { width: 5 },  // B No
    { width: 24 }, // C Offered Desc
    { width: 7 },  // D Qty
    { width: 12 }, // E Supplier Unit Cost (and E3 Markup %)
    { width: 13 }, // F Supplier Total
    { width: 8 },  // G Boxes
    { width: 8 },  // H L cm
    { width: 8 },  // I W cm
    { width: 8 },  // J H cm
    { width: 10 }, // K CBM
    { width: 9 },  // L Kg/Box
    { width: 11 }, // M Gross KG
    { width: 12 }, // N 10% Price
    { width: 12 }, // O 15% Price
    { width: 12 }, // P 18% Price
    { width: 12 }, // Q 20% Price
    { width: 14 }, // R ACTIVE UNIT PRICE (Feeds PI!)
  ];

  wsPC.mergeCells('B2:R2');
  wsPC.getCell('B2').value = 'PRICE CALC — INTERNAL COSTING & MARKUP SCENARIOS (CONFIDENTIAL — DO NOT SEND TO CUSTOMER)';
  wsPC.getCell('B2').font = { name: fontName, size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  wsPC.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF991B1B' } };
  wsPC.getCell('B2').alignment = { vertical: 'middle', indent: 1 };
  wsPC.getRow(2).height = 28;

  // Controls Row 3: E3 = Selected Markup %, I3 = Rate/CBM, L3 = Rate/KG, O3 = Doc Charge %
  wsPC.getRow(3).height = 22;
  wsPC.getCell('B3').value = 'ACTIVE MARKUP (E3):';
  wsPC.mergeCells('B3:D3');
  wsPC.getCell('B3').font = { name: fontName, size: 10, bold: true };
  wsPC.getCell('E3').value = includeCurrentData ? job.freight.selectedMarkupPct : 0.2;
  wsPC.getCell('E3').numFmt = '0.##%';
  wsPC.getCell('E3').font = { name: fontName, size: 11, bold: true, color: { argb: theme.excel.inputFont } };
  wsPC.getCell('E3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };

  wsPC.getCell('H3').value = 'Rate/CBM:';
  wsPC.getCell('I3').value = includeCurrentData ? job.freight.ratePerCbm : 185;
  wsPC.getCell('I3').numFmt = '#,##0.00';
  wsPC.getCell('I3').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.inputFont } };

  wsPC.getCell('K3').value = 'Rate/KG:';
  wsPC.getCell('L3').value = includeCurrentData ? job.freight.ratePerKg : 2.4;
  wsPC.getCell('L3').numFmt = '#,##0.00';
  wsPC.getCell('L3').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.inputFont } };

  wsPC.getCell('N3').value = 'Doc Fee %:';
  wsPC.getCell('O3').value = includeCurrentData ? job.freight.docChargePct : 0.1;
  wsPC.getCell('O3').numFmt = '0.##%';
  wsPC.getCell('O3').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.inputFont } };

  // Controls Row 4: E4 = Internal VAT % (Baked into Unit Prices Cols N..R; set 0% to remove VAT)
  const effectiveVat = includeCurrentData
    ? job.controls.includeVat !== false
      ? job.controls.vatRate
      : 0
    : 0.05;
  wsPC.getRow(4).height = 22;
  wsPC.getCell('B4').value = 'INTERNAL VAT (E4 — 0% to remove):';
  wsPC.mergeCells('B4:D4');
  wsPC.getCell('B4').font = { name: fontName, size: 9, bold: true, color: { argb: 'FF991B1B' } };
  wsPC.getCell('E4').value = effectiveVat;
  wsPC.getCell('E4').numFmt = '0.##%';
  wsPC.getCell('E4').font = { name: fontName, size: 11, bold: true, color: { argb: theme.excel.inputFont } };
  wsPC.getCell('E4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
  wsPC.mergeCells('H4:R4');
  wsPC.getCell('H4').value =
    'Internal VAT (E4) is folded directly into Unit Prices (Cols N–R). End-user sheets (PI / QT / CI) NEVER show a VAT line.';
  wsPC.getCell('H4').font = { name: fontName, size: 9, italic: true, color: { argb: 'FF64748B' } };

  const pcHeaders = [
    'No.', 'Offered Description', 'Qty', 'Supplier Cost', 'Supplier Total',
    'Boxes', 'L (cm)', 'W (cm)', 'H (cm)', 'CBM', 'Kg/Box', 'Gross KG',
    '10% Price', '15% Price', '18% Price', '20% Price', 'Active Unit Price (R)',
  ];
  wsPC.getRow(16).height = 28;
  pcHeaders.forEach((h, i) => {
    const cell = wsPC.getCell(16, i + 2);
    cell.value = h;
    cell.font = { name: fontName, size: 9, bold: true, color: { argb: theme.excel.headerFont } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: i === 16 ? theme.excel.accentFill : theme.excel.headerFill } };
    cell.alignment = { vertical: 'middle', horizontal: i <= 1 ? 'left' : 'right', wrapText: true };
  });

  for (let i = 0; i < 20; i++) {
    const r = 17 + i;
    const item = includeCurrentData ? job.items[i] : null;
    wsPC.getRow(r).height = 22;

    wsPC.getCell(r, 2).value = { formula: `IF('PROFORMA INVOICE'!B${r}="","",'PROFORMA INVOICE'!B${r})` };
    wsPC.getCell(r, 3).value = { formula: `IF('PROFORMA INVOICE'!D${r}="","",'PROFORMA INVOICE'!D${r})` };
    wsPC.getCell(r, 4).value = { formula: `IF('PROFORMA INVOICE'!F${r}="","",'PROFORMA INVOICE'!F${r})` };
    wsPC.getCell(r, 5).value = item && item.supplierUnitCost !== null ? item.supplierUnitCost : null;
    wsPC.getCell(r, 6).value = { formula: `IF(OR(D${r}="",E${r}=""),"",D${r}*E${r})` };
    wsPC.getCell(r, 7).value = item && item.boxes !== null ? item.boxes : null;
    wsPC.getCell(r, 8).value = item && item.boxLengthCm !== null ? item.boxLengthCm : null;
    wsPC.getCell(r, 9).value = item && item.boxWidthCm !== null ? item.boxWidthCm : null;
    wsPC.getCell(r, 10).value = item && item.boxHeightCm !== null ? item.boxHeightCm : null;
    wsPC.getCell(r, 11).value = { formula: `IF(OR(G${r}="",H${r}="",I${r}="",J${r}=""),"",(G${r}*H${r}*I${r}*J${r})/1000000)` };
    wsPC.getCell(r, 12).value = item && item.kgPerBox !== null ? item.kgPerBox : null;
    wsPC.getCell(r, 13).value = { formula: `IF(OR(G${r}="",L${r}=""),"",G${r}*L${r})` };
    // Markup + Internal VAT folded into Unit Price: ROUND(Cost * (1 + Markup) * (1 + $E$4), 2)
    wsPC.getCell(r, 14).value = { formula: `IF(E${r}="","",ROUND(E${r}*1.10*(1+$E$4),2))` };
    wsPC.getCell(r, 15).value = { formula: `IF(E${r}="","",ROUND(E${r}*1.15*(1+$E$4),2))` };
    wsPC.getCell(r, 16).value = { formula: `IF(E${r}="","",ROUND(E${r}*1.18*(1+$E$4),2))` };
    wsPC.getCell(r, 17).value = { formula: `IF(E${r}="","",ROUND(E${r}*1.20*(1+$E$4),2))` };
    wsPC.getCell(r, 18).value = { formula: `IF(E${r}="","",ROUND(E${r}*(1+$E$3)*(1+$E$4),2))` };

    for (let c = 2; c <= 18; c++) {
      const cell = wsPC.getCell(r, c);
      const isInput = c === 5 || (c >= 7 && c <= 10) || c === 12;
      cell.font = {
        name: fontName,
        size: 9,
        bold: c === 18,
        color: { argb: isInput ? theme.excel.inputFont : theme.excel.formulaFont },
      };
      if (c === 18) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.subHeaderFill } };
      } else if (i % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.zebraFill } };
      }
      cell.border = { bottom: thinBorder(theme.excel.borderColor) };
      cell.alignment = { vertical: 'middle', horizontal: c <= 3 ? 'left' : 'right' };
    }
    [5, 6, 12, 13, 14, 15, 16, 17, 18].forEach((c) => (wsPC.getCell(r, c).numFmt = '#,##0.00'));
    wsPC.getCell(r, 11).numFmt = '0.0000';
  }

  // Freight Summary Block on PRICE CALC rows 38-41 (R41 feeds PI!H40)
  wsPC.getCell('F37').value = { formula: 'SUM(F17:F36)' };
  wsPC.getCell('F37').numFmt = '#,##0.00';
  wsPC.getCell('K37').value = { formula: 'SUM(K17:K36)' };
  wsPC.getCell('K37').numFmt = '0.0000';
  wsPC.getCell('M37').value = { formula: 'SUM(M17:M36)' };
  wsPC.getCell('M37').numFmt = '#,##0.00';

  const freightRows: [number, string, string][] = [
    [38, 'Base Freight = MAX(CBM*Rate, KG*Rate)', 'MAX(K37*I3, M37*L3)'],
    [39, 'Documentation Charge (Default 10%)', 'R38*O3'],
    [40, 'Freight + Doc Charge Subtotal', 'R38+R39'],
    [41, 'Customer Freight (With Selected Markup E3)', 'ROUND(R40*(1+E3),2)'],
  ];
  freightRows.forEach(([r, lbl, formula]) => {
    wsPC.getRow(r).height = 22;
    wsPC.mergeCells(r, 14, r, 17);
    wsPC.getCell(r, 14).value = lbl;
    wsPC.getCell(r, 14).alignment = { horizontal: 'right', vertical: 'middle' };
    wsPC.getCell(r, 14).font = { name: fontName, size: 9, bold: r === 41 };
    wsPC.getCell(r, 18).value = { formula };
    wsPC.getCell(r, 18).numFmt = '#,##0.00';
    wsPC.getCell(r, 18).font = { name: fontName, size: 10, bold: r === 41, color: { argb: theme.excel.strongBorderColor } };
    if (r === 41) {
      wsPC.getCell(r, 18).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.totalFill } };
    }
  });

  // Append any custom columns dynamically to PI (after Total Price so core cols B..H + L5 + PRICE CALC stay 100% intact)
  const templateCustomCols = includeCurrentData ? job.customColumns || [] : [];
  if (templateCustomCols.length > 0) {
    templateCustomCols.forEach((cc, cIdx) => {
      const colNum = 14 + cIdx; // Col N onwards on PI (right of K:L controls)
      wsPI.getColumn(colNum).width = 16;
      const hdr = wsPI.getCell(16, colNum);
      hdr.value = cc.label;
      hdr.font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
      hdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.accentFill } };
      hdr.alignment = { vertical: 'middle', horizontal: 'left' };
      for (let i = 0; i < 20; i++) {
        const r = 17 + i;
        const val = job.items[i]?.customFields?.[cc.id] || '';
        const cell = wsPI.getCell(r, colNum);
        cell.value = val;
        cell.font = { name: fontName, size: 10, color: { argb: theme.excel.inputFont } };
        if (i % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.zebraFill } };
        cell.border = { bottom: thinBorder(theme.excel.borderColor) };
      }
    });
  }

  // Map each custom column ID to its fixed storage letter on PROFORMA INVOICE (Col N+)
  const customPiColLetter = new Map<string, string>();
  templateCustomCols.forEach((cc, cIdx) => {
    customPiColLetter.set(cc.id, colToLetter(14 + cIdx));
  });

  // --- SHEETS 4..8: QUOTATION, SUPPLIER PO, PACKING LIST, DELIVERY NOTE, COMMERCIAL INVOICE ---
  const downstreamDocs: DocPrefix[] = ['QT', 'PO', 'PL', 'DN', 'CI'];
  for (const prefix of downstreamDocs) {
    const meta = DOC_META[prefix];
    const isLogistics = prefix === 'PL' || prefix === 'DN';
    const isPO = prefix === 'PO';
    const extraCount = templateCustomCols.length;
    const ws = wb.addWorksheet(meta.sheetName, {
      views: [{ showGridLines: false, state: 'frozen', ySplit: 16 }],
      pageSetup: { paperSize: 9, orientation: isLogistics || extraCount > 0 ? 'landscape' : 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    interface DsColSpec {
      key: string;
      header: string;
      width: number;
      align: 'left' | 'center' | 'right';
      isInput?: boolean;
      numFmt?: string;
      piCustomLetter?: string;
    }

    const dsCustom = (pos: CustomColumnPosition): DsColSpec[] =>
      filterCustomCols(templateCustomCols, pos).map((cc) => ({
        key: `custom:${cc.id}`,
        header: cc.label,
        width: 16,
        align: 'left',
        piCustomLetter: customPiColLetter.get(cc.id),
      }));

    const dsSpecs: DsColSpec[] = [
      { key: 'no', header: 'No.', width: 6, align: 'center' },
      ...dsCustom('after_no'),
      { key: 'requested', header: 'Requested Description', width: isLogistics ? 26 : 28, align: 'left' },
      ...dsCustom('after_requested'),
      { key: 'offered', header: 'Offered Description', width: isLogistics ? 30 : 32, align: 'left' },
      ...dsCustom('after_offered'),
      { key: 'uom', header: 'UoM', width: isLogistics ? 8 : 9, align: 'center' },
      { key: 'qty', header: 'Qty', width: isLogistics ? 8 : 9, align: 'right' },
      ...dsCustom('after_qty'),
      ...(isLogistics
        ? ([
            { key: 'carton', header: 'Carton No.', width: 15, align: 'center', isInput: true },
            { key: 'netWt', header: 'Net Wt (KG)', width: 14, align: 'right', isInput: true, numFmt: '#,##0.00' },
            { key: 'grossWt', header: 'Gross Wt (KG)', width: 14, align: 'right', isInput: true, numFmt: '#,##0.00' },
          ] as DsColSpec[])
        : isPO
        ? ([
            { key: 'price', header: 'Unit Cost', width: 15, align: 'right', numFmt: '#,##0.00' },
            { key: 'total', header: 'Total Cost', width: 16, align: 'right', numFmt: '#,##0.00' },
          ] as DsColSpec[])
        : ([
            { key: 'price', header: 'Unit Price', width: 15, align: 'right', numFmt: '#,##0.00' },
            { key: 'total', header: 'Total Price', width: 16, align: 'right', numFmt: '#,##0.00' },
          ] as DsColSpec[])),
      ...dsCustom('end'),
    ];

    ws.columns = [{ width: 2 }, ...dsSpecs.map((s) => ({ width: s.width }))];

    const docNo =
      includeCurrentData && job.issuedDocs[prefix] && !job.issuedDocs[prefix]!.isVoided
        ? job.issuedDocs[prefix]!.docNumber
        : '';
    styleBannerHeader(wb, ws, theme, company, meta.title, docNo, dsSpecs.length);

    ws.getCell('B8').value = meta.recipientLabel;
    ws.getCell('B8').font = { name: fontName, size: 9, bold: true, color: { argb: 'FF64748B' } };
    ws.getCell('C8').value = isPO
      ? includeCurrentData ? job.controls.supplierName : ''
      : { formula: `IF('PROFORMA INVOICE'!C8="OTHER",'PROFORMA INVOICE'!C9,'PROFORMA INVOICE'!C8)` };
    ws.getCell('C8').font = { name: fontName, size: 10, bold: true };

    const dsColNum = (key: string) => dsSpecs.findIndex((s) => s.key === key) + 2;
    const qtyCol = dsColNum('qty');
    const priceCol = dsColNum('price');
    const totalCol = dsColNum('total');
    const qtyLet = colToLetter(qtyCol);
    const priceLet = priceCol >= 2 ? colToLetter(priceCol) : '';
    const totalLet = totalCol >= 2 ? colToLetter(totalCol) : '';

    ws.getRow(16).height = 25;
    dsSpecs.forEach((spec, i) => {
      const cell = ws.getCell(16, i + 2);
      cell.value = spec.header;
      cell.font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
      cell.alignment = { vertical: 'middle', horizontal: spec.align };
    });

    for (let i = 0; i < 20; i++) {
      const r = 17 + i;
      const item = includeCurrentData ? job.items[i] : null;
      ws.getRow(r).height = 22;

      dsSpecs.forEach((spec, idx) => {
        const c = idx + 2;
        const cell = ws.getCell(r, c);
        if (spec.key === 'no') {
          cell.value = { formula: `IF('PROFORMA INVOICE'!B${r}="","",'PROFORMA INVOICE'!B${r})` };
        } else if (spec.piCustomLetter) {
          cell.value = {
            formula: `IF('PROFORMA INVOICE'!${spec.piCustomLetter}${r}="","",'PROFORMA INVOICE'!${spec.piCustomLetter}${r})`,
          };
        } else if (spec.key === 'requested') {
          cell.value = { formula: `IF('PROFORMA INVOICE'!C${r}="","",'PROFORMA INVOICE'!C${r})` };
        } else if (spec.key === 'offered') {
          cell.value = { formula: `IF('PROFORMA INVOICE'!D${r}="","",'PROFORMA INVOICE'!D${r})` };
        } else if (spec.key === 'uom') {
          cell.value = { formula: `IF('PROFORMA INVOICE'!E${r}="","",'PROFORMA INVOICE'!E${r})` };
        } else if (spec.key === 'qty') {
          cell.value = { formula: `IF('PROFORMA INVOICE'!F${r}="","",'PROFORMA INVOICE'!F${r})` };
        } else if (spec.key === 'carton') {
          cell.value = item ? item.cartonNo : '';
        } else if (spec.key === 'netWt') {
          cell.value = item && item.netWtKg !== null ? item.netWtKg : null;
        } else if (spec.key === 'grossWt') {
          cell.value = item && item.grossWtKg !== null ? item.grossWtKg : null;
        } else if (spec.key === 'price') {
          cell.value = isPO
            ? { formula: `IF('PRICE CALC'!E${r}="","",'PRICE CALC'!E${r})` }
            : { formula: `IF('PROFORMA INVOICE'!G${r}="","",'PROFORMA INVOICE'!G${r})` };
        } else if (spec.key === 'total') {
          cell.value = { formula: `IF(OR(${qtyLet}${r}="",${priceLet}${r}=""),"",${qtyLet}${r}*${priceLet}${r})` };
        }

        if (spec.numFmt) cell.numFmt = spec.numFmt;
        cell.font = {
          name: fontName,
          size: 10,
          color: { argb: spec.isInput ? theme.excel.inputFont : theme.excel.formulaFont },
        };
        if (i % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.zebraFill } };
        cell.border = { bottom: thinBorder(theme.excel.borderColor) };
        cell.alignment = { vertical: 'middle', horizontal: spec.align };
      });
    }

    if (!isLogistics && !isPO) {
      ws.getCell(38, priceCol).value = 'SUBTOTAL';
      ws.getCell(38, totalCol).value = { formula: `SUM(${totalLet}17:${totalLet}36)` };
      ws.getCell(39, priceCol).value = 'FREIGHT';
      ws.getCell(39, totalCol).value = { formula: "'PROFORMA INVOICE'!H39" };
      ws.getCell(40, priceCol).value = 'GRAND TOTAL';
      ws.getCell(40, totalCol).value = { formula: `${totalLet}38+${totalLet}39` };
      [38, 39, 40].forEach((r) => {
        ws.getCell(r, totalCol).numFmt = '#,##0.00';
        ws.getCell(r, priceCol).font = { name: fontName, size: 10, bold: r === 40 };
        ws.getCell(r, totalCol).font = { name: fontName, size: 10, bold: r === 40 };
      });
    } else if (isPO) {
      ws.getCell(38, priceCol).value = 'TOTAL PO COST';
      ws.getCell(38, totalCol).value = { formula: `SUM(${totalLet}17:${totalLet}36)` };
      ws.getCell(38, totalCol).numFmt = '#,##0.00';
      ws.getCell(38, priceCol).font = { name: fontName, size: 10, bold: true };
      ws.getCell(38, totalCol).font = { name: fontName, size: 10, bold: true };
    }
  }

  // --- SHEET 9: CONTROLS (End Users + Server Paths + Lossless Job Metadata & Custom Column Layout) ---
  const wsCtrl = wb.addWorksheet('CONTROLS', { views: [{ showGridLines: true }] });
  wsCtrl.columns = [
    { width: 38 },
    { width: 4 },
    { width: 26 },
    { width: 46 },
    { width: 4 },
    { width: 22 },
    { width: 22 },
    { width: 18 },
  ];
  wsCtrl.getCell('A1').value = `END USERS (${company.endUsers.length}-ITEM LIST)`;
  wsCtrl.getCell('A1').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
  wsCtrl.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
  company.endUsers.forEach((u, idx) => {
    wsCtrl.getCell(2 + idx, 1).value = u;
  });

  wsCtrl.getCell('C1').value = 'SETTING';
  wsCtrl.getCell('D1').value = 'VALUE';
  wsCtrl.getCell('C1').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
  wsCtrl.getCell('D1').font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
  wsCtrl.getCell('C1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
  wsCtrl.getCell('D1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };

  wsCtrl.getCell('C2').value = 'ctl_RegistryPath';
  wsCtrl.getCell('D2').value = '\\\\Server\\Procurement\\Registry\\Master_Registry.xlsx';
  wsCtrl.getCell('C3').value = 'ctl_IssuedFolder';
  wsCtrl.getCell('D3').value = '\\\\Server\\Procurement\\Issued\\';
  wsCtrl.getCell('C4').value = 'Style Variant';
  wsCtrl.getCell('D4').value = theme.name;

  if (includeCurrentData) {
    const metaRows: [string, string | number][] = [
      ['meta_jobName', job.jobName],
      ['meta_endUserContact', job.controls.endUserContact],
      ['meta_endUserAddress', job.controls.endUserAddress],
      ['meta_rfqRef', job.controls.rfqRef],
      ['meta_customerPoRef', job.controls.customerPoRef],
      ['meta_paymentTerms', job.controls.paymentTerms],
      ['meta_deliveryLeadTime', job.controls.deliveryLeadTime],
      ['meta_supplierName', job.controls.supplierName],
      ['meta_supplierAddress', job.controls.supplierAddress],
      ['meta_supplierRef', job.controls.supplierRef],
      ['meta_shipDate', job.controls.shipDate],
      ['meta_portOfLoading', job.controls.portOfLoading],
      ['meta_deliveryDate', job.controls.deliveryDate],
      ['meta_transportRef', job.controls.transportRef],
      ['meta_receivedBy', job.controls.receivedBy],
      ['meta_remarks', job.controls.remarks],
      ['meta_authorizedSignatory', job.controls.authorizedSignatory],
      ['meta_counterpartySignatory', job.controls.counterpartySignatory],
      ['meta_includeVat', job.controls.includeVat !== false ? 'true' : 'false'],
      ['meta_vatRate', job.controls.vatRate],
      ['meta_columnOrder', job.columnOrder ? JSON.stringify(job.columnOrder) : ''],
    ];
    metaRows.forEach(([k, v], idx) => {
      wsCtrl.getCell(6 + idx, 3).value = k;
      wsCtrl.getCell(6 + idx, 4).value = v;
    });
  }

  // Persist Custom Column layout (ID, Label, Position) so re-import restores exact positions
  wsCtrl.getCell('F1').value = 'CUSTOM_COL_ID';
  wsCtrl.getCell('G1').value = 'CUSTOM_COL_LABEL';
  wsCtrl.getCell('H1').value = 'CUSTOM_COL_POSITION';
  ['F1', 'G1', 'H1'].forEach((ref) => {
    wsCtrl.getCell(ref).font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
    wsCtrl.getCell(ref).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
  });
  templateCustomCols.forEach((cc, idx) => {
    wsCtrl.getCell(2 + idx, 6).value = cc.id;
    wsCtrl.getCell(2 + idx, 7).value = cc.label;
    wsCtrl.getCell(2 + idx, 8).value = cc.position || 'after_no';
  });

  const buffer = await wb.xlsx.writeBuffer();
  const filename = includeCurrentData
    ? `GG_Job_${job.jobName.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 28)}_${theme.id}.xlsx`
    : `GG_Job_Template_${theme.id}.xlsx`;
  await triggerDownload(buffer as ArrayBuffer, filename);
  return filename;
}

/**
 * 3. MASTER_REGISTRY.XLSX EXPORT (REGISTRY + PREFIXES sheets)
 */
export async function exportMasterRegistryXlsx(
  entries: RegistryEntry[],
  themeOrId: ThemeVariantId | ThemePalette,
  company: CompanyProfile
): Promise<string> {
  const theme = resolveExcelTheme(themeOrId);
  const fontName = theme.fontName || 'Calibri';
  const wb = new ExcelJS.Workbook();
  wb.creator = company.name;
  wb.created = new Date();

  const ws = wb.addWorksheet('REGISTRY', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 4 }],
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  ws.columns = [
    { width: 2 },
    { width: 10 }, // Seq
    { width: 18 }, // Document No
    { width: 10 }, // Prefix
    { width: 26 }, // Customer / Supplier
    { width: 28 }, // Job Reference
    { width: 10 }, // Currency
    { width: 16 }, // Amount
    { width: 18 }, // Issued By
    { width: 18 }, // Issued Date
    { width: 12 }, // Status
  ];

  ws.mergeCells('B2:K2');
  ws.getCell('B2').value = `${company.name} — MASTER DOCUMENT REGISTRY`;
  ws.getCell('B2').font = { name: fontName, size: 14, bold: true, color: { argb: theme.excel.headerFont } };
  ws.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.headerFill } };
  ws.getCell('B2').alignment = { vertical: 'middle', indent: 1 };
  ws.getRow(2).height = 28;

  const cols = ['Seq', 'Document No.', 'Prefix', 'Customer / Party', 'Job Reference', 'Currency', 'Amount', 'Issued By', 'Date', 'Status'];
  ws.getRow(4).height = 24;
  cols.forEach((h, i) => {
    const cell = ws.getCell(4, i + 2);
    cell.value = h;
    cell.font = { name: fontName, size: 10, bold: true, color: { argb: theme.excel.headerFont } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.accentFill } };
    cell.alignment = { vertical: 'middle', horizontal: i === 0 || i === 2 || i === 5 || i === 9 ? 'center' : i === 6 ? 'right' : 'left' };
  });

  const sorted = [...entries].sort((a, b) => a.sequence - b.sequence);
  sorted.forEach((entry, idx) => {
    const r = 5 + idx;
    ws.getRow(r).height = 22;
    const vals = [
      entry.sequence,
      entry.docNumber,
      entry.prefix,
      entry.customer,
      entry.jobName,
      entry.currency,
      entry.amount,
      entry.issuedBy,
      entry.issuedAt.slice(0, 10),
      entry.status,
    ];
    vals.forEach((v, cIdx) => {
      const cell = ws.getCell(r, cIdx + 2);
      cell.value = v;
      cell.font = {
        name: fontName,
        size: 10,
        bold: cIdx === 1 || cIdx === 9,
        color: { argb: entry.status === 'VOID' ? 'FFDC2626' : 'FF0F172A' },
      };
      if (idx % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.excel.zebraFill } };
      cell.border = { bottom: thinBorder(theme.excel.borderColor) };
      cell.alignment = {
        vertical: 'middle',
        horizontal: cIdx === 0 || cIdx === 2 || cIdx === 5 || cIdx === 9 ? 'center' : cIdx === 6 ? 'right' : 'left',
      };
      if (cIdx === 6) cell.numFmt = '#,##0.00';
    });
  });

  // PREFIXES sheet
  const wsPref = wb.addWorksheet('PREFIXES');
  wsPref.columns = [{ width: 12 }, { width: 28 }, { width: 42 }];
  wsPref.getRow(1).values = ['Prefix', 'Document Type', 'Format'];
  wsPref.getRow(1).font = { bold: true };
  (['PI', 'QT', 'PO', 'CI', 'PL', 'DN'] as DocPrefix[]).forEach((p, idx) => {
    wsPref.getRow(2 + idx).values = [p, DOC_META[p].title, `${p}-GG-SEQ-YEAR`];
  });

  const buffer = await wb.xlsx.writeBuffer();
  const filename = 'Master_Registry.xlsx';
  await triggerDownload(buffer as ArrayBuffer, filename);
  return filename;
}

/**
 * 4. IMPORT / OPEN .XLSX DIRECTLY INSIDE THE APP
 * Supports loading either a Grand Group Job workbook (PROFORMA INVOICE + PRICE CALC + PL/DN)
 * OR a Master_Registry.xlsx file to synchronize sequence numbers offline.
 */
function cellToString(val: ExcelJS.CellValue): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if ('result' in val && val.result !== undefined && val.result !== null) {
      return String(val.result).trim();
    }
    if ('richText' in val && Array.isArray(val.richText)) {
      return val.richText.map((t) => t.text).join('').trim();
    }
    if ('text' in val && typeof val.text === 'string') {
      return val.text.trim();
    }
    return '';
  }
  return String(val).trim();
}

function cellToNumber(val: ExcelJS.CellValue): number | null {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return Number.isNaN(val) ? null : val;
  if (typeof val === 'object' && 'result' in val) {
    const res = val.result;
    if (typeof res === 'number' && !Number.isNaN(res)) return res;
    if (typeof res === 'string' && res.trim() !== '') {
      const n = Number(res.replace(/,/g, ''));
      return Number.isNaN(n) ? null : n;
    }
    return null;
  }
  const n = Number(String(val).replace(/,/g, '').trim());
  return Number.isNaN(n) ? null : n;
}

export type WorkbookImportResult =
  | { kind: 'registry'; entries: RegistryEntry[]; importedCount: number }
  | { kind: 'job'; job: ProcurementJob; itemCount: number }
  | { kind: 'supplier_quote'; quote: ParsedSupplierQuote };

export async function importWorkbookFromXlsx(
  file: File,
  knownEndUsers: string[] = END_USER_OPTIONS
): Promise<WorkbookImportResult> {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  // 1. Check if it's a Master_Registry workbook
  const regSheet = wb.getWorksheet('REGISTRY');
  if (regSheet && !wb.getWorksheet('PROFORMA INVOICE')) {
    const entries: RegistryEntry[] = [];
    regSheet.eachRow((row, rowIdx) => {
      if (rowIdx < 5) return;
      const seq = cellToNumber(row.getCell(2).value);
      const docNo = cellToString(row.getCell(3).value);
      const prefix = cellToString(row.getCell(4).value).toUpperCase() as DocPrefix;
      if (seq && docNo && ['PI', 'QT', 'PO', 'CI', 'PL', 'DN'].includes(prefix)) {
        const statusRaw = cellToString(row.getCell(11).value).toUpperCase();
        entries.push({
          sequence: Math.floor(seq),
          docNumber: docNo,
          prefix,
          year: Number(docNo.split('-').pop()) || new Date().getFullYear(),
          jobId: `imported-${seq}`,
          jobName: cellToString(row.getCell(6).value) || 'Imported Registry Entry',
          customer: cellToString(row.getCell(5).value) || '—',
          currency: cellToString(row.getCell(7).value) || 'USD',
          amount: cellToNumber(row.getCell(8).value) ?? 0,
          issuedBy: cellToString(row.getCell(9).value) || 'Excel Import',
          issuedAt: cellToString(row.getCell(10).value) || new Date().toISOString().slice(0, 10),
          status: statusRaw === 'VOID' ? 'VOID' : 'ISSUED',
        });
      }
    });
    if (entries.length > 0) {
      return { kind: 'registry', entries, importedCount: entries.length };
    }
  }

  // 2. If this workbook is NOT a Grand Group template, parse it as an any-format Supplier Quote
  const wsPI = wb.getWorksheet('PROFORMA INVOICE');
  const wsPC = wb.getWorksheet('PRICE CALC');
  const wsPL = wb.getWorksheet('PACKING LIST');
  const wsCtrl = wb.getWorksheet('CONTROLS');

  if (!wsPI && !wsPC) {
    const quote = parseGenericSupplierXlsx(wb, file.name);
    return { kind: 'supplier_quote', quote };
  }

  // 3. Parse as a Grand Group Job workbook
  const cleanName = file.name.replace(/\.xlsx?$/i, '').replace(/[_-]+/g, ' ');
  const job = createBlankJob(cleanName);

  // Read lossless metadata and custom column definitions from CONTROLS sheet if present
  const validPositions: CustomColumnPosition[] = [
    'after_no',
    'after_requested',
    'after_offered',
    'after_qty',
    'end',
  ];
  const ctrlColLayout = new Map<string, { id: string; position: CustomColumnPosition }>();
  let hasMetaIncludeVat = false;

  if (wsCtrl) {
    const metaMap = new Map<string, ExcelJS.CellValue>();
    for (let r = 6; r <= 30; r++) {
      const key = cellToString(wsCtrl.getCell(r, 3).value);
      if (key.startsWith('meta_')) {
        metaMap.set(key, wsCtrl.getCell(r, 4).value);
      }
    }
    const getStr = (k: string) =>
      metaMap.has(k) ? cellToString(metaMap.get(k) ?? null) : undefined;

    const mJobName = getStr('meta_jobName');
    if (mJobName) job.jobName = mJobName;
    if (metaMap.has('meta_endUserContact')) job.controls.endUserContact = getStr('meta_endUserContact') ?? '';
    if (metaMap.has('meta_endUserAddress')) job.controls.endUserAddress = getStr('meta_endUserAddress') ?? '';
    if (metaMap.has('meta_rfqRef')) job.controls.rfqRef = getStr('meta_rfqRef') ?? '';
    if (metaMap.has('meta_customerPoRef')) job.controls.customerPoRef = getStr('meta_customerPoRef') ?? '';
    if (metaMap.has('meta_paymentTerms')) job.controls.paymentTerms = getStr('meta_paymentTerms') ?? '';
    if (metaMap.has('meta_deliveryLeadTime')) job.controls.deliveryLeadTime = getStr('meta_deliveryLeadTime') ?? '';
    if (metaMap.has('meta_supplierName')) job.controls.supplierName = getStr('meta_supplierName') ?? '';
    if (metaMap.has('meta_supplierAddress')) job.controls.supplierAddress = getStr('meta_supplierAddress') ?? '';
    if (metaMap.has('meta_supplierRef')) job.controls.supplierRef = getStr('meta_supplierRef') ?? '';
    if (metaMap.has('meta_shipDate')) job.controls.shipDate = getStr('meta_shipDate') ?? '';
    if (metaMap.has('meta_portOfLoading')) job.controls.portOfLoading = getStr('meta_portOfLoading') ?? '';
    if (metaMap.has('meta_deliveryDate')) job.controls.deliveryDate = getStr('meta_deliveryDate') ?? '';
    if (metaMap.has('meta_transportRef')) job.controls.transportRef = getStr('meta_transportRef') ?? '';
    if (metaMap.has('meta_receivedBy')) job.controls.receivedBy = getStr('meta_receivedBy') ?? '';
    if (metaMap.has('meta_remarks')) job.controls.remarks = getStr('meta_remarks') ?? '';
    if (metaMap.has('meta_authorizedSignatory')) {
      job.controls.authorizedSignatory = getStr('meta_authorizedSignatory') ?? '';
    }
    if (metaMap.has('meta_counterpartySignatory')) {
      job.controls.counterpartySignatory = getStr('meta_counterpartySignatory') ?? '';
    }
    if (metaMap.has('meta_includeVat')) {
      hasMetaIncludeVat = true;
      job.controls.includeVat = getStr('meta_includeVat')?.toLowerCase() !== 'false';
    }
    if (metaMap.has('meta_vatRate')) {
      const mv = cellToNumber(metaMap.get('meta_vatRate') ?? null);
      if (mv !== null && mv >= 0 && mv <= 1) {
        job.controls.vatRate = mv;
      }
    }
    if (metaMap.has('meta_columnOrder')) {
      const rawOrder = getStr('meta_columnOrder');
      if (rawOrder) {
        try {
          const parsedOrder = JSON.parse(rawOrder);
          if (Array.isArray(parsedOrder)) {
            job.columnOrder = parsedOrder.filter((k): k is string => typeof k === 'string');
          }
        } catch {
          // Ignore invalid JSON
        }
      }
    }

    for (let r = 2; r <= 5; r++) {
      const colId = cellToString(wsCtrl.getCell(r, 6).value);
      const colLabel = cellToString(wsCtrl.getCell(r, 7).value);
      const colPos = cellToString(wsCtrl.getCell(r, 8).value) as CustomColumnPosition;
      if (colLabel) {
        ctrlColLayout.set(colLabel.toLowerCase(), {
          id: colId || `col_${colLabel.toLowerCase().replace(/[^a-z0-9]+/g, '_')}_${r}`,
          position: validPositions.includes(colPos) ? colPos : 'after_no',
        });
      }
    }
  }

  if (wsPI) {
    const toVal = cellToString(wsPI.getCell('C8').value);
    if (toVal) {
      if (knownEndUsers.includes(toVal) || END_USER_OPTIONS.includes(toVal)) {
        job.controls.endUserSelect = toVal;
      } else {
        job.controls.endUserSelect = 'OTHER';
        job.controls.customEndUser = toVal;
      }
    }
    const customName = cellToString(wsPI.getCell('C9').value);
    if (customName && job.controls.endUserSelect === 'OTHER') {
      job.controls.customEndUser = customName;
    }
    const rfqVal = cellToString(wsPI.getCell('C10').value);
    if (rfqVal && !job.controls.rfqRef) job.controls.rfqRef = rfqVal;
    const poVal = cellToString(wsPI.getCell('C11').value);
    if (poVal && !job.controls.customerPoRef) job.controls.customerPoRef = poVal;

    // Read K3:L7 controls dynamically by label (handles both current VAT-free PI and legacy v1 PI!L5 VAT)
    for (let r = 3; r <= 7; r++) {
      const lbl = cellToString(wsPI.getCell(r, 11).value).toLowerCase();
      const valCell = wsPI.getCell(r, 12).value;
      if (lbl.includes('currency')) {
        const curr = cellToString(valCell);
        if (curr) job.controls.currency = curr;
      } else if (lbl.includes('incoterm')) {
        const inco = cellToString(valCell);
        if (inco) job.controls.incoterm = inco;
      } else if (lbl.includes('vat')) {
        const vat = cellToNumber(valCell);
        if (vat !== null && vat >= 0 && vat <= 1 && !hasMetaIncludeVat) {
          job.controls.vatRate = vat;
          job.controls.includeVat = vat > 0;
        }
      } else if (lbl.includes('destination')) {
        const dest = cellToString(valCell);
        if (dest) job.controls.destination = dest;
      } else if (lbl.includes('payment')) {
        const pay = cellToString(valCell);
        if (pay && !job.controls.paymentTerms) job.controls.paymentTerms = pay;
      } else if (lbl.includes('lead')) {
        const lead = cellToString(valCell);
        if (lead && !job.controls.deliveryLeadTime) job.controls.deliveryLeadTime = lead;
      }
    }

    // Parse rows 17..36 (standard template rows)
    let foundRows = 0;
    for (let i = 0; i < 20; i++) {
      const r = 17 + i;
      const reqDesc = cellToString(wsPI.getCell(r, 3).value);
      const offDesc = cellToString(wsPI.getCell(r, 4).value);
      const uom = cellToString(wsPI.getCell(r, 5).value);
      const qty = cellToNumber(wsPI.getCell(r, 6).value);

      if (reqDesc || offDesc || qty !== null) {
        foundRows += 1;
        job.items[i] = {
          ...job.items[i],
          requestedDescription: reqDesc,
          offeredDescription: offDesc || reqDesc,
          uom: uom || 'PCS',
          qty,
        };
      }
    }

    // Fallback: if workbook is an issued values-only export (headers on row 11, items start at row 12)
    if (foundRows === 0) {
      let reqCol = 3;
      let offCol = 4;
      let uomCol = 5;
      let qtyCol = 6;
      let priceCol = 7;
      const totalCols = Math.min(20, wsPI.columnCount || 15);
      for (let c = 2; c <= totalCols; c++) {
        const h = cellToString(wsPI.getCell(11, c).value).toLowerCase();
        if (h.includes('requested')) reqCol = c;
        else if (h.includes('offered')) offCol = c;
        else if (h === 'uom' || h.includes('unit of')) uomCol = c;
        else if (h === 'qty' || h.includes('quantity')) qtyCol = c;
        else if (h.includes('unit price') || h.includes('unit cost')) priceCol = c;
      }

      for (let i = 0; i < 20; i++) {
        const r = 12 + i;
        const reqDesc = cellToString(wsPI.getCell(r, reqCol).value);
        const offDesc = cellToString(wsPI.getCell(r, offCol).value);
        const uom = cellToString(wsPI.getCell(r, uomCol).value);
        const qty = cellToNumber(wsPI.getCell(r, qtyCol).value);
        const price = cellToNumber(wsPI.getCell(r, priceCol).value);
        if ((reqDesc || offDesc) && qty !== null && qty > 0) {
          foundRows += 1;
          job.items[i] = {
            ...job.items[i],
            requestedDescription: reqDesc,
            offeredDescription: offDesc || reqDesc,
            uom: uom || 'PCS',
            qty,
            supplierUnitCost: price !== null ? Number((price / 1.2).toFixed(2)) : null,
          };
        }
      }
    }

    // Read any custom columns stored in Col N+ (colNum 14..17) on template PI, restoring exact positions from CONTROLS
    const importedCustomCols: CustomColumnDef[] = [];
    for (let cNum = 14; cNum <= 17; cNum++) {
      const hdr = cellToString(wsPI.getCell(16, cNum).value);
      if (hdr) {
        const savedMeta = ctrlColLayout.get(hdr.toLowerCase());
        const id =
          savedMeta?.id || `col_${hdr.toLowerCase().replace(/[^a-z0-9]+/g, '_')}_${cNum}`;
        const position: CustomColumnPosition = savedMeta?.position || 'after_no';
        importedCustomCols.push({ id, label: hdr, position });
        for (let i = 0; i < 20; i++) {
          const val = cellToString(wsPI.getCell(17 + i, cNum).value);
          if (val) {
            job.items[i] = {
              ...job.items[i],
              customFields: { ...(job.items[i].customFields || {}), [id]: val },
            };
          }
        }
      }
    }
    if (importedCustomCols.length > 0) {
      job.customColumns = importedCustomCols;
    }
  }

  if (wsPC) {
    const e3 = cellToNumber(wsPC.getCell('E3').value);
    if (e3 !== null && e3 >= 0 && e3 <= 5) {
      job.freight.selectedMarkupPct = e3;
    }
    // Read internal VAT from PRICE CALC!E4 (if present)
    const e4 = cellToNumber(wsPC.getCell('E4').value);
    if (e4 !== null && e4 >= 0 && e4 <= 1) {
      if (e4 > 0) {
        job.controls.includeVat = true;
        job.controls.vatRate = e4;
      } else {
        job.controls.includeVat = false;
      }
    }

    const rateCbm = cellToNumber(wsPC.getCell('I3').value);
    if (rateCbm !== null && rateCbm >= 0) job.freight.ratePerCbm = rateCbm;
    const rateKg = cellToNumber(wsPC.getCell('L3').value);
    if (rateKg !== null && rateKg >= 0) job.freight.ratePerKg = rateKg;
    const docPct = cellToNumber(wsPC.getCell('O3').value);
    if (docPct !== null && docPct >= 0 && docPct <= 1) job.freight.docChargePct = docPct;

    for (let i = 0; i < 20; i++) {
      const r = 17 + i;
      const cost = cellToNumber(wsPC.getCell(r, 5).value);
      const boxes = cellToNumber(wsPC.getCell(r, 7).value);
      const l = cellToNumber(wsPC.getCell(r, 8).value);
      const w = cellToNumber(wsPC.getCell(r, 9).value);
      const h = cellToNumber(wsPC.getCell(r, 10).value);
      const kg = cellToNumber(wsPC.getCell(r, 12).value);

      const updates: Partial<JobItemRow> = {};
      if (cost !== null && cost >= 0) updates.supplierUnitCost = cost;
      if (boxes !== null && boxes >= 0) updates.boxes = boxes;
      if (l !== null && l >= 0) updates.boxLengthCm = l;
      if (w !== null && w >= 0) updates.boxWidthCm = w;
      if (h !== null && h >= 0) updates.boxHeightCm = h;
      if (kg !== null && kg >= 0) updates.kgPerBox = kg;
      if (Object.keys(updates).length > 0) {
        job.items[i] = { ...job.items[i], ...updates };
      }
    }
  }

  if (wsPL) {
    // Locate Carton No., Net Wt, Gross Wt dynamically by Row 16 headers (resilient to custom column shifts!)
    let cartonCol = 7;
    let netCol = 8;
    let grossCol = 9;
    const plMaxCols = Math.min(20, wsPL.columnCount || 15);
    for (let c = 2; c <= plMaxCols; c++) {
      const h = cellToString(wsPL.getCell(16, c).value).toLowerCase();
      if (h.includes('carton')) cartonCol = c;
      else if (h.includes('net wt') || h.includes('net weight')) netCol = c;
      else if (h.includes('gross wt') || h.includes('gross weight')) grossCol = c;
    }

    for (let i = 0; i < 20; i++) {
      const r = 17 + i;
      const carton = cellToString(wsPL.getCell(r, cartonCol).value);
      const net = cellToNumber(wsPL.getCell(r, netCol).value);
      const gross = cellToNumber(wsPL.getCell(r, grossCol).value);
      if (carton || net !== null || gross !== null) {
        job.items[i] = {
          ...job.items[i],
          cartonNo: carton,
          netWtKg: net !== null && net >= 0 ? net : null,
          grossWtKg: gross !== null && gross >= 0 ? gross : null,
        };
      }
    }
  }

  const itemCount = job.items.filter(
    (it) => it.requestedDescription.trim() || it.offeredDescription.trim() || (it.qty ?? 0) > 0
  ).length;

  return { kind: 'job', job, itemCount };
}

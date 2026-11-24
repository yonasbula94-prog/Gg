import { describe, expect, test } from 'bun:test';
import ExcelJS from 'exceljs';
import { deflateSync } from 'node:zlib';
import {
  guessColumnRolesForHeaderRow,
  parseFlexibleNumber,
  parseGenericSupplierXlsx,
  parsePastedSupplierQuote,
  parseSupplierPdfFile,
} from '../src/utils/supplierQuoteParser';

const summary = (text: string) =>
  parsePastedSupplierQuote(text, 'test').items.map((i) => [i.code, i.description, i.uom, i.qty, i.unitPrice]);

describe('number parsing', () => {
  test.each([
    ['12,50', 12.5],
    ['1,250', 1250],
    ['1.234,50', 1234.5],
    ['1,234.50', 1234.5],
    ['€12,50', 12.5],
    ['USD 1,000.00', 1000],
    ['42.50', 42.5],
  ])('%s -> %p', (raw, expected) => {
    expect(parseFlexibleNumber(raw)).toBe(expected);
  });
});

describe('pasted text quotes', () => {
  test('tab separated with part number', () => {
    expect(summary('1\tPN-4402\tHydraulic Seal Kit\tPCS\t10\t42.50\t425.00')).toEqual([
      ['PN-4402', 'Hydraulic Seal Kit', 'PCS', 10, 42.5],
    ]);
  });

  test('space aligned columns', () => {
    expect(summary('PN-8810   Pressure Gauge 0-16 Bar   PCS   5   68.00   340.00')).toEqual([
      ['PN-8810', 'Pressure Gauge 0-16 Bar', 'PCS', 5, 68],
    ]);
  });

  test('pipe delimited', () => {
    expect(summary('| PN-1 | Seal | PCS | 2 | 10.00 | 20.00 |')).toEqual([['PN-1', 'Seal', 'PCS', 2, 10]]);
  });

  test('European decimal commas are read as decimals', () => {
    expect(summary('3  Gasket kit  PCS  4  12,50  50,00')[0][4]).toBe(12.5);
  });

  test('quantity first, no serial number, is not dropped', () => {
    expect(summary('10  Hydraulic seal kit  PCS  42.50  425.00')).toEqual([['', 'Hydraulic seal kit', 'PCS', 10, 42.5]]);
  });

  test('UoM before the quantity is recognised (KG) and CAS number kept', () => {
    expect(summary('Sodium chloride 7647-14-5 KG 25 3.10 77.50')).toEqual([
      ['7647-14-5', 'Sodium chloride 7647-14-5', 'KG', 25, 3.1],
    ]);
  });

  test('totals and payment lines are not items', () => {
    const items = summary('Hydraulic seal kit   PCS   2   10.00   20.00\nTotal   20.00\nPayment Terms: 50% advance');
    expect(items).toEqual([['', 'Hydraulic seal kit', 'PCS', 2, 10]]);
  });

  test('currency comes from price lines, not an incidental mention', () => {
    const q = parsePastedSupplierQuote('Note: approx. USD 1,000 equivalent\n1  Seal kit  PCS  4  AED 42.50  AED 170.00', 't');
    expect(q.currency).toBe('AED');
  });

  test('incoterm and payment terms are read from metadata lines', () => {
    const q = parsePastedSupplierQuote('Incoterm: CIF\nPayment Terms: 30% advance\n1  Seal kit  PCS  4  42.50  170.00', 't');
    expect(q.incoterm).toBe('CIF');
    expect(q.paymentTerms).toBe('30% advance');
  });
});

describe('header detection', () => {
  test('metadata cells never become a unit price column', () => {
    const roles = guessColumnRolesForHeaderRow(['Incoterm: CIF', 'Currency: EUR', 'Payment Terms: 30% advance']);
    expect(roles.roles).not.toContain('unitPrice');
  });

  test('standard header maps each column', () => {
    const { roles } = guessColumnRolesForHeaderRow(['No', 'Item Code', 'Description', 'UoM', 'Qty', 'Unit Price', 'Amount']);
    expect(roles).toEqual(['ignore', 'code', 'description', 'uom', 'qty', 'unitPrice', 'lineTotal']);
  });
});

describe('XLSX quotes', () => {
  test('title and metadata rows above the table, totals below', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Quote');
    ws.addRow(['ACME Industrial Supplies LLC']);
    ws.addRow(['Quotation No: AQ-2026-881', '', 'Date: 2026-11-10']);
    ws.addRow(['Incoterm: CIF', 'Currency: EUR', 'Payment Terms: 30% advance']);
    ws.addRow([]);
    ws.addRow(['No', 'Item Code', 'Description', 'UoM', 'Qty', 'Unit Price', 'Amount']);
    ws.addRow([1, 'PN-4402', 'Hydraulic Seal Kit', 'SET', 10, 42.5, 425]);
    ws.addRow([2, 'PN-8810', 'Pressure Gauge 0-16 Bar', 'PCS', 5, 68, 340]);
    ws.addRow(['', '', 'Total', '', '', '', 765]);
    const q = parseGenericSupplierXlsx(wb as never, 'acme.xlsx');
    expect(q.currency).toBe('EUR');
    expect(q.supplierRef).toBe('AQ-2026-881');
    expect(q.items.map((i) => [i.code, i.uom, i.qty, i.unitPrice])).toEqual([
      ['PN-4402', 'SET', 10, 42.5],
      ['PN-8810', 'PCS', 5, 68],
    ]);
  });
});

describe('PDF quotes', () => {
  const pdfWith = (stream: Buffer) =>
    Buffer.concat([
      Buffer.from(`%PDF-1.4\n1 0 obj << /Type /Catalog >> endobj\n4 0 obj << /Length ${stream.length} /Filter /FlateDecode >>\nstream\n`, 'latin1'),
      stream,
      Buffer.from('\nendstream\nendobj\ntrailer << >>\n%%EOF', 'latin1'),
    ]);

  test('compressed content stream is decoded and parsed', async () => {
    const content = [
      'BT /F1 10 Tf 50 700 Td (PN-4402 Hydraulic Seal Kit PCS 10 42.50 425.00) Tj ET',
      'BT /F1 10 Tf 50 680 Td (PN-8810 Pressure Gauge 0-16 Bar PCS 5 68.00 340.00) Tj ET',
    ].join('\n');
    const file = new File([pdfWith(deflateSync(Buffer.from(content, 'latin1')))], 'quote.pdf', { type: 'application/pdf' });
    const q = await parseSupplierPdfFile(file);
    expect(q.items.map((i) => [i.code, i.description, i.uom, i.qty, i.unitPrice])).toEqual([
      ['PN-4402', 'Hydraulic Seal Kit', 'PCS', 10, 42.5],
      ['PN-8810', 'Pressure Gauge 0-16 Bar', 'PCS', 5, 68],
    ]);
    expect(q.isScannedPdf).toBe(false);
  });
});

import ExcelJS from 'exceljs';
import { CURRENCY_OPTIONS, INCOTERM_OPTIONS, UOM_OPTIONS } from '../constants/themes';

export type ColumnRole =
  | 'ignore'
  | 'code'
  | 'description'
  | 'uom'
  | 'qty'
  | 'unitPrice'
  | 'lineTotal';

export interface RawExtractedTable {
  name: string;
  rows: string[][];
}

export interface DetectedTableMapping {
  tableIndex: number;
  headerRowIndex: number;
  columnRoles: ColumnRole[];
}

export interface ParsedQuoteItem {
  id: string;
  selected: boolean;
  code: string; // Part No. / CAS No. / SKU if detected
  description: string;
  uom: string;
  qty: number | null;
  unitPrice: number | null;
}

export interface ParsedSupplierQuote {
  sourceType: 'xlsx' | 'pdf' | 'text';
  filename: string;
  supplierName: string;
  supplierRef: string;
  currency: string;
  incoterm: string;
  paymentTerms: string;
  deliveryLeadTime: string;
  detectedCodeLabel: string; // e.g. 'Part No.' or 'CAS No.'
  items: ParsedQuoteItem[];
  rawExtractedText?: string;
  rawTables?: RawExtractedTable[];
  detectedMapping?: DetectedTableMapping;
  isScannedPdf?: boolean;
}

function cellToText(val: ExcelJS.CellValue): string {
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

export function parseFlexibleNumber(raw: string | number | null | undefined): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number') return Number.isFinite(raw) ? raw : null;

  let cleaned = String(raw)
    .trim()
    .replace(/^(USD|AED|EUR|GBP|SAR|CNY|RMB|INR|US\$|\$|€|£|¥)\s*/i, '')
    .replace(/\s*(USD|AED|EUR|GBP|SAR|CNY|RMB|INR|\/pc|\/pcs|\/ea|\/unit|\/set|\/kg|\/ltr|\/m|\/mtr|\/box|\/pack)$/i, '')
    .replace(/\s+/g, '')
    .trim();

  if (!cleaned) return null;

  // Handle European decimal formats: "1.250,50" -> "1250.50" or "45,50" -> "45.50"
  if (/^-?\d{1,3}(\.\d{3})+,\d{1,2}$/.test(cleaned)) {
    cleaned = cleaned.replace(/\./g, '').replace(',', '.');
  } else if (/^-?\d+,\d{1,2}$/.test(cleaned) && !cleaned.includes('.')) {
    cleaned = cleaned.replace(',', '.');
  } else {
    cleaned = cleaned.replace(/,/g, '');
  }

  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function normalizeUom(raw: string): string {
  const up = raw.trim().toUpperCase().replace(/\./g, '');
  if (!up) return 'PCS';
  if ((UOM_OPTIONS as readonly string[]).includes(up)) return up;
  if (['PC', 'EA', 'EACH', 'NOS', 'NO', 'UNIT', 'UNITS', 'PIECES', 'PIECE', 'STK', 'STUCK', 'PCE'].includes(up)) return 'PCS';
  if (['SETS', 'KIT', 'KITS', 'SATZ'].includes(up)) return 'SET';
  if (['KGS', 'KILO', 'KILOS', 'KG'].includes(up)) return 'KG';
  if (['LITRE', 'LITRES', 'LITER', 'LITERS', 'L', 'LTR', 'LTRS'].includes(up)) return 'LTR';
  if (['M', 'MTR', 'METER', 'METERS', 'METRE', 'METRES', 'MTRS'].includes(up)) return 'MTR';
  if (['ROLLS', 'DRUM', 'DRUMS', 'REEL', 'COIL'].includes(up)) return 'ROLL';
  if (['BOXES', 'CTN', 'CARTON', 'CARTONS', 'CASE', 'BX'].includes(up)) return 'BOX';
  if (['PK', 'PACK', 'PACKS', 'PKT', 'BAG', 'BAGS', 'BOTTLE', 'BTL', 'CAN', 'TIN'].includes(up)) return 'PACK';
  if (['PAIR', 'PAIRS', 'PR'].includes(up)) return 'PAIR';
  return 'PCS';
}

function extractMetadataFromText(lines: string[], filename: string): {
  supplierName: string;
  supplierRef: string;
  currency: string;
  incoterm: string;
  paymentTerms: string;
  deliveryLeadTime: string;
} {
  const fullText = lines.join('\n');

  // Currency: prefer lines that carry a price (so an incidental "USD" mention elsewhere does not win)
  const priceLines = lines.filter((l) => /\d+[.,]\d{2}\b/.test(l)).join("\n");
  const findCurrency = (text: string): string | undefined =>
    CURRENCY_OPTIONS.find((c) => new RegExp(`\\b${c}\\b`, "i").test(text));
  let currency = findCurrency(priceLines) ?? findCurrency(fullText) ?? "USD";
  if (currency === "USD" && /€/.test(priceLines || fullText)) currency = "EUR";
  else if (currency === "USD" && /£/.test(priceLines || fullText)) currency = "GBP";

  // Incoterm
  let incoterm = '';
  for (const inc of INCOTERM_OPTIONS) {
    if (new RegExp(`\\b${inc}\\b`).test(fullText)) {
      incoterm = inc;
      break;
    }
  }

  // Quote Ref
  let supplierRef = '';
  const refMatch = fullText.match(
    /(?:quote\s*(?:no|#|ref|number)|quotation\s*(?:no|#|ref)|offer\s*(?:no|#|ref)|our\s*ref|ref\.?\s*(?:no)?)\s*[:#-]?\s*([A-Z0-9][A-Z0-9/_-]{2,28})/i
  );
  if (refMatch) supplierRef = refMatch[1].trim();

  // Payment terms
  let paymentTerms = '';
  const payLine = lines.find((l) => /payment\s*terms?|terms\s*of\s*payment|payment\s*conditions?/i.test(l));
  if (payLine) {
    paymentTerms = payLine
      .replace(/.*?(?:payment\s*terms?|terms\s*of\s*payment|payment\s*conditions?)\s*[:|-]?\s*/i, '')
      .trim();
  }
  if (!paymentTerms) {
    const payMatch = fullText.match(
      /(\d+%\s*(?:advance|deposit|tt|t\/t|with\s*po|before\s*shipment)[^;\n]{0,65}|net\s*\d+\s*days?|100%\s*(?:advance|tt|t\/t|lc|l\/c)[^;\n]{0,45})/i
    );
    if (payMatch) paymentTerms = payMatch[1].trim();
  }

  // Delivery / Lead time
  let deliveryLeadTime = '';
  const delLine = lines.find(
    (l) =>
      /(?:delivery|lead\s*time|dispatch|shipment\s*time|availability)/i.test(l) &&
      /\d+|stock|ex-stock|immediate|week|day|month/i.test(l)
  );
  if (delLine) {
    deliveryLeadTime = delLine
      .replace(/.*?(?:delivery\s*(?:time|period|terms)?|lead\s*time|dispatch|availability)\s*[:|-]?\s*/i, '')
      .trim()
      .slice(0, 75);
  }

  // Supplier name
  let supplierName = '';
  const supLine = lines.find((l) => /^(?:from|supplier|vendor|company|seller)\s*[:|-]/i.test(l));
  if (supLine) {
    supplierName = supLine.replace(/^(?:from|supplier|vendor|company|seller)\s*[:|-]\s*/i, '').trim();
  }
  if (!supplierName) {
    const companyLine = lines
      .slice(0, 12)
      .find(
        (l) =>
          l.length >= 4 &&
          l.length <= 75 &&
          !/grand\s*group|quotation|proforma|invoice|date|page|tel|phone|email|www\.|http/i.test(l) &&
          /(?:llc|ltd|inc|corp|fzc|fze|gmbh|ag|s\.a\.|co\.|trading|supplies|industrial|medical|pharma|chem|tech|group|international|engineering|equipment)/i.test(
            l
          )
      );
    supplierName =
      companyLine || filename.replace(/\.(xlsx|pdf|txt)$/i, '').replace(/[_-]+/g, ' ').trim();
  }

  return {
    supplierName: supplierName.slice(0, 80),
    supplierRef: supplierRef.slice(0, 40),
    currency,
    incoterm,
    paymentTerms: paymentTerms.slice(0, 95),
    deliveryLeadTime,
  };
}

function isSummaryOrNoiseRow(desc: string): boolean {
  const s = desc.trim().toLowerCase();
  if (!s) return true;
  return /^(sub\s*-?\s*total|total|grand\s*total|net\s*total|gross\s*total|vat|tax|gst|freight|shipping|delivery\s*charge|handling|discount|bank|terms|notes?|remarks?|validity|payment|thank\s*you|prepared\s*by|authorized|page\s*\d)/i.test(
    s
  );
}

/**
 * Classifies column headers into procurement roles with broad real-world & multilingual synonym coverage.
 */
export function guessColumnRolesForHeaderRow(headerCells: string[]): {
  roles: ColumnRole[];
  score: number;
  codeLabel: string;
} {
  const roles: ColumnRole[] = headerCells.map(() => 'ignore');
  let hasCode = false;
  let hasDesc = false;
  let hasUom = false;
  let hasQty = false;
  let hasUnitPrice = false;
  let hasTotal = false;
  let codeLabel = 'Part No.';

  headerCells.forEach((raw, idx) => {
    const low = raw.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!low) return;

    // Metadata labels such as "Incoterm: CIF" or "Payment Terms: 30%" are not column headers
    if (/^[a-z][a-z ./-]{1,30}:/.test(low)) return;

    // Skip pure serial / line-no headers so they aren't confused with Part No.
    if (/^(#|s\.?n\.?|sr\.?\s*no\.?|sl\.?\s*no\.?|no\.?|pos\.?|line|item\s*#?)$/.test(low)) {
      return;
    }

    if (
      !hasCode &&
      /\b(cas\s*(no\.?|#|num|number)?|part\s*(no\.?|#|num|number|code)|p\/n|pn|mpn|mfr\.?\s*part|oem|sku|item\s*(code|no\.?)|stock\s*(code|no\.?)|material\s*(code|no\.?)|product\s*(code|no\.?)|catalog|cat\.?\s*no|model\s*(no\.?|#)?|hs\s*code|artikel\s*(nr|no)?|ref\.?\s*(no\.?|#)|code)\b/i.test(
        low
      ) &&
      !/description|name|bezeichnung/i.test(low)
    ) {
      roles[idx] = 'code';
      hasCode = true;
      codeLabel = /cas/i.test(low) ? 'CAS No.' : /hs/i.test(low) ? 'HS Code' : 'Part No.';
    } else if (
      !hasDesc &&
      /\b(description|desc\.?|item\s*name|product|material|particulars|specification|spec\.?|goods|designation|details|artikel|bezeichnung|item\s*description|offered|requested)\b/i.test(
        low
      ) &&
      !/price|cost|rate|total|amount|qty|quantity|code$|no\.?$|#$/i.test(low)
    ) {
      roles[idx] = 'description';
      hasDesc = true;
    } else if (
      !hasUom &&
      /\b(uom|u\/m|u\.o\.m|unit|units|packing|pack|einheit|unite|measure)\b/i.test(low) &&
      !/price|cost|rate|total|amount/i.test(low)
    ) {
      roles[idx] = 'uom';
      hasUom = true;
    } else if (
      !hasQty &&
      /\b(qty|q'ty|quantity|qnty|order\s*qty|pcs|nos|menge|qte|quantite|count)\b/i.test(low)
    ) {
      roles[idx] = 'qty';
      hasQty = true;
    } else if (
      !hasUnitPrice &&
      /\b(unit\s*price|u\/price|u\.?\s*price|unit\s*cost|unit\s*rate|rate|price\s*per|net\s*price|einzelpreis|preis|prix\s*unitaire|price|cost)\b/i.test(
        low
      ) &&
      !/total|amount|ext\.?|sum|gesamt/i.test(low)
    ) {
      roles[idx] = 'unitPrice';
      hasUnitPrice = true;
    } else if (
      !hasTotal &&
      /\b(total|amount|ext\.?\s*price|extended|line\s*total|net\s*amount|value|gesamtpreis|montant|sum)\b/i.test(
        low
      )
    ) {
      roles[idx] = 'lineTotal';
      hasTotal = true;
    }
  });

  const score =
    (hasDesc ? 3 : 0) +
    (hasQty ? 2 : 0) +
    (hasUnitPrice ? 3 : 0) +
    (hasTotal ? 2 : 0) +
    (hasUom ? 1 : 0) +
    (hasCode ? 1 : 0);

  return { roles, score, codeLabel };
}

/**
 * Extracts structured quote items from ANY 2D table given a header row index & column role assignments.
 * Used both by automatic extraction AND by the interactive Column Mapper Assistant!
 */
export function extractItemsFromMappedTable(
  rows: string[][],
  headerRowIndex: number,
  columnRoles: ColumnRole[],
  mergeWrappedLines = true
): { items: ParsedQuoteItem[]; detectedCodeLabel: string } {
  const descIdx = columnRoles.indexOf('description');
  const codeIdx = columnRoles.indexOf('code');
  const uomIdx = columnRoles.indexOf('uom');
  const qtyIdx = columnRoles.indexOf('qty');
  const unitPriceIdx = columnRoles.indexOf('unitPrice');
  const totalIdx = columnRoles.indexOf('lineTotal');

  const headerRow = headerRowIndex >= 0 && headerRowIndex < rows.length ? rows[headerRowIndex] : [];
  const codeHeader = codeIdx >= 0 ? (headerRow[codeIdx] || '').toLowerCase() : '';
  const detectedCodeLabel = /cas/i.test(codeHeader)
    ? 'CAS No.'
    : /hs/i.test(codeHeader)
    ? 'HS Code'
    : 'Part No.';

  const items: ParsedQuoteItem[] = [];
  const startRow = Math.max(0, headerRowIndex + 1);

  for (let r = startRow; r < rows.length && items.length < 20; r++) {
    const row = rows[r] || [];
    const desc = descIdx >= 0 ? (row[descIdx] || '').trim() : '';
    const code = codeIdx >= 0 ? (row[codeIdx] || '').trim() : '';
    const uomRaw = uomIdx >= 0 ? (row[uomIdx] || '').trim() : '';
    const qty = qtyIdx >= 0 ? parseFlexibleNumber(row[qtyIdx]) : null;
    let unitPrice = unitPriceIdx >= 0 ? parseFlexibleNumber(row[unitPriceIdx]) : null;
    const total = totalIdx >= 0 ? parseFlexibleNumber(row[totalIdx]) : null;

    const wholeLine = row.join(' ').trim();
    if (!wholeLine) continue;

    // Skip totals/summary/footer rows
    if (isSummaryOrNoiseRow(desc || wholeLine) && !code) {
      continue;
    }

    // Wrapped multi-line continuation description (text only, no qty/price/code)
    if (
      mergeWrappedLines &&
      items.length > 0 &&
      desc &&
      !code &&
      qty === null &&
      unitPrice === null &&
      total === null &&
      desc.length <= 120
    ) {
      const prev = items[items.length - 1];
      prev.description = `${prev.description} ${desc}`.trim();
      continue;
    }

    // Auto-derive Unit Price from Line Total / Qty if supplier only provided a Total column
    if (unitPrice === null && total !== null && qty !== null && qty > 0) {
      unitPrice = Math.round((total / qty) * 100) / 100;
    }

    if ((desc || code) && (qty !== null || unitPrice !== null || total !== null)) {
      items.push({
        id: `map-row-${r}`,
        selected: true,
        code,
        description: desc || code,
        uom: normalizeUom(uomRaw),
        qty: qty !== null && qty > 0 ? qty : 1,
        unitPrice,
      });
    }
  }

  return { items, detectedCodeLabel };
}

/**
 * 1. ANY-FORMAT SUPPLIER EXCEL (.XLSX) PARSER (with Raw Sheet Grid Capture for Visual Column Mapper)
 */
export function parseGenericSupplierXlsx(
  wb: ExcelJS.Workbook,
  filename: string
): ParsedSupplierQuote {
  const rawTables: RawExtractedTable[] = [];
  const allTextLines: string[] = [];

  for (const ws of wb.worksheets) {
    const maxRow = Math.min(ws.rowCount || 80, 160);
    const maxCol = Math.min(ws.columnCount || 18, 24);
    const grid: string[][] = [];
    let usedCols = 0;

    for (let r = 1; r <= maxRow; r++) {
      const row = ws.getRow(r);
      const vals: string[] = [];
      let nonEmpty = false;
      for (let c = 1; c <= maxCol; c++) {
        const txt = cellToText(row.getCell(c).value);
        vals.push(txt);
        if (txt) {
          nonEmpty = true;
          if (c > usedCols) usedCols = c;
        }
      }
      if (nonEmpty) {
        grid.push(vals);
        if (grid.length <= 35 || grid.length > maxRow - 20) {
          allTextLines.push(vals.filter(Boolean).join('  '));
        }
      }
    }

    if (grid.length > 0 && usedCols > 0) {
      rawTables.push({
        name: ws.name || `Sheet ${rawTables.length + 1}`,
        rows: grid.map((r) => r.slice(0, usedCols)),
      });
    }
  }

  let bestTableIdx = 0;
  let bestHeaderIdx = 0;
  let bestRoles: ColumnRole[] = rawTables[0]?.rows[0]?.map(() => 'ignore') || [];
  let bestScore = -1;
  let bestCodeLabel = 'Part No.';

  rawTables.forEach((table, tIdx) => {
    const scanLimit = Math.min(table.rows.length, 40);
    for (let r = 0; r < scanLimit; r++) {
      const single = guessColumnRolesForHeaderRow(table.rows[r]);
      if (single.score > bestScore) {
        bestScore = single.score;
        bestTableIdx = tIdx;
        bestHeaderIdx = r;
        bestRoles = single.roles;
        bestCodeLabel = single.codeLabel;
      }

      // Also test combined 2-row stacked headers (e.g., Row 1: "Unit" / Row 2: "Price (USD)")
      if (r + 1 < scanLimit) {
        const nextRow = table.rows[r + 1];
        const nextNumericCount = nextRow.filter((c) => parseFlexibleNumber(c) !== null).length;
        if (nextNumericCount === 0) {
          const merged = table.rows[r].map((cell, cIdx) =>
            `${cell} ${nextRow[cIdx] || ''}`.trim()
          );
          const dual = guessColumnRolesForHeaderRow(merged);
          if (dual.score > bestScore) {
            bestScore = dual.score;
            bestTableIdx = tIdx;
            bestHeaderIdx = r + 1;
            bestRoles = dual.roles;
            bestCodeLabel = dual.codeLabel;
          }
        }
      }
    }
  });

  let items: ParsedQuoteItem[] = [];
  let detectedCodeLabel = bestCodeLabel;

  if (rawTables.length > 0 && bestScore >= 4) {
    const extracted = extractItemsFromMappedTable(
      rawTables[bestTableIdx].rows,
      bestHeaderIdx,
      bestRoles,
      true
    );
    items = extracted.items;
    detectedCodeLabel = extracted.detectedCodeLabel;
  }

  // Fallback to smart text parser if column headers were unusual
  if (items.length === 0 && rawTables.length > 0) {
    const textFallback = rawTables[0].rows.map((r) => r.filter(Boolean).join('\t')).join('\n');
    const parsed = parsePastedSupplierQuote(textFallback, filename);
    return {
      ...parsed,
      sourceType: 'xlsx',
      rawTables,
      detectedMapping: {
        tableIndex: bestTableIdx,
        headerRowIndex: Math.max(0, bestHeaderIdx),
        columnRoles: bestRoles,
      },
    };
  }

  const meta = extractMetadataFromText(allTextLines, filename);
  return {
    sourceType: 'xlsx',
    filename,
    ...meta,
    detectedCodeLabel,
    items,
    rawTables,
    detectedMapping:
      rawTables.length > 0
        ? {
            tableIndex: bestTableIdx,
            headerRowIndex: bestHeaderIdx,
            columnRoles: bestRoles,
          }
        : undefined,
  };
}

/**
 * 2. UPGRADED PURE CLIENT-SIDE OFFLINE PDF TEXT & TABLE EXTRACTOR
 * Supports FlateDecode, /ObjStm compressed object streams, ToUnicode CMaps, CTM (cm/q/Q),
 * dynamic X-column clustering into 2D tables, and wrapped line continuation merging.
 */
async function inflateBytes(input: Uint8Array): Promise<Uint8Array | null> {
  if (typeof DecompressionStream === 'undefined') return null;
  for (const format of ['deflate', 'deflate-raw'] as CompressionFormat[]) {
    try {
      const ds = new DecompressionStream(format);
      const writer = ds.writable.getWriter();
      writer.write(new Uint8Array(input));
      writer.close();
      const chunks: Uint8Array[] = [];
      const reader = ds.readable.getReader();
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        if (value) chunks.push(value);
      }
      const totalLen = chunks.reduce((s, c) => s + c.length, 0);
      const out = new Uint8Array(totalLen);
      let off = 0;
      for (const c of chunks) {
        out.set(c, off);
        off += c.length;
      }
      if (out.length > 0) return out;
    } catch {
      // Try next deflate format
    }
  }
  return null;
}

function bytesToLatin1(bytes: Uint8Array): string {
  let out = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    out += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return out;
}

function parseToUnicodeMaps(texts: string[]): Map<string, string> {
  const map = new Map<string, string>();
  for (const txt of texts) {
    if (!txt.includes('beginbfchar') && !txt.includes('beginbfrange')) continue;

    const bfcharBlocks = txt.matchAll(/beginbfchar([\s\S]*?)endbfchar/g);
    for (const block of bfcharBlocks) {
      const pairs = block[1].matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g);
      for (const p of pairs) {
        const src = p[1].toUpperCase();
        const dstHex = p[2];
        let decoded = '';
        for (let i = 0; i + 3 < dstHex.length; i += 4) {
          decoded += String.fromCharCode(parseInt(dstHex.slice(i, i + 4), 16));
        }
        if (decoded) map.set(src, decoded);
      }
    }

    const bfrangeBlocks = txt.matchAll(/beginbfrange([\s\S]*?)endbfrange/g);
    for (const block of bfrangeBlocks) {
      // Form 1: <start> <end> <dstStart>
      const ranges = block[1].matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g);
      for (const rg of ranges) {
        const start = parseInt(rg[1], 16);
        const end = parseInt(rg[2], 16);
        const dstStart = parseInt(rg[3], 16);
        const padLen = rg[1].length;
        if (end - start <= 512) {
          for (let code = start; code <= end; code++) {
            const key = code.toString(16).toUpperCase().padStart(padLen, '0');
            map.set(key, String.fromCharCode(dstStart + (code - start)));
          }
        }
      }
      // Form 2: <start> <end> [<dst1> <dst2> ...]
      const arrayRanges = block[1].matchAll(
        /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[([\s\S]*?)\]/g
      );
      for (const rg of arrayRanges) {
        const start = parseInt(rg[1], 16);
        const padLen = rg[1].length;
        const hexes = Array.from(rg[3].matchAll(/<([0-9A-Fa-f]+)>/g));
        hexes.forEach((h, idx) => {
          const key = (start + idx).toString(16).toUpperCase().padStart(padLen, '0');
          const dstHex = h[1];
          let decoded = '';
          for (let i = 0; i + 3 < dstHex.length; i += 4) {
            decoded += String.fromCharCode(parseInt(dstHex.slice(i, i + 4), 16));
          }
          if (decoded) map.set(key, decoded);
        });
      }
    }
  }
  return map;
}

function decodePdfLiteralString(raw: string): string {
  let out = '';
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (ch === '\\' && i + 1 < raw.length) {
      const next = raw[i + 1];
      if (next === 'n') { out += '\n'; i++; }
      else if (next === 'r') { out += '\r'; i++; }
      else if (next === 't') { out += '\t'; i++; }
      else if (next === '(' || next === ')' || next === '\\') { out += next; i++; }
      else if (/[0-7]/.test(next)) {
        const oct = raw.slice(i + 1, i + 4).match(/^[0-7]{1,3}/)?.[0] || '';
        out += String.fromCharCode(parseInt(oct, 8));
        i += oct.length;
      } else {
        i++;
      }
    } else {
      out += ch;
    }
  }
  return out;
}

function decodePdfHexString(hexRaw: string, cmap: Map<string, string>): string {
  const hex = hexRaw.replace(/\s+/g, '').toUpperCase();
  if (!hex) return '';
  if (cmap.size > 0) {
    if (hex.length % 4 === 0) {
      let decoded = '';
      let hits = 0;
      for (let i = 0; i < hex.length; i += 4) {
        const key = hex.slice(i, i + 4);
        const mapped = cmap.get(key);
        if (mapped) {
          decoded += mapped;
          hits++;
        } else {
          const code = parseInt(key, 16);
          if (code >= 32 && code < 127) decoded += String.fromCharCode(code);
        }
      }
      if (hits > 0) return decoded;
    }
    if (hex.length % 2 === 0) {
      let decoded = '';
      let hits = 0;
      for (let i = 0; i < hex.length; i += 2) {
        const key = hex.slice(i, i + 2);
        const mapped = cmap.get(key);
        if (mapped) {
          decoded += mapped;
          hits++;
        } else {
          const code = parseInt(key, 16);
          if (code >= 32 && code < 127) decoded += String.fromCharCode(code);
        }
      }
      if (hits > 0) return decoded;
    }
  }
  if (hex.startsWith('FEFF')) {
    let out = '';
    for (let i = 4; i + 3 < hex.length; i += 4) {
      out += String.fromCharCode(parseInt(hex.slice(i, i + 4), 16));
    }
    return out;
  }
  let out = '';
  for (let i = 0; i + 1 < hex.length; i += 2) {
    const code = parseInt(hex.slice(i, i + 2), 16);
    if (code >= 32 && code <= 126) out += String.fromCharCode(code);
  }
  return out;
}

interface PdfToken {
  x: number;
  y: number;
  text: string;
}

type Matrix6 = [number, number, number, number, number, number];

function multiplyMatrix(m1: Matrix6, m2: Matrix6): Matrix6 {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

function extractTokensFromContentStream(stream: string, cmap: Map<string, string>): PdfToken[] {
  const tokens: PdfToken[] = [];
  const ctmStack: Matrix6[] = [];
  let ctm: Matrix6 = [1, 0, 0, 1, 0, 0];
  let inText = false;
  let textX = 0;
  let textY = 0;
  let leading = 12;

  const opRegex =
    /\bq\b|\bQ\b|\bBT\b|\bET\b|\bT\*|(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(cm|Tm)|(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(Td|TD)|(-?\d+(?:\.\d+)?)\s+TL|(\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]+>|\[(?:[^\[\]]*)\])\s*(Tj|TJ|'|"|)/g;

  let match: RegExpExecArray | null;
  while ((match = opRegex.exec(stream)) !== null) {
    const full = match[0];
    if (full === 'q') {
      ctmStack.push([...ctm]);
    } else if (full === 'Q') {
      const prev = ctmStack.pop();
      if (prev) ctm = prev;
    } else if (full === 'BT') {
      inText = true;
      textX = 0;
      textY = 0;
    } else if (full === 'ET') {
      inText = false;
    } else if (match[7] === 'cm') {
      const m: Matrix6 = [
        parseFloat(match[1]),
        parseFloat(match[2]),
        parseFloat(match[3]),
        parseFloat(match[4]),
        parseFloat(match[5]),
        parseFloat(match[6]),
      ];
      ctm = multiplyMatrix(ctm, m);
    } else if (match[7] === 'Tm' && inText) {
      textX = parseFloat(match[5]);
      textY = parseFloat(match[6]);
    } else if ((match[10] === 'Td' || match[10] === 'TD') && inText) {
      const tx = parseFloat(match[8]);
      const ty = parseFloat(match[9]);
      textX += tx;
      textY += ty;
      if (match[10] === 'TD') leading = -ty;
    } else if (match[11] !== undefined && inText) {
      leading = parseFloat(match[11]);
    } else if (full === 'T*' && inText) {
      textY -= leading;
    } else if (match[12] && match[13] && inText) {
      const arg = match[12];
      const op = match[13];
      if (op === "'" || op === '"') textY -= leading;

      let text = '';
      if (arg.startsWith('(')) {
        text = decodePdfLiteralString(arg.slice(1, -1));
      } else if (arg.startsWith('<')) {
        text = decodePdfHexString(arg.slice(1, -1), cmap);
      } else if (arg.startsWith('[')) {
        const inner = arg.slice(1, -1);
        const parts = inner.matchAll(/\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]+>|-?\d+(?:\.\d+)?/g);
        for (const p of parts) {
          const tok = p[0];
          if (tok.startsWith('(')) text += decodePdfLiteralString(tok.slice(1, -1));
          else if (tok.startsWith('<')) text += decodePdfHexString(tok.slice(1, -1), cmap);
          else if (parseFloat(tok) < -180) text += ' ';
        }
      }

      const cleaned = text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').trim();
      if (cleaned) {
        const gx = ctm[0] * textX + ctm[2] * textY + ctm[4];
        const gy = ctm[1] * textX + ctm[3] * textY + ctm[5];
        tokens.push({ x: gx, y: Math.round(gy / 3) * 3, text: cleaned });
      }
    }
  }

  return tokens;
}

/**
 * Reconstructs a clean 2D table (rows x cols) from positioned PDF tokens by clustering X-coordinates.
 */
function buildPdfLinesAndGrid(tokens: PdfToken[]): { lines: string[]; grid: string[][] } {
  if (tokens.length === 0) return { lines: [], grid: [] };

  // Group tokens by Y row (top-to-bottom)
  const rowMap = new Map<number, { x: number; text: string }[]>();
  for (const t of tokens) {
    // Merge into existing Y bucket within +/- 3pt
    let matchedY: number | null = null;
    for (const existingY of rowMap.keys()) {
      if (Math.abs(existingY - t.y) <= 3.5) {
        matchedY = existingY;
        break;
      }
    }
    const yKey = matchedY !== null ? matchedY : t.y;
    const arr = rowMap.get(yKey) || [];
    arr.push({ x: t.x, text: t.text });
    rowMap.set(yKey, arr);
  }

  const sortedYs = Array.from(rowMap.keys()).sort((a, b) => b - a);

  // Merge adjacent fragments on the same line that are within 14pt into cohesive cells
  const cellRows: { x: number; text: string }[][] = [];
  const lines: string[] = [];

  for (const y of sortedYs) {
    const items = rowMap.get(y)!.sort((a, b) => a.x - b.x);
    const mergedCells: { x: number; text: string }[] = [];
    let lineText = '';

    for (let i = 0; i < items.length; i++) {
      const cur = items[i];
      if (i === 0) {
        mergedCells.push({ x: cur.x, text: cur.text });
        lineText = cur.text;
      } else {
        const prev = items[i - 1];
        const approxPrevEnd = prev.x + prev.text.length * 4.2;
        const gap = cur.x - approxPrevEnd;
        if (gap > 14) {
          mergedCells.push({ x: cur.x, text: cur.text });
          lineText += '\t' + cur.text;
        } else {
          mergedCells[mergedCells.length - 1].text += (gap > 1.5 ? ' ' : '') + cur.text;
          lineText += (gap > 1.5 ? ' ' : '') + cur.text;
        }
      }
    }

    if (lineText.trim()) {
      lines.push(lineText.trim());
      cellRows.push(mergedCells);
    }
  }

  // Cluster column X anchors across rows that have >= 2 cells
  const allXs: number[] = [];
  cellRows.forEach((row) => {
    if (row.length >= 2) {
      row.forEach((c) => allXs.push(c.x));
    }
  });
  allXs.sort((a, b) => a - b);

  const colAnchors: number[] = [];
  for (const x of allXs) {
    if (colAnchors.length === 0 || x - colAnchors[colAnchors.length - 1] > 28) {
      colAnchors.push(x);
    }
  }

  const colCount = Math.max(1, Math.min(colAnchors.length, 12));
  const grid: string[][] = cellRows.map((row) => {
    if (colCount <= 1 || row.length === 1) {
      const padded = new Array(colCount).fill('');
      padded[0] = row.map((c) => c.text).join(' ');
      return padded;
    }
    const cells = new Array(colCount).fill('');
    for (const c of row) {
      let bestCol = 0;
      let bestDist = Infinity;
      for (let i = 0; i < colCount; i++) {
        const dist = Math.abs(c.x - colAnchors[i]);
        if (dist < bestDist) {
          bestDist = dist;
          bestCol = i;
        }
      }
      cells[bestCol] = cells[bestCol] ? `${cells[bestCol]} ${c.text}` : c.text;
    }
    return cells;
  });

  return { lines, grid };
}

export async function parseSupplierPdfFile(file: File): Promise<ParsedSupplierQuote> {
  const buf = new Uint8Array(await file.arrayBuffer());
  const rawLatin1 = bytesToLatin1(buf);

  const decodedStreams: string[] = [];
  let idx = 0;
  while (idx < rawLatin1.length) {
    const streamPos = rawLatin1.indexOf('stream', idx);
    if (streamPos === -1) break;
    const endPos = rawLatin1.indexOf('endstream', streamPos + 6);
    if (endPos === -1) break;

    const dictSnippet = rawLatin1.slice(Math.max(0, streamPos - 350), streamPos);
    let dataStart = streamPos + 6;
    if (rawLatin1[dataStart] === '\r' && rawLatin1[dataStart + 1] === '\n') dataStart += 2;
    else if (rawLatin1[dataStart] === '\n' || rawLatin1[dataStart] === '\r') dataStart += 1;

    let dataEnd = endPos;
    while (
      dataEnd > dataStart &&
      (rawLatin1[dataEnd - 1] === '\n' || rawLatin1[dataEnd - 1] === '\r')
    ) {
      dataEnd--;
    }

    const slice = buf.subarray(dataStart, dataEnd);
    if (/\/FlateDecode/i.test(dictSnippet)) {
      const inflated = await inflateBytes(slice);
      if (inflated) {
        decodedStreams.push(bytesToLatin1(inflated));
      }
    } else if (!/\/DCTDecode|\/JPXDecode|\/CCITTFaxDecode/i.test(dictSnippet)) {
      decodedStreams.push(bytesToLatin1(slice));
    }

    idx = endPos + 9;
  }

  const cmap = parseToUnicodeMaps(decodedStreams);
  const allLines: string[] = [];
  const allGridRows: string[][] = [];

  for (const s of decodedStreams) {
    if (s.includes('BT') && (s.includes('Tj') || s.includes('TJ'))) {
      const tokens = extractTokensFromContentStream(s, cmap);
      const { lines, grid } = buildPdfLinesAndGrid(tokens);
      allLines.push(...lines);
      allGridRows.push(...grid);
    }
  }

  // Normalize grid row width across pages
  const maxCols = allGridRows.reduce((m, r) => Math.max(m, r.length), 0);
  const normalizedGrid =
    maxCols > 0
      ? allGridRows.map((r) => {
          const copy = [...r];
          while (copy.length < maxCols) copy.push('');
          return copy;
        })
      : [];

  const rawText = allLines.join('\n').trim();
  const textParsed = parsePastedSupplierQuote(rawText, file.name);

  // Also test header-based column detection on the reconstructed PDF 2D table!
  let bestHeaderIdx = 0;
  let bestRoles: ColumnRole[] = maxCols > 0 ? new Array(maxCols).fill('ignore') : [];
  let bestScore = -1;
  let mappedItems: ParsedQuoteItem[] = [];
  let mappedCodeLabel = textParsed.detectedCodeLabel;

  if (normalizedGrid.length > 0 && maxCols >= 3) {
    const scanLimit = Math.min(normalizedGrid.length, 35);
    for (let r = 0; r < scanLimit; r++) {
      const guess = guessColumnRolesForHeaderRow(normalizedGrid[r]);
      if (guess.score > bestScore) {
        bestScore = guess.score;
        bestHeaderIdx = r;
        bestRoles = guess.roles;
        mappedCodeLabel = guess.codeLabel;
      }
    }
    if (bestScore >= 4) {
      const res = extractItemsFromMappedTable(normalizedGrid, bestHeaderIdx, bestRoles, true);
      mappedItems = res.items;
      mappedCodeLabel = res.detectedCodeLabel;
    }
  }

  // Prefer whichever strategy captured more priced line items
  const mappedPriced = mappedItems.filter((i) => i.unitPrice !== null).length;
  const textPriced = textParsed.items.filter((i) => i.unitPrice !== null).length;
  const useMapped = mappedPriced > textPriced || (mappedPriced === textPriced && mappedItems.length > textParsed.items.length);

  const rawTables: RawExtractedTable[] =
    normalizedGrid.length > 0
      ? [{ name: 'PDF Reconstructed Table', rows: normalizedGrid }]
      : [];

  return {
    ...textParsed,
    sourceType: 'pdf',
    items: useMapped ? mappedItems : textParsed.items,
    detectedCodeLabel: useMapped ? mappedCodeLabel : textParsed.detectedCodeLabel,
    rawExtractedText: rawText,
    rawTables,
    detectedMapping:
      rawTables.length > 0
        ? {
            tableIndex: 0,
            headerRowIndex: bestHeaderIdx,
            columnRoles: bestRoles,
          }
        : undefined,
    isScannedPdf: rawText.length < 30,
  };
}

/**
 * Splits raw pasted text into a clean 2D grid for the Visual Column Mapper Assistant.
 */
export function parseDelimitedTextToTable(
  rawText: string,
  delimiter: 'auto' | 'tab' | 'spaces' | 'pipe' | 'comma' = 'auto'
): string[][] {
  const lines = rawText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  if (lines.length === 0) return [];

  const splitRow = (line: string): string[] => {
    if (delimiter === 'tab') return line.split('\t').map((c) => c.trim());
    if (delimiter === 'pipe') return line.split('|').map((c) => c.trim()).filter((_, idx, arr) => !(idx === 0 && arr[0] === ''));
    if (delimiter === 'comma') return line.split(',').map((c) => c.trim());
    if (delimiter === 'spaces') return line.split(/\s{2,}/).map((c) => c.trim());
    // Auto
    if (line.includes('\t')) return line.split('\t').map((c) => c.trim());
    if (line.includes('|')) return line.split('|').map((c) => c.trim()).filter(Boolean);
    const multiSpace = line.split(/\s{2,}/).map((c) => c.trim()).filter(Boolean);
    if (multiSpace.length >= 2) return multiSpace;
    return line.split(',').map((c) => c.trim());
  };

  const rows = lines.map(splitRow);
  const maxCols = Math.min(16, rows.reduce((m, r) => Math.max(m, r.length), 1));
  return rows.map((r) => {
    const out = r.slice(0, maxCols);
    while (out.length < maxCols) out.push('');
    return out;
  });
}

/**
 * 3. SMART TEXT / TABULAR QUOTE PARSER (Used for PDFs, Fallback Excel, & Pasted Quotes)
 */
export function parsePastedSupplierQuote(
  rawText: string,
  filename = 'Supplier Quote'
): ParsedSupplierQuote {
  const lines = rawText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const meta = extractMetadataFromText(lines, filename);
  const items: ParsedQuoteItem[] = [];
  const hasCasHint = /\bcas\b/i.test(rawText);
  const detectedCodeLabel = hasCasHint ? 'CAS No.' : 'Part No.';

  for (let i = 0; i < lines.length && items.length < 20; i++) {
    const line = lines[i];
    if (isSummaryOrNoiseRow(line)) continue;
    if (
      /\b(description|unit\s*price|qty|quantity)\b/i.test(line) &&
      /\b(total|amount|uom|unit|rate|price)\b/i.test(line)
    ) {
      continue;
    }

    // Strategy A: Tab-separated or multi-space/pipe columns
    const cols = line
      .split(/\t|\s{2,}|\|/)
      .map((c) => c.trim())
      .filter(Boolean);

    if (cols.length >= 3) {
      const numIndices: { idx: number; val: number }[] = [];
      cols.forEach((c, idx) => {
        const n = parseFlexibleNumber(c);
        if (n !== null && n > 0) numIndices.push({ idx, val: n });
      });

      if (numIndices.length >= 2) {
        const dataNums =
          numIndices[0].idx === 0 && Number.isInteger(numIndices[0].val) && numIndices[0].val === items.length + 1
            ? numIndices.slice(1)
            : numIndices;

        if (dataNums.length >= 2) {
          let qty: number | null = null;
          let unitPrice: number | null = null;

          if (dataNums.length >= 3) {
            const a = dataNums[dataNums.length - 3].val;
            const b = dataNums[dataNums.length - 2].val;
            const c = dataNums[dataNums.length - 1].val;
            if (Math.abs(a * b - c) <= Math.max(1, c * 0.03)) {
              qty = a;
              unitPrice = b;
            } else {
              qty = dataNums[0].val;
              unitPrice = b;
            }
          } else {
            qty = dataNums[0].val;
            unitPrice = dataNums[1].val;
          }

          const numSet = new Set(numIndices.map((n) => n.idx));
          let uom = 'PCS';
          let code = '';
          const descParts: string[] = [];

          cols.forEach((c, idx) => {
            if (numSet.has(idx)) return;
            if (
              /^(pcs?|ea|each|nos?|set|sets|kg|kgs|ltr|mtr|roll|box|ctn|pack|pair|kit|unit|units|drum|bag|btl)\.?$/i.test(
                c
              )
            ) {
              uom = normalizeUom(c);
              return;
            }
            if (!code && /^[A-Z0-9][A-Z0-9._/-]{2,22}$/i.test(c) && /\d/.test(c) && cols.length >= 4) {
              code = c;
              return;
            }
            descParts.push(c);
          });

          const description = descParts.join(' — ').trim();
          if (description.length >= 3 && !isSummaryOrNoiseRow(description)) {
            items.push({
              id: `txt-row-${i}`,
              selected: true,
              code,
              description,
              uom,
              qty,
              unitPrice,
            });
            continue;
          }
        }
      }
    }

    // Strategy B: Single-space PDF row heuristic: "<item desc> <qty> <uom?> <unitPrice> <total?>"
    const singleMatch = line.match(
      /^(?:\d{1,2}[.)]?\s+)?(.{4,115}?)\s+(?:(PCS|PC|EA|EACH|NOS|SET|KG|LTR|MTR|ROLL|BOX|PACK|PAIR|UNIT|UNITS|KIT|DRUM)\s+)?(\d+(?:[.,]\d+)?)\s+(?:(PCS|PC|EA|EACH|NOS|SET|KG|LTR|MTR|ROLL|BOX|PACK|PAIR|UNIT|UNITS|KIT|DRUM)\s+)?(?:USD|AED|EUR|GBP|SAR|\$|€|£)?\s*([\d,.]+[.,]\d{2})(?:\s+(?:USD|AED|EUR|GBP|SAR|\$|€|£)?\s*([\d,.]+[.,]\d{2}))?$/i
    );
    if (singleMatch) {
      let desc = singleMatch[1].trim();
      let uomRaw = singleMatch[4] || singleMatch[2] || '';
      const qty = parseFlexibleNumber(singleMatch[3]);
      const unitPrice = parseFlexibleNumber(singleMatch[5]);
      // A UoM word left at the end of the description ("... Seal Kit PCS") is moved to the UoM column
      const tailUom = desc.match(/\s(PCS|PC|EA|EACH|NOS|SET|KG|LTR|MTR|ROLL|BOX|PACK|PAIR|UNIT|UNITS|KIT|DRUM)$/i);
      if (tailUom && !uomRaw) {
        uomRaw = tailUom[1];
        desc = desc.slice(0, -tailUom[0].length).trim();
      }
      // A part number at the start ("PN-4402 Hydraulic Seal Kit") is the code. Needs a letter and a digit so "100 mm pipe" is not a code.
      let code = '';
      const casMatch = desc.match(/\b(\d{2,7}-\d{2}-\d)\b/);
      const lead = desc.match(/^([A-Z0-9][A-Z0-9._\/-]{2,22})\s+(?=\S)/i);
      if (casMatch) code = casMatch[1];
      else if (lead && /[A-Z]/i.test(lead[1]) && /\d/.test(lead[1])) {
        code = lead[1];
        desc = desc.slice(lead[0].length).trim();
      }
      if (desc && !isSummaryOrNoiseRow(desc) && qty && unitPrice !== null) {
        items.push({
          id: `pdf-row-${i}`,
          selected: true,
          code,
          description: desc,
          uom: normalizeUom(uomRaw || 'PCS'),
          qty,
          unitPrice,
        });
      }
    }
  }

  const rawGrid = parseDelimitedTextToTable(rawText, 'auto');
  const rawTables: RawExtractedTable[] =
    rawGrid.length > 0 ? [{ name: 'Pasted Text Table', rows: rawGrid }] : [];

  return {
    sourceType: 'text',
    filename,
    ...meta,
    detectedCodeLabel,
    items,
    rawExtractedText: rawText,
    rawTables,
  };
}

import ExcelJS from 'exceljs';
import { CURRENCY_OPTIONS, INCOTERM_OPTIONS, UOM_OPTIONS } from '../constants/themes';

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

function parseFlexibleNumber(raw: string | number | null | undefined): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number') return Number.isFinite(raw) ? raw : null;
  const cleaned = String(raw)
    .replace(/^(USD|AED|EUR|GBP|SAR|\$|€|£)\s*/i, '')
    .replace(/\s*(USD|AED|EUR|GBP|SAR)$/i, '')
    .replace(/,/g, '')
    .trim();
  if (!cleaned || !/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function normalizeUom(raw: string): string {
  const up = raw.trim().toUpperCase().replace(/\./g, '');
  if (!up) return 'PCS';
  if ((UOM_OPTIONS as readonly string[]).includes(up)) return up;
  if (['PC', 'EA', 'EACH', 'NOS', 'NO', 'UNIT', 'UNITS', 'PIECES', 'PIECE'].includes(up)) return 'PCS';
  if (['SETS', 'KIT', 'KITS'].includes(up)) return 'SET';
  if (['KGS', 'KILO', 'KILOS'].includes(up)) return 'KG';
  if (['LITRE', 'LITRES', 'LITER', 'LITERS', 'L', 'LTR', 'LTRS'].includes(up)) return 'LTR';
  if (['M', 'MTR', 'METER', 'METERS', 'METRE', 'METRES'].includes(up)) return 'MTR';
  if (['ROLLS', 'DRUM', 'DRUMS', 'REEL'].includes(up)) return 'ROLL';
  if (['BOXES', 'CTN', 'CARTON', 'CARTONS', 'CASE'].includes(up)) return 'BOX';
  if (['PK', 'PACK', 'PACKS', 'PKT', 'BAG', 'BAGS', 'BOTTLE', 'BTL'].includes(up)) return 'PACK';
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

  // Currency
  let currency = 'USD';
  for (const c of CURRENCY_OPTIONS) {
    if (new RegExp(`\\b${c}\\b`, 'i').test(fullText)) {
      currency = c;
      break;
    }
  }
  if (currency === 'USD' && /€/.test(fullText)) currency = 'EUR';
  else if (currency === 'USD' && /£/.test(fullText)) currency = 'GBP';

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
    /(?:quote\s*(?:no|#|ref|number)|quotation\s*(?:no|#|ref)|our\s*ref|ref\.?\s*(?:no)?)\s*[:#-]?\s*([A-Z0-9][A-Z0-9/_-]{2,26})/i
  );
  if (refMatch) supplierRef = refMatch[1].trim();

  // Payment terms
  let paymentTerms = '';
  const payLine = lines.find((l) => /payment\s*terms?|terms\s*of\s*payment/i.test(l));
  if (payLine) {
    paymentTerms = payLine.replace(/.*?(?:payment\s*terms?|terms\s*of\s*payment)\s*[:|-]?\s*/i, '').trim();
  }
  if (!paymentTerms) {
    const payMatch = fullText.match(
      /(\d+%\s*(?:advance|deposit|tt|t\/t|with\s*po|before\s*shipment)[^;\n]{0,60}|net\s*\d+\s*days?|100%\s*(?:advance|tt|t\/t|lc|l\/c)[^;\n]{0,40})/i
    );
    if (payMatch) paymentTerms = payMatch[1].trim();
  }

  // Delivery / Lead time
  let deliveryLeadTime = '';
  const delLine = lines.find((l) => /(?:delivery|lead\s*time|dispatch|shipment\s*time)/i.test(l) && /\d+|stock|ex-stock|immediate|week|day/i.test(l));
  if (delLine) {
    deliveryLeadTime = delLine
      .replace(/.*?(?:delivery\s*(?:time|period|terms)?|lead\s*time|dispatch)\s*[:|-]?\s*/i, '')
      .trim()
      .slice(0, 70);
  }

  // Supplier name: look for explicit label or first company-like header line
  let supplierName = '';
  const supLine = lines.find((l) => /^(?:from|supplier|vendor|company)\s*[:|-]/i.test(l));
  if (supLine) {
    supplierName = supLine.replace(/^(?:from|supplier|vendor|company)\s*[:|-]\s*/i, '').trim();
  }
  if (!supplierName) {
    const companyLine = lines
      .slice(0, 10)
      .find(
        (l) =>
          l.length >= 4 &&
          l.length <= 75 &&
          !/grand\s*group|quotation|proforma|invoice|date|page|tel|phone|email|www\.|http/i.test(l) &&
          /(?:llc|ltd|inc|corp|fzc|fze|gmbh|co\.|trading|supplies|industrial|medical|pharma|chem|tech|group|international)/i.test(l)
      );
    supplierName = companyLine || filename.replace(/\.(xlsx|pdf|txt)$/i, '').replace(/[_-]+/g, ' ').trim();
  }

  return {
    supplierName: supplierName.slice(0, 80),
    supplierRef: supplierRef.slice(0, 40),
    currency,
    incoterm,
    paymentTerms: paymentTerms.slice(0, 90),
    deliveryLeadTime,
  };
}

function isSummaryOrNoiseRow(desc: string): boolean {
  const s = desc.trim().toLowerCase();
  if (!s) return true;
  return /^(sub\s*-?\s*total|total|grand\s*total|net\s*total|vat|tax|freight|shipping|delivery\s*charge|discount|bank|terms|notes?|remarks?|validity|payment|thank\s*you|prepared\s*by|authorized)/i.test(
    s
  );
}

/**
 * 1. ANY-FORMAT SUPPLIER EXCEL (.XLSX) PARSER
 */
export function parseGenericSupplierXlsx(
  wb: ExcelJS.Workbook,
  filename: string
): ParsedSupplierQuote {
  interface ColMap {
    sheet: ExcelJS.Worksheet;
    headerRow: number;
    score: number;
    descCol: number;
    codeCol: number;
    codeLabel: string;
    uomCol: number;
    qtyCol: number;
    unitPriceCol: number;
    totalCol: number;
  }

  let best: ColMap | null = null;
  const allTextLines: string[] = [];

  for (const ws of wb.worksheets) {
    const maxScanRow = Math.min(ws.rowCount || 40, 45);
    const maxCol = Math.min(ws.columnCount || 18, 22);

    for (let r = 1; r <= maxScanRow; r++) {
      const row = ws.getRow(r);
      const rowTexts: string[] = [];
      let descCol = -1;
      let codeCol = -1;
      let codeLabel = 'Part No.';
      let uomCol = -1;
      let qtyCol = -1;
      let unitPriceCol = -1;
      let totalCol = -1;

      for (let c = 1; c <= maxCol; c++) {
        const txt = cellToText(row.getCell(c).value);
        if (!txt) continue;
        rowTexts.push(txt);
        const low = txt.toLowerCase().replace(/\s+/g, ' ');

        if (
          codeCol === -1 &&
          /\b(cas\s*(no\.?|#|num|number)?|part\s*(no\.?|#|num|number|code)|p\/n|pn|mpn|mfr\.?\s*part|oem|sku|item\s*(code|no\.?|#)|stock\s*(code|no\.?|#)|material\s*(code|no\.?|#)|product\s*(code|no\.?|#)|catalog|cat\.?\s*no|model\s*(no\.?|#)?|hs\s*code|ref\.?\s*(no\.?|#))\b/i.test(
            low
          )
        ) {
          codeCol = c;
          codeLabel = /cas/i.test(low) ? 'CAS No.' : /hs/i.test(low) ? 'HS Code' : 'Part No.';
        } else if (descCol === -1 && /\b(description|item\s*name|product|material|particulars|specification|goods|designation|details|item)\b/i.test(low) && !/code|no\.?$|#$/i.test(low)) {
          descCol = c;
        } else if (uomCol === -1 && /\b(uom|u\/m|unit|units|packing|pack)\b/i.test(low) && !/price|cost|rate|total/i.test(low)) {
          uomCol = c;
        } else if (qtyCol === -1 && /\b(qty|q'ty|quantity|qnty|order\s*qty|pcs|nos)\b/i.test(low)) {
          qtyCol = c;
        } else if (
          unitPriceCol === -1 &&
          /\b(unit\s*price|u\/price|u\.?\s*price|unit\s*cost|unit\s*rate|rate|price\s*per|net\s*price|exw|fob|cif|price)\b/i.test(low) &&
          !/total|amount|ext/i.test(low)
        ) {
          unitPriceCol = c;
        } else if (totalCol === -1 && /\b(total|amount|ext\.?\s*price|extended|line\s*total|net\s*amount|value)\b/i.test(low)) {
          totalCol = c;
        }
      }

      if (rowTexts.length > 0 && r <= 20) {
        allTextLines.push(rowTexts.join('  '));
      }

      const score =
        (descCol !== -1 ? 3 : 0) +
        (qtyCol !== -1 ? 2 : 0) +
        (unitPriceCol !== -1 ? 3 : 0) +
        (totalCol !== -1 ? 1 : 0) +
        (uomCol !== -1 ? 1 : 0) +
        (codeCol !== -1 ? 1 : 0);

      if (score >= 4 && (!best || score > best.score)) {
        best = {
          sheet: ws,
          headerRow: r,
          score,
          descCol,
          codeCol,
          codeLabel,
          uomCol,
          qtyCol,
          unitPriceCol,
          totalCol,
        };
      }
    }
  }

  const items: ParsedQuoteItem[] = [];
  let detectedCodeLabel = 'Part No.';

  if (best) {
    detectedCodeLabel = best.codeLabel;
    const ws = best.sheet;
    const maxRow = Math.min(ws.rowCount || 120, best.headerRow + 80);

    for (let r = best.headerRow + 1; r <= maxRow && items.length < 20; r++) {
      const row = ws.getRow(r);
      const desc = best.descCol !== -1 ? cellToText(row.getCell(best.descCol).value) : '';
      const code = best.codeCol !== -1 ? cellToText(row.getCell(best.codeCol).value) : '';
      const uomRaw = best.uomCol !== -1 ? cellToText(row.getCell(best.uomCol).value) : '';
      const qty = best.qtyCol !== -1 ? parseFlexibleNumber(cellToText(row.getCell(best.qtyCol).value)) : null;
      let unitPrice =
        best.unitPriceCol !== -1
          ? parseFlexibleNumber(cellToText(row.getCell(best.unitPriceCol).value))
          : null;
      const total =
        best.totalCol !== -1
          ? parseFlexibleNumber(cellToText(row.getCell(best.totalCol).value))
          : null;

      // Capture footer text lines for payment/delivery terms
      if (!desc && !code && qty === null && unitPrice === null) {
        const lineJoin = [];
        for (let c = 1; c <= 10; c++) {
          const t = cellToText(row.getCell(c).value);
          if (t) lineJoin.push(t);
        }
        if (lineJoin.length > 0) allTextLines.push(lineJoin.join(' '));
        continue;
      }

      if (isSummaryOrNoiseRow(desc) && !code) {
        const lineJoin = [];
        for (let c = 1; c <= 10; c++) {
          const t = cellToText(row.getCell(c).value);
          if (t) lineJoin.push(t);
        }
        if (lineJoin.length > 0) allTextLines.push(lineJoin.join(' '));
        continue;
      }

      if (unitPrice === null && total !== null && qty !== null && qty > 0) {
        unitPrice = Math.round((total / qty) * 100) / 100;
      }

      if ((desc || code) && (qty !== null || unitPrice !== null)) {
        items.push({
          id: `q-row-${r}`,
          selected: true,
          code,
          description: desc || code,
          uom: normalizeUom(uomRaw),
          qty: qty && qty > 0 ? qty : 1,
          unitPrice,
        });
      }
    }
  } else {
    // Fallback: extract text rows across first worksheet and run line parser
    const ws = wb.worksheets[0];
    if (ws) {
      const lines: string[] = [];
      ws.eachRow((row) => {
        const vals: string[] = [];
        row.eachCell((cell) => {
          const t = cellToText(cell.value);
          if (t) vals.push(t);
        });
        if (vals.length > 0) lines.push(vals.join('\t'));
      });
      return {
        ...parsePastedSupplierQuote(lines.join('\n'), filename),
        sourceType: 'xlsx',
      };
    }
  }

  const meta = extractMetadataFromText(allTextLines, filename);
  return {
    sourceType: 'xlsx',
    filename,
    ...meta,
    detectedCodeLabel,
    items,
  };
}

/**
 * 2. PURE CLIENT-SIDE OFFLINE PDF TEXT EXTRACTOR (FlateDecode + ToUnicode CMap + Line Grouping)
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

    // beginbfchar ... endbfchar
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

    // beginbfrange ... endbfrange
    const bfrangeBlocks = txt.matchAll(/beginbfrange([\s\S]*?)endbfrange/g);
    for (const block of bfrangeBlocks) {
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
    // Try 4-hex-digit (2-byte CID) lookup first, then 2-hex-digit
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
  // Plain ASCII/UTF-16BE fallback
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

function extractLinesFromContentStream(stream: string, cmap: Map<string, string>): string[] {
  const tokens: { y: number; x: number; text: string }[] = [];
  const btBlocks = stream.matchAll(/BT([\s\S]*?)ET/g);

  for (const block of btBlocks) {
    const body = block[1];
    let curX = 0;
    let curY = 0;
    let leading = 12;

    // Scan operators inside BT...ET
    const opRegex =
      /(\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]+>|\[(?:[^\[\]]*)\])\s*(Tj|TJ|'|"|)|(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+Tm|(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(Td|TD)|(-?\d+(?:\.\d+)?)\s+TL|\bT\*/g;

    let match: RegExpExecArray | null;
    while ((match = opRegex.exec(body)) !== null) {
      if (match[9] !== undefined) {
        // Tm
        curX = parseFloat(match[7]);
        curY = parseFloat(match[8]);
      } else if (match[12] === 'Td' || match[12] === 'TD') {
        const tx = parseFloat(match[10]);
        const ty = parseFloat(match[11]);
        curX += tx;
        curY += ty;
        if (match[12] === 'TD') leading = -ty;
      } else if (match[13] !== undefined) {
        leading = parseFloat(match[13]);
      } else if (match[0] === 'T*') {
        curY -= leading;
      } else if (match[1] && match[2]) {
        const arg = match[1];
        const op = match[2];
        if (op === "'" || op === '"') curY -= leading;

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
          tokens.push({ y: Math.round(curY / 3) * 3, x: curX, text: cleaned });
        }
      }
    }
  }

  // Group tokens by approximate Y line (top to bottom) and sort left-to-right by X
  const lineMap = new Map<number, { x: number; text: string }[]>();
  for (const t of tokens) {
    const bucket = lineMap.get(t.y) || [];
    bucket.push({ x: t.x, text: t.text });
    lineMap.set(t.y, bucket);
  }

  const sortedYs = Array.from(lineMap.keys()).sort((a, b) => b - a);
  const lines: string[] = [];
  for (const y of sortedYs) {
    const rowItems = lineMap.get(y)!.sort((a, b) => a.x - b.x);
    let line = '';
    for (let i = 0; i < rowItems.length; i++) {
      if (i > 0) {
        const gap = rowItems[i].x - rowItems[i - 1].x;
        line += gap > 25 ? '\t' : ' ';
      }
      line += rowItems[i].text;
    }
    const trimmed = line.trim();
    if (trimmed) lines.push(trimmed);
  }
  return lines;
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

    // Inspect preceding dictionary for FlateDecode
    const dictSnippet = rawLatin1.slice(Math.max(0, streamPos - 300), streamPos);
    let dataStart = streamPos + 6;
    if (rawLatin1[dataStart] === '\r' && rawLatin1[dataStart + 1] === '\n') dataStart += 2;
    else if (rawLatin1[dataStart] === '\n' || rawLatin1[dataStart] === '\r') dataStart += 1;

    let dataEnd = endPos;
    while (dataEnd > dataStart && (rawLatin1[dataEnd - 1] === '\n' || rawLatin1[dataEnd - 1] === '\r')) {
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
  for (const s of decodedStreams) {
    if (s.includes('BT') && (s.includes('Tj') || s.includes('TJ'))) {
      allLines.push(...extractLinesFromContentStream(s, cmap));
    }
  }

  const rawText = allLines.join('\n').trim();
  const parsed = parsePastedSupplierQuote(rawText, file.name);
  return {
    ...parsed,
    sourceType: 'pdf',
    rawExtractedText: rawText,
    isScannedPdf: rawText.length < 30,
  };
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
    // Skip table header rows
    if (/\b(description|unit\s*price|qty|quantity)\b/i.test(line) && /\b(total|amount|uom|unit|rate|price)\b/i.test(line)) {
      continue;
    }

    // Strategy A: Tab-separated or multi-space/pipe columns
    const cols = line
      .split(/\t|\s{2,}|\|/)
      .map((c) => c.trim())
      .filter(Boolean);

    if (cols.length >= 3) {
      // Find numeric cells from right to left
      const numIndices: { idx: number; val: number }[] = [];
      cols.forEach((c, idx) => {
        const n = parseFlexibleNumber(c);
        if (n !== null && n > 0) numIndices.push({ idx, val: n });
      });

      if (numIndices.length >= 2) {
        // Ignore leading serial number (1, 2, 3...)
        const dataNums =
          numIndices[0].idx === 0 && numIndices[0].val <= 50 && Number.isInteger(numIndices[0].val)
            ? numIndices.slice(1)
            : numIndices;

        if (dataNums.length >= 2) {
          let qty: number | null = null;
          let unitPrice: number | null = null;

          if (dataNums.length >= 3) {
            // Typical pattern: Qty, Unit Price, Total
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
            // 2 numbers: Qty and Unit Price (or Total)
            qty = dataNums[0].val;
            unitPrice = dataNums[1].val;
          }

          // Collect non-numeric text columns
          const numSet = new Set(numIndices.map((n) => n.idx));
          let uom = 'PCS';
          let code = '';
          const descParts: string[] = [];

          cols.forEach((c, idx) => {
            if (numSet.has(idx)) return;
            if (/^(pcs?|ea|each|nos?|set|sets|kg|kgs|ltr|mtr|roll|box|ctn|pack|pair|kit|unit|units|drum|bag|btl)\.?$/i.test(c)) {
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
      /^(?:\d{1,2}[.)]?\s+)?(.{5,110}?)\s+(\d+(?:\.\d+)?)\s+(?:(PCS|PC|EA|EACH|NOS|SET|KG|LTR|MTR|ROLL|BOX|PACK|PAIR|UNIT|UNITS|KIT|DRUM)\s+)?(?:USD|AED|EUR|GBP|SAR|\$|€|£)?\s*([\d,]+\.\d{2})(?:\s+(?:USD|AED|EUR|GBP|SAR|\$|€|£)?\s*([\d,]+\.\d{2}))?$/i
    );
    if (singleMatch) {
      const desc = singleMatch[1].trim();
      const qty = parseFlexibleNumber(singleMatch[2]);
      const uom = normalizeUom(singleMatch[3] || 'PCS');
      const unitPrice = parseFlexibleNumber(singleMatch[4]);
      if (desc && !isSummaryOrNoiseRow(desc) && qty && unitPrice !== null) {
        // Pull CAS or Part code if embedded at start
        const casMatch = desc.match(/\b(\d{2,7}-\d{2}-\d)\b/);
        items.push({
          id: `pdf-row-${i}`,
          selected: true,
          code: casMatch ? casMatch[1] : '',
          description: desc,
          uom,
          qty,
          unitPrice,
        });
      }
    }
  }

  return {
    sourceType: 'text',
    filename,
    ...meta,
    detectedCodeLabel,
    items,
    rawExtractedText: rawText,
  };
}

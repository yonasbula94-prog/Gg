import { MARKUP_RATES } from '../constants/themes';
import {
  ComputedRow,
  DocPrefix,
  JobCalculationSummary,
  JobItemRow,
  MarkupRate,
  ProcurementJob,
  ValidationIssue,
} from '../types/procurement';

export function round2(val: number): number {
  if (!Number.isFinite(val)) return 0;
  return Math.round((val + Number.EPSILON) * 100) / 100;
}

export function round4(val: number): number {
  if (!Number.isFinite(val)) return 0;
  return Math.round((val + Number.EPSILON) * 10000) / 10000;
}

function safeNonNeg(val: number | null | undefined): number | null {
  if (val === null || val === undefined || typeof val !== 'number' || !Number.isFinite(val)) {
    return null;
  }
  return val < 0 ? 0 : val;
}

function clampRate(val: number | undefined, fallback: number, max: number): number {
  if (typeof val !== 'number' || !Number.isFinite(val)) return fallback;
  if (val < 0) return 0;
  if (val > max) return max;
  return val;
}

export function getResolvedEndUser(job: ProcurementJob): string {
  if (job.controls.endUserSelect === 'OTHER') {
    return (job.controls.customEndUser || '').trim();
  }
  return (job.controls.endUserSelect || '').trim();
}

export function isRowEmpty(row: JobItemRow): boolean {
  const hasCustom = Object.values(row.customFields || {}).some((v) => v && v.trim() !== '');
  return (
    !row.requestedDescription.trim() &&
    !row.offeredDescription.trim() &&
    !row.uom.trim() &&
    !hasCustom &&
    (row.qty === null || !Number.isFinite(row.qty) || row.qty === 0)
  );
}

/**
 * Computes all pricing, freight, internal VAT, and logistics figures:
 * - Markup & Internal VAT folded directly into customer Unit Price:
 *   Unit Price = ROUND(Supplier Cost * (1 + Markup %) * (1 + Effective VAT %), 2)
 *   where Effective VAT % = includeVat ? vatRate : 0
 * - End-user documents (PI / QT / CI) NEVER show a separate VAT row:
 *   Customer Goods Subtotal = SUM(Qty * Active Unit Price)
 *   Customer Grand Total = Customer Goods Subtotal + Customer Freight
 * - Internal PRICE CALC tracks Goods Excl. VAT, Embedded VAT, and true Net Gross Profit (Excl. VAT)
 * - Freight = MAX(Total CBM * Rate/CBM, Total Gross KG * Rate/KG)
 *   + Documentation charge (default 10%) FIRST, then selected markup % applied on top.
 */
export function calculateJob(job: ProcurementJob): JobCalculationSummary {
  const activeMarkup = clampRate(job.freight.selectedMarkupPct, 0.2, 5);
  const rawVatRate = clampRate(job.controls.vatRate, 0.05, 1);
  const includeVat = job.controls.includeVat !== false;
  const effectiveVatRate = includeVat ? rawVatRate : 0;
  const vatMultiplier = 1 + effectiveVatRate;

  let activeRowCount = 0;
  let totalQty = 0;
  let totalSupplierCost = 0;
  let totalBoxes = 0;
  let totalCbm = 0;
  let totalGrossKgCalc = 0;
  let goodsSubtotalExVat = 0;
  let goodsSubtotal = 0;
  let plTotalNetWt = 0;
  let plTotalGrossWt = 0;

  const scenarioGoodsExVat: Record<MarkupRate, number> = {
    0.1: 0,
    0.15: 0,
    0.18: 0,
    0.2: 0,
  };
  const scenarioGoodsInclVat: Record<MarkupRate, number> = {
    0.1: 0,
    0.15: 0,
    0.18: 0,
    0.2: 0,
  };

  const rows: ComputedRow[] = job.items.map((item, idx) => {
    const empty = isRowEmpty(item);
    const qty = safeNonNeg(item.qty);
    const cost = safeNonNeg(item.supplierUnitCost);

    if (!empty) {
      activeRowCount += 1;
      if (qty !== null) totalQty += qty;
    }

    const supplierTotalCost =
      !empty && qty !== null && cost !== null ? round2(qty * cost) : null;
    if (supplierTotalCost !== null) {
      totalSupplierCost += supplierTotalCost;
    }

    // All selling unit prices include internal VAT when includeVat is ON
    const unitPrice10 = !empty && cost !== null ? round2(cost * 1.1 * vatMultiplier) : null;
    const unitPrice15 = !empty && cost !== null ? round2(cost * 1.15 * vatMultiplier) : null;
    const unitPrice18 = !empty && cost !== null ? round2(cost * 1.18 * vatMultiplier) : null;
    const unitPrice20 = !empty && cost !== null ? round2(cost * 1.2 * vatMultiplier) : null;

    if (!empty && qty !== null && cost !== null) {
      for (const rate of MARKUP_RATES) {
        const uExVat = round2(cost * (1 + rate));
        const uInclVat = round2(cost * (1 + rate) * vatMultiplier);
        scenarioGoodsExVat[rate] += round2(qty * uExVat);
        scenarioGoodsInclVat[rate] += round2(qty * uInclVat);
      }
    }

    const activeUnitPriceExVat =
      !empty && cost !== null ? round2(cost * (1 + activeMarkup)) : null;
    const activeUnitPrice =
      !empty && cost !== null ? round2(cost * (1 + activeMarkup) * vatMultiplier) : null;
    const activeTotalPrice =
      !empty && qty !== null && activeUnitPrice !== null
        ? round2(qty * activeUnitPrice)
        : null;

    if (activeTotalPrice !== null) {
      goodsSubtotal += activeTotalPrice;
    }
    if (!empty && qty !== null && activeUnitPriceExVat !== null) {
      goodsSubtotalExVat += round2(qty * activeUnitPriceExVat);
    }

    const boxes = safeNonNeg(item.boxes);
    const l = safeNonNeg(item.boxLengthCm) ?? 0;
    const w = safeNonNeg(item.boxWidthCm) ?? 0;
    const h = safeNonNeg(item.boxHeightCm) ?? 0;
    const kgPerBox = safeNonNeg(item.kgPerBox) ?? 0;

    const cbm =
      !empty && boxes !== null && l > 0 && w > 0 && h > 0
        ? round4((boxes * l * w * h) / 1_000_000)
        : null;
    const grossKgCalc =
      !empty && boxes !== null && kgPerBox > 0 ? round2(boxes * kgPerBox) : null;

    const safeNetWt = safeNonNeg(item.netWtKg);
    const safeGrossWt = safeNonNeg(item.grossWtKg);

    if (!empty) {
      if (boxes !== null) totalBoxes += boxes;
      if (cbm !== null) totalCbm += cbm;
      if (grossKgCalc !== null) totalGrossKgCalc += grossKgCalc;
      if (safeNetWt !== null) plTotalNetWt += safeNetWt;
      if (safeGrossWt !== null) plTotalGrossWt += safeGrossWt;
    }

    return {
      rowNumber: item.rowNumber,
      index: idx + 1,
      isEmpty: empty,
      requestedDescription: item.requestedDescription,
      offeredDescription: item.offeredDescription,
      customFields: item.customFields || {},
      uom: item.uom,
      qty,
      supplierUnitCost: cost,
      supplierTotalCost,
      unitPrice10,
      unitPrice15,
      unitPrice18,
      unitPrice20,
      activeUnitPrice,
      activeTotalPrice,
      boxes,
      cbm,
      grossKgCalc,
      cartonNo: item.cartonNo,
      netWtKg: safeNetWt,
      grossWtKg: safeGrossWt,
    };
  });

  totalSupplierCost = round2(totalSupplierCost);
  totalCbm = round4(totalCbm);
  totalGrossKgCalc = round2(totalGrossKgCalc);
  goodsSubtotalExVat = round2(goodsSubtotalExVat);
  goodsSubtotal = round2(goodsSubtotal);
  plTotalNetWt = round2(plTotalNetWt);
  plTotalGrossWt = round2(plTotalGrossWt);

  // Freight calculation
  const ratePerCbm = safeNonNeg(job.freight.ratePerCbm) ?? 0;
  const ratePerKg = safeNonNeg(job.freight.ratePerKg) ?? 0;
  const docChargePct = clampRate(job.freight.docChargePct, 0.1, 1);

  const freightByCbm = round2(totalCbm * ratePerCbm);
  const freightByKg = round2(totalGrossKgCalc * ratePerKg);
  const baseFreight = Math.max(freightByCbm, freightByKg);
  const freightBasis: 'CBM' | 'KG' | 'NONE' =
    baseFreight === 0 ? 'NONE' : freightByCbm >= freightByKg ? 'CBM' : 'KG';

  const docChargeAmount = round2(baseFreight * docChargePct);
  const freightWithDoc = round2(baseFreight + docChargeAmount);
  const freightMarkupAmount = round2(freightWithDoc * activeMarkup);
  const customerFreight = round2(freightWithDoc + freightMarkupAmount);

  // Embedded internal VAT (already inside Unit Prices & goodsSubtotal — never added again!)
  const vatAmount =
    effectiveVatRate > 0 ? round2(Math.max(0, goodsSubtotal - goodsSubtotalExVat)) : 0;
  const grandTotal = round2(goodsSubtotal + customerFreight);

  // True Net Gross Profit (Excl. VAT pass-through)
  const grossProfit = round2(goodsSubtotalExVat - totalSupplierCost + freightMarkupAmount);

  const scenarioTotals = {} as JobCalculationSummary['scenarioTotals'];
  for (const rate of MARKUP_RATES) {
    const sGoodsExVat = round2(scenarioGoodsExVat[rate]);
    const sGoods = round2(scenarioGoodsInclVat[rate]);
    const sVat = effectiveVatRate > 0 ? round2(Math.max(0, sGoods - sGoodsExVat)) : 0;
    const sFreightMarkup = round2(freightWithDoc * rate);
    const sFreight = round2(freightWithDoc + sFreightMarkup);
    scenarioTotals[rate] = {
      goodsSubtotalExVat: sGoodsExVat,
      goodsSubtotal: sGoods,
      vatAmount: sVat,
      customerFreight: sFreight,
      grandTotal: round2(sGoods + sFreight),
      grossProfit: round2(sGoodsExVat - totalSupplierCost + sFreightMarkup),
    };
  }

  return {
    rows,
    activeRowCount,
    totalQty: round2(totalQty),
    totalSupplierCost,
    totalBoxes,
    totalCbm,
    totalGrossKgCalc,
    freightByCbm,
    freightByKg,
    freightBasis,
    baseFreight,
    docChargeAmount,
    freightWithDoc,
    freightMarkupAmount,
    customerFreight,
    effectiveVatRate,
    goodsSubtotalExVat,
    vatAmount,
    goodsSubtotal,
    grandTotal,
    scenarioTotals,
    grossProfit,
    plTotalNetWt,
    plTotalGrossWt,
  };
}

/**
 * Validates a document before claiming a number from Master Registry.
 */
export function validateForIssue(
  job: ProcurementJob,
  prefix: DocPrefix,
  calc: JobCalculationSummary
): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  if (job.issuedDocs[prefix] && !job.issuedDocs[prefix]?.isVoided) {
    issues.push({
      field: 'issued',
      message: `${prefix} has already been issued (${job.issuedDocs[prefix]?.docNumber}). Create a new job to re-issue.`,
    });
    return issues;
  }

  const endUser = getResolvedEndUser(job);
  if (prefix === 'PO') {
    if (!job.controls.supplierName.trim()) {
      issues.push({
        field: 'supplierName',
        message: 'Cannot issue yet: enter Supplier Name on Supplier PO.',
      });
    }
  } else {
    if (!endUser) {
      issues.push({
        field: 'endUser',
        message: 'Cannot issue yet: select an End User (or type custom name for OTHER).',
      });
    }
  }

  if (calc.activeRowCount === 0) {
    issues.push({
      field: 'items',
      message: 'Cannot issue yet: enter at least one line item in rows 17–36.',
    });
    return issues;
  }

  for (const row of calc.rows) {
    if (row.isEmpty) continue;

    if (!row.requestedDescription.trim() && !row.offeredDescription.trim()) {
      issues.push({
        rowNumber: row.rowNumber,
        field: 'description',
        message: `Cannot issue yet: row ${row.rowNumber} is missing item description.`,
      });
    }
    if (!row.uom.trim()) {
      issues.push({
        rowNumber: row.rowNumber,
        field: 'uom',
        message: `Cannot issue yet: row ${row.rowNumber} is missing UoM.`,
      });
    }
    if (row.qty === null || row.qty <= 0) {
      issues.push({
        rowNumber: row.rowNumber,
        field: 'qty',
        message: `Cannot issue yet: row ${row.rowNumber} must have Qty greater than 0.`,
      });
    }

    if (prefix === 'QT' || prefix === 'PI' || prefix === 'CI' || prefix === 'PO') {
      if (row.supplierUnitCost === null || row.supplierUnitCost <= 0) {
        issues.push({
          rowNumber: row.rowNumber,
          field: 'supplierUnitCost',
          message: `Cannot issue yet: row ${row.rowNumber} is missing Supplier Unit Cost on PRICE CALC.`,
        });
      }
    }

    if (prefix === 'PL' || prefix === 'DN') {
      if (!row.cartonNo.trim()) {
        issues.push({
          rowNumber: row.rowNumber,
          field: 'cartonNo',
          message: `Cannot issue yet: row ${row.rowNumber} is missing Carton No.`,
        });
      }
      if (row.grossWtKg === null || row.grossWtKg <= 0) {
        issues.push({
          rowNumber: row.rowNumber,
          field: 'grossWtKg',
          message: `Cannot issue yet: row ${row.rowNumber} is missing Gross Wt (KG).`,
        });
      }
    }
  }

  if (prefix === 'PL' && !job.controls.shipDate.trim()) {
    issues.push({
      field: 'shipDate',
      message: 'Cannot issue yet: enter Ship Date on Packing List.',
    });
  }

  if (prefix === 'DN') {
    if (!job.controls.deliveryDate.trim()) {
      issues.push({
        field: 'deliveryDate',
        message: 'Cannot issue yet: enter Delivery Date on Delivery Note.',
      });
    }
    if (!job.controls.transportRef.trim()) {
      issues.push({
        field: 'transportRef',
        message: 'Cannot issue yet: enter Transport Ref. on Delivery Note.',
      });
    }
  }

  return issues;
}

export function formatMoney(amount: number | null, currency?: string): string {
  if (amount === null || !Number.isFinite(amount)) return '';
  const formatted = amount.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return currency ? `${currency} ${formatted}` : formatted;
}

export function formatNumber(val: number | null, decimals = 2): string {
  if (val === null || !Number.isFinite(val)) return '';
  return val.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function formatPct(rate: number): string {
  if (!Number.isFinite(rate)) return '0%';
  const pct = Math.round((rate * 100 + Number.EPSILON) * 100) / 100;
  return `${pct}%`;
}

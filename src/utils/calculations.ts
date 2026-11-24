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
  return Math.round((val + Number.EPSILON) * 100) / 100;
}

export function round4(val: number): number {
  return Math.round((val + Number.EPSILON) * 10000) / 10000;
}

export function getResolvedEndUser(job: ProcurementJob): string {
  if (job.controls.endUserSelect === 'OTHER') {
    return job.controls.customEndUser.trim();
  }
  return job.controls.endUserSelect.trim();
}

export function isRowEmpty(row: JobItemRow): boolean {
  return (
    !row.requestedDescription.trim() &&
    !row.offeredDescription.trim() &&
    !row.uom.trim() &&
    (row.qty === null || row.qty === 0)
  );
}

/**
 * Computes all pricing, freight, VAT, and logistics figures following Grand Group FZC's
 * confirmed template specification:
 * - Markup on cost: Unit Price = Supplier Cost * (1 + %)
 * - CBM = Boxes * L * W * H / 1,000,000
 * - Base Freight = MAX(Total CBM * Rate/CBM, Total Gross KG * Rate/KG)
 * - Documentation charge (default 10%) added to freight FIRST, then selected markup % added on top
 * - VAT applied to goods SUBTOTAL only (default 5%, 0% for zero-rated)
 * - Freight is always a separate line AFTER VAT
 * - GRAND TOTAL = SUBTOTAL + VAT + FREIGHT
 */
export function calculateJob(job: ProcurementJob): JobCalculationSummary {
  const activeMarkup = job.freight.selectedMarkupPct;

  let activeRowCount = 0;
  let totalQty = 0;
  let totalSupplierCost = 0;
  let totalBoxes = 0;
  let totalCbm = 0;
  let totalGrossKgCalc = 0;
  let goodsSubtotal = 0;
  let plTotalNetWt = 0;
  let plTotalGrossWt = 0;

  const scenarioGoods: Record<MarkupRate, number> = {
    0.1: 0,
    0.15: 0,
    0.18: 0,
    0.2: 0,
  };

  const rows: ComputedRow[] = job.items.map((item, idx) => {
    const empty = isRowEmpty(item);
    const qty = item.qty !== null && !Number.isNaN(item.qty) ? item.qty : null;
    const cost =
      item.supplierUnitCost !== null && !Number.isNaN(item.supplierUnitCost)
        ? item.supplierUnitCost
        : null;

    if (!empty) {
      activeRowCount += 1;
      if (qty !== null) totalQty += qty;
    }

    const supplierTotalCost =
      !empty && qty !== null && cost !== null ? round2(qty * cost) : null;
    if (supplierTotalCost !== null) {
      totalSupplierCost += supplierTotalCost;
    }

    const unitPrice10 = !empty && cost !== null ? round2(cost * 1.1) : null;
    const unitPrice15 = !empty && cost !== null ? round2(cost * 1.15) : null;
    const unitPrice18 = !empty && cost !== null ? round2(cost * 1.18) : null;
    const unitPrice20 = !empty && cost !== null ? round2(cost * 1.2) : null;

    if (!empty && qty !== null && cost !== null) {
      scenarioGoods[0.1] += round2(qty * round2(cost * 1.1));
      scenarioGoods[0.15] += round2(qty * round2(cost * 1.15));
      scenarioGoods[0.18] += round2(qty * round2(cost * 1.18));
      scenarioGoods[0.2] += round2(qty * round2(cost * 1.2));
    }

    const activeUnitPrice =
      !empty && cost !== null ? round2(cost * (1 + activeMarkup)) : null;
    const activeTotalPrice =
      !empty && qty !== null && activeUnitPrice !== null
        ? round2(qty * activeUnitPrice)
        : null;

    if (activeTotalPrice !== null) {
      goodsSubtotal += activeTotalPrice;
    }

    const boxes = item.boxes !== null && !Number.isNaN(item.boxes) ? item.boxes : null;
    const l = item.boxLengthCm ?? 0;
    const w = item.boxWidthCm ?? 0;
    const h = item.boxHeightCm ?? 0;
    const kgPerBox = item.kgPerBox ?? 0;

    const cbm =
      !empty && boxes !== null && l > 0 && w > 0 && h > 0
        ? round4((boxes * l * w * h) / 1_000_000)
        : null;
    const grossKgCalc =
      !empty && boxes !== null && kgPerBox > 0 ? round2(boxes * kgPerBox) : null;

    if (!empty) {
      if (boxes !== null) totalBoxes += boxes;
      if (cbm !== null) totalCbm += cbm;
      if (grossKgCalc !== null) totalGrossKgCalc += grossKgCalc;
      if (item.netWtKg !== null && !Number.isNaN(item.netWtKg)) {
        plTotalNetWt += item.netWtKg;
      }
      if (item.grossWtKg !== null && !Number.isNaN(item.grossWtKg)) {
        plTotalGrossWt += item.grossWtKg;
      }
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
      netWtKg: item.netWtKg,
      grossWtKg: item.grossWtKg,
    };
  });

  totalSupplierCost = round2(totalSupplierCost);
  totalCbm = round4(totalCbm);
  totalGrossKgCalc = round2(totalGrossKgCalc);
  goodsSubtotal = round2(goodsSubtotal);
  plTotalNetWt = round2(plTotalNetWt);
  plTotalGrossWt = round2(plTotalGrossWt);

  // Freight calculation
  const freightByCbm = round2(totalCbm * (job.freight.ratePerCbm || 0));
  const freightByKg = round2(totalGrossKgCalc * (job.freight.ratePerKg || 0));
  const baseFreight = Math.max(freightByCbm, freightByKg);
  const freightBasis: 'CBM' | 'KG' | 'NONE' =
    baseFreight === 0 ? 'NONE' : freightByCbm >= freightByKg ? 'CBM' : 'KG';

  const docChargeAmount = round2(baseFreight * (job.freight.docChargePct || 0));
  const freightWithDoc = round2(baseFreight + docChargeAmount);
  const freightMarkupAmount = round2(freightWithDoc * activeMarkup);
  const customerFreight = round2(freightWithDoc + freightMarkupAmount);

  // VAT on goods subtotal ONLY
  const vatRate = job.controls.vatRate ?? 0.05;
  const vatAmount = round2(goodsSubtotal * vatRate);
  const grandTotal = round2(goodsSubtotal + vatAmount + customerFreight);

  // Profit = Customer Goods + Customer Freight - Supplier Goods - FreightWithDoc
  const grossProfit = round2(goodsSubtotal - totalSupplierCost + freightMarkupAmount);

  const scenarioTotals = {} as JobCalculationSummary['scenarioTotals'];
  for (const rate of MARKUP_RATES) {
    const sGoods = round2(scenarioGoods[rate]);
    const sVat = round2(sGoods * vatRate);
    const sFreightMarkup = round2(freightWithDoc * rate);
    const sFreight = round2(freightWithDoc + sFreightMarkup);
    scenarioTotals[rate] = {
      goodsSubtotal: sGoods,
      vatAmount: sVat,
      customerFreight: sFreight,
      grandTotal: round2(sGoods + sVat + sFreight),
      grossProfit: round2(sGoods - totalSupplierCost + sFreightMarkup),
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
    goodsSubtotal,
    vatAmount,
    grandTotal,
    scenarioTotals,
    grossProfit,
    plTotalNetWt,
    plTotalGrossWt,
  };
}

/**
 * Validates a document before claiming a number from Master Registry.
 * Matches VBA rule: "Validate before claiming a number" and reports exact row errors
 * like "Cannot issue yet: row 17 — missing quantity".
 */
export function validateForIssue(
  job: ProcurementJob,
  prefix: DocPrefix,
  calc: JobCalculationSummary
): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  // Already issued check (unless voided, though per spec re-issue means new job)
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
  if (amount === null || Number.isNaN(amount)) return '';
  const formatted = amount.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  return currency ? `${currency} ${formatted}` : formatted;
}

export function formatNumber(val: number | null, decimals = 2): string {
  if (val === null || Number.isNaN(val)) return '';
  return val.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function formatPct(rate: number): string {
  const pct = Math.round((rate * 100 + Number.EPSILON) * 100) / 100;
  return `${pct}%`;
}

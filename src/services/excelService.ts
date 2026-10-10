import ExcelJS from "exceljs";
import type { PagoDoc, Factura, Concepto } from "./parser";

// ─── Palette ────────────────────────────────────────────────────────────────
const HEADER_FILL  = "1F4E79";
const SUB_FILL     = "2E75B6";
const ALT_FILL     = "D6E4F0";
const WHITE_FILL   = "FFFFFF";
const TOTAL_FILL   = "BDD7EE";
const BORDER_COLOR = "AAAAAA";
const MONEY_FMT    = "#,##0.00";

type HAlign = "left" | "center" | "right";

// ─── Helpers ─────────────────────────────────────────────────────────────────

function solidFill(hex: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
}

function thinBorder(): Partial<ExcelJS.Borders> {
  const s: Partial<ExcelJS.Border> = { style: "thin", color: { argb: "FF" + BORDER_COLOR } };
  return { left: s, right: s, top: s, bottom: s };
}

function titleRow(ws: ExcelJS.Worksheet, text: string, numCols: number): void {
  ws.getRow(1).height = 26;
  const cell = ws.getCell("A1");
  cell.value = text;
  cell.fill = solidFill(HEADER_FILL);
  cell.font = { name: "Arial", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.border = thinBorder();
  ws.mergeCells(1, 1, 1, numCols);
}

function headerCell(
  row: ExcelJS.Row,
  col: number,
  text: string,
  align: HAlign = "center"
): void {
  const cell = row.getCell(col);
  cell.value = text;
  cell.fill = solidFill(SUB_FILL);
  cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
  cell.alignment = { horizontal: align, vertical: "middle", wrapText: true };
  cell.border = thinBorder();
}

function dataCell(
  row: ExcelJS.Row,
  col: number,
  value: ExcelJS.CellValue,
  opts: {
    fmt?: string;
    fill?: string;
    bold?: boolean;
    align?: HAlign;
  } = {}
): void {
  const cell = row.getCell(col);
  cell.value = value;
  cell.fill = solidFill(opts.fill ?? WHITE_FILL);
  cell.font = { name: "Arial", size: 10, bold: opts.bold ?? false };
  cell.alignment = { horizontal: opts.align ?? "left", vertical: "middle" };
  cell.border = thinBorder();
  if (opts.fmt) cell.numFmt = opts.fmt;
}

function totalRow(
  ws: ExcelJS.Worksheet,
  rowNum: number,
  labelCol: number,
  sumCols: number[],
  dataStart: number,
  numCols: number,
  omitGlobalFmt: boolean = false
): void {
  const row = ws.getRow(rowNum);
  row.height = 18;

  for (let c = 1; c <= numCols; c++) {
    if (c === labelCol) {
      dataCell(row, c, "TOTAL", { fill: TOTAL_FILL, bold: true, align: "right" });
    } else if (sumCols.includes(c)) {
      const colLetter = ws.getColumn(c).letter;
      const formula = `SUM(${colLetter}${dataStart}:${colLetter}${rowNum - 1})`;
      const cell = row.getCell(c);
      cell.value = { formula };
      cell.fill = solidFill(TOTAL_FILL);
      cell.font = { name: "Arial", size: 10, bold: true };
      cell.alignment = { horizontal: "right", vertical: "middle" };
      cell.border = thinBorder();
      if (!omitGlobalFmt) cell.numFmt = MONEY_FMT;
    } else {
      dataCell(row, c, null, { fill: TOTAL_FILL });
    }
  }
}

function setColWidths(ws: ExcelJS.Worksheet, widths: number[]): void {
  widths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });
}

// ─── Sheet builders ──────────────────────────────────────────────────────────

function buildPagosSheet(wb: ExcelJS.Workbook, pagos: PagoDoc[]): void {
  const ws = wb.addWorksheet("Complementos de Pago");
  const NUM_COLS = 20;
  const DATA_START = 3;

  titleRow(ws, "COMPLEMENTOS DE PAGO — CFDI", NUM_COLS);

  const hRow = ws.getRow(2);
  hRow.height = 32;
  const headers: [number, string, HAlign][] = [
    [1, "Folio Fiscal (UUID)", "left"],
    [2, "Serie / Folio", "left"],
    [3, "RFC Emisor", "left"],
    [4, "Emisor", "left"],
    [5, "RFC Receptor", "left"],
    [6, "Receptor", "left"],
    [7, "Fecha de Emisión", "left"],
    [8, "Fecha de Pago", "left"],
    [9, "Forma de Pago", "left"],
    [10, "IdDocumento Relacionado", "left"],
    [11, "Saldo Anterior ($)", "center"],
    [12, "Importe Pagado ($)", "center"],
    [13, "Base 16% ($)", "center"],
    [14, "IVA 16% ($)", "center"],
    [15, "IVA Retenido ($)", "center"],
    [16, "IEPS / YEPS ($)", "center"],
    [17, "Exento ($)", "center"],
    [18, "Suma Cuadre ($)", "center"],
    [19, "Saldo Insoluto ($)", "center"],
    [20, "Archivo Fuente", "left"],
  ];
  headers.forEach(([c, t, a]) => headerCell(hRow, c, t, a));

  pagos.forEach((p, i) => {
    const row = ws.getRow(DATA_START + i);
    row.height = 18;
    const fill = i % 2 === 0 ? WHITE_FILL : ALT_FILL;
    const sumaCuadre = p.baseIVA16 + p.iva16 + (p.ieps ?? 0) - (p.ivaRetenido ?? 0) + p.exento;

    dataCell(row, 1, p.uuid, { fill });
    dataCell(row, 2, p.serieFolio, { fill });
    dataCell(row, 3, p.rfcEmisor ?? "", { fill });
    dataCell(row, 4, p.emisor, { fill });
    dataCell(row, 5, p.rfcReceptor ?? "", { fill });
    dataCell(row, 6, p.receptor, { fill });
    dataCell(row, 7, p.fechaEmision ? p.fechaEmision.slice(0, 10) : "", { fill });
    dataCell(row, 8, p.fechaPago ? p.fechaPago.slice(0, 10) : "", { fill });
    dataCell(row, 9, p.formaPago, { fill });
    dataCell(row, 10, p.docRelacionado, { fill });
    dataCell(row, 11, p.saldoAnterior, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 12, p.impPagado, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 13, p.baseIVA16, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 14, p.iva16, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 15, p.ivaRetenido ?? 0, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 16, p.ieps ?? 0, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 17, p.exento, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 18, sumaCuadre, { fill, fmt: MONEY_FMT, align: "right", bold: true });
    dataCell(row, 19, p.saldoInsoluto, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 20, p.fileName ?? "", { fill });
  });

  if (pagos.length > 0) {
    const totalRowNum = DATA_START + pagos.length;
    totalRow(ws, totalRowNum, 10, [11, 12, 13, 14, 15, 16, 17, 18, 19], DATA_START, NUM_COLS);
  }

  ws.views = [{ state: "frozen", ySplit: 2, topLeftCell: "A3" }];
  setColWidths(ws, [43, 15, 15, 37, 15, 25, 15, 15, 29, 43, 19, 19, 19, 19, 19, 19, 19, 19, 19, 25]);
}

function buildFacturasSheet(wb: ExcelJS.Workbook, facturas: Factura[]): void {
  const ws = wb.addWorksheet("Facturas");
  const NUM_COLS = 18;
  const DATA_START = 3;

  titleRow(ws, "FACTURAS DE INGRESO / EGRESO — CFDI", NUM_COLS);

  const hRow = ws.getRow(2);
  hRow.height = 32;
  const headers: [number, string, HAlign][] = [
    [1, "Folio Fiscal (UUID)", "left"],
    [2, "UUID Relacionado", "left"],
    [3, "Tipo", "center"],
    [4, "Serie / Folio", "left"],
    [5, "RFC Emisor", "left"],
    [6, "Emisor", "left"],
    [7, "RFC Receptor", "left"],
    [8, "Receptor", "left"],
    [9, "Fecha", "left"],
    [10, "Forma de Pago", "left"],
    [11, "Método Pago", "center"],
    [12, "Moneda", "center"],
    [13, "Subtotal ($)", "center"],
    [14, "IVA 16% ($)", "center"],
    [15, "IVA Retenido ($)", "center"],
    [16, "IEPS / YEPS ($)", "center"],
    [17, "Total ($)", "center"],
    [18, "Archivo Fuente", "left"],
  ];
  headers.forEach(([c, t, a]) => headerCell(hRow, c, t, a));

  facturas.forEach((f, i) => {
    const row = ws.getRow(DATA_START + i);
    row.height = 18;
    const fill = i % 2 === 0 ? WHITE_FILL : ALT_FILL;
    dataCell(row, 1, f.uuid, { fill });
    dataCell(row, 2, f.cfdiRelacionado, { fill });
    dataCell(row, 3, f.tipo, { fill, align: "center" });
    dataCell(row, 4, f.serieFolio, { fill });
    dataCell(row, 5, f.rfcEmisor ?? "", { fill });
    dataCell(row, 6, f.emisor, { fill });
    dataCell(row, 7, f.rfcReceptor ?? "", { fill });
    dataCell(row, 8, f.receptor, { fill });
    dataCell(row, 9, f.fecha ? f.fecha.slice(0, 10) : "", { fill });
    dataCell(row, 10, f.formaPago, { fill });
    dataCell(row, 11, f.metodoPago, { fill, align: "center" });
    dataCell(row, 12, f.moneda, { fill, align: "center" });
    dataCell(row, 13, f.subtotal, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 14, f.iva16, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 15, f.ivaRetenido ?? 0, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 16, f.ieps ?? 0, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 17, f.total, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 18, f.fileName ?? "", { fill });
  });

  if (facturas.length > 0) {
    const totalRowNum = DATA_START + facturas.length;
    totalRow(ws, totalRowNum, 12, [13, 14, 15, 16, 17], DATA_START, NUM_COLS);
  }

  ws.views = [{ state: "frozen", ySplit: 2, topLeftCell: "A3" }];
  setColWidths(ws, [43, 43, 9, 15, 15, 37, 15, 25, 13, 23, 11, 9, 17, 17, 17, 17, 17, 25]);
}

function buildConceptosSheet(wb: ExcelJS.Workbook, conceptos: Concepto[]): void {
  const ws = wb.addWorksheet("Conceptos");
  const NUM_COLS = 12;
  const DATA_START = 3;

  titleRow(ws, "CONCEPTOS — CFDI", NUM_COLS);

  const hRow = ws.getRow(2);
  hRow.height = 32;
  const headers: [number, string, HAlign][] = [
    [1, "Folio Fiscal (UUID)", "left"],
    [2, "Tipo", "center"],
    [3, "Emisor", "left"],
    [4, "Fecha", "left"],
    [5, "Descripción", "left"],
    [6, "Cantidad", "center"],
    [7, "Unidad", "left"],
    [8, "Valor Unitario ($)", "center"],
    [9, "Importe ($)", "center"],
    [10, "IVA 16% ($)", "center"],
    [11, "IEPS / YEPS ($)", "center"],
    [12, "Total ($)", "center"],
  ];
  headers.forEach(([c, t, a]) => headerCell(hRow, c, t, a));

  conceptos.forEach((c, i) => {
    const row = ws.getRow(DATA_START + i);
    row.height = 18;
    const fill = i % 2 === 0 ? WHITE_FILL : ALT_FILL;
    dataCell(row, 1, c.uuid, { fill });
    dataCell(row, 2, c.tipo, { fill, align: "center" });
    dataCell(row, 3, c.emisor, { fill });
    dataCell(row, 4, c.fecha ? c.fecha.slice(0, 10) : "", { fill });
    dataCell(row, 5, c.descripcion, { fill });
    dataCell(row, 6, c.cantidad, { fill, align: "right" });
    dataCell(row, 7, c.unidad, { fill });
    dataCell(row, 8, c.valorUnitario, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 9, c.importe, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 10, c.iva16, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 11, c.ieps ?? 0, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 12, c.importe + c.iva16 + (c.ieps ?? 0), { fill, fmt: MONEY_FMT, align: "right" });
  });

  if (conceptos.length > 0) {
    const totalRowNum = DATA_START + conceptos.length;
    totalRow(ws, totalRowNum, 8, [9, 10, 11, 12], DATA_START, NUM_COLS);
  }

  ws.views = [{ state: "frozen", ySplit: 2, topLeftCell: "A3" }];
  setColWidths(ws, [43, 9, 37, 13, 49, 11, 11, 19, 17, 17, 17, 17]);
}

function buildNotasCreditoSheets(wb: ExcelJS.Workbook, facturas: Factura[], conceptos: Concepto[]): void {
  const ncs = facturas.filter((f) => f.tipo === "E");
  if (ncs.length === 0) return;

  const ws = wb.addWorksheet("Notas de Crédito");
  const NUM_COLS = 10;
  const DATA_START = 3;

  titleRow(ws, "NOTAS DE CRÉDITO — CFDI EGRESO", NUM_COLS);

  const hRow = ws.getRow(2);
  hRow.height = 32;
  const headers: [number, string, HAlign][] = [
    [1, "Folio Fiscal (UUID)", "left"],
    [2, "UUID Relacionado", "left"],
    [3, "Serie / Folio", "left"],
    [4, "Emisor", "left"],
    [5, "Receptor", "left"],
    [6, "Fecha", "left"],
    [7, "Subtotal ($)", "center"],
    [8, "Descuento ($)", "center"],
    [9, "IVA 16% ($)", "center"],
    [10, "Total ($)", "center"],
  ];
  headers.forEach(([c, t, a]) => headerCell(hRow, c, t, a));

  ncs.forEach((nc, i) => {
    const row = ws.getRow(DATA_START + i);
    row.height = 18;
    const fill = i % 2 === 0 ? WHITE_FILL : ALT_FILL;

    dataCell(row, 1, nc.uuid, { fill });
    dataCell(row, 2, nc.cfdiRelacionado, { fill });
    dataCell(row, 3, nc.serieFolio, { fill });
    dataCell(row, 4, nc.emisor, { fill });
    dataCell(row, 5, nc.receptor, { fill });
    dataCell(row, 6, nc.fecha ? nc.fecha.slice(0, 10) : "", { fill });
    dataCell(row, 7, nc.subtotal, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 8, nc.descuento || 0, { fill, fmt: MONEY_FMT, align: "right" });

    // Formulas
    const subL = ws.getColumn(7).letter;
    const descL = ws.getColumn(8).letter;
    const ivaL = ws.getColumn(9).letter;
    const rIdx = DATA_START + i;
    
    // IVA formula: =ROUND((G - H) * 0.16, 2)
    const ivaCell = row.getCell(9);
    ivaCell.value = { formula: `ROUND((${subL}${rIdx}-${descL}${rIdx})*0.16, 2)` };
    ivaCell.fill = solidFill(fill);
    ivaCell.font = { name: "Arial", size: 10 };
    ivaCell.alignment = { horizontal: "right", vertical: "middle" };
    ivaCell.border = thinBorder();
    ivaCell.numFmt = MONEY_FMT;

    // Total formula: =ROUND((G - H) + I, 2)
    const totCell = row.getCell(10);
    totCell.value = { formula: `ROUND((${subL}${rIdx}-${descL}${rIdx})+${ivaL}${rIdx}, 2)` };
    totCell.fill = solidFill(fill);
    totCell.font = { name: "Arial", size: 10, bold: true };
    totCell.alignment = { horizontal: "right", vertical: "middle" };
    totCell.border = thinBorder();
    totCell.numFmt = MONEY_FMT;
  });

  const totalRowNum = DATA_START + ncs.length;
  totalRow(ws, totalRowNum, 6, [7, 8, 9, 10], DATA_START, NUM_COLS);

  ws.views = [{ state: "frozen", ySplit: 2, topLeftCell: "A3" }];
  setColWidths(ws, [43, 43, 14, 30, 24, 12, 14, 14, 14, 14]);

  // Conceptos NC Sheet
  const encUUIDs = new Set(ncs.map((nc) => nc.uuid));
  const conceptosNC = conceptos.filter((c) => encUUIDs.has(c.uuid));

  const wsC = wb.addWorksheet("Conceptos NC");
  const C_NUM_COLS = 11;
  
  titleRow(wsC, "CONCEPTOS NOTAS DE CRÉDITO", C_NUM_COLS);

  const hcRow = wsC.getRow(2);
  hcRow.height = 32;
  const cHeaders: [number, string, HAlign][] = [
    [1, "Folio Fiscal (UUID)", "left"],
    [2, "Serie / Folio", "left"],
    [3, "Emisor", "left"],
    [4, "Fecha", "left"],
    [5, "Clave Prod/Serv", "center"],
    [6, "Descripción", "left"],
    [7, "Cantidad", "center"],
    [8, "Unidad", "left"],
    [9, "Valor Unitario ($)", "center"],
    [10, "Descuento ($)", "center"],
    [11, "Importe ($)", "center"],
  ];
  cHeaders.forEach(([c, t, a]) => headerCell(hcRow, c, t, a));

  conceptosNC.forEach((c, i) => {
    const row = wsC.getRow(DATA_START + i);
    row.height = 18;
    const fill = i % 2 === 0 ? WHITE_FILL : ALT_FILL;

    dataCell(row, 1, c.uuid, { fill });
    dataCell(row, 2, c.serieFolio, { fill });
    dataCell(row, 3, c.emisor, { fill });
    dataCell(row, 4, c.fecha ? c.fecha.slice(0, 10) : "", { fill });
    dataCell(row, 5, c.claveProdServ, { fill, align: "center" });
    dataCell(row, 6, c.descripcion, { fill });
    dataCell(row, 7, c.cantidad, { fill, align: "right" });
    dataCell(row, 8, c.unidad, { fill });
    dataCell(row, 9, c.valorUnitario, { fill, fmt: MONEY_FMT, align: "right" });
    dataCell(row, 10, c.descuento || 0, { fill, fmt: MONEY_FMT, align: "right" });

    const cantL = wsC.getColumn(7).letter;
    const vuL = wsC.getColumn(9).letter;
    const descL = wsC.getColumn(10).letter;
    const rIdx = DATA_START + i;

    // Importe formula: =ROUND((G * I) - J, 2)
    const impCell = row.getCell(11);
    impCell.value = { formula: `ROUND((${cantL}${rIdx}*${vuL}${rIdx})-${descL}${rIdx}, 2)` };
    impCell.fill = solidFill(fill);
    impCell.font = { name: "Arial", size: 10, bold: true };
    impCell.alignment = { horizontal: "right", vertical: "middle" };
    impCell.border = thinBorder();
    impCell.numFmt = MONEY_FMT;
  });

  if (conceptosNC.length > 0) {
    const cTotalRowNum = DATA_START + conceptosNC.length;
    totalRow(wsC, cTotalRowNum, 10, [11], DATA_START, C_NUM_COLS);
  }

  wsC.views = [{ state: "frozen", ySplit: 2, topLeftCell: "A3" }];
  setColWidths(wsC, [43, 14, 30, 12, 16, 40, 11, 11, 16, 16, 16]);
}

// ─── Public API ──────────────────────────────────────────────────────────────

function getExcelColumnName(colNum: number): string {
  let temp = colNum;
  let columnName = "";
  while (temp > 0) {
    let rem = temp % 26;
    if (rem === 0) {
      columnName = "Z" + columnName;
      temp = Math.floor(temp / 26) - 1;
    } else {
      columnName = String.fromCharCode((rem - 1) + 65) + columnName;
      temp = Math.floor(temp / 26);
    }
  }
  return columnName;
}

function buildIepsBreakdownSheet(wb: ExcelJS.Workbook, pagos: PagoDoc[], facturas: Factura[]): void {
  const ws = wb.addWorksheet("Desglose de IEPS");
  
  // Aggregate data by rates
  const rateAggregation: { [rate: string]: { total: number; count: number; types: Set<string> } } = {};
  const documentsWithIeps: { 
    uuid: string; 
    type: string; 
    rfc: string; 
    emisor: string; 
    ieps: number; 
    iepsBreakdown: { [rate: string]: number };
    isFactura: boolean;
  }[] = [];

  // Parse facturas
  facturas.forEach((f) => {
    const docIeps = f.ieps ?? 0;
    if (docIeps > 0 || (f.iepsBreakdown && Object.keys(f.iepsBreakdown).length > 0)) {
      const breakdown = f.iepsBreakdown || {};
      
      // If no granular breakdown, treat as "Monto General"
      if (Object.keys(breakdown).length === 0) {
        const fallbackRate = "Monto General";
        if (!rateAggregation[fallbackRate]) {
          rateAggregation[fallbackRate] = { total: 0, count: 0, types: new Set() };
        }
        rateAggregation[fallbackRate].total += docIeps;
        rateAggregation[fallbackRate].count += 1;
        rateAggregation[fallbackRate].types.add(f.tipo === "E" ? "EGRESO" : "INGRESO");
      } else {
        Object.entries(breakdown).forEach(([rateStr, amount]) => {
          if (!rateAggregation[rateStr]) {
            rateAggregation[rateStr] = { total: 0, count: 0, types: new Set() };
          }
          rateAggregation[rateStr].total += amount;
          rateAggregation[rateStr].count += 1;
          rateAggregation[rateStr].types.add(f.tipo === "E" ? "EGRESO" : "INGRESO");
        });
      }

      documentsWithIeps.push({
        uuid: f.uuid,
        type: f.tipo === "E" ? "EGRESO (Nota de Crédito)" : "INGRESO",
        rfc: f.rfcEmisor || "",
        emisor: f.emisor,
        ieps: docIeps,
        iepsBreakdown: f.iepsBreakdown || {},
        isFactura: true
      });
    }
  });

  // Parse pagos
  pagos.forEach((p) => {
    const docIeps = p.ieps ?? 0;
    if (docIeps > 0 || (p.iepsBreakdown && Object.keys(p.iepsBreakdown).length > 0)) {
      const breakdown = p.iepsBreakdown || {};

      if (Object.keys(breakdown).length === 0) {
        const fallbackRate = "Monto General (Pago)";
        if (!rateAggregation[fallbackRate]) {
          rateAggregation[fallbackRate] = { total: 0, count: 0, types: new Set() };
        }
        rateAggregation[fallbackRate].total += docIeps;
        rateAggregation[fallbackRate].count += 1;
        rateAggregation[fallbackRate].types.add("COMPLEMENTO");
      } else {
        Object.entries(breakdown).forEach(([rateStr, amount]) => {
          if (!rateAggregation[rateStr]) {
            rateAggregation[rateStr] = { total: 0, count: 0, types: new Set() };
          }
          rateAggregation[rateStr].total += amount;
          rateAggregation[rateStr].count += 1;
          rateAggregation[rateStr].types.add("COMPLEMENTO");
        });
      }

      documentsWithIeps.push({
        uuid: p.uuid,
        type: `PAGO (Rel: ${p.docRelacionado ? p.docRelacionado.substring(0, 8) : "-"}...)`,
        rfc: p.rfcEmisor || "",
        emisor: p.emisor,
        ieps: docIeps,
        iepsBreakdown: p.iepsBreakdown || {},
        isFactura: false
      });
    }
  });

  // Unique rates present in dataset
  const uniqueRates = Object.keys(rateAggregation).sort();

  // Title Row (expanded to match the number of rate columns)
  const numIepsCols = 5 + Math.max(1, uniqueRates.length);
  titleRow(ws, "DESGLOSE DE IMPUESTO ESPECIAL SOBRE PRODUCCIÓN Y SERVICIOS (IEPS) POR TASAS Y CUOTAS", Math.max(6, numIepsCols));

  // Section 1: Consolidated Rates summary table
  const groupLabelCell = ws.getCell("A3");
  groupLabelCell.value = "1. RESUMEN CONSOLIDADO POR PORCENTAJE / CUOTA";
  groupLabelCell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FF" + HEADER_FILL } };

  const hRowRates = ws.getRow(4);
  hRowRates.height = 24;
  headerCell(hRowRates, 1, "Clasificación Tasa % / Cuota", "left");
  headerCell(hRowRates, 2, "Tipo Comprobantes Encontrados", "left");
  headerCell(hRowRates, 3, "Cantidad Documentos", "center");
  headerCell(hRowRates, 4, "Monto Total IEPS ($)", "right");

  const aggList = Object.entries(rateAggregation).sort((a, b) => b[1].total - a[1].total);
  aggList.forEach(([rate, data], idx) => {
    const rowNum = 5 + idx;
    const r = ws.getRow(rowNum);
    const fill = idx % 2 === 0 ? WHITE_FILL : ALT_FILL;
    dataCell(r, 1, rate, { fill, bold: true });
    dataCell(r, 2, Array.from(data.types).join(", "), { fill });
    dataCell(r, 3, data.count, { fill, align: "center" });
    dataCell(r, 4, data.total, { fill, fmt: MONEY_FMT, align: "right" });
  });

  const ratesTotalRowY = 5 + aggList.length;
  const tr = ws.getRow(ratesTotalRowY);
  dataCell(tr, 1, "TOTAL IEPS CONSOLIDADO", { fill: TOTAL_FILL, bold: true });
  dataCell(tr, 2, "", { fill: TOTAL_FILL });
  dataCell(tr, 3, { formula: `SUM(C5:C${ratesTotalRowY - 1})` }, { fill: TOTAL_FILL, align: "center", bold: true });
  dataCell(tr, 4, { formula: `SUM(D5:D${ratesTotalRowY - 1})` }, { fill: TOTAL_FILL, fmt: MONEY_FMT, align: "right", bold: true });

  // Section 2: Line-by-Line Document Audit Log
  const detailsTitleY = ratesTotalRowY + 3;
  const detailLabelCell = ws.getCell(`A${detailsTitleY}`);
  detailLabelCell.value = `2. BITÁCORA DETALLADA DE COMPROBANTES CON IEPS (${documentsWithIeps.length} DOCUMENTOS)`;
  detailLabelCell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FF" + HEADER_FILL } };

  const hRowDocs = ws.getRow(detailsTitleY + 1);
  hRowDocs.height = 24;
  headerCell(hRowDocs, 1, "Folio Fiscal (UUID)", "left");
  headerCell(hRowDocs, 2, "Tipo de Flujo", "center");
  headerCell(hRowDocs, 3, "RFC Emisor", "left");
  headerCell(hRowDocs, 4, "Nombre / Razón Social Emisor", "left");
  headerCell(hRowDocs, 5, "Monto IEPS Registrado ($)", "right");

  // Generate dynamic columns for rates
  uniqueRates.forEach((rateVal, k) => {
    headerCell(hRowDocs, 6 + k, `Tasa: ${rateVal}`, "right");
  });

  const docStartRow = detailsTitleY + 2;
  documentsWithIeps.forEach((doc, idx) => {
    const rowNum = docStartRow + idx;
    const r = ws.getRow(rowNum);
    const fill = idx % 2 === 0 ? WHITE_FILL : ALT_FILL;

    dataCell(r, 1, doc.uuid, { fill });
    dataCell(r, 2, doc.type, { fill, align: "center" });
    dataCell(r, 3, doc.rfc, { fill });
    dataCell(r, 4, doc.emisor, { fill });
    dataCell(r, 5, doc.ieps, { fill, fmt: MONEY_FMT, align: "right" });

    // Output dynamic columns value
    uniqueRates.forEach((rateStr, k) => {
      let val = 0;
      if (doc.iepsBreakdown && doc.iepsBreakdown[rateStr] !== undefined) {
        val = doc.iepsBreakdown[rateStr];
      } else if (rateStr === "Monto General" && Object.keys(doc.iepsBreakdown || {}).length === 0 && doc.isFactura) {
        val = doc.ieps;
      } else if (rateStr === "Monto General (Pago)" && Object.keys(doc.iepsBreakdown || {}).length === 0 && !doc.isFactura) {
        val = doc.ieps;
      }
      dataCell(r, 6 + k, val, { fill, fmt: MONEY_FMT, align: "right" });
    });
  });

  const docsTotalRowY = docStartRow + documentsWithIeps.length;
  if (documentsWithIeps.length > 0) {
    const dtr = ws.getRow(docsTotalRowY);
    dataCell(dtr, 1, "SUMA CUADRE DOCUMENTAL", { fill: TOTAL_FILL, bold: true });
    dataCell(dtr, 2, "", { fill: TOTAL_FILL });
    dataCell(dtr, 3, "", { fill: TOTAL_FILL });
    dataCell(dtr, 4, "", { fill: TOTAL_FILL });
    dataCell(dtr, 5, { formula: `SUM(E${docStartRow}:E${docsTotalRowY - 1})` }, { fill: TOTAL_FILL, fmt: MONEY_FMT, align: "right", bold: true });

    // Sum each dynamic column
    uniqueRates.forEach((rateStr, k) => {
      const colLetterStr = getExcelColumnName(6 + k);
      dataCell(dtr, 6 + k, { formula: `SUM(${colLetterStr}${docStartRow}:${colLetterStr}${docsTotalRowY - 1})` }, { fill: TOTAL_FILL, fmt: MONEY_FMT, align: "right", bold: true });
    });
  }

  const baseColWidths = [43, 26, 17, 37, 24];
  const dynamicColWidths = uniqueRates.map(() => 20);
  setColWidths(ws, [...baseColWidths, ...dynamicColWidths]);
}

export async function generarExcel(
  pagos: PagoDoc[],
  facturas: Factura[],
  conceptos: Concepto[]
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "CFDI Analizador";
  wb.created = new Date();

  buildPagosSheet(wb, pagos);
  buildFacturasSheet(wb, facturas);
  buildConceptosSheet(wb, conceptos);
  buildNotasCreditoSheets(wb, facturas, conceptos);
  buildIepsBreakdownSheet(wb, pagos, facturas);

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "cfdi_analisis.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

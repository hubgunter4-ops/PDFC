import JSZip from "jszip";
import { descripcionFormaPago } from "./formas-pago";

export interface IEPSBreakdown {
  [rate: string]: number;
}

export interface PagoDoc {
  uuid: string;
  serieFolio: string;
  emisor: string;
  receptor: string;
  fechaPago: string;
  formaPago: string;
  docRelacionado: string;
  saldoAnterior: number;
  impPagado: number;
  saldoInsoluto: number;
  baseIVA16: number;
  iva16: number;
  exento: number;
  rfcEmisor?: string;
  rfcReceptor?: string;
  nombreEmisor?: string;
  nombreReceptor?: string;
  fileName?: string;
  ivaRetenido?: number;
  ieps?: number;
  iepsBreakdown?: IEPSBreakdown;
  fechaEmision?: string;
}

export interface Factura {
  uuid: string;
  tipo: string;
  serieFolio: string;
  emisor: string;
  receptor: string;
  fecha: string;
  formaPago: string;
  metodoPago: string;
  moneda: string;
  subtotal: number;
  descuento: number;
  iva16: number;
  total: number;
  cfdiRelacionado: string;
  rfcEmisor?: string;
  rfcReceptor?: string;
  nombreEmisor?: string;
  nombreReceptor?: string;
  fileName?: string;
  ivaRetenido?: number;
  ieps?: number;
  iepsBreakdown?: IEPSBreakdown;
}

export interface Concepto {
  uuid: string;
  tipo: string;
  emisor: string;
  fecha: string;
  claveProdServ: string;
  descripcion: string;
  cantidad: number;
  unidad: string;
  valorUnitario: number;
  descuento: number;
  importe: number;
  iva16: number;
  serieFolio: string;
  ieps?: number;
  iepsBreakdown?: IEPSBreakdown;
}

export interface ParseResult {
  pagos: PagoDoc[];
  facturas: Factura[];
  conceptos: Concepto[];
}

export function formatIEPSRate(tasaOCuotaStr: string, tipoFactor?: string): string {
  if (!tasaOCuotaStr) return "S/T (Sin Tasa)";
  const val = parseFloat(tasaOCuotaStr);
  if (isNaN(val)) return tasaOCuotaStr;
  if (tipoFactor === "Cuota") {
    return `Cuota ($${parseFloat(val.toFixed(6))})`;
  }
  const pct = val * 100;
  return `${parseFloat(pct.toFixed(4))}%`;
}

function attr(el: Element | null, name: string): string {
  if (!el) return "";
  return el.getAttribute(name) ?? "";
}

function numAttr(el: Element | null, name: string): number {
  const v = attr(el, name);
  return v ? parseFloat(v) : 0;
}

function getNodes(parent: Document | Element, nodeName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS("*", nodeName));
}

function getNode(parent: Document | Element, nodeName: string): Element | null {
  return parent.getElementsByTagNameNS("*", nodeName)[0] || null;
}

export function parsearXML(text: string, fileName?: string): ParseResult {
  const parser = new DOMParser();
  const doc = parser.parseFromString(text, "application/xml");

  const comprobante = getNode(doc, "Comprobante") || doc.documentElement;
  if (!comprobante || !comprobante.localName?.includes("Comprobante")) {
    return { pagos: [], facturas: [], conceptos: [] };
  }

  const tipo = attr(comprobante, "TipoDeComprobante");
  const serie = attr(comprobante, "Serie");
  const folio = attr(comprobante, "Folio");
  const fecha = attr(comprobante, "Fecha");
  const formaPago = attr(comprobante, "FormaPago");
  const metodoPago = attr(comprobante, "MetodoPago");
  const moneda = attr(comprobante, "Moneda");
  const total = numAttr(comprobante, "Total");
  const serieFolio = [serie, folio].filter(Boolean).join("-");

  const tfd = getNode(doc, "TimbreFiscalDigital");
  const uuid = attr(tfd, "UUID");

  const emisorEl = getNode(doc, "Emisor");
  const receptorEl = getNode(doc, "Receptor");
  const rfcEmisor = attr(emisorEl, "Rfc");
  const rfcReceptor = attr(receptorEl, "Rfc");
  const nombreEmisor = attr(emisorEl, "Nombre") || rfcEmisor || "Sin Nombre";
  const nombreReceptor = attr(receptorEl, "Nombre") || rfcReceptor || "Sin Nombre";
  
  const emisor = nombreEmisor;
  const receptor = nombreReceptor;

  const pagos: PagoDoc[] = [];
  const facturas: Factura[] = [];
  const conceptos: Concepto[] = [];

  if (tipo === "P") {
    const pagoNodes = getNodes(doc, "Pago");
    pagoNodes.forEach((pago) => {
      const fechaPago = attr(pago, "FechaPago");
      const formaP = attr(pago, "FormaDePagoP");
      const doctos = getNodes(pago, "DoctoRelacionado");
      doctos.forEach((d) => {
        const impPagado = numAttr(d, "ImpPagado");
        let base16 = 0;
        let iva16 = 0;
        let ivaRetenido = 0;
        let ieps = 0;
        const iepsBreakdown: IEPSBreakdown = {};

        // CFDI 4.0: ImpuestosDR/TrasladosDR/TrasladoDR
        const traslados = getNodes(d, "TrasladoDR");
        traslados.forEach((t) => {
          const impuesto = attr(t, "ImpuestoDR") || attr(t, "Impuesto"); // 002 is IVA, 003 is IEPS
          const tasa = attr(t, "TasaOCuotaDR") || attr(t, "TasaOCuota");
          const tf = attr(t, "TipoFactorDR") || attr(t, "TipoFactor");
          const importe = numAttr(t, "ImporteDR") || numAttr(t, "Importe");
          if (impuesto === "002" && (tasa === "0.160000" || tasa === "0.16")) {
            base16 += numAttr(t, "BaseDR") || numAttr(t, "Base");
            iva16 += numAttr(t, "ImporteDR") || numAttr(t, "Importe");
          } else if (impuesto === "003" && importe > 0) {
            ieps += importe;
            const rateLabel = formatIEPSRate(tasa, tf);
            iepsBreakdown[rateLabel] = (iepsBreakdown[rateLabel] || 0) + importe;
          }
        });

        // CFDI 4.0: ImpuestosDR/RetencionesDR/RetencionDR
        const retenciones = getNodes(d, "RetencionDR");
        retenciones.forEach((r) => {
          const impuesto = attr(r, "ImpuestoDR") || attr(r, "Impuesto");
          if (impuesto === "002") {
            ivaRetenido += numAttr(r, "ImporteDR") || numAttr(r, "Importe");
          }
        });

        const exento = Math.max(0, impPagado - (base16 + iva16 + ieps - ivaRetenido));

        pagos.push({
          uuid,
          serieFolio,
          emisor,
          receptor,
          fechaPago,
          formaPago: descripcionFormaPago(formaP),
          docRelacionado: attr(d, "IdDocumento"),
          saldoAnterior: numAttr(d, "ImpSaldoAnt"),
          impPagado,
          saldoInsoluto: numAttr(d, "ImpSaldoInsoluto"),
          baseIVA16: base16,
          iva16: iva16,
          exento: exento,
          rfcEmisor,
          rfcReceptor,
          nombreEmisor,
          nombreReceptor,
          fileName,
          ivaRetenido,
          ieps,
          iepsBreakdown,
          fechaEmision: fecha,
        });
      });
    });
  } else if (tipo === "I" || tipo === "E") {
    const subtotal = numAttr(comprobante, "SubTotal");
    const descuento = numAttr(comprobante, "Descuento");
    let iva16 = 0;
    let ieps = 0;
    const iepsBreakdown: IEPSBreakdown = {};
    
    // Extract IVA 16% and IEPS from global taxes
    const traslados = getNodes(doc, "Traslado");
    traslados.forEach(t => {
      // Check if it's a child of Impuestos (global) or from Conceptos (we want global here for the row)
      const isGlobal = t.parentElement?.parentElement?.localName?.includes("Comprobante") || 
                       t.parentElement?.localName?.includes("Impuestos");
      
      const impuesto = attr(t, "Impuesto");
      const tasa = attr(t, "TasaOCuota");
      const tf = attr(t, "TipoFactor");
      const importe = numAttr(t, "Importe");
      if (isGlobal) {
        if (impuesto === "002" && (tasa === "0.160000" || tasa === "0.16")) {
          iva16 += importe;
        } else if (impuesto === "003" && importe > 0) {
          ieps += importe;
          const rateLabel = formatIEPSRate(tasa, tf);
          iepsBreakdown[rateLabel] = (iepsBreakdown[rateLabel] || 0) + importe;
        }
      }
    });

    // Fallback: If global IVA 16% is 0, sum of concept taxes
    if (iva16 === 0) {
      const conceptsList = getNodes(doc, "Concepto");
      conceptsList.forEach((c) => {
        const cTraslados = getNodes(c, "Traslado");
        cTraslados.forEach(ct => {
          const impuesto = attr(ct, "Impuesto");
          const tasa = attr(ct, "TasaOCuota");
          if (impuesto === "002" && (tasa === "0.160000" || tasa === "0.16")) {
            iva16 += numAttr(ct, "Importe");
          }
        });
      });
    }

    const conceptIepsBreakdown: IEPSBreakdown = {};
    const conceptsList = getNodes(doc, "Concepto");
    conceptsList.forEach((c) => {
      const cTraslados = getNodes(c, "Traslado");
      cTraslados.forEach(ct => {
        const impuesto = attr(ct, "Impuesto");
        const tasa = attr(ct, "TasaOCuota");
        const tf = attr(ct, "TipoFactor");
        const importe = numAttr(ct, "Importe");
        if (impuesto === "003" && importe > 0) {
          const rateLabel = formatIEPSRate(tasa, tf);
          conceptIepsBreakdown[rateLabel] = (conceptIepsBreakdown[rateLabel] || 0) + importe;
        }
      });
    });

    if (ieps === 0 || Object.keys(iepsBreakdown).length === 0) {
      Object.entries(conceptIepsBreakdown).forEach(([r, amt]) => {
        iepsBreakdown[r] = (iepsBreakdown[r] || 0) + amt;
        ieps += amt;
      });
    }

    let ivaRetenido = 0;
    const retenciones = getNodes(doc, "Retencion");
    retenciones.forEach(r => {
      const isGlobal = r.parentElement?.parentElement?.localName?.includes("Comprobante") || 
                       r.parentElement?.localName?.includes("Impuestos");
      const impuesto = attr(r, "Impuesto");
      if (isGlobal && impuesto === "002") {
        ivaRetenido += numAttr(r, "Importe");
      }
    });

    if (ivaRetenido === 0) {
      const conceptsList = getNodes(doc, "Concepto");
      conceptsList.forEach((c) => {
        const cRetenciones = getNodes(c, "Retencion");
        cRetenciones.forEach(cr => {
          const impuesto = attr(cr, "Impuesto");
          if (impuesto === "002") {
            ivaRetenido += numAttr(cr, "Importe");
          }
        });
      });
    }

    const cfdiRelNodes = getNodes(doc, "CfdiRelacionado");
    const cfdiRelacionado = cfdiRelNodes.map(r => attr(r, "UUID")).filter(Boolean).join(", ");

    facturas.push({
      uuid,
      tipo,
      serieFolio,
      emisor,
      receptor,
      fecha,
      formaPago: descripcionFormaPago(formaPago),
      metodoPago,
      moneda,
      subtotal,
      descuento,
      iva16,
      total,
      cfdiRelacionado,
      rfcEmisor,
      rfcReceptor,
      nombreEmisor,
      nombreReceptor,
      fileName,
      ivaRetenido,
      ieps,
      iepsBreakdown,
    });

    const conceptoNodes = getNodes(doc, "Concepto");
    conceptoNodes.forEach((c) => {
      let cIva16 = 0;
      let cIeps = 0;
      const cIepsBreakdown: IEPSBreakdown = {};
      const cTraslados = getNodes(c, "Traslado");
      cTraslados.forEach(ct => {
        const impuesto = attr(ct, "Impuesto");
        const tasa = attr(ct, "TasaOCuota");
        const tf = attr(ct, "TipoFactor");
        const importe = numAttr(ct, "Importe");
        if (impuesto === "002" && (tasa === "0.160000" || tasa === "0.16")) {
          cIva16 += numAttr(ct, "Importe");
        } else if (impuesto === "003" && importe > 0) {
          cIeps += importe;
          const rateLabel = formatIEPSRate(tasa, tf);
          cIepsBreakdown[rateLabel] = (cIepsBreakdown[rateLabel] || 0) + importe;
        }
      });

      conceptos.push({
        uuid,
        tipo,
        emisor,
        fecha,
        claveProdServ: attr(c, "ClaveProdServ"),
        descripcion: attr(c, "Descripcion"),
        cantidad: numAttr(c, "Cantidad"),
        unidad: attr(c, "ClaveUnidad"),
        valorUnitario: numAttr(c, "ValorUnitario"),
        descuento: numAttr(c, "Descuento"),
        importe: numAttr(c, "Importe"),
        iva16: cIva16,
        serieFolio,
        ieps: cIeps,
        iepsBreakdown: cIepsBreakdown,
      });
    });
  }

  return { pagos, facturas, conceptos };
}

export function readFileAsText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target?.result as string);
    reader.onerror = reject;
    reader.readAsText(file, "utf-8");
  });
}

export async function parseFiles(files: File[]): Promise<ParseResult> {
  const result: ParseResult = { pagos: [], facturas: [], conceptos: [] };
  for (const file of files) {
    if (file.name.toLowerCase().endsWith(".zip")) {
      const zip = new JSZip();
      const contents = await zip.loadAsync(file);
      for (const [path, zipEntry] of Object.entries(contents.files)) {
        if (!zipEntry.dir && path.toLowerCase().endsWith(".xml")) {
          const text = await zipEntry.async("text");
          const nameOnly = path.split("/").pop() || path;
          const r = parsearXML(text, nameOnly);
          result.pagos.push(...r.pagos);
          result.facturas.push(...r.facturas);
          result.conceptos.push(...r.conceptos);
        }
      }
    } else if (file.name.toLowerCase().endsWith(".xml")) {
      const text = await readFileAsText(file);
      const r = parsearXML(text, file.name);
      result.pagos.push(...r.pagos);
      result.facturas.push(...r.facturas);
      result.conceptos.push(...r.conceptos);
    }
  }
  return result;
}

export function enriquecerPagosConFacturas(pagos: PagoDoc[], facturas: Factura[]): PagoDoc[] {
  // Map of facturas by lowercase UUID for O(1) matching
  const facturasMap = new Map<string, Factura>();
  facturas.forEach(f => {
    facturasMap.set(f.uuid.toLowerCase().trim(), f);
  });

  return pagos.map(pago => {
    const docRelId = pago.docRelacionado.toLowerCase().trim();
    const matched = facturasMap.get(docRelId);

    if (matched) {
      const totalFact = matched.total;
      
      // Calculate proportion
      const proportion = totalFact > 0 ? (pago.impPagado / totalFact) : 0;
      
      // Base IVA 16% proportion: Subtotal - Descuento scaled by the payment proportion
      let baseCalculada = (matched.subtotal - (matched.descuento || 0)) * proportion;
      let ivaCalculado = matched.iva16 * proportion;
      let ivaRetenidoCalculado = (matched.ivaRetenido || 0) * proportion;
      let iepsCalculado = (matched.ieps || 0) * proportion;

      // Ensure precision matching (ROUND to 2 decimals or match standard MX formulas)
      baseCalculada = Math.round(baseCalculada * 100) / 100;
      ivaCalculado = Math.round(ivaCalculado * 100) / 100;
      ivaRetenidoCalculado = Math.round(ivaRetenidoCalculado * 100) / 100;
      iepsCalculado = Math.round(iepsCalculado * 100) / 100;
      const exentoCalculado = Math.max(0, Math.round((pago.impPagado - (baseCalculada + ivaCalculado + iepsCalculado - ivaRetenidoCalculado)) * 100) / 100);

      const iepsBreakdownCalculado: IEPSBreakdown = {};
      if (matched.iepsBreakdown) {
        Object.entries(matched.iepsBreakdown).forEach(([rate, amt]) => {
          const scaled = Math.round(amt * proportion * 100) / 100;
          if (scaled > 0) {
            iepsBreakdownCalculado[rate] = scaled;
          }
        });
      }

      // If pago already has iepsBreakdown but it is empty/unset, use calculation
      const finalIepsBreakdown = (pago.iepsBreakdown && Object.keys(pago.iepsBreakdown).length > 0)
        ? pago.iepsBreakdown
        : iepsBreakdownCalculado;

      return {
        ...pago,
        baseIVA16: baseCalculada,
        iva16: ivaCalculado,
        exento: exentoCalculado,
        ivaRetenido: ivaRetenidoCalculado,
        ieps: iepsCalculado,
        iepsBreakdown: finalIepsBreakdown,
      };
    } else {
      // If we don't have a matching loaded invoice, but the payment doesn't have IVA from CFDI 4.0 (e.g. CFDI 3.3)
      // we intelligently assume 16% is standard, helping calculate a realistic estimate
      if (pago.baseIVA16 === 0 && pago.iva16 === 0 && pago.impPagado > 0) {
        const baseCalculada = Math.round((pago.impPagado / 1.16) * 100) / 100;
        const ivaCalculado = Math.round((pago.impPagado - baseCalculada) * 100) / 100;
        return {
          ...pago,
          baseIVA16: baseCalculada,
          iva16: ivaCalculado,
          exento: 0,
        };
      }
    }
    return pago;
  });
}

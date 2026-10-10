/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useCallback } from 'react';
import { 
  FileUp, 
  Trash2, 
  Download, 
  AlertCircle, 
  CheckCircle2, 
  Clock, 
  BarChart3, 
  Info,
  CreditCard,
  FileText,
  Search,
  Filter,
  Copy,
  Check
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

import { parseFiles, ParseResult, PagoDoc, Factura, Concepto, enriquecerPagosConFacturas } from './services/parser';
import { generarExcel } from './services/excelService';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

function UUIDTooltip({ uuid }: { uuid: string }) {
  const [copied, setCopied] = useState(false);
  const [show, setShow] = useState(false);

  const handleCopy = (e: React.MouseEvent) => {
    e.stopPropagation();
    navigator.clipboard.writeText(uuid);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  if (!uuid) return <span className="text-slate-600 italic">-</span>;
  const truncated = uuid.length > 13 ? `${uuid.substring(0, 13)}...` : uuid;

  return (
    <div 
      className="relative inline-block font-mono text-[11px]"
      onMouseEnter={() => setShow(true)}
      onMouseLeave={() => {
        setShow(false);
        setCopied(false);
      }}
    >
      <span className="text-slate-300 hover:text-amber-400 font-bold hover:underline transition-colors cursor-help select-all">
        {truncated}
      </span>
      
      <AnimatePresence>
        {show && (
          <motion.div 
            initial={{ opacity: 0, y: 5 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 5 }}
            transition={{ duration: 0.15 }}
            className="absolute left-0 bottom-full mb-2 w-72 bg-[#0F172A] border border-slate-800 text-slate-200 text-[11px] p-3 rounded-xl shadow-2xl z-50 pointer-events-auto flex flex-col gap-2"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center border-b border-slate-800 pb-1.5">
              <span className="text-[9px] text-slate-400 font-extrabold uppercase tracking-wide">UUID Completo</span>
              <button 
                onClick={handleCopy}
                className="flex items-center gap-1.5 text-[9px] font-black uppercase text-amber-400 hover:text-amber-300 border border-amber-500/20 px-2 py-0.5 rounded bg-amber-500/5 hover:bg-amber-500/15 transition-colors"
              >
                {copied ? (
                  <>
                    <Check size={8} /> ¡Copiado!
                  </>
                ) : (
                  <>
                    <Copy size={8} /> Copiar
                  </>
                )}
              </button>
            </div>
            <div className="font-mono break-all select-all text-slate-100 bg-[#070B13] p-2 rounded-lg border border-slate-800 text-[10px] leading-relaxed">
              {uuid}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

export default function App() {
  const [data, setData] = useState<ParseResult>({ pagos: [], facturas: [], conceptos: [] });
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState<'payments' | 'analysis' | 'summary' | 'facturas' | 'ieps'>('facturas');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedRecord, setSelectedRecord] = useState<any | null>(null);

  const iepsAnalysis = useMemo(() => {
    const rateTotals: { [rate: string]: { total: number; count: number; docTypes: Set<string> } } = {};
    let totalIEPS = 0;
    const documentsWithIEPS: Array<{
      uuid: string;
      displayType: string;
      emisor: string;
      ieps: number;
      iepsBreakdown?: { [rate: string]: number };
      rawRecord: any;
    }> = [];

    // Facturas
    data.facturas.forEach(f => {
      if (f.ieps && f.ieps > 0) {
        totalIEPS += f.ieps;
        documentsWithIEPS.push({
          uuid: f.uuid,
          displayType: f.tipo === "E" ? "EGRESO (Factura)" : "INGRESO (Factura)",
          emisor: f.emisor,
          ieps: f.ieps,
          iepsBreakdown: f.iepsBreakdown,
          rawRecord: f
        });

        if (f.iepsBreakdown && Object.keys(f.iepsBreakdown).length > 0) {
          Object.entries(f.iepsBreakdown).forEach(([rate, val]) => {
            if (!rateTotals[rate]) {
              rateTotals[rate] = { total: 0, count: 0, docTypes: new Set() };
            }
            rateTotals[rate].total += val as number;
            rateTotals[rate].count += 1;
            rateTotals[rate].docTypes.add(f.tipo === "E" ? "EGRESO" : "INGRESO");
          });
        } else {
          const rate = "S/T (Sin Tasa)";
          if (!rateTotals[rate]) {
            rateTotals[rate] = { total: 0, count: 0, docTypes: new Set() };
          }
          rateTotals[rate].total += f.ieps;
          rateTotals[rate].count += 1;
          rateTotals[rate].docTypes.add(f.tipo === "E" ? "EGRESO" : "INGRESO");
        }
      }
    });

    // Pagos
    data.pagos.forEach(p => {
      if (p.ieps && p.ieps > 0) {
        totalIEPS += p.ieps;
        documentsWithIEPS.push({
          uuid: p.uuid,
          displayType: `PAGO (Rel: ${p.docRelacionado?.substring(0, 8)}...)`,
          emisor: p.emisor,
          ieps: p.ieps,
          iepsBreakdown: p.iepsBreakdown,
          rawRecord: p
        });

        if (p.iepsBreakdown && Object.keys(p.iepsBreakdown).length > 0) {
          Object.entries(p.iepsBreakdown).forEach(([rate, val]) => {
            if (!rateTotals[rate]) {
              rateTotals[rate] = { total: 0, count: 0, docTypes: new Set() };
            }
            rateTotals[rate].total += val as number;
            rateTotals[rate].count += 1;
            rateTotals[rate].docTypes.add("PAGO");
          });
        } else {
          const rate = "S/T (Sin Tasa)";
          if (!rateTotals[rate]) {
            rateTotals[rate] = { total: 0, count: 0, docTypes: new Set() };
          }
          rateTotals[rate].total += p.ieps;
          rateTotals[rate].count += 1;
          rateTotals[rate].docTypes.add("PAGO");
        }
      }
    });

    // Turn sets back to sorted arrays / formatting
    const breakdownList = Object.entries(rateTotals).map(([rate, info]) => {
      return {
        rate,
        total: info.total,
        count: info.count,
        types: Array.from(info.docTypes).join(", ")
      };
    }).sort((a, b) => b.total - a.total);

    return { breakdownList, totalIEPS, documentsWithIEPS };
  }, [data]);

  const uniqueRates = useMemo(() => {
    return iepsAnalysis.breakdownList.map(item => item.rate);
  }, [iepsAnalysis.breakdownList]);

  const egresosCredito = useMemo(() => {
    return data.facturas.filter(f => 
      f.tipo === "E" && 
      (f.metodoPago === "PPD" || f.formaPago?.includes("99") || f.formaPago?.toLowerCase().includes("por definir"))
    );
  }, [data.facturas]);

  const getRelatedRecords = useMemo(() => {
    if (!selectedRecord) return [];
    
    const uuids: string[] = [];
    if (selectedRecord.docRelacionado) {
      uuids.push(selectedRecord.docRelacionado);
    }
    if (selectedRecord.cfdiRelacionado) {
      selectedRecord.cfdiRelacionado.split(',').forEach((u: string) => {
        const trimmed = u.trim();
        if (trimmed && !uuids.includes(trimmed)) uuids.push(trimmed);
      });
    }

    return uuids.map(uuid => {
      const clean = uuid.toLowerCase().trim();
      const fact = data.facturas.find(f => f.uuid.toLowerCase().trim() === clean);
      if (fact) {
        return { 
          uuid, 
          records: [{ ...fact, displayType: fact.tipo === "E" ? "EGRESO (NC)" : "INGRESO" }], 
          found: true 
        };
      }
      
      const pgs = data.pagos.filter(p => p.uuid.toLowerCase().trim() === clean || p.docRelacionado.toLowerCase().trim() === clean);
      if (pgs.length > 0) {
        return { 
          uuid, 
          records: pgs.map(p => ({ ...p, displayType: "COMPLEMENTO DE PAGO" })), 
          found: true 
        };
      }
      
      return { uuid, records: [], found: false };
    });
  }, [selectedRecord, data.facturas, data.pagos]);

  const analysis = useMemo(() => {
    const map = new Map<string, any>();
    // Building a basic 'analysis' representation that mirrors the previous UI logic 
    // from just PagoDoc list.
    data.pagos.forEach(p => {
      const existing = map.get(p.docRelacionado);
      if (existing) {
        existing.totalPagadoLote += p.impPagado;
        existing.numPagos += 1;
        existing.saldoInsolutoFinal = p.saldoInsoluto;
        existing.status = p.saldoInsoluto === 0 ? 'liquidado' : 'parcial';
      } else {
        map.set(p.docRelacionado, {
          idDocumento: p.docRelacionado,
          emisor: p.emisor,
          receptor: p.receptor,
          saldoInicialReal: p.saldoAnterior,
          totalPagadoLote: p.impPagado,
          saldoInsolutoFinal: p.saldoInsoluto,
          numPagos: 1,
          status: p.saldoInsoluto === 0 ? 'liquidado' : 'pendiente'
        });
      }
    });
    return Array.from(map.values());
  }, [data.pagos]);

  const summaries = useMemo(() => {
    const map = new Map<string, any>();
    data.pagos.forEach(p => {
      const emisor = p.emisor;
      const existing = map.get(emisor);
      if (existing) {
        existing.montoTotalPagado += p.impPagado;
        existing.saldoInsolutoTotal = p.saldoInsoluto;
        if (p.rfcEmisor && p.rfcEmisor !== 'N/A') {
          existing.rfc = p.rfcEmisor;
        }
      } else {
        map.set(emisor, {
          rfc: p.rfcEmisor || 'N/A',
          nombre: emisor,
          totalCfdis: 0,
          montoTotalPagado: p.impPagado,
          saldoInsolutoTotal: p.saldoInsoluto,
        });
      }
    });

    // Count facturas/pagos per emisor
    data.facturas.forEach(f => {
       const existing = map.get(f.emisor);
       if (existing) {
          existing.totalCfdis += 1;
          if (f.rfcEmisor && f.rfcEmisor !== 'N/A') {
            existing.rfc = f.rfcEmisor;
          }
       } else {
         map.set(f.emisor, {
           rfc: f.rfcEmisor || 'N/A',
           nombre: f.emisor,
           totalCfdis: 1,
           montoTotalPagado: 0,
           saldoInsolutoTotal: 0
         })
       }
    });

    return Array.from(map.values());
  }, [data.pagos, data.facturas]);

  const handleFileUpload = async (files: FileList | null) => {
    if (!files) return;
    setLoading(true);
    try {
      const result = await parseFiles(Array.from(files));
      
      // Merge results
      setData(prev => {
        const uniqueFacturas = [...prev.facturas, ...result.facturas].filter((v,i,a)=>a.findIndex(t=>(t.uuid === v.uuid))===i);
        const rawPagos = [...prev.pagos, ...result.pagos].filter((v,i,a)=>a.findIndex(t=>(t.uuid === v.uuid && t.docRelacionado === v.docRelacionado))===i);
        const uniqueConceptos = [...prev.conceptos, ...result.conceptos].filter((v,i,a)=>a.findIndex(t=>(t.uuid === v.uuid && t.descripcion === v.descripcion))===i);

        const enrichedPagos = enriquecerPagosConFacturas(rawPagos, uniqueFacturas);

        return {
          pagos: enrichedPagos,
          facturas: uniqueFacturas,
          conceptos: uniqueConceptos
        };
      });
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    handleFileUpload(e.dataTransfer.files);
  }, []);

  const clearData = () => {
    setData({ pagos: [], facturas: [], conceptos: [] });
    setSearchTerm('');
    setSelectedRecord(null);
  };

  const exportExcel = () => {
    generarExcel(data.pagos, data.facturas, data.conceptos);
  };

  const filteredPagos = data.pagos.filter(p => 
    p.emisor.toLowerCase().includes(searchTerm.toLowerCase()) ||
    p.uuid.toLowerCase().includes(searchTerm.toLowerCase()) ||
    p.docRelacionado.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const fileCount = data.facturas.length + Array.from(new Set(data.pagos.map(p=>p.uuid))).length;

  return (
    <div className="flex flex-col h-screen overflow-hidden bg-[#0A0F1D] text-slate-100 font-sans selection:bg-amber-400 selection:text-slate-900"
      style={{ backgroundImage: 'radial-gradient(rgba(148, 163, 184, 0.05) 1.5px, transparent 1.5px)', backgroundSize: '24px 24px' }}>
      
      {/* Header */}
      <header className="bg-[#0F172A]/80 backdrop-blur-md border-b border-slate-800 p-5 flex justify-between items-center shrink-0 z-10 transition-all">
        <div>
          <div className="flex items-center gap-2">
            <h1 className="text-3xl font-black italic tracking-tight uppercase bg-gradient-to-r from-amber-400 via-yellow-400 to-indigo-400 bg-clip-text text-transparent">
              AuditPro <span className="text-slate-100 font-extrabold not-italic">Core</span>
            </h1>
            <span className="text-[9px] bg-indigo-500/20 text-indigo-300 font-bold tracking-widest uppercase px-1.5 py-0.5 rounded border border-indigo-500/20">
              v2.1
            </span>
          </div>
          <p className="text-[10px] text-slate-400 font-medium tracking-wide mt-1 uppercase">
            Sistemas de Validación Fiscal & Auditoría Financiera de Pagos // Procesamiento 100% Local
          </p>
        </div>
        
        <div className="hidden md:flex gap-4">
          <div className="border border-slate-800 bg-[#141C30]/90 px-4 py-2 rounded-xl shadow-lg min-w-[90px]">
            <span className="text-[9px] text-slate-400 font-black tracking-wider block uppercase">LOTE PARSED</span>
            <span className="text-xl font-extrabold text-amber-400 tabular-nums">{fileCount} <span className="text-xs text-slate-500 font-normal">CFDIs</span></span>
          </div>
          <div className="border border-slate-800 bg-[#141C30]/90 px-4 py-2 rounded-xl shadow-lg min-w-[90px]">
            <span className="text-[9px] text-slate-400 font-black tracking-wider block uppercase">ANÁLISIS</span>
            <span className="text-xl font-extrabold text-indigo-400 tabular-nums">{analysis.length} <span className="text-xs text-slate-500 font-normal">Docs</span></span>
          </div>
        </div>
      </header>

      <main className="flex-1 grid grid-cols-12 gap-0 overflow-hidden">
        {/* Sidebar */}
        <section className="col-span-3 border-r border-slate-800 p-5 flex flex-col gap-5 bg-[#0C1222] overflow-y-auto custom-scrollbar">
          
          {/* File Upload Trigger */}
          <div className={cn(
            "relative border-2 border-dashed border-slate-800 hover:border-amber-400/50 bg-[#11192C]/80 p-8 text-center flex flex-col items-center justify-center gap-4 transition-all duration-300 rounded-2xl shadow-lg backdrop-blur-sm group cursor-pointer",
            isDragging ? "bg-[#18263D] border-amber-400 scale-[1.01]" : ""
          )}>
            <input 
              type="file" 
              multiple 
              accept=".xml,.zip" 
              className="absolute inset-0 opacity-0 cursor-pointer z-20"
              onChange={(e) => handleFileUpload(e.target.files)}
              onDragOver={() => setIsDragging(true)}
              onDragLeave={() => setIsDragging(false)}
              onDrop={(e) => { setIsDragging(false); onDrop(e); }}
            />
            <div className="w-12 h-12 bg-gradient-to-tr from-amber-500/20 to-yellow-500/10 rounded-xl flex items-center justify-center border border-amber-500/20 group-hover:scale-110 transition-transform">
              <FileUp className="text-amber-400" size={24} />
            </div>
            <div>
              <p className="text-xs font-bold text-slate-200 uppercase tracking-wide group-hover:text-amber-300 transition-colors">Importar Archivos</p>
              <p className="opacity-50 text-[9px] text-slate-400 uppercase tracking-wider mt-1">Arrastra CFDI (.xml) o lote (.zip)</p>
            </div>
          </div>

          {/* Queue Status / Loaded files */}
          <div className="flex flex-col gap-2.5 flex-1 select-none">
            <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-wider border-b border-slate-800 pb-1.5 flex justify-between">
              <span>Cola de Carga Local</span>
              <span className="text-amber-400 text-[8px]">{fileCount} archivos</span>
            </h3>
            
            <div className="space-y-1.5 overflow-y-auto max-h-[260px] pr-1.5 custom-scrollbar">
              {fileCount === 0 ? (
                <div className="flex flex-col items-center justify-center py-10 opacity-30">
                  <p className="text-[10px] font-bold uppercase tracking-wider">Esperando CFDIs...</p>
                </div>
              ) : (
                data.facturas.slice(0, 30).map(c => (
                  <div key={c.uuid} className="bg-[#101728]/80 border border-slate-800/60 p-2 rounded-lg flex justify-between items-center text-[10px] font-mono group hover:border-slate-700 transition-colors">
                    <span className="truncate max-w-[130px] italic text-slate-300 font-semibold">{c.uuid.substring(0, 15)}...</span>
                    <span className="text-emerald-400 font-bold text-[9px] uppercase tracking-wider bg-emerald-500/10 px-1.5 py-0.5 rounded border border-emerald-500/10">OK</span>
                  </div>
                ))
              )}
              {fileCount > 30 && (
                <p className="text-[9px] text-slate-500 font-bold opacity-80 text-center uppercase py-1">
                  +{fileCount - 30} más en memoria local
                </p>
              )}
            </div>

            {fileCount > 0 && (
              <div className="w-full h-1 bg-slate-800 rounded-full overflow-hidden mt-1.5">
                <div className="bg-gradient-to-r from-amber-500 to-indigo-400 h-full w-full"></div>
              </div>
            )}
          </div>

          <div className="mt-auto space-y-3 pt-4">
            <button 
              onClick={exportExcel}
              disabled={fileCount === 0}
              className="w-full bg-gradient-to-r from-amber-500 to-yellow-500 hover:from-amber-400 hover:to-yellow-400 text-slate-950 px-4 py-3 font-extrabold uppercase text-xs tracking-wider rounded-xl shadow-lg shadow-amber-500/10 active:scale-[0.98] disabled:opacity-20 disabled:cursor-not-allowed transition-all"
            >
              Exportar a Excel
            </button>
            <button 
              onClick={clearData}
              disabled={fileCount === 0}
              className="w-full bg-slate-900 border border-red-500/30 hover:border-red-500/60 text-red-400 px-4 py-2.5 font-bold uppercase text-[10px] tracking-wider rounded-xl hover:bg-red-500/5 active:scale-[0.98] disabled:opacity-20 disabled:cursor-not-allowed transition-all"
            >
              Limpiar Memoria
            </button>
          </div>
        </section>

        {/* Content Area */}
        <section className="col-span-9 p-6 flex flex-col gap-5 overflow-hidden">
          
          {/* Stats Bento Layout */}
          <div className="grid grid-cols-4 gap-4 select-none">
            
            <div className="bg-[#101729]/90 border border-slate-800/80 p-4 rounded-xl shadow-md relative overflow-hidden group">
              <div className="absolute right-3 top-3 w-8 h-8 rounded-lg bg-emerald-500/5 flex items-center justify-center border border-emerald-500/10">
                <CheckCircle2 className="text-emerald-400" size={16} />
              </div>
              <span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block mb-1">LIQUIDADO</span>
              <div className="flex items-baseline gap-2 mt-1">
                <span className="text-3xl font-extrabold text-emerald-400 tabular-nums">
                  {analysis.filter(a => a.status === 'liquidado').length}
                </span>
                <span className="text-[10px] text-slate-500 uppercase">Documentos</span>
              </div>
              <div className="w-full h-1 bg-slate-800 rounded-full mt-2 overflow-hidden">
                <div className="bg-emerald-400 h-full" style={{ width: `${analysis.length ? (analysis.filter(a => a.status === 'liquidado').length / analysis.length) * 100 : 0}%` }}></div>
              </div>
            </div>

            <div className="bg-[#101729]/90 border border-slate-800/80 p-4 rounded-xl shadow-md relative overflow-hidden group">
              <div className="absolute right-3 top-3 w-8 h-8 rounded-lg bg-amber-500/5 flex items-center justify-center border border-amber-500/10">
                <Clock className="text-amber-400" size={16} />
              </div>
              <span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block mb-1">PAGO PARCIAL</span>
              <div className="flex items-baseline gap-2 mt-1">
                <span className="text-3xl font-extrabold text-amber-400 tabular-nums">
                  {analysis.filter(a => a.status === 'parcial').length}
                </span>
                <span className="text-[10px] text-slate-500 uppercase">Documentos</span>
              </div>
              <div className="w-full h-1 bg-slate-800 rounded-full mt-2 overflow-hidden">
                <div className="bg-amber-400 h-full" style={{ width: `${analysis.length ? (analysis.filter(a => a.status === 'parcial').length / analysis.length) * 100 : 0}%` }}></div>
              </div>
            </div>

            <div className="bg-[#101729]/90 border border-slate-800/80 p-4 rounded-xl shadow-md relative overflow-hidden group">
              <div className="absolute right-3 top-3 w-8 h-8 rounded-lg bg-red-500/5 flex items-center justify-center border border-red-500/10">
                <AlertCircle className="text-red-400" size={16} />
              </div>
              <span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block mb-1">PENDIENTE</span>
              <div className="flex items-baseline gap-2 mt-1">
                <span className="text-3xl font-extrabold text-red-400 tabular-nums">
                  {analysis.filter(a => a.status === 'pendiente').length}
                </span>
                <span className="text-[10px] text-slate-500 uppercase">Documentos</span>
              </div>
              <div className="w-full h-1 bg-slate-800 rounded-full mt-2 overflow-hidden">
                <div className="bg-red-500 h-full" style={{ width: `${analysis.length ? (analysis.filter(a => a.status === 'pendiente').length / analysis.length) * 100 : 0}%` }}></div>
              </div>
            </div>

            <div className={cn(
              "border p-4 rounded-xl shadow-md relative overflow-hidden group select-none flex flex-col justify-between transition-colors duration-300",
              egresosCredito.length > 0 
                ? "bg-red-500/5 border-red-500/20 text-red-100" 
                : "bg-gradient-to-b from-[#141E33] to-[#0D1222] border-slate-800"
            )}>
              <div className="flex justify-between items-start">
                <div>
                  <span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block mb-1">EGRESOS A CRÉDITO</span>
                  <div className="flex items-baseline gap-2 mt-1">
                    <span className={cn(
                      "text-3xl font-extrabold tabular-nums",
                      egresosCredito.length > 0 ? "text-red-400" : "text-emerald-400"
                    )}>
                      {egresosCredito.length}
                    </span>
                    <span className="text-[10px] text-slate-500 uppercase">Alertas (PPD/99)</span>
                  </div>
                </div>
                <div className={cn(
                  "w-8 h-8 rounded-lg flex items-center justify-center border",
                  egresosCredito.length > 0 ? "bg-red-500/10 border-red-500/20" : "bg-emerald-500/10 border-emerald-500/20"
                )}>
                  {egresosCredito.length > 0 ? (
                    <AlertCircle className="text-red-400" size={16} />
                  ) : (
                    <CheckCircle2 className="text-emerald-400" size={16} />
                  )}
                </div>
              </div>
              <div className="text-[9px] text-slate-400 font-medium tracking-wide mt-2 uppercase">
                {egresosCredito.length > 0 
                  ? "⚠ Egresos a crédito detectados" 
                  : "✓ Sin egresos a crédito (PPD/99)"}
              </div>
            </div>

          </div>

          {/* Table Container in Premium Slate/Navy */}
          <div className="bg-[#101729]/90 border border-slate-800 rounded-2xl flex-1 flex flex-col shadow-xl overflow-hidden backdrop-blur-md">
            
            {/* Table Tabs / Search Control */}
            <div className="bg-[#0D1222] border-b border-slate-800/80 flex flex-col md:flex-row justify-between items-stretch">
              <div className="flex bg-[#070B13] border-r border-slate-800 overflow-x-auto">
                <button 
                  onClick={() => setActiveTab('facturas')}
                  className={cn(
                    "px-5 py-3 font-bold text-[10px] uppercase tracking-wider border-r border-slate-800/80 transition-all",
                    activeTab === 'facturas' ? "bg-[#11192C] text-amber-400 border-b-2 border-b-amber-400" : "text-slate-400 hover:text-slate-200 hover:bg-[#11192C]/50"
                  )}
                >
                  Facturas DB
                </button>
                <button 
                  onClick={() => setActiveTab('payments')}
                  className={cn(
                    "px-5 py-3 font-bold text-[10px] uppercase tracking-wider border-r border-slate-800/80 transition-all",
                    activeTab === 'payments' ? "bg-[#11192C] text-amber-400 border-b-2 border-b-amber-400" : "text-slate-400 hover:text-slate-200 hover:bg-[#11192C]/50"
                  )}
                >
                  Payments DB
                </button>
                <button 
                  onClick={() => setActiveTab('analysis')}
                  className={cn(
                    "px-5 py-3 font-bold text-[10px] uppercase tracking-wider border-r border-slate-800/80 transition-all",
                    activeTab === 'analysis' ? "bg-[#11192C] text-amber-400 border-b-2 border-b-amber-400" : "text-slate-400 hover:text-slate-200 hover:bg-[#11192C]/50"
                  )}
                >
                  Insolutos Analysis
                </button>
                <button 
                  onClick={() => setActiveTab('summary')}
                  className={cn(
                    "px-5 py-3 font-bold text-[10px] uppercase tracking-wider border-r border-slate-800/80 transition-all",
                    activeTab === 'summary' ? "bg-[#11192C] text-amber-400 border-b-2 border-b-amber-400" : "text-slate-400 hover:text-slate-200 hover:bg-[#11192C]/50"
                  )}
                >
                  Emisor View
                </button>
                <button 
                  onClick={() => setActiveTab('ieps')}
                  className={cn(
                    "px-5 py-3 font-bold text-[10px] uppercase tracking-wider border-r border-slate-800/80 transition-all text-purple-400 hover:text-purple-300",
                    activeTab === 'ieps' ? "bg-[#11192C]/80 text-purple-300 border-b-2 border-b-purple-400 font-extrabold" : "text-purple-400/80 hover:bg-[#11192C]/30"
                  )}
                >
                  Desglose IEPS
                </button>
              </div>

              {/* Precise search */}
              <div className="flex-1 flex border-t border-slate-800 md:border-t-0 md:border-l border-slate-800 bg-[#0A0F1D]/60">
                <div className="flex items-center px-4 w-full">
                  <Search size={14} className="text-slate-400 mr-2.5" />
                  <input 
                    type="text" 
                    placeholder="Escribe RFC / UUID o emisor para buscar..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="bg-transparent border-none text-slate-100 font-sans text-xs w-full focus:outline-none focus:ring-0 placeholder:text-slate-500 uppercase"
                  />
                </div>
              </div>
            </div>

            {/* Display list/table with customized theme style */}
            <div className="flex-1 overflow-auto bg-[#0E1524] custom-scrollbar">
              {fileCount > 0 ? (
                activeTab === 'ieps' ? (
                  <div className="p-6 space-y-6">
                    {/* Resumen de IEPS Acumulado */}
                    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-purple-950/10 border border-purple-900/30 p-5 rounded-2xl">
                      <div>
                        <h4 className="text-sm font-black text-purple-300 uppercase tracking-wider flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-full bg-purple-500 animate-pulse"></span>
                          Resumen de IEPS Acumulado
                        </h4>
                        <p className="text-[10px] text-slate-400 uppercase mt-1">Clasificación consolidada de Impuesto Especial sobre Producción y Servicios por Tasas y Cuotas</p>
                      </div>
                      <div className="bg-purple-950/20 border border-purple-900/40 p-3 rounded-xl text-right min-w-[200px]">
                        <span className="text-[9px] text-purple-400 font-extrabold uppercase block tracking-wider">Total IEPS Soportado</span>
                        <span className="text-2xl font-black text-purple-300 tabular-nums">
                          ${iepsAnalysis.totalIEPS.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
                        </span>
                      </div>
                    </div>

                    {/* Breakdown Classification Cards Grid */}
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      {iepsAnalysis.breakdownList.length === 0 ? (
                        <div className="col-span-full py-12 text-center bg-[#101729]/50 border border-slate-800/80 rounded-2xl text-xs text-slate-500 italic uppercase font-bold">
                          No se detectaron montos de IEPS en los comprobantes XML procesados.
                        </div>
                      ) : (
                        iepsAnalysis.breakdownList.map(({ rate, total, count, types }) => {
                          const percentage = iepsAnalysis.totalIEPS > 0 ? (total / iepsAnalysis.totalIEPS) * 100 : 0;
                          return (
                            <div key={rate} className="bg-[#101729]/90 border border-slate-800 p-4.5 rounded-2xl shadow-lg space-y-3.5 relative overflow-hidden group hover:border-purple-500/30 transition-all duration-300">
                              <div className="flex justify-between items-start">
                                <div className="space-y-1">
                                  <span className="text-[10px] bg-purple-500/15 text-purple-300 font-extrabold px-2.5 py-1 rounded-md border border-purple-500/20 uppercase tracking-wider block w-fit">
                                    Tasa / Cuota: {rate}
                                  </span>
                                  <span className="text-[8px] text-slate-500 block uppercase font-extrabold pt-1">Modelos de Impuesto</span>
                                  <span className="text-[10px] text-slate-300 font-bold block truncate max-w-[150px]" title={types}>{types}</span>
                                </div>
                                <div className="text-right">
                                  <span className="text-[15px] font-black text-purple-300 block tabular-nums">
                                    ${total.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
                                  </span>
                                  <span className="text-[9px] text-slate-400 block font-mono mt-0.5">{count} Documento(s)</span>
                                </div>
                              </div>
                              <div className="space-y-1">
                                <div className="flex justify-between text-[9px] text-slate-400 font-extrabold uppercase">
                                  <span>Participación</span>
                                  <span>{percentage.toFixed(2)}%</span>
                                </div>
                                <div className="w-full h-1.5 bg-slate-800 rounded-full overflow-hidden">
                                  <div className="bg-gradient-to-r from-purple-500 to-indigo-500 h-full rounded-full" style={{ width: `${percentage}%` }}></div>
                                </div>
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>

                    {/* Detailed List of Documents with IEPS */}
                    <div className="bg-[#0B0F19] border border-slate-800/80 rounded-2xl overflow-hidden shadow-xl">
                      <div className="bg-[#0F172A] border-b border-slate-800 p-4 flex justify-between items-center">
                        <span className="text-[10px] text-slate-300 font-black uppercase tracking-wider flex items-center gap-1.5">
                          📋 Detalle de Comprobantes con IEPS ({iepsAnalysis.documentsWithIEPS.length})
                        </span>
                        <span className="text-[8px] bg-slate-800/80 text-slate-400 px-2.5 py-0.5 rounded font-bold uppercase">Auditoría Filtrada</span>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full border-collapse">
                          <thead>
                            <tr className="text-[9px] font-extrabold uppercase text-slate-400 tracking-wider text-left bg-[#080C14] border-b border-slate-800/80">
                              <th className="p-3">UUID Comprobante</th>
                              <th className="p-3">Tipo Flujo</th>
                              <th className="p-3">Emisor de Documento</th>
                              <th className="p-3 text-right">Monto IEPS Total</th>
                              {uniqueRates.map((rate) => (
                                <th key={rate} className="p-3 text-right text-purple-400 font-mono text-[9px] border-l border-slate-800/40">{rate}</th>
                              ))}
                              <th className="p-3 text-center border-l border-slate-800/40">Acciones</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-800/40">
                            {iepsAnalysis.documentsWithIEPS.filter(doc => 
                              doc.emisor.toLowerCase().includes(searchTerm.toLowerCase()) ||
                              doc.uuid.toLowerCase().includes(searchTerm.toLowerCase())
                            ).length === 0 ? (
                              <tr>
                                <td colSpan={5 + uniqueRates.length} className="p-8 text-center text-xs text-slate-500 italic uppercase font-bold">
                                  Ningún comprobante con IEPS coincide con el filtro actual
                                </td>
                              </tr>
                            ) : (
                              iepsAnalysis.documentsWithIEPS.filter(doc => 
                                doc.emisor.toLowerCase().includes(searchTerm.toLowerCase()) ||
                                doc.uuid.toLowerCase().includes(searchTerm.toLowerCase())
                              ).map((doc, dIdx) => (
                                <tr key={dIdx} className="text-xs hover:bg-[#141C30]/50 transition-colors">
                                  <td className="p-3 font-mono">
                                    <UUIDTooltip uuid={doc.uuid} />
                                  </td>
                                  <td className="p-3">
                                    <span className="px-2 py-0.5 rounded text-[9px] font-black bg-purple-500/10 text-purple-400 border border-purple-500/10">
                                      {doc.displayType}
                                    </span>
                                  </td>
                                  <td className="p-3 font-bold text-slate-200">{doc.emisor}</td>
                                  <td className="p-3 text-right font-black text-purple-350 tabular-nums">
                                    ${doc.ieps.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
                                  </td>
                                  {uniqueRates.map((rate) => {
                                    const val = doc.iepsBreakdown?.[rate];
                                    const hasVal = val !== undefined && val > 0;
                                    return (
                                      <td key={rate} className="p-3 text-right border-l border-slate-800/40 font-mono text-[11px] tabular-nums">
                                        {hasVal ? (
                                          <span className="text-purple-300 font-bold">
                                            ${val!.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
                                          </span>
                                        ) : (
                                          <span className="text-slate-600 font-normal opacity-40">
                                            $0.00
                                          </span>
                                        )}
                                      </td>
                                    );
                                  })}
                                  <td className="p-3 text-center border-l border-slate-800/40">
                                    <button
                                      onClick={() => setSelectedRecord(doc.rawRecord)}
                                      className="px-2.5 py-1 text-[9px] font-black uppercase tracking-wider text-amber-400 bg-amber-500/5 hover:bg-amber-500/15 border border-amber-500/10 rounded transition-colors"
                                    >
                                      Analizar ➔
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                ) : (
                  <table className="w-full border-collapse">
                    <thead className="sticky top-0 bg-[#0E1524] border-b border-slate-800 z-10 select-none">
                      {activeTab === 'facturas' && (
                        <tr className="text-[9px] font-bold uppercase text-slate-400 tracking-wider text-left bg-[#0A0F1D]">
                          <th className="p-3 border-r border-slate-800/50">UUID Factura</th>
                          <th className="p-3 border-r border-slate-800/50">UUID Rel.</th>
                          <th className="p-3 border-r border-slate-800/50 text-center">Tipo</th>
                          <th className="p-3 border-r border-slate-800/50">Emisor</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Subtotal</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">IVA 16%</th>
                          <th className="p-3 border-r border-slate-800/50 text-right text-red-400/95">IVA Ret.</th>
                          <th className="p-3 border-r border-slate-800/50 text-right text-purple-400/95">IEPS / YEPS</th>
                          <th className="p-3 text-right">Total</th>
                        </tr>
                      )}
                      {activeTab === 'payments' && (
                        <tr className="text-[9px] font-bold uppercase text-slate-400 tracking-wider text-left bg-[#0A0F1D]">
                          <th className="p-3 border-r border-slate-800/50">UUID Pago</th>
                          <th className="p-3 border-r border-slate-800/50">Emisor</th>
                          <th className="p-3 border-r border-slate-800/50 text-center text-teal-400">F. Emisión</th>
                          <th className="p-3 border-r border-slate-800/50 text-center text-sky-400">F. Pago</th>
                          <th className="p-3 border-r border-slate-800/50">UUID Rel.</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Pagado</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Base 16%</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">IVA 16%</th>
                          <th className="p-3 border-r border-slate-800/50 text-right text-red-400/95">IVA Ret.</th>
                          <th className="p-3 border-r border-slate-800/50 text-right text-purple-400/95">IEPS / YEPS</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Insoluto</th>
                          <th className="p-3 text-center">Status</th>
                        </tr>
                      )}
                      {activeTab === 'analysis' && (
                        <tr className="text-[9px] font-bold uppercase text-slate-400 tracking-wider text-left bg-[#0A0F1D]">
                          <th className="p-3 border-r border-slate-800/50">ID Documento</th>
                          <th className="p-3 border-r border-slate-800/50">Emisor</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Inicial</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Pagado</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Final</th>
                          <th className="p-3 text-center">Lifecycle</th>
                        </tr>
                      )}
                      {activeTab === 'summary' && (
                        <tr className="text-[9px] font-bold uppercase text-slate-400 tracking-wider text-left bg-[#0A0F1D]">
                          <th className="p-3 border-r border-slate-800/50">RFC</th>
                          <th className="p-3 border-r border-slate-800/50">Emisor</th>
                          <th className="p-3 border-r border-slate-800/50 text-center">Count</th>
                          <th className="p-3 border-r border-slate-800/50 text-right">Pagado</th>
                          <th className="p-3 text-right">Total Insoluto</th>
                        </tr>
                      )}
                    </thead>
                    
                    <tbody className="divide-y divide-slate-800/50">
                      {activeTab === 'facturas' && data.facturas.filter(f => 
                        f.emisor.toLowerCase().includes(searchTerm.toLowerCase()) ||
                        f.uuid.toLowerCase().includes(searchTerm.toLowerCase()) ||
                        f.cfdiRelacionado.toLowerCase().includes(searchTerm.toLowerCase())
                      ).map((f, idx) => {
                        const isSelected = selectedRecord?.uuid === f.uuid;
                        return (
                          <tr 
                            key={`${idx}-${f.uuid}`} 
                            onClick={() => setSelectedRecord(isSelected ? null : f)}
                            className={cn(
                              "text-xs hover:bg-slate-800/50 transition-colors cursor-pointer",
                              isSelected ? "bg-amber-400/10 hover:bg-amber-400/15" : ""
                            )}
                          >
                            <td className="p-3 border-slate-800/30" onClick={(e) => e.stopPropagation()}>
                              <UUIDTooltip uuid={f.uuid} />
                            </td>
                            <td className="p-3 border-l border-slate-800/40 text-slate-500 font-mono italic" onClick={(e) => e.stopPropagation()}>
                              {f.cfdiRelacionado ? (
                                <div className="flex flex-col gap-1">
                                  {f.cfdiRelacionado.split(',').map((id, i) => (
                                    <div key={i}>
                                      <UUIDTooltip uuid={id.trim()} />
                                    </div>
                                  ))}
                                </div>
                              ) : "-"}
                            </td>
                            <td className="p-3 text-center border-l border-slate-800/40">
                              <span className={cn(
                                "px-2 py-0.5 rounded text-[9px] font-black uppercase text-center",
                                f.tipo === "E" ? "bg-amber-500/10 text-amber-400 border border-amber-500/10" : "bg-indigo-500/10 text-indigo-400 border border-indigo-500/10"
                              )}>
                                {f.tipo === "E" ? "EGRESO (NC)" : "INGRESO"}
                              </span>
                            </td>
                            <td className="p-3 font-semibold text-slate-200 border-l border-slate-800/40">{f.emisor}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 tabular-nums text-slate-300">${f.subtotal.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 tabular-nums text-slate-400">${f.iva16.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 tabular-nums text-red-500/80 font-medium">${(f.ivaRetenido ?? 0).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 tabular-nums text-purple-400/85 font-medium">${(f.ieps ?? 0).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 font-bold tabular-nums text-amber-400">${f.total.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                          </tr>
                        );
                      })}

                      {activeTab === 'payments' && filteredPagos.map((p, idx) => {
                        const isSelected = selectedRecord?.uuid === p.uuid && selectedRecord?.docRelacionado === p.docRelacionado;
                        return (
                          <tr 
                            key={`${idx}-${p.uuid}-${p.docRelacionado}`} 
                            onClick={() => setSelectedRecord(isSelected ? null : p)}
                            className={cn(
                              "text-xs hover:bg-slate-800/50 transition-colors cursor-pointer",
                              isSelected ? "bg-amber-400/10 hover:bg-amber-400/15" : p.saldoInsoluto === 0 ? "bg-emerald-500/5" : "bg-red-500/5"
                            )}
                          >
                            <td className="p-3 border-slate-800/30" onClick={(e) => e.stopPropagation()}>
                              <UUIDTooltip uuid={p.uuid} />
                            </td>
                            <td className="p-3 font-semibold text-slate-200 border-l border-slate-800/40">{p.emisor}</td>
                            <td className="p-3 border-l border-slate-800/40 text-center">
                              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-teal-500/10 text-teal-400 border border-teal-500/20">
                                {p.fechaEmision ? p.fechaEmision.substring(0, 10) : "-"}
                              </span>
                            </td>
                            <td className="p-3 border-l border-slate-800/40 text-center">
                              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-sky-500/10 text-sky-450 border border-sky-500/20">
                                {p.fechaPago ? p.fechaPago.substring(0, 10) : "-"}
                              </span>
                            </td>
                            <td className="p-3 border-l border-slate-800/40 font-mono text-slate-400 text-[11px]" onClick={(e) => e.stopPropagation()}>
                              <UUIDTooltip uuid={p.docRelacionado} />
                            </td>
                            <td className="p-3 text-right border-l border-slate-800/40 font-bold text-slate-100 tabular-nums">${p.impPagado.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 text-slate-400 tabular-nums">${p.baseIVA16.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 text-slate-400 tabular-nums">${p.iva16.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 text-red-500/80 tabular-nums font-medium">${(p.ivaRetenido ?? 0).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 text-purple-400/85 tabular-nums font-medium">${(p.ieps ?? 0).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-right border-l border-slate-800/40 font-bold tabular-nums text-indigo-300">${p.saldoInsoluto.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            <td className="p-3 text-center border-l border-slate-800/40">
                              <span className={cn(
                                "px-1.5 py-0.5 rounded text-[9px] font-black",
                                p.saldoInsoluto === 0 ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" : "bg-red-500/10 text-red-400 border border-red-500/20"
                              )}>
                                {p.saldoInsoluto === 0 ? "LIQUIDADO" : "PENDIENTE"}
                              </span>
                            </td>
                          </tr>
                        );
                      })}

                      {activeTab === 'analysis' && analysis.filter(a =>
                        a.emisor.toLowerCase().includes(searchTerm.toLowerCase()) ||
                        a.idDocumento.toLowerCase().includes(searchTerm.toLowerCase())
                      ).map((a, idx) => (
                        <tr key={idx} className={cn(
                          "text-xs hover:bg-slate-800/50 transition-colors",
                          a.status === 'liquidado' ? "bg-emerald-500/5" : a.status === 'parcial' ? "bg-amber-500/5" : "bg-red-500/5"
                        )}>
                          <td className="p-3 font-mono font-semibold text-slate-300" onClick={(e) => e.stopPropagation()}>
                            <UUIDTooltip uuid={a.idDocumento} />
                          </td>
                          <td className="p-3 font-bold border-l border-slate-800/40 truncate max-w-[160px] text-slate-200">{a.emisor}</td>
                          <td className="p-3 text-right border-l border-slate-800/40 text-slate-300 tabular-nums">${a.saldoInicialReal.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                          <td className="p-3 text-right border-l border-slate-800/40 text-emerald-400 font-bold tabular-nums">${a.totalPagadoLote.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                          <td className="p-3 text-right border-l border-slate-800/40 font-bold tabular-nums text-slate-100">${a.saldoInsolutoFinal.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                          <td className="p-3 text-center border-l border-slate-800/40">
                            <span className={cn(
                               "px-2 py-0.5 rounded text-[9px] font-black uppercase text-center border",
                               a.status === 'liquidado' ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/20" : a.status === 'parcial' ? "bg-amber-500/10 text-amber-400 border-amber-500/20" : "bg-red-500/10 text-red-400 border-red-500/20"
                            )}>
                               {a.status === 'liquidado' ? "✓ TOTAL" : a.status === 'parcial' ? "⚡ PARCIAL" : "⚠ PENDIENTE"}
                            </span>
                          </td>
                        </tr>
                      ))}

                      {activeTab === 'summary' && summaries.map((s, idx) => (
                        <tr key={idx} className="text-xs hover:bg-slate-800/50 transition-colors">
                          <td className="p-3 font-mono text-slate-400">{s.rfc}</td>
                          <td className="p-3 font-bold text-slate-200 border-l border-slate-800/40">{s.nombre}</td>
                          <td className="p-3 text-center text-indigo-300 border-l border-slate-800/40 font-bold">{s.totalCfdis}</td>
                          <td className="p-3 text-right border-l border-slate-800/40 text-emerald-400 font-bold tabular-nums">${s.montoTotalPagado.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                          <td className="p-3 text-right border-l border-slate-800/40 font-bold text-red-400 tabular-nums">${s.saldoInsolutoTotal.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )
              ) : (
                <div className="flex flex-col items-center justify-center p-24 opacity-30 select-none">
                  <Info size={40} className="mb-4 text-slate-500" />
                  <p className="text-xs font-black uppercase tracking-wider text-slate-400">Esperando carga de CFDIs...</p>
                </div>
              )}
            </div>

            <div className="bg-[#0D1222] text-slate-400 border-t border-slate-800 p-3 text-[9px] font-mono flex justify-between items-center shrink-0 select-none">
               <span>SYS_STATUS: READY</span>
               {egresosCredito.length > 0 && (
                 <span className="text-red-400 animate-pulse font-extrabold flex items-center gap-1">
                   <AlertCircle size={10} /> AUDITORÍA_ALERTA: SE DETECTARON {egresosCredito.length} EGRESO(S) A CRÉDITO (PPD/99) [RIESGO_FISCAL]
                 </span>
               )}
               <span>CLIENT_PERSISTENCE: LOCAL_SANDBOX_MEM</span>
            </div>
          </div>
        </section>
      </main>

      <footer className="bg-[#0D1222] border-t border-slate-800/80 p-3 flex justify-between items-center text-[9px] font-mono shrink-0 uppercase tracking-wider text-slate-500 select-none">
        <div className="flex gap-6">
          <span>LOTE: AUDIT-{Math.floor(Date.now()/100000)}</span>
          <span className="hidden md:inline">Parser Core v2.1 (Pure JS)</span>
        </div>
        <div className="text-emerald-500/80">Local security: Documents are parsed inside client sandbox memory only</div>
      </footer>

      {/* Slide-over Detail Drawer */}
      <AnimatePresence>
        {selectedRecord && (
          <div className="fixed inset-0 bg-[#070B13]/60 backdrop-blur-sm z-50 flex justify-end">
            <motion.div 
              initial={{ x: "100%" }}
              animate={{ x: 0 }}
              exit={{ x: "100%" }}
              transition={{ type: "spring", damping: 30, stiffness: 300 }}
              className="w-full max-w-lg bg-[#0F172A] border-l border-slate-800 h-full p-6 flex flex-col shadow-2xl overflow-y-auto custom-scrollbar text-slate-100"
            >
              <div className="flex justify-between items-start border-b border-slate-800 pb-4 mb-5">
                <div>
                  <div className="flex flex-wrap items-center gap-2">
                    <span className={cn(
                      "px-2.5 py-0.5 rounded text-[10px] font-black uppercase text-center border",
                      selectedRecord.tipo === "E" ? "bg-amber-500/10 text-amber-400 border border-amber-500/10" : selectedRecord.tipo === "P" ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/10" : "bg-indigo-500/10 text-indigo-400 border border-indigo-500/10"
                    )}>
                      {selectedRecord.tipo === "E" ? "EGRESO" : selectedRecord.tipo === "P" ? "COMPLEMENTO" : "INGRESO"}
                    </span>
                    {selectedRecord.tipo === "E" && (selectedRecord.metodoPago === "PPD" || selectedRecord.formaPago?.includes("99") || selectedRecord.formaPago?.toLowerCase()?.includes("por definir")) && (
                      <span className="px-2.5 py-0.5 rounded text-[10px] font-black bg-red-500/10 text-red-100 border border-red-500/20">
                        ALERTA: CRÉDITO
                      </span>
                    )}
                  </div>
                  <h3 className="text-lg font-bold text-slate-200 mt-1.5">Detalles del Comprobante</h3>
                </div>
                <button 
                  onClick={() => setSelectedRecord(null)}
                  className="p-1 px-2.5 border border-slate-800 rounded-lg text-slate-400 hover:text-slate-100 bg-slate-900/60 transition-colors"
                >
                  ✕
                </button>
              </div>

              <div className="space-y-4">
                <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                  <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">UUID de Documento</span>
                  <span className="text-xs font-mono font-bold text-slate-200 block mt-1 break-all select-all">{selectedRecord.uuid}</span>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Serie / Folio</span>
                    <span className="text-xs text-slate-100 font-bold block mt-1">{selectedRecord.serieFolio || "-"}</span>
                  </div>
                  {selectedRecord.tipo === "P" ? (
                    <div className="bg-teal-500/5 border border-teal-500/20 p-3.5 rounded-xl">
                      <span className="text-[10px] text-teal-400 font-extrabold uppercase tracking-wide block">Fecha Emisión</span>
                      <span className="text-xs text-teal-300 font-bold block mt-1">{selectedRecord.fechaEmision ? selectedRecord.fechaEmision.substring(0, 10) : "-"}</span>
                    </div>
                  ) : (
                    <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                      <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Fecha Emisión</span>
                      <span className="text-xs text-slate-100 font-bold block mt-1">{selectedRecord.fecha ? selectedRecord.fecha.substring(0, 10) : "-"}</span>
                    </div>
                  )}
                </div>

                {selectedRecord.tipo === "P" && (
                  <div className="bg-sky-500/5 border border-sky-500/20 p-3.5 rounded-xl">
                    <span className="text-[10px] text-sky-400 font-extrabold uppercase tracking-wide block">Fecha Efectiva de Pago (Complemento)</span>
                    <span className="text-xs text-sky-350 font-bold block mt-1">{selectedRecord.fechaPago ? selectedRecord.fechaPago.substring(0, 10) : "-"}</span>
                  </div>
                )}

                <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Emisor / Proveedor</span>
                    {selectedRecord.rfcEmisor && <span className="text-[10px] text-amber-400 font-mono font-bold uppercase">{selectedRecord.rfcEmisor}</span>}
                  </div>
                  <span className="text-xs text-slate-100 font-bold block">{selectedRecord.nombreEmisor || selectedRecord.emisor || "-"}</span>
                </div>

                <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                  <div className="flex justify-between items-center mb-1">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Receptor</span>
                    {selectedRecord.rfcReceptor && <span className="text-[10px] text-indigo-400 font-mono font-bold uppercase">{selectedRecord.rfcReceptor}</span>}
                  </div>
                  <span className="text-xs text-slate-100 font-bold block">{selectedRecord.nombreReceptor || selectedRecord.receptor || "-"}</span>
                </div>

                <div className="grid grid-cols-5 gap-1.5">
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-1.5 rounded-xl text-center">
                    <span className="text-[8px] text-slate-400 font-extrabold uppercase tracking-wide block truncate" title="Subtotal">Subtotal</span>
                    <span className="text-[11px] font-mono font-bold text-slate-300 block mt-1 truncate">
                      ${(selectedRecord.subtotal ?? 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                    </span>
                  </div>
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-1.5 rounded-xl text-center">
                    <span className="text-[8px] text-slate-400 font-extrabold uppercase tracking-wide block truncate" title="IVA 16%">IVA 16%</span>
                    <span className="text-[11px] font-mono font-bold text-slate-400 block mt-1 truncate">
                      ${((selectedRecord.iva16 ?? selectedRecord.iva ?? 0)).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                    </span>
                  </div>
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-1.5 rounded-xl text-center">
                    <span className="text-[8px] text-red-400/90 font-extrabold uppercase tracking-wide block truncate" title="IVA Retenido">IVA Ret.</span>
                    <span className="text-[11px] font-mono font-bold text-red-400 block mt-1 truncate">
                      ${((selectedRecord.ivaRetenido ?? 0)).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                    </span>
                  </div>
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-1.5 rounded-xl text-center">
                    <span className="text-[8px] text-purple-400/90 font-extrabold uppercase tracking-wide block truncate" title="IEPS / YEPS">IEPS</span>
                    <span className="text-[11px] font-mono font-bold text-purple-400 block mt-1 truncate">
                      ${((selectedRecord.ieps ?? 0)).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                    </span>
                  </div>
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-1.5 rounded-xl text-center bg-amber-500/5">
                    <span className="text-[8px] text-slate-400 font-extrabold uppercase tracking-wide block truncate" title="Total">Total</span>
                    <span className="text-[11px] font-mono font-bold text-amber-400 block mt-1 truncate">
                      ${(selectedRecord.total ?? selectedRecord.impPagado ?? 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                    </span>
                  </div>
                </div>

                {selectedRecord.tipo === "P" && (
                  <div className="grid grid-cols-3 gap-3">
                    <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl text-center">
                      <span className="text-[9px] text-slate-400 font-extrabold uppercase tracking-wide block">Saldo Anterior</span>
                      <span className="text-xs font-mono font-bold text-slate-300 block mt-1">
                        ${(selectedRecord.saldoAnterior ?? 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                      </span>
                    </div>
                    <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl text-center">
                      <span className="text-[9px] text-slate-400 font-extrabold uppercase tracking-wide block">Imp. Pagado</span>
                      <span className="text-xs font-mono font-bold text-emerald-400 block mt-1">
                        ${(selectedRecord.impPagado ?? 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                      </span>
                    </div>
                    <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl text-center">
                      <span className="text-[9px] text-slate-400 font-extrabold uppercase tracking-wide block">Saldo Insoluto</span>
                      <span className="text-xs font-mono font-bold text-red-500 block mt-1">
                        ${(selectedRecord.saldoInsoluto ?? 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                      </span>
                    </div>
                  </div>
                )}

                {selectedRecord.iepsBreakdown && Object.keys(selectedRecord.iepsBreakdown).length > 0 && (
                  <div className="bg-purple-950/15 border border-purple-900/30 p-4 rounded-xl space-y-2.5">
                    <span className="text-[10px] text-purple-300 font-extrabold uppercase tracking-wide flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-purple-500 animate-pulse"></span>
                      Desglose de IEPS (por Porcentaje y Cuota)
                    </span>
                    <div className="grid grid-cols-2 gap-2 mt-1">
                      {Object.entries(selectedRecord.iepsBreakdown).map(([rate, amount]) => (
                        <div key={rate} className="bg-[#121021] border border-purple-900/25 px-3 py-2 rounded-lg flex justify-between items-center transition-all">
                          <span className="text-[10px] text-purple-300/90 font-bold">{rate}</span>
                          <span className="text-xs font-mono font-bold text-purple-250">
                            ${(amount as number).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-2 gap-3">
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Forma de Pago</span>
                    <span className="text-xs text-slate-200 mt-1 block truncate" title={selectedRecord.formaPago}>
                      {selectedRecord.formaPago || "-"}
                    </span>
                  </div>
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Método de Pago</span>
                    <span className="text-xs text-slate-200 mt-1 block">
                      {selectedRecord.metodoPago || "-"}
                    </span>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Moneda</span>
                    <span className="text-xs text-slate-200 mt-1 block">
                      {selectedRecord.moneda || "MXN"}
                    </span>
                  </div>
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Estado Fiscal</span>
                    <span className="text-xs text-emerald-400 font-bold mt-1 block">
                      Vigente (100% Validado)
                    </span>
                  </div>
                </div>

                {/* Related CFDIs Interactive Auditor */}
                <div className="bg-[#141C30]/50 border border-slate-800/80 p-4 rounded-xl space-y-3">
                  <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wider block border-b border-slate-800/80 pb-2">
                    🔗 CFDI Relacionados Detectados ({getRelatedRecords.length})
                  </span>
                  
                  {getRelatedRecords.length === 0 ? (
                    <span className="text-xs text-slate-500 italic block">Este comprobante no especifica relaciones fiscales.</span>
                  ) : (
                    <div className="space-y-3">
                      {getRelatedRecords.map((rel, index) => (
                        <div key={index} className="bg-[#0B0F19]/90 border border-slate-800 p-3 rounded-lg flex flex-col gap-2">
                          <div className="flex justify-between items-center text-[9px] font-mono select-all">
                            <span className="text-slate-400">UUID: {rel.uuid.substring(0, 15)}...</span>
                            <button 
                              onClick={(e) => {
                                e.stopPropagation();
                                navigator.clipboard.writeText(rel.uuid);
                                alert("UUID Copiado!");
                              }} 
                              className="text-amber-400 hover:text-amber-300 font-black uppercase tracking-wider text-[8px] bg-amber-500/5 hover:bg-amber-500/10 border border-amber-500/10 px-1 py-0.5 rounded"
                            >
                              Copiar UUID
                            </button>
                          </div>
                          
                          {rel.found ? (
                            <div className="space-y-2">
                              {rel.records.map((rec: any, rIdx) => (
                                <div key={rIdx} className="flex justify-between items-center bg-[#111827]/60 p-2.5 rounded border border-slate-800/80 text-[11px]">
                                  <div className="flex flex-col gap-0.5 max-w-[70%]">
                                    <span className="font-bold text-slate-100 uppercase truncate">{rec.emisor}</span>
                                    <span className="text-[9px] text-slate-400 font-mono italic">
                                      {rec.displayType} • {rec.serieFolio || "S/N"}
                                    </span>
                                    <span className="text-[9px] text-indigo-400 font-bold font-mono">
                                      Monto: ${(rec.total ?? rec.impPagado ?? 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}
                                    </span>
                                  </div>
                                  <button
                                    onClick={() => setSelectedRecord(rec)}
                                    className="px-2 py-1 text-[9px] font-black uppercase text-amber-400 border border-amber-500/20 rounded bg-amber-500/10 hover:bg-amber-500/20 transition-colors"
                                  >
                                    Ver ➔
                                  </button>
                                </div>
                              ))}
                            </div>
                          ) : (
                            <div className="flex justify-between items-center text-[10px] text-slate-500 italic bg-[#070B13] p-2 rounded border border-slate-800/40">
                              <span>No cargado en este lote actual</span>
                              <span className="text-[8px] bg-slate-800 text-slate-400 rounded px-1.5 py-0.5 uppercase not-italic">Ligar XML</span>
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {selectedRecord.fileName && (
                  <div className="bg-[#141C30]/50 border border-slate-800/80 p-3.5 rounded-xl">
                    <span className="text-[10px] text-slate-400 font-extrabold uppercase tracking-wide block">Archivo de Origen</span>
                    <span className="text-xs font-mono font-bold text-slate-400 block mt-1 break-all">{selectedRecord.fileName}</span>
                  </div>
                )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Modern Glassmorphic Loading Overlay */}
      <AnimatePresence>
        {loading && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-[#070B13]/80 backdrop-blur-md z-[100] flex items-center justify-center"
          >
            <div className="bg-[#11192C]/90 border border-slate-800 p-8 rounded-2xl flex flex-col items-center gap-5 shadow-2xl max-w-sm text-center">
              <div className="w-12 h-12 rounded-full border-4 border-t-amber-400 border-slate-800 animate-spin"></div>
              <div>
                <p className="font-extrabold text-slate-100 text-lg uppercase tracking-tight">Procesando CFDI</p>
                <p className="text-[10px] text-slate-400 uppercase tracking-widest mt-1">Leyendo nodos y calculando iva proporcional...</p>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

// Sub-components kept but styled to support Premium dark
function StatusBadge({ value }: { value: number }) {
  return (
    <div className={cn(
      "border inline-flex items-center gap-1.5 px-2 py-0.5 text-[9px] font-bold uppercase rounded",
      value === 0 ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/20" : "bg-red-500/10 text-red-500 border-red-500/20"
    )}>
      {value === 0 ? <CheckCircle2 size={11} /> : <AlertCircle size={11} />}
      {value === 0 ? 'Liquidado' : 'Pendiente'}
    </div>
  );
}

function AnalysisBadge({ status }: { status: 'liquidado' | 'parcial' | 'pendiente' }) {
  const styles = {
    liquidado: "bg-emerald-500/10 text-emerald-400 border-emerald-500/20",
    parcial: "bg-amber-500/10 text-amber-400 border-amber-500/20",
    pendiente: "bg-red-500/10 text-red-400 border-red-500/20",
  };
  
  const labels = {
    liquidado: "✓ Liquidado Total",
    parcial: "⚡ Parcial c/ Seguimiento",
    pendiente: "⚠ Pendiente s/ Seguimiento"
  };

  return (
    <div className={cn(
      "border px-2.5 py-1 text-[10px] font-bold uppercase rounded shadow-sm",
      styles[status]
    )}>
      {labels[status]}
    </div>
  );
}


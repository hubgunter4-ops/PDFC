import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import {
  Activity,
  AlertTriangle,
  BarChart3,
  Building2,
  CalendarDays,
  CheckCircle2,
  ChevronRight,
  Clock3,
  CreditCard,
  Download,
  FileSpreadsheet,
  FileText,
  LayoutDashboard,
  LockKeyhole,
  Search,
  ShieldCheck,
  Trash2,
  Upload,
  WalletCards,
  X,
} from 'lucide-react';
import {
  enriquecerPagosConFacturas,
  parseFiles,
  type Concepto,
  type Factura,
  type PagoDoc,
  type ParseResult,
} from './services/parser';

type ViewKey = 'summary' | 'invoices' | 'payments' | 'balances' | 'ieps' | 'alerts';
type DetailRecord = Factura | PagoDoc;

const NAV_ITEMS: Array<{ id: ViewKey; label: string; icon: typeof LayoutDashboard }> = [
  { id: 'summary', label: 'Resumen', icon: LayoutDashboard },
  { id: 'invoices', label: 'CFDI', icon: FileText },
  { id: 'payments', label: 'Pagos', icon: CreditCard },
  { id: 'balances', label: 'Saldos', icon: Clock3 },
  { id: 'ieps', label: 'IEPS', icon: BarChart3 },
  { id: 'alerts', label: 'Revisión', icon: AlertTriangle },
];

const VIEW_COPY: Record<ViewKey, { eyebrow: string; title: string; subtitle: string }> = {
  summary: { eyebrow: 'Visión general', title: 'Resumen de actividad', subtitle: 'Lectura rápida de los comprobantes y movimientos que importaste.' },
  invoices: { eyebrow: 'Documentos', title: 'CFDI', subtitle: 'Consulta facturas y notas de crédito leídas desde tus archivos.' },
  payments: { eyebrow: 'Movimientos', title: 'Complementos de pago', subtitle: 'Revisa los pagos y documentos relacionados que aparecen en el XML.' },
  balances: { eyebrow: 'Seguimiento', title: 'Saldos reportados', subtitle: 'Agrupación informativa a partir de los saldos de complementos importados.' },
  ieps: { eyebrow: 'Impuestos trasladados', title: 'Desglose IEPS', subtitle: 'Importes y tasas identificados en los comprobantes importados.' },
  alerts: { eyebrow: 'Revisión humana', title: 'Observaciones', subtitle: 'Señales basadas en campos del XML; no son dictámenes ni validaciones fiscales.' },
};

const EMPTY_DATA: ParseResult = { pagos: [], facturas: [], conceptos: [] };

function mergeUnique<T>(current: T[], incoming: T[], keyOf: (item: T) => string): T[] {
  const records = new Map(current.map((item) => [keyOf(item), item]));
  incoming.forEach((item) => records.set(keyOf(item), item));
  return [...records.values()];
}

function money(value: number, currency = 'MXN'): string {
  try {
    return new Intl.NumberFormat('es-MX', { style: 'currency', currency, minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number.isFinite(value) ? value : 0);
  } catch {
    return `${currency} ${rawAmount(value)}`;
  }
}

function rawAmount(value: number): string {
  return new Intl.NumberFormat('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number.isFinite(value) ? value : 0);
}

function dateLabel(value?: string): string {
  if (!value) return '—';
  const date = new Date(`${value.slice(0, 10)}T12:00:00`);
  return Number.isNaN(date.getTime()) ? value.slice(0, 10) : new Intl.DateTimeFormat('es-MX', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
}

function shortUuid(value?: string): string {
  return value ? `${value.slice(0, 8)}…${value.slice(-4)}` : '—';
}

function invoicePaymentStatus(invoice: Factura, payments: PagoDoc[]): { label: string; tone: string } {
  const related = payments.filter((payment) => payment.docRelacionado.toLowerCase() === invoice.uuid.toLowerCase());
  if (!related.length) return { label: 'Sin complemento', tone: 'muted' };
  const latest = [...related].sort((a, b) => (a.fechaPago || '').localeCompare(b.fechaPago || '')).at(-1);
  if ((latest?.saldoInsoluto ?? 0) > 0) return { label: 'Saldo reportado', tone: 'gold' };
  return { label: 'Saldo 0 reportado', tone: 'green' };
}

function App() {
  const [data, setData] = useState<ParseResult>(EMPTY_DATA);
  const [activeView, setActiveView] = useState<ViewKey>('summary');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedRecord, setSelectedRecord] = useState<DetailRecord | null>(null);
  const [loadedFiles, setLoadedFiles] = useState<string[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);

  const fileCount = data.facturas.length + data.pagos.length;
  const creditInvoices = useMemo(() => data.facturas.filter((invoice) => invoice.metodoPago === 'PPD' || invoice.formaPago?.includes('99')), [data.facturas]);
  const balanceRows = useMemo(() => {
    const grouped = new Map<string, { uuid: string; issuer: string; initial: number; paid: number; balance: number; latestDate: string }>();
    data.pagos.forEach((payment) => {
      const key = payment.docRelacionado || payment.uuid;
      const row = grouped.get(key);
      if (!row) {
        grouped.set(key, { uuid: key, issuer: payment.emisor, initial: payment.saldoAnterior, paid: payment.impPagado, balance: payment.saldoInsoluto, latestDate: payment.fechaPago || '' });
      } else {
        row.paid += payment.impPagado;
        if ((payment.fechaPago || '') >= row.latestDate) {
          row.balance = payment.saldoInsoluto;
          row.latestDate = payment.fechaPago || '';
        }
      }
    });
    return [...grouped.values()].sort((a, b) => b.balance - a.balance);
  }, [data.pagos]);

  const iepsRows = useMemo(() => {
    const rows: Array<{ key: string; uuid: string; issuer: string; kind: string; rate: string; amount: number; currency?: string; source: DetailRecord }> = [];
    const collect = (record: DetailRecord, kind: string, currency?: string) => {
      const breakdown = Object.entries(record.iepsBreakdown || {});
      if (breakdown.length) {
        breakdown.forEach(([rate, amount], index) => rows.push({ key: `${record.uuid}-${kind}-${index}-${rate}`, uuid: record.uuid, issuer: record.emisor, kind, rate, amount: Number(amount) || 0, currency, source: record }));
      } else if ((record.ieps || 0) > 0) {
        rows.push({ key: `${record.uuid}-${kind}-unrated`, uuid: record.uuid, issuer: record.emisor, kind, rate: 'Sin tasa desglosada', amount: record.ieps || 0, currency, source: record });
      }
    };
    data.facturas.forEach((invoice) => collect(invoice, invoice.tipo === 'E' ? 'CFDI egreso' : 'CFDI ingreso', invoice.moneda || 'MXN'));
    data.pagos.forEach((payment) => collect(payment, 'Complemento de pago'));
    return rows;
  }, [data.facturas, data.pagos]);

  const openBalances = balanceRows.filter((row) => row.balance > 0).length;
  const invoiceCurrencies = [...new Set(data.facturas.map((invoice) => invoice.moneda || 'MXN'))].sort();
  const copy = VIEW_COPY[activeView];

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setSelectedRecord(null);
      if (event.key === '/' && !event.ctrlKey && !event.metaKey && !event.altKey && !['INPUT', 'TEXTAREA'].includes((event.target as HTMLElement | null)?.tagName || '')) {
        event.preventDefault();
        searchInputRef.current?.focus();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  const importFiles = async (files: FileList | File[]) => {
    const incoming = Array.from(files).filter((file) => /\.(xml|zip)$/i.test(file.name));
    if (!incoming.length) {
      setMessage('Selecciona archivos XML o ZIP que contengan XML.');
      return;
    }
    setLoading(true);
    setMessage('Leyendo archivos localmente…');
    try {
      const parsed = await parseFiles(incoming);
      setData((current) => {
        const facturas = mergeUnique(current.facturas, parsed.facturas, (item) => item.uuid.toLowerCase());
        const pagos = mergeUnique(current.pagos, parsed.pagos, (item) => `${item.uuid.toLowerCase()}|${item.docRelacionado.toLowerCase()}`);
        const conceptos = mergeUnique(current.conceptos, parsed.conceptos, (item) => `${item.uuid.toLowerCase()}|${item.descripcion}`);
        return { facturas, pagos: enriquecerPagosConFacturas(pagos, facturas), conceptos };
      });
      setLoadedFiles((current) => [...new Set([...current, ...incoming.map((file) => file.name)])]);
      const count = parsed.facturas.length + parsed.pagos.length;
      setMessage(count ? `Lectura terminada: ${parsed.facturas.length} CFDI y ${parsed.pagos.length} registros de pago encontrados.` : 'No se encontraron CFDI o complementos reconocibles en los archivos seleccionados.');
      if (count) setActiveView('summary');
    } catch (error) {
      console.error(error);
      setMessage('No se pudo leer el lote. Comprueba que los archivos XML/ZIP estén íntegros e inténtalo de nuevo.');
    } finally {
      setLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const clearData = () => {
    setData(EMPTY_DATA);
    setLoadedFiles([]);
    setSearchTerm('');
    setSelectedRecord(null);
    setMessage('Se retiraron los datos de esta sesión de memoria.');
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const exportWorkbook = async () => {
    try {
      const { generarExcel } = await import('./services/excelService');
      await generarExcel(data.pagos, data.facturas, data.conceptos);
      setMessage('Se generó el libro de Excel localmente.');
    } catch (error) {
      console.error(error);
      setMessage('No se pudo generar el archivo Excel.');
    }
  };

  const matches = (values: Array<string | number | undefined>) => values.some((value) => String(value ?? '').toLowerCase().includes(searchTerm.trim().toLowerCase()));
  const filteredInvoices = data.facturas.filter((invoice) => matches([invoice.serieFolio, invoice.uuid, invoice.emisor, invoice.rfcEmisor, invoice.receptor, invoice.concepto, invoice.tipo]));
  const filteredPayments = data.pagos.filter((payment) => matches([payment.serieFolio, payment.uuid, payment.emisor, payment.rfcEmisor, payment.docRelacionado, payment.fechaPago]));
  const filteredBalances = balanceRows.filter((row) => matches([row.uuid, row.issuer]));
  const filteredIeps = iepsRows.filter((row) => matches([row.rate, row.uuid, row.issuer, row.kind]));
  const recentInvoices = [...filteredInvoices].sort((a, b) => (b.fecha || '').localeCompare(a.fecha || '')).slice(0, 5);
  const inputCount = loadedFiles.length;

  const renderNavigation = (mobile = false) => (
    <nav className={mobile ? 'mobile-nav' : 'primary-nav'} aria-label={mobile ? 'Navegación móvil' : 'Navegación principal'}>
      {NAV_ITEMS.map(({ id, label, icon: Icon }) => {
        const count = id === 'invoices' ? data.facturas.length : id === 'payments' ? data.pagos.length : id === 'balances' ? openBalances : id === 'alerts' ? creditInvoices.length : undefined;
        return (
          <button key={id} className={`nav-item ${activeView === id ? 'is-active' : ''}`} onClick={() => setActiveView(id)} aria-current={activeView === id ? 'page' : undefined}>
            <span className="nav-icon"><Icon size={17} aria-hidden="true" /></span>
            <span>{label}</span>
            {count !== undefined && <span className="nav-count">{String(count).padStart(2, '0')}</span>}
          </button>
        );
      })}
    </nav>
  );

  const renderInvoiceTable = (invoices: Factura[], compact = false) => (
    <div className="table-scroll">
      <table className="data-table">
        <thead><tr><th>Folio · fecha</th><th>Emisor</th><th>Concepto</th><th>Tipo</th><th>Pago reportado</th><th className="numeric">Total</th></tr></thead>
        <tbody>
          {invoices.map((invoice) => {
            const status = invoicePaymentStatus(invoice, data.pagos);
            return <tr key={invoice.uuid} onClick={() => setSelectedRecord(invoice)} className="clickable-row">
              <td><strong>{invoice.serieFolio || shortUuid(invoice.uuid)}</strong><small>{dateLabel(invoice.fecha)}</small></td>
              <td><strong>{invoice.emisor || 'Emisor no identificado'}</strong><small>RFC · {invoice.rfcEmisor || '—'}</small></td>
              {!compact && <td>{invoice.cfdiRelacionado ? 'CFDI relacionado' : 'Comprobante importado'}</td>}
              <td><span className={`badge ${invoice.tipo === 'E' ? 'badge-violet' : 'badge-blue'}`}>{invoice.tipo === 'E' ? 'Egreso' : 'Ingreso'}</span></td>
              <td><span className={`badge badge-${status.tone}`}>{status.label}</span></td>
              <td className="numeric money">{money(invoice.total, invoice.moneda || 'MXN')}</td>
            </tr>;
          })}
        </tbody>
      </table>
    </div>
  );

  const renderPaymentsTable = (payments: PagoDoc[]) => (
    <div className="table-scroll"><table className="data-table">
      <thead><tr><th>Complemento · fecha</th><th>Emisor</th><th>Documento relacionado</th><th className="numeric">Importe pagado</th><th className="numeric">Saldo reportado</th><th>Estado del complemento</th></tr></thead>
      <tbody>{payments.map((payment, index) => <tr key={`${payment.uuid}-${payment.docRelacionado}-${index}`} onClick={() => setSelectedRecord(payment)} className="clickable-row">
        <td><strong>{payment.serieFolio || shortUuid(payment.uuid)}</strong><small>{dateLabel(payment.fechaPago)}</small></td>
        <td><strong>{payment.emisor || 'Emisor no identificado'}</strong><small>RFC · {payment.rfcEmisor || '—'}</small></td>
        <td className="mono">{shortUuid(payment.docRelacionado)}</td>
        <td className="numeric money">{rawAmount(payment.impPagado)}</td>
        <td className="numeric money">{rawAmount(payment.saldoInsoluto)}</td>
        <td><span className={`badge ${payment.saldoInsoluto > 0 ? 'badge-gold' : 'badge-green'}`}>{payment.saldoInsoluto > 0 ? 'Saldo pendiente' : 'Saldo 0 reportado'}</span></td>
      </tr>)}</tbody>
    </table></div>
  );

  const renderBalancesTable = () => <div className="table-scroll"><table className="data-table">
    <thead><tr><th>Documento relacionado</th><th>Emisor</th><th className="numeric">Saldo anterior registrado</th><th className="numeric">Abonos importados</th><th className="numeric">Último saldo insoluto</th><th>Último movimiento</th></tr></thead>
    <tbody>{filteredBalances.map((row) => <tr key={row.uuid}>
      <td className="mono">{shortUuid(row.uuid)}</td><td>{row.issuer}</td><td className="numeric money">{rawAmount(row.initial)}</td><td className="numeric money">{rawAmount(row.paid)}</td><td className="numeric money">{rawAmount(row.balance)}</td><td>{dateLabel(row.latestDate)}</td>
    </tr>)}</tbody>
  </table></div>;

  const renderIepsTable = () => <div className="table-scroll"><table className="data-table">
    <thead><tr><th>Documento</th><th>Origen</th><th>Tasa / etiqueta</th><th className="numeric">Importe tal como aparece</th></tr></thead>
    <tbody>{filteredIeps.map((row) => <tr key={row.key} onClick={() => setSelectedRecord(row.source)} className="clickable-row"><td className="mono">{shortUuid(row.uuid)}<small>{row.issuer}</small></td><td>{row.kind}</td><td>{row.rate}</td><td className="numeric money">{row.currency ? money(row.amount, row.currency) : rawAmount(row.amount)}</td></tr>)}</tbody>
  </table></div>;

  const renderMainPanel = () => {
    if (activeView === 'summary') return (
      <div className="dashboard-grid">
        <section className="panel activity-panel">
          <PanelHeader title="CFDI recientes" subtitle={data.facturas.length ? 'Registros extraídos de los archivos de esta sesión' : 'Importa archivos para poblar el resumen'} action={data.facturas.length ? <button className="text-action" onClick={() => setActiveView('invoices')}>Ver todos <ChevronRight size={14} /></button> : undefined} />
          {recentInvoices.length ? renderInvoiceTable(recentInvoices, true) : <EmptyState icon={<FileText />} title="Aún no hay CFDI en esta sesión" detail="Elige uno o varios XML o un ZIP para comenzar. El análisis se ejecuta en esta pestaña." />}
        </section>
        <aside className="side-stack">
          <section className="panel batch-panel">
            <div className="panel-heading-row"><div><h2>Lote de trabajo</h2><p>Archivos leídos en esta sesión</p></div><span className="local-status"><CheckCircle2 size={15} /> Local</span></div>
            <div className="batch-progress"><span style={{ width: inputCount ? '100%' : '0%' }} /></div>
            <div className="batch-meta"><span>Archivos seleccionados</span><strong>{inputCount}</strong></div>
            <div className="batch-meta"><span>Registros extraídos</span><strong>{fileCount}</strong></div>
            <div className="batch-meta"><span>Monedas en CFDI</span><strong>{invoiceCurrencies.join(', ') || '—'}</strong></div>
            {loadedFiles.length > 0 && <ul className="file-list">{loadedFiles.slice(-4).map((name) => <li key={name} title={name}><FileText size={13} />{name}</li>)}</ul>}
          </section>
          <section className="panel note-panel"><div className="note-icon"><ShieldCheck size={17} /></div><div><h2>Procesamiento local</h2><p>Los XML y ZIP se leen en el navegador. No se envían a un servidor ni se consulta al SAT.</p></div></section>
          <section className="panel note-panel note-panel-gold"><div className="note-icon"><AlertTriangle size={17} /></div><div><h2>Lectura, no dictamen</h2><p>Los campos y saldos reflejan el contenido importado. Confirma cualquier conclusión con fuentes y revisión profesional.</p></div></section>
        </aside>
      </div>
    );

    if (activeView === 'invoices') return <section className="panel table-panel"><PanelHeader title="Comprobantes importados" subtitle={`${filteredInvoices.length} de ${data.facturas.length} registros`} />{filteredInvoices.length ? renderInvoiceTable(filteredInvoices) : <EmptyState icon={<FileText />} title="No hay CFDI para mostrar" detail={searchTerm ? 'Prueba otro término de búsqueda.' : 'Importa XML o ZIP con CFDI de ingreso o egreso.'} />}</section>;
    if (activeView === 'payments') return <section className="panel table-panel"><PanelHeader title="Complementos de pago importados" subtitle={`${filteredPayments.length} de ${data.pagos.length} registros · importes sin conversión de moneda`} />{filteredPayments.length ? renderPaymentsTable(filteredPayments) : <EmptyState icon={<CreditCard />} title="No hay complementos para mostrar" detail={searchTerm ? 'Prueba otro término de búsqueda.' : 'Los complementos se listarán aquí cuando se encuentren en los XML.'} />}</section>;
    if (activeView === 'balances') return <section className="panel table-panel"><PanelHeader title="Resumen de saldos de complementos" subtitle={`${filteredBalances.length} documentos relacionados; importes de origen sin conversión`} />{filteredBalances.length ? renderBalancesTable() : <EmptyState icon={<Clock3 />} title="Aún no hay saldos importados" detail="Carga complementos de pago para agrupar los saldos reportados por documento." />}</section>;
    if (activeView === 'ieps') return <section className="panel table-panel"><PanelHeader title="Desglose IEPS por documento" subtitle="Cada importe se conserva por separado en su moneda de origen; no se agregan tasas de distintos CFDI" />{filteredIeps.length ? renderIepsTable() : <EmptyState icon={<BarChart3 />} title="No se identificaron importes IEPS" detail={searchTerm ? 'Prueba otro término de búsqueda.' : 'Los registros con IEPS aparecerán aquí si el XML incluye el desglose.'} />}</section>;

    const openPayments = data.pagos.filter((payment) => payment.saldoInsoluto > 0);
    return <div className="review-grid">
      <section className="panel table-panel"><PanelHeader title="Campos que conviene revisar" subtitle="Indicadores simples extraídos del XML; requieren interpretación humana" />
        <div className="review-list">
          <ReviewItem icon={<CreditCard size={17} />} title="CFDI con método PPD o forma 99" detail="El campo indica pago diferido o forma de pago por definir. La vista no determina cumplimiento ni vencimiento." count={creditInvoices.length} tone="gold" onClick={() => setActiveView('invoices')} />
          <ReviewItem icon={<Clock3 size={17} />} title="Complementos con saldo insoluto positivo" detail="El XML reporta un saldo mayor que cero; confirma secuencia y vigencia con documentación de respaldo." count={openPayments.length} tone="red" onClick={() => setActiveView('payments')} />
          <ReviewItem icon={<Activity size={17} />} title="Importes IEPS identificados" detail="Se encontraron campos de IEPS en comprobantes. No se verifica la tasa ni su tratamiento fiscal." count={iepsRows.length} tone="blue" onClick={() => setActiveView('ieps')} />
        </div>
      </section>
      <aside className="panel review-callout"><AlertTriangle size={19} /><h2>Señal no equivale a incumplimiento</h2><p>Estos conteos solo reflejan valores parseados. No comprueban vigencia, listas oficiales, deducibilidad ni autenticidad ante la autoridad.</p></aside>
    </div>;
  };

  return (
    <div className="auditpro-app">
      <aside className="sidebar">
        <a className="brand" href="#resumen" onClick={(event) => { event.preventDefault(); setActiveView('summary'); }} aria-label="AUDITPRO, ir al resumen">
          <span className="brand-mark"><span>△</span></span><span className="brand-word"><strong><span>AUDIT</span>PRO</strong><small>CONTROL FISCAL</small></span><span className="brand-version">CORE <i>2.2</i></span>
        </a>
        <div className="workspace-switcher"><span className="workspace-avatar">P</span><span className="workspace-copy"><small>ESPACIO DE TRABAJO</small><strong>PDFC · Local</strong></span></div>
        <p className="nav-label">REVISIÓN</p>
        {renderNavigation()}
        <div className="sidebar-spacer" />
        <div className="sidebar-note"><span className="local-pulse" /><div><strong>Procesamiento local</strong><small>Sin conexión externa</small></div><LockKeyhole size={14} className="lock-icon" /></div>
        <div className="sidebar-profile"><span className="profile-avatar">AP</span><div><strong>Espacio privado</strong><small>Los datos viven en esta sesión</small></div></div>
      </aside>

      <div className="main-shell">
        <header className="topbar">
          <div className="topbar-context"><span className="topbar-dot" /><strong>REVISIÓN LOCAL</strong><span className="topbar-divider" /><span>{inputCount ? `${inputCount} archivo${inputCount === 1 ? '' : 's'} en sesión` : 'Sin lote cargado'}</span></div>
          <div className="topbar-actions"><span className="local-chip"><span />DATOS EN TU SESIÓN</span><span className="avatar-button" aria-label="Procesamiento local">AP</span></div>
        </header>
        <div className="disclaimer"><span className="disclaimer-symbol">i</span><span><strong>Alcance de lectura.</strong> Los archivos se procesan en el navegador. Esta herramienta no valida CFDI ante el SAT ni ofrece asesoría fiscal.</span></div>

        <main className="page-content">
          <div className="page-heading">
            <div><p className="eyebrow">{copy.eyebrow}</p><h1>{copy.title}</h1><p className="heading-sub">{copy.subtitle}</p></div>
            <div className="heading-actions">
              {activeView !== 'summary' && <label className="search-box"><Search size={15} /><input ref={searchInputRef} value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} placeholder="Buscar folio, emisor, UUID…" aria-label="Buscar en esta vista" /><kbd>/</kbd></label>}
              <button className="button button-primary" onClick={() => fileInputRef.current?.click()} disabled={loading}><Upload size={15} />{loading ? 'Leyendo…' : 'Importar archivos'}</button>
              <input ref={fileInputRef} className="visually-hidden" type="file" multiple accept=".xml,.zip,application/xml,application/zip" onChange={(event) => event.target.files && importFiles(event.target.files)} aria-label="Seleccionar archivos XML o ZIP" />
            </div>
          </div>

          <section className="metrics-grid" aria-label="Resumen de registros importados">
            <MetricCard icon={<FileText />} label="CFDI en sesión" value={data.facturas.length} foot={`${data.facturas.filter((invoice) => invoice.tipo === 'E').length} egresos identificados`} tone="blue" />
            <MetricCard icon={<WalletCards />} label="Complementos de pago" value={data.pagos.length} foot="Importes tal como aparecen en el XML" tone="green" />
            <MetricCard icon={<Clock3 />} label="Saldos positivos" value={openBalances} foot="Según último complemento por documento" tone="gold" />
            <MetricCard icon={<BarChart3 />} label="Desgloses IEPS" value={iepsRows.length} foot={`${new Set(iepsRows.map((row) => row.rate)).size} tasas/etiquetas; sin agregación`} tone="violet" />
          </section>

          {message && <div className={`toast-message ${message.startsWith('No ') ? 'is-error' : ''}`} role="status">{message}<button onClick={() => setMessage('')} aria-label="Cerrar mensaje"><X size={14} /></button></div>}

          <div className={`import-strip ${isDragging ? 'is-dragging' : ''}`} onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }} onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setIsDragging(false); }} onDrop={(event) => { event.preventDefault(); setIsDragging(false); void importFiles(event.dataTransfer.files); }}>
            <span className="import-icon"><Upload size={18} /></span><div><strong>{isDragging ? 'Suelta los archivos para leerlos' : 'Añade CFDI XML o un lote ZIP'}</strong><small>Se procesan en memoria en este navegador. No se cargan a un servidor.</small></div><button className="button button-secondary" onClick={() => fileInputRef.current?.click()} disabled={loading}>Seleccionar archivos <ChevronRight size={15} /></button>
          </div>

          {renderMainPanel()}

          {data.facturas.length > 0 || data.pagos.length > 0 ? <div className="below-actions"><span><LockKeyhole size={13} /> Los datos cargados permanecen solo en esta sesión de navegador.</span><div><button className="button button-secondary" onClick={() => void exportWorkbook()}><Download size={15} />Exportar Excel</button><button className="button button-quiet" onClick={clearData}><Trash2 size={14} />Limpiar sesión</button></div></div> : null}
        </main>
        <footer className="app-footer"><span><i className="trust-dot" />PARSEO EN CLIENTE</span><span className="footer-note">Los valores reflejan XML importados; sin consulta externa ni validación oficial.</span><span className="footer-lock"><LockKeyhole size={12} /> SIN BACKEND</span></footer>
      </div>

      {renderNavigation(true)}
      {selectedRecord && <RecordDrawer record={selectedRecord} facturas={data.facturas} onClose={() => setSelectedRecord(null)} onSelectInvoice={(invoice) => setSelectedRecord(invoice)} />}
    </div>
  );
}

function PanelHeader({ title, subtitle, action }: { title: string; subtitle: string; action?: ReactNode }) {
  return <header className="panel-header"><div><h2 className="panel-title">{title}</h2><p className="panel-subtitle">{subtitle}</p></div>{action}</header>;
}

function MetricCard({ icon, label, value, foot, tone }: { icon: ReactNode; label: string; value: string | number; foot: string; tone: string }) {
  return <article className={`metric-card metric-${tone}`}><div className="metric-top"><span>{label}</span><span className="metric-icon">{icon}</span></div><div className="metric-value">{value}</div><p className="metric-foot"><span className="mini-dot" />{foot}</p></article>;
}

function EmptyState({ icon, title, detail }: { icon: ReactNode; title: string; detail: string }) {
  return <div className="empty-state"><span className="empty-icon">{icon}</span><strong>{title}</strong><p>{detail}</p></div>;
}

function ReviewItem({ icon, title, detail, count, tone, onClick }: { icon: ReactNode; title: string; detail: string; count: number; tone: string; onClick: () => void }) {
  return <button className="review-item" onClick={onClick}><span className={`review-icon tone-${tone}`}>{icon}</span><span className="review-copy"><strong>{title}</strong><small>{detail}</small></span><span className={`review-count tone-${tone}`}>{count}</span><ChevronRight size={16} /></button>;
}

function RecordDrawer({ record, facturas, onClose, onSelectInvoice }: { record: DetailRecord; facturas: Factura[]; onClose: () => void; onSelectInvoice: (invoice: Factura) => void }) {
  const isPayment = 'impPagado' in record;
  const invoice = isPayment ? null : record;
  const payment = isPayment ? record : null;
  const linkedInvoice = payment ? facturas.find((item) => item.uuid.toLowerCase() === payment.docRelacionado.toLowerCase()) : null;
  const amount = payment ? payment.impPagado : invoice?.total || 0;
  const uuid = record.uuid;
  const issuer = record.nombreEmisor || record.emisor;
  const receiver = record.nombreReceptor || record.receptor;
  const date = payment ? payment.fechaEmision : invoice?.fecha;

  return <div className="drawer-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
    <aside className="record-drawer" role="dialog" aria-modal="true" aria-labelledby="drawer-title">
      <div className="drawer-header"><div><p className="eyebrow">Ficha de lectura · {isPayment ? 'Complemento de pago' : 'CFDI'}</p><h2 id="drawer-title">{record.serieFolio || shortUuid(uuid)}</h2><span className="mono drawer-uuid">UUID · {uuid}</span></div><button className="icon-button" onClick={onClose} aria-label="Cerrar ficha"><X size={18} /></button></div>
      <div className="drawer-amount"><span>{isPayment ? 'Importe pagado en el complemento · sin conversión' : 'Total del comprobante'}</span><strong>{payment ? rawAmount(amount) : money(amount, invoice?.moneda || 'MXN')}</strong><small>Importe leído del archivo; no es una validación fiscal.</small></div>
      <section className="detail-section"><h3><Building2 size={14} /> Emisor y receptor</h3><DetailField label="Emisor" value={issuer} meta={record.rfcEmisor ? `RFC · ${record.rfcEmisor}` : undefined} /><DetailField label="Receptor" value={receiver} meta={record.rfcReceptor ? `RFC · ${record.rfcReceptor}` : undefined} /></section>
      <section className="detail-section"><h3><CalendarDays size={14} /> Datos del comprobante</h3><div className="detail-grid"><DetailField label="Fecha de emisión" value={dateLabel(date)} />{payment && <DetailField label="Fecha de pago" value={dateLabel(payment.fechaPago)} />}{invoice && <><DetailField label="Tipo CFDI" value={invoice.tipo === 'E' ? 'Egreso' : 'Ingreso'} /><DetailField label="Método de pago" value={invoice.metodoPago || '—'} /></>}{payment && <DetailField label="Documento relacionado" value={shortUuid(payment.docRelacionado)} />}</div></section>
      <section className="detail-section"><h3><FileSpreadsheet size={14} /> Importes leídos</h3><div className="detail-grid"><DetailField label="Subtotal / saldo anterior" value={payment ? rawAmount(payment.saldoAnterior) : money(invoice?.subtotal || 0, invoice?.moneda || 'MXN')} /><DetailField label="IVA 16%" value={payment ? rawAmount(record.iva16 || 0) : money(record.iva16 || 0, invoice?.moneda || 'MXN')} /><DetailField label="IVA retenido" value={payment ? rawAmount(record.ivaRetenido || 0) : money(record.ivaRetenido || 0, invoice?.moneda || 'MXN')} /><DetailField label="IEPS" value={payment ? rawAmount(record.ieps || 0) : money(record.ieps || 0, invoice?.moneda || 'MXN')} />{payment && <><DetailField label="Saldo insoluto en XML" value={rawAmount(payment.saldoInsoluto)} /><DetailField label="Forma de pago" value={payment.formaPago || '—'} /></>}</div></section>
      {record.iepsBreakdown && Object.keys(record.iepsBreakdown).length > 0 && <section className="detail-section"><h3>Desglose IEPS leído</h3>{Object.entries(record.iepsBreakdown).map(([rate, value]) => <div className="detail-line" key={rate}><span>{rate}</span><strong>{invoice ? money(Number(value), invoice.moneda || 'MXN') : rawAmount(Number(value))}</strong></div>)}</section>}
      {linkedInvoice && <button className="linked-record" onClick={() => onSelectInvoice(linkedInvoice)}><span><strong>Ver CFDI relacionado</strong><small>{linkedInvoice.serieFolio || shortUuid(linkedInvoice.uuid)}</small></span><ChevronRight size={17} /></button>}
      {!linkedInvoice && invoice?.cfdiRelacionado && <section className="detail-section"><h3>UUID relacionados</h3><p className="related-uuid">{invoice.cfdiRelacionado}</p></section>}
      <div className="drawer-caveat"><ShieldCheck size={16} /><span>Se muestran los datos extraídos del XML. No se confirmó vigencia, autenticidad, situación ante el SAT ni deducibilidad.</span></div>
    </aside>
  </div>;
}

function DetailField({ label, value, meta }: { label: string; value: string; meta?: string }) {
  return <div className="detail-field"><span>{label}</span><strong>{value || '—'}</strong>{meta && <small>{meta}</small>}</div>;
}

export default App;

import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Download, FileSpreadsheet, AlertTriangle, CheckCircle2, Loader2, RotateCcw, ArrowRight, Info, Copy, Filter, Hash, AlertCircle } from 'lucide-react';

// ---------- helpers ----------
const col = (letters) => {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
};
const clean = (v) => String(v ?? '').replace(/ /g, '').replace(/\s/g, '').replace(/-/g, '');
const trimUpper = (v) => String(v ?? '').trim().toUpperCase();
const commodity = (v) => {
  const c = trimUpper(v);
  if (c === 'CN') return 'corn';
  if (c === 'BN') return 'soybeans';
  return '';
};
const isCornOrBeans = (r, g) => {
  const c = trimUpper(g(r,'O'));
  return c === 'CN' || c === 'BN';
};
const fmtDate = (v) => {
  if (v === null || v === undefined || v === '') return '';
  let d = null;
  if (v instanceof Date) d = v;
  else if (typeof v === 'number') {
    const p = XLSX.SSF.parse_date_code(v);
    if (p) d = new Date(p.y, p.m - 1, p.d);
  } else {
    const parsed = new Date(v);
    if (!isNaN(parsed)) d = parsed;
  }
  if (!d) return String(v);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${mm}/${dd}/${d.getFullYear()}`;
};
const trimStr = (v) => String(v ?? '').replace(/ /g, ' ').trim();

// ---------- validation helpers ----------
const findDuplicateImportIds = (rows, templateName) => {
  const seen = {};
  rows.forEach((r, i) => {
    const id = String(r.import_id ?? '');
    if (!id) return;
    if (!seen[id]) seen[id] = [];
    seen[id].push({ idx: i, ticket: r.ticket_number || r['source.ticket_number'] || '' });
  });
  const warnings = [];
  Object.entries(seen).filter(([, v]) => v.length > 1).forEach(([id, entries]) => {
    entries.forEach(e => {
      warnings.push({
        category: 'Duplicate Import ID',
        template: templateName,
        ticket_number: e.ticket,
        field: 'import_id',
        value: id,
        detail: `Appears ${entries.length}× in ${templateName}`,
      });
    });
  });
  return warnings;
};

const findQuantityIssues = (rows, templateName) => {
  const warnings = [];
  const qtyField = rows[0] && 'quantity' in rows[0] ? 'quantity' : 'source.quantity';
  rows.forEach(r => {
    const raw = r[qtyField];
    const num = Number(raw);
    const ticket = r.ticket_number || r['source.ticket_number'] || '';
    if (raw === '' || raw === null || raw === undefined) {
      warnings.push({ category: 'Invalid Quantity', template: templateName, ticket_number: ticket, field: qtyField, value: '(blank)', detail: 'Quantity is missing' });
    } else if (isNaN(num)) {
      warnings.push({ category: 'Invalid Quantity', template: templateName, ticket_number: ticket, field: qtyField, value: String(raw), detail: 'Quantity is not a number' });
    } else if (num < 0) {
      warnings.push({ category: 'Invalid Quantity', template: templateName, ticket_number: ticket, field: qtyField, value: String(raw), detail: 'Negative quantity' });
    } else if (num === 0) {
      warnings.push({ category: 'Invalid Quantity', template: templateName, ticket_number: ticket, field: qtyField, value: '0', detail: 'Zero quantity' });
    }
  });
  return warnings;
};

const findMissingFields = (rows, templateName, requiredFields) => {
  const warnings = [];
  rows.forEach(r => {
    const ticket = r.ticket_number || r['source.ticket_number'] || '';
    requiredFields.forEach(f => {
      const v = r[f];
      if (v === '' || v === null || v === undefined) {
        warnings.push({ category: 'Missing Field', template: templateName, ticket_number: ticket, field: f, value: '(blank)', detail: `Required field "${f}" is empty` });
      }
    });
  });
  return warnings;
};

// ---------- core transform ----------
const processWorkbook = (wb) => {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase() !== 'duplicates') || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, cellDates: true, defval: null });
  const dataRows = rows.slice(2).filter(r => {
    if (!r || r.length === 0) return false;
    const ticket = r[col('E')];
    return ticket !== null && ticket !== undefined && String(ticket).trim() !== '';
  });
  const g = (r, L) => r[col(L)];
  const isIn  = (r) => trimUpper(g(r,'A')) === 'INBOUND';
  const isOut = (r) => trimUpper(g(r,'A')) === 'OUTBOUND';
  const isReg = (r) => trimUpper(g(r,'D')) === 'REGULAR';
  const isXfr = (r) => trimUpper(g(r,'D')) === 'TRANSFER';
  const isCommodity = (r) => isCornOrBeans(r, g);

  // Count excluded rows for reporting
  const excludedCommodity = dataRows.filter(r => !isCommodity(r)).length;

  // ---- Inbound (INBOUND + REGULAR + corn/soybeans) ----
  const inbound = dataRows.filter(r => isIn(r) && isReg(r) && isCommodity(r)).map(r => ({
    import_id: clean(g(r,'G')) + clean(g(r,'F')) + clean(g(r,'BV')),
    type: commodity(g(r,'O')),
    date: fmtDate(g(r,'M')),
    quantity: g(r,'BY') ?? '',
    quantity_unit: 'bu',
    ticket_number: trimStr(g(r,'E')),
    source_id: trimStr(g(r,'BV')),
    source_name: trimStr(g(r,'J')),
    target_id: trimStr(g(r,'B')),
    target_enterprise_id: 'grain-coop',
  }));

  // ---- Outbound (OUTBOUND + REGULAR + corn/soybeans) ----
  const outbound = dataRows.filter(r => isOut(r) && isReg(r) && isCommodity(r)).map(r => ({
    import_id: clean(g(r,'G')) + clean(g(r,'F')) + clean(g(r,'BV')),
    type: commodity(g(r,'O')),
    date: fmtDate(g(r,'M')),
    quantity: g(r,'BY') ?? '',
    quantity_unit: 'bu',
    ticket_number: trimStr(g(r,'E')),
    source_id: trimStr(g(r,'B')),
    source_name: trimStr(g(r,'C')),
    source_enterprise_id: 'grain-coop',
    target_id: trimStr(g(r,'I')),
    target_name: trimStr(g(r,'J')),
  }));

  // ---- Internal Transfer (OUTBOUND + TRANSFER + corn/soybeans, matched via BOL) ----
  const bolMap = {};
  dataRows.filter(r => isIn(r) && isXfr(r) && isCommodity(r)).forEach(r => {
    const bol = trimUpper(g(r,'W'));
    if (bol) bolMap[bol] = g(r,'E');
  });

  const bolWarnings = [];
  const transfer = dataRows.filter(r => isOut(r) && isXfr(r) && isCommodity(r)).map(r => {
    const bolRaw = g(r,'W');
    const bol = trimUpper(bolRaw);
    const matched = bol ? bolMap[bol] : null;
    if (!matched) {
      bolWarnings.push({
        category: 'BOL Unmatched',
        template: 'Internal Transfer',
        ticket_number: trimStr(g(r,'E')),
        field: 'target.ticket_number',
        value: bolRaw ? String(bolRaw).trim() : '(blank)',
        detail: bol ? 'No matching inbound-transfer with this BOL' : 'BOL blank on outbound-transfer',
      });
    }
    return {
      import_id: trimStr(g(r,'G')),
      type: commodity(g(r,'O')),
      'source.id': trimStr(g(r,'B')),
      'source.enterprise_id': 'grain-coop',
      'source.ticket_number': trimStr(g(r,'E')),
      'source.date': fmtDate(g(r,'M')),
      'source.quantity': g(r,'BY') ?? '',
      'source.quantity_unit': 'bu',
      'target.id': trimStr(g(r,'AU')),
      'target.enterprise_id': 'grain-coop',
      'target.ticket_number': trimStr(matched),
      'target.date': fmtDate(g(r,'Z')),
      'target.quantity': g(r,'BY') ?? '',
      'target.quantity_unit': 'bu',
    };
  });

  // ---- External Partner Transfer (OUTBOUND + REGULAR + corn/soybeans → ext-partner) ----
  const extTransfer = dataRows.filter(r => isOut(r) && isReg(r) && isCommodity(r)).map(r => {
    const bolRaw = g(r,'W');
    const bolClean = bolRaw ? String(bolRaw).trim() : '';
    if (!bolClean) {
      bolWarnings.push({
        category: 'BOL Blank',
        template: 'External Partner Transfer',
        ticket_number: trimStr(g(r,'E')),
        field: 'target.ticket_number',
        value: '(blank)',
        detail: 'BOL blank — target.ticket_number cannot be populated',
      });
    }
    return {
      import_id: clean(g(r,'G')) + clean(g(r,'F')) + clean(g(r,'BV')),
      type: commodity(g(r,'O')),
      'source.enterprise_id': 'grain-coop',
      'source.id': trimStr(g(r,'B')),
      'source.ticket_number': trimStr(g(r,'E')),
      'source.date': fmtDate(g(r,'M')),
      'source.quantity': g(r,'BY') ?? '',
      'source.quantity_unit': 'bu',
      'target.enterprise_id': 'ext-partner',
      'target.id': 'ext-partner-01',
      'target.ticket_number': bolClean,
      'target.date': fmtDate(g(r,'Z')),
      'target.quantity': g(r,'BY') ?? '',
      'target.quantity_unit': 'bu',
    };
  });

  // ---- Validation warnings ----
  const inReq   = ['import_id','type','date','ticket_number','source_id','target_id'];
  const outReq  = ['import_id','type','date','ticket_number','source_id','target_id'];
  const xfrReq  = ['import_id','type','source.id','source.ticket_number','source.date','target.id'];
  const extReq  = ['import_id','type','source.id','source.ticket_number','source.date'];

  const allWarnings = [
    ...bolWarnings,
    ...findDuplicateImportIds(inbound,     'Inbound'),
    ...findDuplicateImportIds(outbound,    'Outbound'),
    ...findDuplicateImportIds(transfer,    'Internal Transfer'),
    ...findDuplicateImportIds(extTransfer, 'External Partner Transfer'),
    ...findQuantityIssues(inbound,         'Inbound'),
    ...findQuantityIssues(outbound,        'Outbound'),
    ...findQuantityIssues(transfer,        'Internal Transfer'),
    ...findQuantityIssues(extTransfer,     'External Partner Transfer'),
    ...findMissingFields(inbound,          'Inbound',                    inReq),
    ...findMissingFields(outbound,         'Outbound',                   outReq),
    ...findMissingFields(transfer,         'Internal Transfer',           xfrReq),
    ...findMissingFields(extTransfer,      'External Partner Transfer',   extReq),
  ];

  // Summarize by category — always include Duplicate Import ID even if 0
  const warningSummary = {
    'Duplicate Import ID': { count: 0, templates: {} },
  };
  allWarnings.forEach(w => {
    if (!warningSummary[w.category]) warningSummary[w.category] = { count: 0, templates: {} };
    warningSummary[w.category].count++;
    if (!warningSummary[w.category].templates[w.template]) warningSummary[w.category].templates[w.template] = 0;
    warningSummary[w.category].templates[w.template]++;
  });

  // Unique import_id counts per template
  const uniqueIds = {
    Inbound:                    new Set(inbound.map(r => r.import_id).filter(Boolean)).size,
    Outbound:                   new Set(outbound.map(r => r.import_id).filter(Boolean)).size,
    'Internal Transfer':        new Set(transfer.map(r => r.import_id).filter(Boolean)).size,
    'External Partner Transfer':new Set(extTransfer.map(r => r.import_id).filter(Boolean)).size,
  };
  const totalUniqueIds = Object.values(uniqueIds).reduce((a, b) => a + b, 0);

  return {
    inbound, outbound, transfer, extTransfer,
    warnings: allWarnings, warningSummary,
    uniqueIds, totalUniqueIds,
    excludedCommodity,
    sheetName, totalRows: dataRows.length,
  };
};

const IN_COLS   = ['import_id','type','date','quantity','quantity_unit','ticket_number','source_id','source_name','target_id','target_enterprise_id'];
const OUT_COLS  = ['import_id','type','date','quantity','quantity_unit','ticket_number','source_id','source_name','source_enterprise_id','target_id','target_name'];
const XFER_COLS = ['import_id','type','source.id','source.enterprise_id','source.ticket_number','source.date','source.quantity','source.quantity_unit','target.id','target.enterprise_id','target.ticket_number','target.date','target.quantity','target.quantity_unit'];
const EXT_COLS  = ['import_id','type','source.enterprise_id','source.id','source.ticket_number','source.date','source.quantity','source.quantity_unit','target.enterprise_id','target.id','target.ticket_number','target.date','target.quantity','target.quantity_unit'];
const WARN_COLS = ['category','template','ticket_number','field','value','detail'];

const toCsv = (rows, cols) => {
  const esc = (v) => {
    const s = String(v ?? '');
    return /[,"\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return [cols.join(','), ...rows.map(r => cols.map(c => esc(r[c])).join(','))].join('\n');
};

const download = (filename, content) => {
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
};

const todayStamp = () => {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
};

// ---------- Warning category config ----------
const WARN_CATEGORY_CONFIG = {
  'Duplicate Import ID': { icon: Copy,          color: 'red',    bg: 'bg-red-50',    border: 'border-red-200',    text: 'text-red-700',    badge: 'bg-red-100 text-red-800' },
  'Invalid Quantity':    { icon: Hash,           color: 'orange', bg: 'bg-orange-50', border: 'border-orange-200', text: 'text-orange-700', badge: 'bg-orange-100 text-orange-800' },
  'Missing Field':       { icon: AlertCircle,    color: 'amber',  bg: 'bg-amber-50',  border: 'border-amber-200',  text: 'text-amber-700',  badge: 'bg-amber-100 text-amber-800' },
  'BOL Unmatched':       { icon: AlertTriangle,  color: 'violet', bg: 'bg-violet-50', border: 'border-violet-200', text: 'text-violet-700', badge: 'bg-violet-100 text-violet-800' },
  'BOL Blank':           { icon: Info,           color: 'blue',   bg: 'bg-blue-50',   border: 'border-blue-200',   text: 'text-blue-700',   badge: 'bg-blue-100 text-blue-800' },
};

// ---------- UI ----------
export default function GrainElevatorTransformer() {
  const [state, setState] = useState('idle');
  const [result, setResult] = useState(null);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [activeTab, setActiveTab] = useState('inbound');
  const [dragging, setDragging] = useState(false);
  const [warnFilter, setWarnFilter] = useState('all');
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    if (!file) return;
    setFileName(file.name);
    setState('processing');
    setError('');
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true });
      const out = processWorkbook(wb);
      setResult(out);
      setState('done');
      setActiveTab(out.inbound.length ? 'inbound' : out.outbound.length ? 'outbound' : 'transfer');
    } catch (e) {
      console.error(e);
      setError(e.message || 'Failed to process file.');
      setState('error');
    }
  };

  const reset = () => {
    setState('idle'); setResult(null); setFileName(''); setError(''); setWarnFilter('all');
    if (inputRef.current) inputRef.current.value = '';
  };

  const downloadAll = () => {
    const stamp = todayStamp();
    download(`IN_SCALE_TICKET_${stamp}.csv`,      toCsv(result.inbound,     IN_COLS));
    download(`OUT_SCALE_TICKET_${stamp}.csv`,     toCsv(result.outbound,    OUT_COLS));
    download(`TRANSFER_INTERNAL_${stamp}.csv`,    toCsv(result.transfer,    XFER_COLS));
    download(`TRANSFER_EXT_PARTNER_${stamp}.csv`, toCsv(result.extTransfer, EXT_COLS));
    if (result.warnings.length) download(`WARNINGS_${stamp}.csv`, toCsv(result.warnings, WARN_COLS));
  };

  const tabs = result ? [
    { id: 'inbound',     label: 'Inbound',                  count: result.inbound.length,     cols: IN_COLS,   rows: result.inbound,     file: `IN_SCALE_TICKET_${todayStamp()}.csv` },
    { id: 'outbound',    label: 'Outbound',                 count: result.outbound.length,    cols: OUT_COLS,  rows: result.outbound,    file: `OUT_SCALE_TICKET_${todayStamp()}.csv` },
    { id: 'transfer',    label: 'Internal Transfer',        count: result.transfer.length,    cols: XFER_COLS, rows: result.transfer,    file: `TRANSFER_INTERNAL_${todayStamp()}.csv` },
    { id: 'extTransfer', label: 'Ext. Partner Transfer',   count: result.extTransfer.length, cols: EXT_COLS,  rows: result.extTransfer, file: `TRANSFER_EXT_PARTNER_${todayStamp()}.csv`, accent: true },
    ...(result.warnings.length ? [{ id: 'warnings', label: 'Warnings', count: result.warnings.length, cols: WARN_COLS, rows: result.warnings, file: `WARNINGS_${todayStamp()}.csv`, warn: true }] : []),
  ] : [];

  const activeTabObj = tabs.find(t => t.id === activeTab);
  const extMissingBol = result ? result.extTransfer.filter(r => !r['target.ticket_number']).length : 0;

  // Filtered warnings for table display
  const filteredWarnings = result && result.warnings
    ? (warnFilter === 'all' ? result.warnings : result.warnings.filter(w => w.category === warnFilter))
    : [];

  const totalIncluded = result ? result.inbound.length + result.outbound.length + result.transfer.length + result.extTransfer.length : 0;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="max-w-6xl mx-auto px-6 py-10">
        {/* Header */}
        <div className="flex items-center gap-3 mb-2">
          <div className="w-10 h-10 rounded-lg bg-slate-900 flex items-center justify-center">
            <FileSpreadsheet className="w-5 h-5 text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-semibold tracking-tight">Grain Elevator Data Transformer</h1>
            <p className="text-sm text-slate-500">Raw scale ticket report → Balance API–ready CSVs</p>
          </div>
        </div>

        {/* Upload */}
        {state === 'idle' && (
          <div
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => { e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files?.[0]); }}
            className={`mt-8 border-2 border-dashed rounded-xl p-12 text-center transition-colors ${dragging ? 'border-slate-900 bg-white' : 'border-slate-300 bg-white hover:border-slate-400'}`}
          >
            <Upload className="w-10 h-10 mx-auto text-slate-400 mb-4" />
            <p className="text-slate-700 font-medium mb-1">Drop your ticket report here</p>
            <p className="text-sm text-slate-500 mb-5">Accepts .xlsx and .xls files</p>
            <button
              onClick={() => inputRef.current?.click()}
              className="px-5 py-2.5 bg-slate-900 text-white rounded-lg font-medium hover:bg-slate-800 transition-colors text-sm"
            >
              Choose file
            </button>
            <input
              ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden"
              onChange={(e) => handleFile(e.target.files?.[0])}
            />
          </div>
        )}

        {state === 'processing' && (
          <div className="mt-8 bg-white rounded-xl p-12 text-center border border-slate-200">
            <Loader2 className="w-8 h-8 mx-auto text-slate-600 animate-spin mb-3" />
            <p className="text-slate-700 font-medium">Processing {fileName}…</p>
          </div>
        )}

        {state === 'error' && (
          <div className="mt-8 bg-white rounded-xl p-8 border border-red-200">
            <div className="flex items-start gap-3">
              <AlertTriangle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
              <div className="flex-1">
                <p className="font-medium text-red-900">Something went wrong</p>
                <p className="text-sm text-red-700 mt-1">{error}</p>
                <button onClick={reset} className="mt-4 px-4 py-2 bg-red-600 text-white rounded-lg text-sm font-medium hover:bg-red-700">
                  Try another file
                </button>
              </div>
            </div>
          </div>
        )}

        {state === 'done' && result && (
          <>
            {/* Summary bar */}
            <div className="mt-8 bg-white rounded-xl border border-slate-200 p-5 flex flex-wrap items-center gap-4">
              <div className="flex items-center gap-2 text-emerald-700">
                <CheckCircle2 className="w-5 h-5" />
                <span className="font-medium text-sm">Processed {fileName}</span>
              </div>
              <div className="text-sm text-slate-500">
                {result.totalRows} source rows <ArrowRight className="inline w-3 h-3 mx-1" />
                {' '}{result.inbound.length} in · {result.outbound.length} out · {result.transfer.length} internal xfer · {result.extTransfer.length} ext. partner xfer
                {result.excludedCommodity > 0 && (
                  <span className="text-slate-400"> · <Filter className="inline w-3 h-3 -mt-0.5" /> {result.excludedCommodity} excluded (non-corn/soybean)</span>
                )}
                {result.warnings.length > 0 && <span className="text-amber-600 font-medium"> · {result.warnings.length} warning{result.warnings.length === 1 ? '' : 's'}</span>}
              </div>
              <div className="ml-auto flex gap-2">
                <button onClick={reset} className="px-3 py-2 text-sm text-slate-600 hover:text-slate-900 flex items-center gap-1.5">
                  <RotateCcw className="w-4 h-4" /> New file
                </button>
                <button onClick={downloadAll} className="px-4 py-2 bg-slate-900 text-white rounded-lg text-sm font-medium hover:bg-slate-800 flex items-center gap-2">
                  <Download className="w-4 h-4" /> Download all CSVs
                </button>
              </div>
            </div>

            {/* Warnings dashboard — always shown */}
            <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-5">
              {Object.entries(result.warningSummary).map(([cat, data]) => {
                const isClean = cat === 'Duplicate Import ID' && data.count === 0;
                const cfg = isClean
                  ? { icon: CheckCircle2, bg: 'bg-emerald-50', border: 'border-emerald-200', text: 'text-emerald-700', badge: 'bg-emerald-100 text-emerald-800' }
                  : (WARN_CATEGORY_CONFIG[cat] || WARN_CATEGORY_CONFIG['Missing Field']);
                const Icon = cfg.icon;
                return (
                  <button
                    key={cat}
                    onClick={() => { if (data.count > 0) { setActiveTab('warnings'); setWarnFilter(warnFilter === cat ? 'all' : cat); } }}
                    className={`${cfg.bg} ${cfg.border} border rounded-lg p-3 text-left transition-all ${data.count > 0 ? 'hover:shadow-sm cursor-pointer' : 'cursor-default'} ${warnFilter === cat && activeTab === 'warnings' ? 'ring-2 ring-offset-1 ring-slate-400' : ''}`}
                  >
                    <div className="flex items-center gap-1.5 mb-1">
                      <Icon className={`w-3.5 h-3.5 ${cfg.text}`} />
                      <span className={`text-xs font-semibold ${cfg.text}`}>{data.count}</span>
                    </div>
                    <p className={`text-xs font-medium ${cfg.text} leading-tight`}>{isClean ? 'Duplicate Import IDs' : cat}</p>
                    <p className="text-[10px] text-slate-500 mt-1 leading-tight">
                      {isClean
                        ? Object.entries(result.uniqueIds).map(([t, c]) => `${t}: ${c}`).join(' · ')
                        : Object.entries(data.templates).map(([t, c]) => `${t}: ${c}`).join(' · ')
                      }
                    </p>
                    {isClean && (
                      <p className="text-[10px] text-emerald-600 font-medium mt-0.5">{result.totalUniqueIds} unique IDs — all clean</p>
                    )}
                  </button>
                );
              })}
            </div>

            {/* External partner BOL notice */}
            {extMissingBol > 0 && (
              <div className="mt-3 bg-blue-50 border border-blue-200 rounded-lg px-4 py-3 flex items-start gap-2.5">
                <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
                <p className="text-sm text-blue-800">
                  <span className="font-medium">{extMissingBol} External Partner transfer row{extMissingBol === 1 ? '' : 's'}</span> have blank BOL — <code className="text-xs bg-blue-100 px-1 py-0.5 rounded">target.ticket_number</code> left empty. These will need the partner's inbound ticket numbers once available.
                </p>
              </div>
            )}

            {/* Tabs */}
            <div className="mt-4 bg-white rounded-xl border border-slate-200 overflow-hidden">
              <div className="border-b border-slate-200 flex items-center overflow-x-auto">
                {tabs.map(t => (
                  <button
                    key={t.id}
                    onClick={() => { setActiveTab(t.id); if (t.id !== 'warnings') setWarnFilter('all'); }}
                    className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors whitespace-nowrap ${activeTab === t.id
                      ? (t.warn ? 'border-amber-500 text-amber-700' : t.accent ? 'border-indigo-500 text-indigo-700' : 'border-slate-900 text-slate-900')
                      : 'border-transparent text-slate-500 hover:text-slate-700'}`}
                  >
                    {t.warn && <AlertTriangle className="w-3.5 h-3.5 inline -mt-0.5 mr-1" />}
                    {t.label}
                    <span className={`ml-1.5 px-1.5 py-0.5 rounded text-xs ${
                      t.warn ? 'bg-amber-100 text-amber-800' :
                      t.accent ? 'bg-indigo-100 text-indigo-700' :
                      'bg-slate-100 text-slate-600'
                    }`}>
                      {t.count}
                    </span>
                  </button>
                ))}
                <div className="ml-auto pr-4 shrink-0">
                  {activeTabObj && activeTabObj.count > 0 && (
                    <button
                      onClick={() => download(activeTabObj.file, toCsv(activeTab === 'warnings' ? filteredWarnings : activeTabObj.rows, activeTabObj.cols))}
                      className="px-3 py-1.5 text-xs font-medium text-slate-700 hover:text-slate-900 flex items-center gap-1.5 border border-slate-200 rounded-md hover:bg-slate-50"
                    >
                      <Download className="w-3.5 h-3.5" /> {activeTabObj.file}
                    </button>
                  )}
                </div>
              </div>

              {/* Warning filter pills (shown when Warnings tab active) */}
              {activeTab === 'warnings' && result.warnings.length > 0 && (
                <div className="px-4 py-2 bg-slate-50 border-b border-slate-200 flex items-center gap-2 overflow-x-auto">
                  <span className="text-xs text-slate-500 shrink-0">Filter:</span>
                  <button
                    onClick={() => setWarnFilter('all')}
                    className={`px-2.5 py-1 rounded-full text-xs font-medium transition-colors ${warnFilter === 'all' ? 'bg-slate-900 text-white' : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'}`}
                  >
                    All ({result.warnings.length})
                  </button>
                  {Object.entries(result.warningSummary).map(([cat, data]) => {
                    const cfg = WARN_CATEGORY_CONFIG[cat] || WARN_CATEGORY_CONFIG['Missing Field'];
                    return (
                      <button
                        key={cat}
                        onClick={() => setWarnFilter(warnFilter === cat ? 'all' : cat)}
                        className={`px-2.5 py-1 rounded-full text-xs font-medium transition-colors whitespace-nowrap ${warnFilter === cat ? cfg.badge : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'}`}
                      >
                        {cat} ({data.count})
                      </button>
                    );
                  })}
                </div>
              )}

              {/* Preview table */}
              <div className="overflow-auto max-h-[520px]">
                {(() => {
                  const displayRows = activeTab === 'warnings' ? filteredWarnings : (activeTabObj ? activeTabObj.rows : []);
                  const displayCols = activeTabObj ? activeTabObj.cols : [];
                  if (displayRows.length > 0) {
                    return (
                      <>
                        <table className="w-full text-xs">
                          <thead className="bg-slate-50 sticky top-0">
                            <tr>
                              {displayCols.map(c => (
                                <th key={c} className="px-3 py-2 text-left font-medium text-slate-600 whitespace-nowrap border-b border-slate-200">{c}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {displayRows.slice(0, 500).map((r, i) => (
                              <tr key={i} className={`${i % 2 ? 'bg-slate-50/50' : ''} ${activeTab === 'warnings' ? 'hover:bg-amber-50/50' : ''}`}>
                                {displayCols.map(c => {
                                  const val = String(r[c] ?? '');
                                  const isEmpty = activeTab === 'extTransfer' && c === 'target.ticket_number' && !val;
                                  const isCategory = activeTab === 'warnings' && c === 'category';
                                  const cfg = isCategory ? WARN_CATEGORY_CONFIG[val] : null;
                                  return (
                                    <td key={c} className={`px-3 py-1.5 whitespace-nowrap border-b border-slate-100 ${
                                      isEmpty ? 'text-amber-500 italic' :
                                      cfg ? cfg.text + ' font-medium' :
                                      'text-slate-700'
                                    }`}>
                                      {isEmpty ? '(pending)' : val}
                                    </td>
                                  );
                                })}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        {displayRows.length > 500 && (
                          <div className="p-3 text-center text-xs text-slate-500 bg-slate-50 border-t border-slate-200">
                            Showing first 500 of {displayRows.length} rows. Full data in downloaded CSV.
                          </div>
                        )}
                      </>
                    );
                  }
                  return <div className="p-10 text-center text-slate-400 text-sm">No rows in this category.</div>;
                })()}
              </div>
            </div>

            {/* Footer note */}
            <p className="mt-4 text-xs text-slate-500">
              Source sheet: <span className="font-mono">{result.sheetName}</span> · Dates MM/DD/YYYY · Quantities are net (Column BY) · Commodity filter: corn + soybeans only · Ext. Partner target: enterprise=ext-partner, id=ext-partner-01
            </p>
          </>
        )}
      </div>
    </div>
  );
}

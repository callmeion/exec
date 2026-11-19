import { useEffect, useMemo, useState, type ReactNode } from 'react'
import * as XLSX from 'xlsx'
import { ResponsiveContainer, LineChart, Line, BarChart, Bar, XAxis, YAxis, Tooltip, Legend, CartesianGrid, Treemap, PieChart, Pie, Cell } from 'recharts'
import { useReactTable, getCoreRowModel, getSortedRowModel, flexRender, type ColumnDef, type SortingState } from '@tanstack/react-table'
import type { Dataset } from './types'
import { parseWorkbook, parseBuffer } from './lib/parser'
import { decryptData } from './lib/crypto'
import { DATA_URL } from './config'
import { F0, type F, type FKey, scope, monthly, inPeriod, groupBy, topN, alerts, forecast, lbl, nextKey, derive, type MD, type Group } from './lib/metrics'

const COLORS = ['#0d9488', '#0ea5e9', '#f59e0b', '#ef4444', '#6366f1', '#8b5cf6', '#84cc16', '#ec4899']
const ACC:Record<string,string> = { sales:'#0d9488', payroll:'#f59e0b', ot:'#ef4444', payrollOt:'#f59e0b', labor:'#6366f1', hc:'#0ea5e9', salesPerHead:'#0d9488', salesPerBaht:'#0d9488', cpp:'#8b5cf6', turnover:'#ef4444' }
const LVL_TH = { RED:'เฝ้าระวังสูง', YELLOW:'ควรติดตาม', GREEN:'ดี', NONE:'ปกติ' } as const
const n0 = (v:number) => Math.round(v).toLocaleString('en-US')
const compact = (v:number) => Intl.NumberFormat('en', { notation: 'compact', maximumFractionDigits: 2 }).format(v)
const p1 = (v:number) => (v * 100).toFixed(1) + '%'
const x2 = (v:number) => v.toFixed(2)
const MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
const PAGES = ['ภาพรวมผู้บริหาร','ยอดขาย','ค่าแรง','ผลิตภาพ','ทรัพยากรบุคคล','พยากรณ์','ศูนย์แจ้งเตือน']

/* ---------- small UI pieces ---------- */
const Card = ({ title, children, className = '' }:{ title?:string; children:ReactNode; className?:string }) => (
  <div className={`rounded-2xl border border-slate-200/70 bg-white p-5 shadow-[0_1px_3px_rgba(15,23,42,.05)] dark:border-slate-800 dark:bg-slate-900 ${className}`}>
    {title && <h3 className="mb-4 flex items-center gap-2 text-sm font-semibold text-slate-700 dark:text-slate-200"><span className="h-3.5 w-1 rounded-full bg-teal-500" />{title}</h3>}{children}
  </div>
)
const Sel = ({ label, value, opts, on }:{ label:string; value:string; opts:(string|number)[]; on:(v:string)=>void }) => (
  <label className="flex min-w-36 flex-col text-xs font-medium text-slate-500">{label}
    <select value={value} onChange={e => on(e.target.value)} className={`mt-1 rounded-lg border px-2.5 py-1.5 text-sm text-slate-800 dark:text-slate-100 ${value ? 'border-teal-500 bg-teal-50 dark:bg-teal-500/10' : 'border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800'}`}>
      <option value="">ทั้งหมด</option>{opts.map(o => <option key={o} value={o}>{o}</option>)}
    </select></label>
)
const Chart = ({ children, h = 280 }:{ children:any; h?:number }) => <div style={{ height: h }}><ResponsiveContainer width="100%" height="100%">{children}</ResponsiveContainer></div>
const grid = <CartesianGrid strokeDasharray="3 3" stroke="#94a3b833" />
const ax = { tick: { fontSize: 12, fill: '#64748b' }, axisLine: false, tickLine: false } as const
const tip = <Tooltip cursor={{ fill: '#94a3b81a' }} formatter={(v:any) => (typeof v === 'number' ? n0(v) : v)} contentStyle={{ borderRadius: 12, border: '1px solid #e2e8f0', boxShadow: '0 8px 24px rgba(15,23,42,.12)', fontSize: 13 }} />
const TreeCell = (p:any) => { const { x, y, width, height, index, name } = p; if (width < 4 || height < 4) return null
  return <g><rect x={x} y={y} width={width} height={height} rx={8} fill={COLORS[index % COLORS.length]} fillOpacity={.88} stroke="#fff" />
    {width > 70 && height > 28 && <text x={x + 9} y={y + 21} fill="#fff" fontSize={12} fontWeight={600}>{String(name).slice(0, Math.floor(width / 8))}</text>}</g> }

function KPI({ label, cur, prev, yoy, fmt, up, note, hint, spark, accent = '#0d9488' }:{ hint?:string; label:string; cur:number; prev?:number; yoy?:number; fmt:(v:number)=>string; up:boolean; note?:string; spark?:{v:number}[]; accent?:string }) {
  const d = (b?:number) => (b ? (cur - b) / Math.abs(b) : null)
  const Badge = ({ t, v }:{ t:string; v:number|null }) => {
    if (v === null) return <span className="rounded-full bg-slate-100 px-2 py-0.5 text-slate-400 dark:bg-slate-800">{t} –</span>
    const good = up ? v >= 0 : v <= 0
    return <span className={`rounded-full px-2 py-0.5 font-medium ${good ? 'bg-emerald-50 text-emerald-700 dark:bg-emerald-500/15 dark:text-emerald-300' : 'bg-red-50 text-red-700 dark:bg-red-500/15 dark:text-red-300'}`}>{t} {v >= 0 ? '▲' : '▼'} {Math.abs(v * 100).toFixed(1)}%</span>
  }
  return <Card className="relative overflow-hidden !p-4"><div className="absolute inset-y-0 left-0 w-1" style={{ background: accent }} />
    <div className="text-xs font-medium text-slate-500">{label}{note && <span className="ml-1 rounded bg-slate-100 px-1 text-[10px] text-slate-400 dark:bg-slate-800">{note}</span>}</div>
    <div className="mt-1 text-[26px] font-bold leading-tight tracking-tight">{cur === undefined ? '–' : fmt(cur)}</div>
    {spark && spark.length > 1 && <div className="-mx-1 h-9"><ResponsiveContainer width="100%" height="100%"><LineChart data={spark}><Line dot={false} strokeWidth={2.5} dataKey="v" stroke={accent} isAnimationActive={false} /></LineChart></ResponsiveContainer></div>}
    <div className="mt-1 flex flex-wrap gap-1.5 text-[11px]"><Badge t="เดือนก่อน" v={d(prev)} /><Badge t="ปีก่อน" v={d(yoy)} /></div>
    {hint && <div className="mt-2 text-[11px] leading-snug text-slate-400">{hint}</div>}</Card>
}

function DataTable<T>({ data, columns }:{ data:T[]; columns:ColumnDef<T, any>[] }) {
  const [sorting, setSorting] = useState<SortingState>([])
  const t = useReactTable({ data, columns, state: { sorting }, onSortingChange: setSorting, getCoreRowModel: getCoreRowModel(), getSortedRowModel: getSortedRowModel() })
  return <div className="overflow-x-auto"><table className="w-full text-sm">
    <thead>{t.getHeaderGroups().map(hg => <tr key={hg.id} className="border-b border-slate-200 text-left text-xs text-slate-500 dark:border-slate-700">
      {hg.headers.map(h => <th key={h.id} className="cursor-pointer px-2 py-1.5" onClick={h.column.getToggleSortingHandler()}>{flexRender(h.column.columnDef.header, h.getContext())}{{ asc: ' ↑', desc: ' ↓' }[h.column.getIsSorted() as string] ?? ''}</th>)}</tr>)}</thead>
    <tbody>{t.getRowModel().rows.map(r => <tr key={r.id} className="border-b border-slate-100 dark:border-slate-800">{r.getVisibleCells().map(c => <td key={c.id} className="px-2 py-1.5">{flexRender(c.column.columnDef.cell, c.getContext())}</td>)}</tr>)}</tbody></table></div>
}

const download = (blob:Blob, name:string) => { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click() }
const uniq = (a:any[]) => [...new Set(a.filter(Boolean))].sort() as string[]

/* ---------- App ---------- */
export default function App() {
  const [ds, setDs] = useState<Dataset | null>(null)
  const [f, setF] = useState<F>(F0)
  const [page, setPage] = useState(0)
  const [dark, setDark] = useState(() => localStorage.getItem('theme') === 'dark')
  const [busy, setBusy] = useState(false)
  const [h, setH] = useState(6)
  const [showF, setShowF] = useState(true)
  const [rankKey, setRankKey] = useState<'labor'|'salesPerHead'|'salesPerBaht'|'cpp'>('labor')

  useEffect(() => { document.documentElement.classList.toggle('dark', dark); localStorage.setItem('theme', dark ? 'dark' : 'light') }, [dark])
  const [enc, setEnc] = useState<ArrayBuffer | null>(null)
  const [pw, setPw] = useState('')
  const [err, setErr] = useState('')
  const [checking, setChecking] = useState(true)

  const unlock = async (password:string, buf = enc) => {
    if (!buf) return; setErr(''); setBusy(true)
    try { const d = await parseBuffer(await decryptData(buf, password), 'ข้อมูลล่าสุด'); setDs(d); sessionStorage.setItem('k', password) }
    catch { sessionStorage.removeItem('k'); setErr('รหัสผ่านไม่ถูกต้อง') }
    setBusy(false)
  }
  const logout = () => { sessionStorage.removeItem('k'); setDs(null); setPw(''); setF(F0) }

  useEffect(() => { (async () => {   // ดึงไฟล์ข้อมูลที่เข้ารหัสจากโฟลเดอร์ data/ ทุกครั้งที่เปิดหน้า
    try {
      const r = await fetch(DATA_URL, { cache: 'no-store' })
      if (r.ok && !(r.headers.get('content-type') || '').includes('html')) {
        const buf = await r.arrayBuffer(); setEnc(buf)
        const saved = sessionStorage.getItem('k'); if (saved) await unlock(saved, buf)   // ไม่ต้องกรอกซ้ำเมื่อรีเฟรชในแท็บเดิม
      }
    } catch { /* ไม่พบไฟล์ */ }
    setChecking(false)
  })() }, [])

  const onFile = async (file?:File) => {
    if (!file) return; setBusy(true)
    await new Promise(r => setTimeout(r, 30))   // let the spinner paint before the heavy parse
    const d = await parseWorkbook(file); setDs(d); setF(F0); setBusy(false)
  }

  const opts = useMemo(() => ds && ({
    year: uniq(ds.sales.map(r => r.Year_BE)), productStd: uniq(ds.sales.map(r => r.Product_Standard)), product: uniq(ds.sales.map(r => r.Product_Name_Raw)),
    market: uniq(ds.sales.map(r => r.Market)), customer: uniq(ds.sales.map(r => r.Customer)), dept: uniq([...ds.sales, ...ds.payroll].map(r => r.Department_Short)),
    line: uniq([...ds.sales, ...ds.payroll].map(r => r.Production_Line)), phase: uniq(ds.sales.map(r => r.Phase)), payCycle: uniq(ds.payroll.map(r => r.Pay_Cycle)),
  }), [ds])

  const calc = useMemo(() => {
    if (!ds) return null
    const sc = scope(ds, f), all = monthly(sc), pred = inPeriod(f)
    const vis = all.filter(m => pred(Math.floor(m.key / 100), m.key % 100))
    const cur = vis[vis.length - 1], idx = cur ? all.findIndex(m => m.key === cur.key) : -1
    return { sc, all, vis, pred, cur, prev: all[idx - 1], yoy: all.find(m => cur && m.key === cur.key - 100), al: alerts(all) }
  }, [ds, f])

  const setFilter = (k:FKey) => (v:string) => setF(p => ({ ...p, [k]: v }))
  const trend = (calc?.vis ?? []).map(m => ({ ...m, name: lbl(m.key) }))
  const exportData = (type:'xlsx'|'csv') => {
    const ws = XLSX.utils.json_to_sheet((calc?.vis ?? []).map(m => ({ Month: lbl(m.key), Sales: m.sales, Payroll: m.payroll, OT: m.ot, 'Labor %': +(m.labor * 100).toFixed(2), Headcount: m.hc, 'Turnover %': +(m.turnover * 100).toFixed(2) })))
    if (type === 'csv') return download(new Blob(['\ufeff' + XLSX.utils.sheet_to_csv(ws)], { type: 'text/csv' }), 'dashboard.csv')
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Monthly'); XLSX.writeFile(wb, 'dashboard.xlsx')
  }

  /* ---------- login / empty state ---------- */
  if (!ds || !calc || !opts) return (
    <div className="grid min-h-screen place-items-center bg-gradient-to-br from-teal-50 via-white to-sky-50 p-6 dark:from-slate-950 dark:via-slate-950 dark:to-slate-900">
      <div className="w-full max-w-sm rounded-3xl border border-slate-200/70 bg-white p-8 text-center shadow-xl dark:border-slate-800 dark:bg-slate-900">
        <div className="mx-auto mb-4 grid h-14 w-14 place-items-center rounded-2xl bg-gradient-to-br from-teal-500 to-teal-700 text-2xl text-white shadow-lg">📊</div>
        <h1 className="text-xl font-bold">แดชบอร์ดผู้บริหาร</h1>
        {checking ? <p className="mt-6 text-sm text-slate-500">กำลังโหลด…</p> : !enc ? <p className="mt-6 rounded-lg bg-amber-50 p-3 text-sm text-amber-700">ยังไม่พบไฟล์ข้อมูล (public/data/data.enc)<br />กรุณาติดต่อผู้ดูแลระบบ</p> : <>
          <p className="mt-1 text-sm text-slate-500">กรุณาใส่รหัสผ่านเพื่อเข้าถึงข้อมูล</p>
          <form className="mt-6 space-y-3" onSubmit={e => { e.preventDefault(); unlock(pw) }}>
            <input type="password" autoFocus value={pw} onChange={e => { setPw(e.target.value); setErr('') }} placeholder="รหัสผ่าน"
              className={`w-full rounded-xl border px-4 py-2.5 text-center text-base focus:outline-none focus:ring-2 focus:ring-teal-500/40 dark:bg-slate-800 ${err ? 'border-red-400' : 'border-slate-300 dark:border-slate-700'}`} />
            {err && <div className="text-sm text-red-500">{err}</div>}
            <button disabled={busy || !pw} className="w-full rounded-xl bg-teal-600 py-2.5 font-semibold text-white hover:bg-teal-700 disabled:opacity-50">{busy ? 'กำลังตรวจสอบ…' : 'เข้าสู่ระบบ'}</button>
          </form></>}
        <p className="mt-6 text-xs text-slate-400">🔒 ข้อมูลถูกเข้ารหัส และประมวลผลในเบราว์เซอร์ของคุณเท่านั้น</p>
      </div></div>)

  const { sc, all, vis, pred, cur, prev, yoy, al } = calc
  const alertNow = al.find(a => cur && a.key === cur.key)
  const lvlCls = { RED: 'bg-red-500', YELLOW: 'bg-amber-500', GREEN: 'bg-emerald-500', NONE: 'bg-slate-500' }
  const dept = groupBy(sc, 'dept', pred), line = groupBy(sc, 'line', pred)


  /* Each KPI uses the latest month in range where ITS inputs exist (payroll/workforce may cover fewer months than sales). */
  const need:Record<string,(m:MD)=>boolean> = { sales: m => m.sales > 0, payroll: m => m.payroll > 0, ot: m => m.payroll > 0, payrollOt: m => m.payroll > 0, hc: m => m.hc > 0, turnover: m => m.avgHc > 0,
    labor: m => m.sales > 0 && m.payroll > 0, salesPerBaht: m => m.sales > 0 && m.payroll > 0, cpp: m => m.qty > 0 && m.payroll > 0, salesPerHead: m => m.sales > 0 && m.hc > 0 }
  const kp = (k:string) => { const ok = need[k]; const c = [...vis].reverse().find(ok); if (!c) return null
    const i = all.findIndex(m => m.key === c.key), pv = all[i - 1], yy = all.find(m => m.key === c.key - 100)
    return { cur: c[k], prev: pv && ok(pv) ? pv[k] : undefined, yoy: yy && ok(yy) ? yy[k] : undefined, note: lbl(c.key) } }
  const HINT:Record<string,string> = { sales:'ยอดขายสุทธิ (Net Sales) ของเดือนล่าสุด', payroll:'ค่าแรงรวมก่อนหักภาษี/ประกัน (Gross Pay)', ot:'ค่าล่วงเวลารวม',
    payrollOt:'ค่าแรงปกติ + ค่าล่วงเวลา (ไม่รวมเบี้ยเลี้ยง)', labor:'ค่าแรงรวม หาร ยอดขาย (ยิ่งต่ำยิ่งดี)', hc:'พนักงานคงเหลือปลายเดือน (HC_End)', salesPerHead:'ยอดขาย หาร จำนวนพนักงาน',
    salesPerBaht:'ยอดขาย หาร ค่าแรง (ยิ่งสูงยิ่งดี)', cpp:'ค่าแรง หาร จำนวนชิ้นที่ขาย (ยิ่งต่ำยิ่งดี)', turnover:'จำนวนลาออก หาร พนักงานเฉลี่ย ((ต้นเดือน+ปลายเดือน)/2)' }
  const kpi = (l:string, k:string, fmt:(v:number)=>string, up:boolean) => { const r = kp(k); const spark = vis.filter(need[k]).map(m => ({ v: m[k] as number })); return <KPI key={k} spark={spark} accent={ACC[k]} label={l} cur={r?.cur} prev={r?.prev} yoy={r?.yoy} fmt={fmt} up={up} hint={HINT[k]} note={r ? (cur && r.note !== lbl(cur.key) ? r.note : undefined) : 'ไม่มีข้อมูล'} /> }

  const ins:{ t:'good'|'bad'|'warn'|'info'; s:string }[] = []
  const ks = kp('sales'); if (ks && ks.prev) { const v = (ks.cur - ks.prev) / ks.prev; ins.push({ t: v >= 0 ? 'good' : 'bad', s: `ยอดขายเดือน ${ks.note} ${v >= 0 ? 'เพิ่มขึ้น' : 'ลดลง'} ${Math.abs(v * 100).toFixed(1)}% จากเดือนก่อน` }) }
  const kl = kp('labor'); if (kl) ins.push({ t: kl.cur > .18 ? 'bad' : 'good', s: `ต้นทุนแรงงานเดือน ${kl.note} อยู่ที่ ${p1(kl.cur)} (${kl.cur > .18 ? 'สูงกว่า' : 'ต่ำกว่า'}เกณฑ์ 18%)` })
  const pmm = [...vis].reverse().find(m => m.payroll > 0); if (pmm) ins.push({ t: pmm.otPct > .15 ? 'bad' : 'good', s: `OT เดือน ${lbl(pmm.key)} คิดเป็น ${p1(pmm.otPct)} ของค่าแรง (เกณฑ์ 15%)` })
  const kt = kp('turnover'); if (kt) ins.push({ t: kt.cur > .10 ? 'bad' : kt.cur > .08 ? 'warn' : 'good', s: `อัตราลาออกเดือน ${kt.note} = ${p1(kt.cur)}` })
  if (alertNow && alertNow.level !== 'NONE') ins.push({ t: alertNow.level === 'RED' ? 'bad' : alertNow.level === 'YELLOW' ? 'warn' : 'good', s: `สถานะ ${lbl(alertNow.key)}: ${LVL_TH[alertNow.level]} — ${alertNow.reasons.join(', ')}` })
  const insTone = { good: 'bg-emerald-50 text-emerald-800 dark:bg-emerald-500/10 dark:text-emerald-300', bad: 'bg-red-50 text-red-800 dark:bg-red-500/10 dark:text-red-300', warn: 'bg-amber-50 text-amber-800 dark:bg-amber-500/10 dark:text-amber-300', info: 'bg-slate-100' }
  const insIcon = { good: '✓', bad: '!', warn: '▲', info: 'i' }

  /* ---------- pages ---------- */
  const summary = cur && <div className="space-y-4">{ins.length > 0 && <div className="flex flex-wrap gap-2">{ins.map((x, i) => <span key={i} className={`flex items-center gap-2 rounded-full px-3 py-1.5 text-sm ${insTone[x.t]}`}><b className="grid h-5 w-5 place-items-center rounded-full bg-white/70 text-xs dark:bg-white/10">{insIcon[x.t]}</b>{x.s}</span>)}</div>}<div className="grid grid-cols-2 gap-4 md:grid-cols-3 xl:grid-cols-5">
    {([['ยอดขายรวม', 'sales', compact, true], ['ค่าแรงรวม (Gross)', 'payroll', compact, false], ['ค่าล่วงเวลา (OT)', 'ot', compact, false], ['ค่าแรงปกติ + OT', 'payrollOt', compact, false],
      ['สัดส่วนต้นทุนแรงงาน %', 'labor', p1, false], ['จำนวนพนักงาน', 'hc', n0, true], ['ยอดขายต่อหัว', 'salesPerHead', compact, true], ['ยอดขายต่อค่าแรง 1 บาท', 'salesPerBaht', x2, true],
      ['ต้นทุนแรงงานต่อชิ้น', 'cpp', x2, false], ['อัตราลาออก %', 'turnover', p1, false]] as const).map(([l, k, fmt, up]) =>
      kpi(l, k, fmt, up))}
  </div></div>

  const sales = <div className="grid gap-4 lg:grid-cols-2">
    <Card title="แนวโน้มยอดขายรายเดือน"><Chart><LineChart data={trend}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={compact} />{tip}<Line dot={false} strokeWidth={2.5} dataKey="sales" stroke={COLORS[0]} /></LineChart></Chart></Card>
    <Card title="ยอดขายตามแผนก (มาก → น้อย)"><Chart><BarChart layout="vertical" data={[...dept].sort((a, b) => b.sales - a.sales)}>{grid}<XAxis {...ax} type="number" tickFormatter={compact} /><YAxis {...ax} type="category" dataKey="name" width={110} />{tip}<Bar radius={[5, 5, 5, 5]} dataKey="sales" fill={COLORS[1]} /></BarChart></Chart></Card>
    <Card title="ยอดขายตามสินค้า (Top 20)" className="lg:row-span-2"><Chart h={560}><BarChart layout="vertical" data={topN(sc.sales.filter(r => pred(r.Year_BE, r.Month_No)), 'Net_Sales_THB', 20, 'Product_Standard')}>{grid}<XAxis {...ax} type="number" tickFormatter={compact} /><YAxis {...ax} type="category" dataKey="name" width={130} />{tip}<Bar radius={[5, 5, 5, 5]} dataKey="value" fill={COLORS[0]} /></BarChart></Chart></Card>
    <Card title="สัดส่วนสินค้า (Treemap ตามกลุ่มสินค้ามาตรฐาน)"><Chart><Treemap data={topN(sc.sales.filter(r => pred(r.Year_BE, r.Month_No)), 'Net_Sales_THB', 30, 'Product_Standard')} dataKey="value" nameKey="name" content={<TreeCell />} /></Chart></Card>
    <Card title="สัดส่วนตลาด"><Chart><PieChart><Pie data={topN(sc.sales.filter(r => pred(r.Year_BE, r.Month_No)), 'Net_Sales_THB', 10, 'Market')} dataKey="value" nameKey="name" innerRadius={60} outerRadius={100} label>{COLORS.map(c => <Cell key={c} fill={c} />)}</Pie>{tip}<Legend /></PieChart></Chart></Card>
    <Card title="อันดับลูกค้า (Top 10)"><Chart><BarChart layout="vertical" data={topN(sc.sales.filter(r => pred(r.Year_BE, r.Month_No)), 'Net_Sales_THB', 10, 'Customer')}>{grid}<XAxis {...ax} type="number" tickFormatter={compact} /><YAxis {...ax} type="category" dataKey="name" width={110} />{tip}<Bar radius={[5, 5, 5, 5]} dataKey="value" fill={COLORS[2]} /></BarChart></Chart></Card>
  </div>

  const pct = (c:number, b:number) => (b ? (c - b) / b : null)
  const pm = [...vis].reverse().find(m => m.payroll > 0)   // latest month that actually has payroll data
  const pmPrev = pm ? all[all.findIndex(m => m.key === pm.key) - 1]?.key : undefined
  const prevG = groupBy(sc, 'dept', (y, m) => y * 100 + m === pmPrev), yoyG = groupBy(sc, 'dept', (y, m) => pm ? y * 100 + m === pm.key - 100 : false)
  const growth = (v:number|null) => v === null ? '–' : <span className={v <= 0 ? 'text-emerald-500' : 'text-red-500'}>{(v * 100).toFixed(1)}%</span>
  const deptRows = groupBy(sc, 'dept', (y, m) => pm ? y * 100 + m === pm.key : false).map(g => ({ ...g,
    mom: pct(g.payroll, prevG.find(x => x.name === g.name)?.payroll || 0), yoy: pct(g.payroll, yoyG.find(x => x.name === g.name)?.payroll || 0) }))
  const deptCols:ColumnDef<any>[] = [{ header: 'แผนก', accessorKey: 'name' }, { header: 'พนักงาน', accessorKey: 'emp', cell: i => n0(i.getValue() as any) },
    { header: 'ค่าแรงปกติ', accessorKey: 'regular', cell: i => n0(i.getValue() as any) }, { header: 'OT', accessorKey: 'ot', cell: i => n0(i.getValue() as any) },
    { header: 'ค่าแรงรวม', accessorKey: 'payroll', cell: i => n0(i.getValue() as any) }, { header: 'MoM', accessorKey: 'mom', cell: i => growth(i.getValue() as any) }, { header: 'YoY', accessorKey: 'yoy', cell: i => growth(i.getValue() as any) }]
  const payroll = <div className="grid gap-4 lg:grid-cols-2">
    <div className="grid grid-cols-3 gap-4 lg:col-span-2">{([['ค่าแรงรวม (Gross)', 'payroll'], ['ค่าล่วงเวลา (OT)', 'ot'], ['ค่าแรงปกติ + OT', 'payrollOt']] as const).map(([l, k]) => kpi(l, k, compact, false))}</div>
    <Card title="แนวโน้มค่าแรงรวม"><Chart><LineChart data={trend}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={compact} />{tip}<Line dot={false} strokeWidth={2.5} dataKey="payroll" stroke={COLORS[0]} /></LineChart></Chart></Card>
    <Card title="แนวโน้มค่าล่วงเวลา (OT)"><Chart><LineChart data={trend}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={compact} />{tip}<Line dot={false} strokeWidth={2.5} dataKey="ot" stroke={COLORS[3]} /></LineChart></Chart></Card>
    <Card title="ค่าแรงตามแผนก"><Chart><BarChart data={dept}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={compact} />{tip}<Bar radius={[5, 5, 5, 5]} dataKey="payroll" fill={COLORS[0]} /></BarChart></Chart></Card>
    <Card title="ค่าแรงตามไลน์ผลิต"><Chart><BarChart data={line}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={compact} />{tip}<Bar radius={[5, 5, 5, 5]} dataKey="payroll" fill={COLORS[1]} /></BarChart></Chart></Card>
    <Card title={`มุมมองรายแผนก — เดือน ${pm ? lbl(pm.key) : 'ไม่มีข้อมูล'}`} className="lg:col-span-2"><DataTable data={deptRows} columns={deptCols} /></Card>
  </div>

  const lowerBetter = rankKey === 'labor' || rankKey === 'cpp'
  const ranked = [...line].filter(g => g.sales > 0 || g.payroll > 0).sort((a, b) => (lowerBetter ? a[rankKey] - b[rankKey] : b[rankKey] - a[rankKey]))
  const rankFmt = rankKey === 'labor' ? p1 : x2
  const rankCard = (title:string, rows:Group[]) => <Card title={title}><Chart h={320}><BarChart layout="vertical" data={rows.map(r => ({ name: r.name, value: r[rankKey] }))}>{grid}<XAxis {...ax} type="number" /><YAxis {...ax} type="category" dataKey="name" width={140} /><Tooltip formatter={(v:any) => rankFmt(v)} /><Bar radius={[5, 5, 5, 5]} dataKey="value" fill={COLORS[1]} /></BarChart></Chart></Card>
  const productivity = <div className="grid gap-4 lg:grid-cols-2">
    <div className="flex gap-2 lg:col-span-2">{([['labor', 'สัดส่วนต้นทุนแรงงาน % (ต่ำ = ดี)'], ['salesPerHead', 'ยอดขายต่อหัว'], ['salesPerBaht', 'ยอดขายต่อค่าแรง 1 บาท'], ['cpp', 'ต้นทุนแรงงานต่อชิ้น (ต่ำ = ดี)']] as const).map(([k, l]) =>
      <button key={k} onClick={() => setRankKey(k)} className={`rounded-full px-3 py-1 text-sm ${rankKey === k ? 'bg-indigo-600 text-white' : 'bg-slate-200 dark:bg-slate-800'}`}>{l}</button>)}</div>
    {rankCard('Top 10 ดีที่สุด (ไลน์ผลิต)', ranked.slice(0, 10))}{rankCard('Top 10 แย่ที่สุด (ไลน์ผลิต)', ranked.slice(-10).reverse())}
    <Card title="ผลิตภาพตามไลน์ผลิต" className="lg:col-span-2"><DataTable data={line} columns={[{ header: 'ไลน์ผลิต', accessorKey: 'name' }, { header: 'ยอดขาย', accessorKey: 'sales', cell: i => n0(i.getValue() as any) },
      { header: 'ต้นทุนแรงงาน %', accessorKey: 'labor', cell: i => p1(i.getValue() as any) }, { header: 'ยอดขาย/หัว', accessorKey: 'salesPerHead', cell: i => n0(i.getValue() as any) },
      { header: 'ยอดขาย/฿ค่าแรง', accessorKey: 'salesPerBaht', cell: i => x2(i.getValue() as any) }, { header: 'ต้นทุน/ชิ้น', accessorKey: 'cpp', cell: i => x2(i.getValue() as any) }]} /></Card>
  </div>

  const hrBar = (title:string, rows:Group[], k:'turnover'|'hires'|'exits', fmt:(v:number)=>string) => <Card title={title}><Chart h={240}><BarChart data={rows}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={fmt} /><Tooltip formatter={(v:any) => fmt(v)} /><Bar radius={[5, 5, 5, 5]} dataKey={k} fill={COLORS[5]} /></BarChart></Chart></Card>
  const hr = <div className="grid gap-4 lg:grid-cols-2">
    {([['แนวโน้มจำนวนพนักงาน', 'hc', COLORS[0]], ['แนวโน้มการรับเข้า', 'hires', COLORS[1]], ['แนวโน้มการลาออก', 'exits', COLORS[3]], ['แนวโน้มอัตราลาออก (ลาออก หาร พนักงานเฉลี่ย)', 'turnover', COLORS[2]]] as const).map(([t, k, c]) =>
      <Card key={k} title={t}><Chart h={240}><LineChart data={trend}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={k === 'turnover' ? p1 : compact} /><Tooltip formatter={(v:any) => (k === 'turnover' ? p1(v) : x2(v))} /><Line dot={false} strokeWidth={2.5} dataKey={k} stroke={c} /></LineChart></Chart></Card>)}
    {hrBar('อัตราลาออกตามแผนก', dept, 'turnover', p1)}{hrBar('อัตราลาออกตามไลน์ผลิต', line, 'turnover', p1)}
    {hrBar('การรับเข้าตามแผนก', dept, 'hires', n0)}{hrBar('การลาออกตามแผนก', dept, 'exits', n0)}
    {hrBar('การรับเข้าตามไลน์ผลิต', line, 'hires', n0)}{hrBar('การลาออกตามไลน์ผลิต', line, 'exits', n0)}
  </div>

  const fcRows = (get:(m:MD)=>number, fcS?:{ma:number[]; sn:number[]}) => {
    const v = all.map(get), fc = fcS ?? forecast(v, h), last = all[all.length - 1]?.key || 0; let k = last
    const hist = all.slice(-24).map(m => ({ name: lbl(m.key), actual: get(m) as number|null, ma: null as number|null, sn: null as number|null }))
    if (hist.length) { hist[hist.length - 1].ma = hist[hist.length - 1].actual; hist[hist.length - 1].sn = hist[hist.length - 1].actual }
    return [...hist, ...fc.ma.map((x, i) => { k = nextKey(k); return { name: lbl(k), actual: null, ma: x, sn: fc.sn[i] } })]
  }
  const salesFc = forecast(all.map(m => m.sales), h), payFc = forecast(all.map(m => m.payroll), h)
  const laborFc = { ma: salesFc.ma.map((s, i) => (s ? payFc.ma[i] / s : 0)), sn: salesFc.sn.map((s, i) => (s ? payFc.sn[i] / s : 0)) }
  const fcChart = (title:string, data:any[], fmt:(v:number)=>string) => <Card title={title}><Chart><LineChart data={data}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={fmt} /><Tooltip formatter={(v:any) => fmt(v)} /><Legend />
    <Line dot={false} strokeWidth={2.5} dataKey="actual" stroke={COLORS[0]} name="ข้อมูลจริง" /><Line dot={false} strokeWidth={2.5} dataKey="ma" stroke={COLORS[1]} strokeDasharray="5 4" name="ค่าเฉลี่ยเคลื่อนที่ 3 เดือน" /><Line dot={false} strokeWidth={2.5} dataKey="sn" stroke={COLORS[2]} strokeDasharray="2 3" name="อิงฤดูกาล (ค่าเดือนเดียวกันปีก่อน)" /></LineChart></Chart></Card>
  const forecastPage = <div className="grid gap-4 lg:grid-cols-2">
    <div className="flex items-center gap-2 lg:col-span-2 text-sm">ช่วงพยากรณ์: {[3, 6, 12].map(x => <button key={x} onClick={() => setH(x)} className={`rounded-full px-3 py-1 ${h === x ? 'bg-indigo-600 text-white' : 'bg-slate-200 dark:bg-slate-800'}`}>{x} เดือน</button>)}
      {all.length < 12 && <span className="text-amber-500">ข้อมูลน้อยกว่า 12 เดือน — วิธีอิงฤดูกาลจะใช้ค่าเฉลี่ยเคลื่อนที่แทน</span>}</div>
    {fcChart('พยากรณ์ยอดขาย', fcRows(m => m.sales, salesFc), compact)}{fcChart('พยากรณ์ค่าแรง', fcRows(m => m.payroll, payFc), compact)}
    {fcChart('พยากรณ์สัดส่วนต้นทุนแรงงาน (ค่าแรง หาร ยอดขาย)', fcRows(m => m.labor, laborFc), p1)}
  </div>

  const alertPage = <Card title="ศูนย์แจ้งเตือน (ทุกเดือนตามตัวกรองปัจจุบัน)"><DataTable data={[...al].reverse().filter(a => pred(Math.floor(a.key / 100), a.key % 100))}
    columns={[{ header: 'เดือน', accessorFn: a => lbl(a.key) }, { header: 'ระดับ', accessorKey: 'level', cell: i => <span className={`rounded px-2 py-0.5 text-xs text-white ${lvlCls[i.getValue() as keyof typeof lvlCls]}`}>{LVL_TH[i.getValue() as keyof typeof LVL_TH]}</span> },
      { header: 'เหตุผล', accessorFn: a => a.reasons.join(' • ') }]} /></Card>

const ICONS = ['🏠', '📈', '💰', '⚙️', '👥', '🔮', '🔔']
  const nF = Object.values(f).filter(Boolean).length
  const pill = alertNow && <span className={`rounded-full px-3 py-1 text-xs font-semibold text-white ${lvlCls[alertNow.level]}`} title={alertNow.reasons.join(', ')}>สถานะ {cur ? lbl(cur.key) : ''}: {LVL_TH[alertNow.level]}</span>
  const btn = 'rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-800 dark:text-slate-200'
  return (
    <div className="min-h-screen lg:flex">
      <aside className="no-print sticky top-0 hidden h-screen w-64 shrink-0 flex-col border-r border-slate-200/70 bg-white p-4 dark:border-slate-800 dark:bg-slate-900 lg:flex">
        <div className="mb-7 flex items-center gap-3"><div className="grid h-10 w-10 place-items-center rounded-xl bg-gradient-to-br from-teal-500 to-teal-700 text-lg text-white shadow">📊</div>
          <div><div className="font-bold leading-tight">แดชบอร์ดผู้บริหาร</div><div className="text-xs text-slate-400">ยอดขาย × ค่าแรง × พนักงาน</div></div></div>
        <nav className="flex flex-col gap-1">{PAGES.map((p, i) => <button key={p} onClick={() => setPage(i)} className={`flex items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm font-medium transition ${page === i ? 'bg-teal-50 text-teal-700 dark:bg-teal-500/15 dark:text-teal-300' : 'text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800'}`}>
          <span className="text-base">{ICONS[i]}</span>{p}{i === 6 && alertNow && alertNow.level !== 'NONE' && <span className={`ml-auto h-2.5 w-2.5 rounded-full ${lvlCls[alertNow.level]}`} />}</button>)}</nav>
        <div className="mt-auto rounded-xl bg-slate-50 p-3 text-xs text-slate-500 dark:bg-slate-800/60"><div className="font-medium text-slate-600 dark:text-slate-300">ไฟล์ที่ใช้งาน</div><div className="mt-0.5 break-all">{ds.fileName}</div>
          <div className="mt-1">🔒 ประมวลผลในเครื่องเท่านั้น</div></div>
      </aside>
      <div className="min-w-0 flex-1">
        <header className="no-print sticky top-0 z-20 border-b border-slate-200/70 bg-white/85 backdrop-blur dark:border-slate-800 dark:bg-slate-900/85">
          <div className="flex flex-wrap items-center gap-2 px-4 py-3 lg:px-6">
            <h1 className="text-lg font-bold">{ICONS[page]} {PAGES[page]}</h1>{pill}
            <div className="ml-auto flex flex-wrap gap-2">
              <button onClick={() => setShowF(!showF)} className={btn}>🔎 ตัวกรอง{nF > 0 && <span className="ml-1.5 rounded-full bg-teal-600 px-1.5 text-xs text-white">{nF}</span>}</button>
              <label className="cursor-pointer rounded-lg bg-teal-600 px-3 py-1.5 text-sm font-medium text-white hover:bg-teal-700">อัปโหลดไฟล์ใหม่<input type="file" accept=".xlsx,.xls" hidden onChange={e => onFile(e.target.files?.[0])} /></label>
              <button onClick={() => exportData('xlsx')} className={btn}>Excel</button><button onClick={() => exportData('csv')} className={btn}>CSV</button><button onClick={() => window.print()} className={btn}>PDF</button>
              <button onClick={logout} className={btn}>ออกจากระบบ</button>
              <button onClick={() => setDark(!dark)} className={btn} title="สลับโหมดสว่าง/มืด">{dark ? '☀️' : '🌙'}</button></div></div>
          <nav className="flex gap-1 overflow-x-auto px-4 pb-2 lg:hidden">{PAGES.map((p, i) => <button key={p} onClick={() => setPage(i)} className={`whitespace-nowrap rounded-full px-3 py-1 text-sm ${page === i ? 'bg-teal-600 text-white' : 'bg-slate-100 dark:bg-slate-800'}`}>{p}</button>)}</nav>
          {showF && <div className="flex gap-3 overflow-x-auto border-t border-slate-100 px-4 py-3 dark:border-slate-800 lg:px-6">
            <Sel label="ปี (พ.ศ.)" value={f.year} opts={opts.year} on={setFilter('year')} />
            <label className="flex min-w-28 flex-col text-xs font-medium text-slate-500">เดือน<select value={f.month} onChange={e => setFilter('month')(e.target.value)} className={`mt-1 rounded-lg border px-2.5 py-1.5 text-sm text-slate-800 dark:text-slate-100 ${f.month ? 'border-teal-500 bg-teal-50 dark:bg-teal-500/10' : 'border-slate-200 bg-white dark:border-slate-700 dark:bg-slate-800'}`}><option value="">ทั้งหมด</option>{MONTHS.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}</select></label>
            <Sel label="กลุ่มสินค้ามาตรฐาน" value={f.productStd} opts={opts.productStd} on={setFilter('productStd')} />
            <Sel label="สินค้า" value={f.product} opts={opts.product} on={setFilter('product')} />
            <Sel label="ตลาด" value={f.market} opts={opts.market} on={setFilter('market')} />
            <Sel label="ลูกค้า" value={f.customer} opts={opts.customer} on={setFilter('customer')} />
            <Sel label="แผนก" value={f.dept} opts={opts.dept} on={setFilter('dept')} />
            <Sel label="ไลน์ผลิต" value={f.line} opts={opts.line} on={setFilter('line')} />
            <Sel label="เฟส" value={f.phase} opts={opts.phase} on={setFilter('phase')} />
            <Sel label="รอบจ่าย" value={f.payCycle} opts={opts.payCycle} on={setFilter('payCycle')} />
            {nF > 0 && <button onClick={() => setF(F0)} className="self-end whitespace-nowrap rounded-lg px-2 py-1.5 text-sm font-medium text-teal-600 hover:bg-teal-50 dark:hover:bg-teal-500/10">ล้างตัวกรอง</button>}
          </div>}
        </header>
      <main className="space-y-4 p-4 lg:p-6">
        {ds.warnings.length > 0 && <div className="rounded-md bg-amber-500/15 p-2 text-sm text-amber-600">{ds.warnings.join(' • ')}</div>}
        {page > 0 && <p className="text-sm text-slate-500">{['', 'ยอดขายตามเดือน สินค้า แผนก ตลาด และลูกค้า ตามตัวกรองด้านบน', 'ค่าแรง ค่าล่วงเวลา และจำนวนพนักงานตามแผนก/ไลน์ผลิต เทียบเดือนก่อน (MoM) และปีก่อน (YoY)', 'จัดอันดับไลน์ผลิตจากต้นทุนแรงงาน ยอดขายต่อหัว ยอดขายต่อค่าแรง และต้นทุนต่อชิ้น (เลือกตัวชี้วัดด้านล่าง)', 'จำนวนพนักงาน การรับเข้า การลาออก และอัตราลาออก = ลาออก หาร พนักงานเฉลี่ย', 'พยากรณ์ด้วยค่าเฉลี่ยเคลื่อนที่ 3 เดือน และวิธีอิงฤดูกาล (ใช้ค่าเดือนเดียวกันของปีก่อน) ประมวลผลในเครื่องทั้งหมด', '🔴 แดง: ลาออก>10% และยอดขายลด / ค่าแรงโตเร็วกว่ายอดขาย / ต้นทุนแรงงาน>18% / OT>15% • 🟡 เหลือง: ลาออก>8% หรือยอดขายลด • 🟢 เขียว: ยอดขายโตและต้นทุนแรงงานลด'][page]}</p>}
        {!vis.length ? <Card>ไม่พบข้อมูลตามตัวกรอง</Card> : [summary, sales, payroll, productivity, hr, forecastPage, alertPage][page]}
        {page === 0 && vis.length > 0 && <div className="grid gap-4 lg:grid-cols-2">
          <Card title="ยอดขาย เทียบ ค่าแรง"><Chart><LineChart data={trend}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={compact} />{tip}<Legend /><Line dot={false} strokeWidth={2.5} dataKey="sales" name="ยอดขาย" stroke={COLORS[0]} /><Line dot={false} strokeWidth={2.5} dataKey="payroll" name="ค่าแรง" stroke={COLORS[2]} /></LineChart></Chart></Card>
          <Card title="ต้นทุนแรงงาน % และ OT %"><Chart><LineChart data={trend.map(m => ({ ...m, ...derive(m) }))}>{grid}<XAxis {...ax} dataKey="name" /><YAxis {...ax} tickFormatter={p1} /><Tooltip formatter={(v:any) => p1(v)} /><Legend /><Line dot={false} strokeWidth={2.5} dataKey="labor" name="ต้นทุนแรงงาน %" stroke={COLORS[2]} /><Line dot={false} strokeWidth={2.5} dataKey="otPct" name="OT % ของค่าแรง" stroke={COLORS[3]} /></LineChart></Chart></Card></div>}
      </main>
      </div>
    </div>)
}

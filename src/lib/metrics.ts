import type { Dataset } from '../types'

export type FKey = 'year'|'month'|'productStd'|'product'|'market'|'customer'|'dept'|'line'|'phase'|'payCycle'
export type F = Record<FKey, string>
export const F0: F = { year:'', month:'', productStd:'', product:'', market:'', customer:'', dept:'', line:'', phase:'', payCycle:'' }
const ok = (f:string, v:any) => !f || String(v) === f
const add = (m:Map<string,number>, k:string, v:number) => m.set(k, (m.get(k) || 0) + v)
export const kk = (y:number, m:number) => y * 100 + m
export const lbl = (k:number) => `${k % 100}/${String(Math.floor(k / 100)).slice(2)}`   // Month/Year(BE short)
export const nextKey = (k:number) => (k % 100 === 12 ? k + 89 : k + 1)

export interface Scoped { sales:Dataset['sales']; payroll:(Dataset['payroll'][0] & {w:number})[]; workforce:(Dataset['workforce'][0] & {w:number})[] }

/** Apply global filters (except time). Shared lines: payroll/headcount are allocated by the filtered share of Quantity_Pcs. */
export function scope(d:Dataset, f:F): Scoped {
  const sales = d.sales.filter(r => ok(f.productStd, r.Product_Standard) && ok(f.product, r.Product_Name_Raw) && ok(f.market, r.Market) &&
    ok(f.customer, r.Customer) && ok(f.dept, r.Department_Short) && ok(f.line, r.Production_Line) && ok(f.phase, r.Phase))
  const scoped = !!(f.productStd || f.product || f.market || f.customer || f.phase)
  const tot = new Map<string, number>(), sub = new Map<string, number>()
  if (scoped) { for (const r of d.sales) add(tot, `${r.Year_BE}|${r.Month_No}|${r.Production_Line}`, r.Quantity_Pcs); for (const r of sales) add(sub, `${r.Year_BE}|${r.Month_No}|${r.Production_Line}`, r.Quantity_Pcs) }
  const sh = (y:number, m:number, l:string) => { if (!scoped) return 1; const k = `${y}|${m}|${l}`; return (sub.get(k) || 0) / (tot.get(k) || 1) }
  return {
    sales,
    payroll: d.payroll.filter(r => ok(f.dept, r.Department_Short) && ok(f.line, r.Production_Line) && ok(f.payCycle, r.Pay_Cycle)).map(r => ({ ...r, w: sh(r.Year_BE, r.Month_No, r.Production_Line) })),
    workforce: d.workforce.filter(r => ok(f.dept, r.Department_Short) && ok(f.line, r.Production_Line)).map(r => ({ ...r, w: sh(r.Year_BE, r.Month_No, r.Production_Line) })),
  }
}

export interface M { key:number; sales:number; qty:number; payroll:number; regular:number; ot:number; hc:number; avgHc:number; hires:number; exits:number }
export const derive = (m:Partial<M>) => ({
  payrollOt: (m.regular || 0) + (m.ot || 0),
  labor: m.sales ? (m.payroll || 0) / m.sales : 0,                    // Gross Payroll / Sales
  salesPerHead: m.hc ? (m.sales || 0) / m.hc : 0,
  salesPerBaht: m.payroll ? (m.sales || 0) / m.payroll : 0,
  cpp: m.qty ? (m.payroll || 0) / m.qty : 0,                          // Payroll / Quantity_Pcs
  turnover: m.avgHc ? (m.exits || 0) / m.avgHc : 0,                   // Exits / ((HC_Start+HC_End)/2)
  otPct: m.payroll ? (m.ot || 0) / m.payroll : 0,
})
export type MD = M & ReturnType<typeof derive>

export function monthly(s:Scoped): MD[] {
  const mp = new Map<number, M>()
  const g = (y:number, m:number) => { const k = kk(y, m); if (!mp.has(k)) mp.set(k, { key:k, sales:0, qty:0, payroll:0, regular:0, ot:0, hc:0, avgHc:0, hires:0, exits:0 }); return mp.get(k)! }
  for (const r of s.sales) { const x = g(r.Year_BE, r.Month_No); x.sales += r.Net_Sales_THB; x.qty += r.Quantity_Pcs }
  for (const r of s.payroll) { const x = g(r.Year_BE, r.Month_No); x.payroll += r.Gross_Pay_THB * r.w; x.regular += r.Regular_Pay_THB * r.w; x.ot += r.OT_THB * r.w }
  for (const r of s.workforce) { const x = g(r.Year_BE, r.Month_No); x.hc += r.HC_End * r.w; x.avgHc += (r.HC_Start + r.HC_End) / 2 * r.w; x.hires += r.New_Hires * r.w; x.exits += r.Exits * r.w }
  return [...mp.values()].sort((a, b) => a.key - b.key).map(m => ({ ...m, ...derive(m) }))
}

export const inPeriod = (f:F) => (y:number, m:number) => (!f.year || +f.year === y) && (!f.month || +f.month === m)

export interface Group { name:string; sales:number; qty:number; payroll:number; regular:number; ot:number; hc:number; emp:number; hires:number; exits:number; turnover:number; labor:number; salesPerHead:number; salesPerBaht:number; cpp:number }
/** Group by department / production line / etc. over a period predicate. */
export function groupBy(s:Scoped, dim:'dept'|'line', pred:(y:number,m:number)=>boolean): Group[] {
  const mp = new Map<string, any>(), months = new Set<number>()
  const g = (n:string) => { if (!mp.has(n)) mp.set(n, { name:n, sales:0, qty:0, payroll:0, regular:0, ot:0, hcSum:0, avgHc:0, emp:0, hires:0, exits:0 }); return mp.get(n) }
  const nm = (r:any) => (dim === 'dept' ? r.Department_Short : r.Production_Line) || '(n/a)'
  for (const r of s.sales) if (pred(r.Year_BE, r.Month_No)) { const x = g(nm(r)); x.sales += r.Net_Sales_THB; x.qty += r.Quantity_Pcs; months.add(kk(r.Year_BE, r.Month_No)) }
  for (const r of s.payroll) if (pred(r.Year_BE, r.Month_No)) { const x = g(nm(r)); x.payroll += r.Gross_Pay_THB * r.w; x.regular += r.Regular_Pay_THB * r.w; x.ot += r.OT_THB * r.w; x.emp += r.Employees_Paid * r.w }
  for (const r of s.workforce) if (pred(r.Year_BE, r.Month_No)) { const x = g(nm(r)); x.hcSum += r.HC_End * r.w; x.avgHc += (r.HC_Start + r.HC_End) / 2 * r.w; x.hires += r.New_Hires * r.w; x.exits += r.Exits * r.w }
  const n = Math.max(months.size, 1)
  return [...mp.values()].map(x => { const hc = x.hcSum / n
    return { name:x.name, sales:x.sales, qty:x.qty, payroll:x.payroll, regular:x.regular, ot:x.ot, hc, emp:x.emp / n, hires:x.hires, exits:x.exits,
      turnover: x.avgHc ? x.exits / x.avgHc : 0, labor: x.sales ? x.payroll / x.sales : 0, salesPerHead: hc ? x.sales / hc : 0, salesPerBaht: x.payroll ? x.sales / x.payroll : 0, cpp: x.qty ? x.payroll / x.qty : 0 } })
}
export function topN<T extends Record<string, any>>(rows:T[], key:string, n=20, name='name') {
  const m = new Map<string, number>(); rows.forEach(r => m.set(r[name], (m.get(r[name]) || 0) + (r[key] || 0)))
  return [...m].map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value).slice(0, n)
}

/* ---------------- Alert engine ---------------- */
export type Level = 'RED'|'YELLOW'|'GREEN'|'NONE'
export interface Alert { key:number; level:Level; reasons:string[] }
const pct = (a:number, b:number) => (b ? (a - b) / Math.abs(b) : 0)
export function alerts(ms:MD[]): Alert[] {
  return ms.map((m, i) => {
    const p = ms[i - 1], sg = p ? pct(m.sales, p.sales) : 0, pg = p ? pct(m.payroll, p.payroll) : 0
    const red:string[] = [], yel:string[] = []
    if (m.turnover > .10 && sg < 0) red.push('Turnover > 10% และ Sales ลดลง')
    if (p && pg > sg) red.push('Payroll โตเร็วกว่า Sales')
    if (m.labor > .18) red.push('Labor Cost > 18%')
    if (m.otPct > .15) red.push('OT > 15% ของ Payroll')
    if (m.turnover > .08) yel.push('Turnover > 8%')
    if (p && sg < 0) yel.push('Sales Growth < 0')
    if (red.length) return { key:m.key, level:'RED' as Level, reasons:red }
    if (yel.length) return { key:m.key, level:'YELLOW' as Level, reasons:yel }
    if (p && sg > 0 && m.labor < p.labor) return { key:m.key, level:'GREEN' as Level, reasons:['Sales โตและ Labor Cost ลดลง'] }
    return { key:m.key, level:'NONE' as Level, reasons:[] }
  })
}

/* ---------------- Forecast engine (local only) ---------------- */
export function forecast(vals:number[], h:number) {
  const ma:number[] = [], sn:number[] = [], hist = [...vals], histSn = [...vals]
  for (let i = 0; i < h; i++) {
    const last = hist.slice(-3), v = last.reduce((a, b) => a + b, 0) / Math.max(last.length, 1)
    ma.push(v); hist.push(v)
    const s = histSn.length >= 12 ? histSn[histSn.length - 12] : v   // seasonal naive; falls back to MA when < 12 months
    sn.push(s); histSn.push(s)
  }
  return { ma, sn }
}

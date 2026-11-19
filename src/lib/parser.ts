import * as XLSX from 'xlsx'
import type { Dataset } from '../types'

const num = (v:any) => (typeof v === 'number' ? v : Number(String(v ?? '').replace(/,/g, '')) || 0)
const str = (v:any) => (v == null ? '' : String(v).trim())
const HEADER_KEYS = /^(Year_BE|Product_Code|Department_Raw)$/

/** Sheets have title rows above the header, so find the header row dynamically. */
function table(wb:XLSX.WorkBook, name:string): Record<string, any>[] {
  const ws = wb.Sheets[name]; if (!ws) return []
  const a = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, raw: true, defval: null })
  const h = a.findIndex(r => r && typeof r[0] === 'string' && HEADER_KEYS.test(r[0].trim()))
  if (h < 0) return []
  const head = (a[h] as string[]).map(x => str(x))
  return a.slice(h + 1).filter(r => r && r.some(v => v !== null && v !== '')).map(r => Object.fromEntries(head.map((k, i) => [k, r[i]])))
}

export const parseWorkbook = async (file:File) => parseBuffer(await file.arrayBuffer(), file.name)

export async function parseBuffer(buf:ArrayBuffer, fileName:string): Promise<Dataset> {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true })
  const warnings:string[] = []
  for (const s of ['Sales_Input','Payroll_Input','Workforce_Input']) if (!wb.Sheets[s]) warnings.push(`ไม่พบ Sheet ${s}`)
  const productMap = table(wb, 'Product_Map'), deptMap = table(wb, 'Dept_Line_Map'), alloc = table(wb, 'Allocation_Rules')
  const byCode = new Map(productMap.map(p => [str(p.Product_Code), p]))

  let unmapped = 0
  const sales = table(wb, 'Sales_Input').filter(r => r.Year_BE).map(r => {
    const m = byCode.get(str(r.Product_Code))
    const std = str(r.Product_Standard) || str(m?.Product_Standard)
    if (!std) unmapped++
    return {
      Year_BE: num(r.Year_BE), Month_No: num(r.Month_No), Market: str(r.Market), Customer: str(r.Customer), Product_Code: str(r.Product_Code),
      Product_Name_Raw: str(r.Product_Name_Raw), Product_Standard: std || str(r.Product_Name_Raw) || '(unmapped)',
      Production_Line: str(r.Production_Line) || str(m?.Production_Line), Department_Short: str(r.Department_Short) || str(m?.Department_Short),
      Phase: str(r.Phase), Quantity_Pcs: num(r.Quantity_Pcs), Net_Sales_THB: num(r.Net_Sales_THB), CN_THB: num(r.CN_THB),
    }
  })
  if (unmapped) warnings.push(`Product mapping ไม่ครบ ${unmapped} แถว`)

  const payroll = table(wb, 'Payroll_Input').filter(r => r.Year_BE).map(r => ({
    Year_BE: num(r.Year_BE), Month_No: num(r.Month_No), Pay_Cycle: str(r.Pay_Cycle), Department_Raw: str(r.Department_Raw),
    Department_Short: str(r.Department_Short), Production_Line: str(r.Production_Line), Employees_Paid: num(r.Employees_Paid),
    Regular_Pay_THB: num(r.Regular_Pay_THB), OT_THB: num(r.OT_THB), Allowance_THB: num(r.Allowance_THB),
    Gross_Pay_THB: num(r.Gross_Pay_THB), Deduction_THB: num(r.Deduction_THB), Net_Pay_THB: num(r.Net_Pay_THB),
  }))
  const workforce = table(wb, 'Workforce_Input').filter(r => r.Year_BE).map(r => ({
    Year_BE: num(r.Year_BE), Month_No: num(r.Month_No), Department_Short: str(r.Department_Short), Production_Line: str(r.Production_Line),
    HC_Start: num(r.HC_Start), New_Hires: num(r.New_Hires), Transfers_In: num(r.Transfers_In), Transfers_Out: num(r.Transfers_Out),
    Exits: num(r.Exits), HC_End: num(r.HC_End),   // Avg_HC / Turnover are recomputed, formula cells are ignored
  }))
  return { sales, payroll, workforce, productMap, deptMap, alloc, warnings, fileName }
}

/* ---- IndexedDB persistence (everything stays in the browser) ---- */
const open = () => new Promise<IDBDatabase>((res, rej) => {
  const r = indexedDB.open('exec-dashboard', 1)
  r.onupgradeneeded = () => r.result.createObjectStore('kv')
  r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error)
})
export async function saveDataset(d:Dataset) {
  const db = await open(); return new Promise<void>((res, rej) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(d, 'dataset'); t.oncomplete = () => res(); t.onerror = () => rej(t.error) })
}
export async function loadDataset(): Promise<Dataset | null> {
  const db = await open(); return new Promise(res => { const q = db.transaction('kv').objectStore('kv').get('dataset'); q.onsuccess = () => res(q.result ?? null); q.onerror = () => res(null) })
}

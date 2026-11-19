export interface SalesRow { Year_BE:number; Month_No:number; Market:string; Customer:string; Product_Code:string; Product_Name_Raw:string; Product_Standard:string; Production_Line:string; Department_Short:string; Phase:string; Quantity_Pcs:number; Net_Sales_THB:number; CN_THB:number }
export interface PayrollRow { Year_BE:number; Month_No:number; Pay_Cycle:string; Department_Raw:string; Department_Short:string; Production_Line:string; Employees_Paid:number; Regular_Pay_THB:number; OT_THB:number; Allowance_THB:number; Gross_Pay_THB:number; Deduction_THB:number; Net_Pay_THB:number }
export interface WorkforceRow { Year_BE:number; Month_No:number; Department_Short:string; Production_Line:string; HC_Start:number; New_Hires:number; Transfers_In:number; Transfers_Out:number; Exits:number; HC_End:number }
export interface Dataset {
  sales:SalesRow[]; payroll:PayrollRow[]; workforce:WorkforceRow[]
  productMap:Record<string,any>[]; deptMap:Record<string,any>[]; alloc:Record<string,any>[]
  warnings:string[]; fileName:string
}

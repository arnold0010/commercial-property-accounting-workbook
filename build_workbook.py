#!/usr/bin/env python3
"""
Harborview Commercial Properties, LLC
Simulated Commercial Property Accounting & Financial Reporting Workbook - FY2025

The whole model is generated from one data set (chart of accounts, opening
balances, 50 transactions, 8 adjusting entries, mortgage amortisation) so the
workbook can never contain a typed-in total that disagrees with the detail.

Sheet chain (Excel formulas, not pasted values):
  Transaction Register -> Journal Entries -> General Ledger -> Trial Balance
  -> Profit & Loss / Balance Sheet / Cash Flow Statement
  -> Management Financial Package
"""

import datetime as dt
import re
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.workbook.properties import CalcProperties

# ----------------------------------------------------------------------------
# 0. helpers
# ----------------------------------------------------------------------------


def r2(x):
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

# ----------------------------------------------------------------------------
# 1. chart of accounts
# ----------------------------------------------------------------------------
# (code, name, type, normal balance, statement)
ACCOUNTS = [
    (1000, "Cash - Operating Bank Account (First Harbor Bank)", "Asset", "Dr", "Balance Sheet"),
    (1100, "Tenant Receivables", "Asset", "Dr", "Balance Sheet"),
    (1200, "Prepaid Insurance", "Asset", "Dr", "Balance Sheet"),
    (1500, "Land", "Asset", "Dr", "Balance Sheet"),
    (1510, "Building & Improvements", "Asset", "Dr", "Balance Sheet"),
    (1515, "Accumulated Depreciation - Building", "Asset", "Cr", "Balance Sheet"),
    (1520, "Furniture, Fixtures & Equipment", "Asset", "Dr", "Balance Sheet"),
    (1525, "Accumulated Depreciation - FF&E", "Asset", "Cr", "Balance Sheet"),
    (2010, "Accounts Payable", "Liability", "Cr", "Balance Sheet"),
    (2020, "Accrued Expenses", "Liability", "Cr", "Balance Sheet"),
    (2030, "Tenant Security Deposits", "Liability", "Cr", "Balance Sheet"),
    (2100, "Mortgage Payable - Rivergate Office Center", "Liability", "Cr", "Balance Sheet"),
    (2110, "Mortgage Payable - Maple Street Retail Shops", "Liability", "Cr", "Balance Sheet"),
    (3010, "Owner's Capital", "Equity", "Cr", "Balance Sheet"),
    (3020, "Owner Distributions", "Equity", "Dr", "Balance Sheet"),
    (4010, "Rental Revenue - Offices (Rivergate)", "Revenue", "Cr", "Profit & Loss"),
    (4020, "Rental Revenue - Retail (Maple Street)", "Revenue", "Cr", "Profit & Loss"),
    (4030, "Expense Recoveries (CAM / NNN)", "Revenue", "Cr", "Profit & Loss"),
    (4040, "Parking Revenue", "Revenue", "Cr", "Profit & Loss"),
    (4050, "Other Property Income (Antenna Lease)", "Revenue", "Cr", "Profit & Loss"),
    (4060, "Late Fee Income", "Revenue", "Cr", "Profit & Loss"),
    (4900, "Interest Income", "Revenue", "Cr", "Profit & Loss"),
    (5010, "Property Taxes", "Expense", "Dr", "Profit & Loss"),
    (5020, "Property Insurance", "Expense", "Dr", "Profit & Loss"),
    (5030, "Utilities", "Expense", "Dr", "Profit & Loss"),
    (5040, "Repairs & Maintenance", "Expense", "Dr", "Profit & Loss"),
    (5050, "Cleaning & Landscaping", "Expense", "Dr", "Profit & Loss"),
    (5060, "Property Management Fees", "Expense", "Dr", "Profit & Loss"),
    (5070, "Depreciation Expense", "Expense", "Dr", "Profit & Loss"),
    (5080, "Interest Expense - Mortgages", "Expense", "Dr", "Profit & Loss"),
    (5090, "Bank Charges", "Expense", "Dr", "Profit & Loss"),
    (5100, "Bad Debt Expense", "Expense", "Dr", "Profit & Loss"),
]
NAME = {a[0]: a[1] for a in ACCOUNTS}
TYPE = {a[0]: a[2] for a in ACCOUNTS}
NORM = {a[0]: a[3] for a in ACCOUNTS}

# ----------------------------------------------------------------------------
# 2. mortgage schedules
# ----------------------------------------------------------------------------


def pmt(principal, annual_rate, n_months):
    r = annual_rate / 12
    return r2(principal * r / (1 - (1 + r) ** -n_months))


def schedule(principal, annual_rate, n_months, n_payments):
    """Payment 1 covers the first month of interest. Returns list of dicts."""
    r = annual_rate / 12
    pay = pmt(principal, annual_rate, n_months)
    bal, out = principal, []
    for _ in range(n_payments):
        interest = r2(bal * r)
        princ = r2(pay - interest)
        bal = r2(bal - princ)
        out.append({"interest": interest, "principal": princ, "payment": pay, "balance": bal})
    return out


LOAN_A = {"name": "Rivergate Office Center", "principal": 1_150_000.00, "rate": 0.0585, "term": 240}
LOAN_B = {"name": "Maple Street Retail Shops", "principal": 1_400_000.00, "rate": 0.0625, "term": 300}

# Jan-1-2025 payment settles December-2024 interest -> 12 payments during 2025
SCHED_A = schedule(LOAN_A["principal"], LOAN_A["rate"], LOAN_A["term"], 12)
# First payment 1-Apr-2025 settles March interest -> 9 payments during 2025
SCHED_B = schedule(LOAN_B["principal"], LOAN_B["rate"], LOAN_B["term"], 9)

# December 2025 interest accrued (AJ-06) on the 1-Dec-2025 closing balances
INT_A_DEC = r2(SCHED_A[-1]["balance"] * LOAN_A["rate"] / 12)
INT_B_DEC = r2(SCHED_B[-1]["balance"] * LOAN_B["rate"] / 12)

# ----------------------------------------------------------------------------
# 3. opening balances (1 January 2025) - debit positive / credit negative
# ----------------------------------------------------------------------------
OPEN = {
    1000: 81_606.25,
    1100: 24_000.00,
    1200: 9_000.00,
    1500: 900_000.00,
    1510: 2_400_000.00,
    1515: -380_000.00,
    1520: 65_000.00,
    1525: -22_000.00,
    2020: -SCHED_A[0]["interest"],     # December 2024 mortgage interest accrued
    2030: -32_000.00,
    2100: -LOAN_A["principal"],
    3010: -1_890_000.00,
}
assert round(sum(OPEN.values()), 2) == 0.0, ("opening balances do not balance", sum(OPEN.values()))

# ----------------------------------------------------------------------------
# 4. transactions (50)
# ----------------------------------------------------------------------------
TX = []


def T(date, kind, party, scope, desc, cash, lines, note=""):
    TX.append({"date": date, "kind": kind, "party": party, "scope": scope,
               "desc": desc, "cash": r2(cash),
               "lines": [(a, r2(d), r2(c), p, m) for a, d, c, p, m in lines],
               "note": note})


# ---- 4a. monthly rent roll billing (12 entries)
for m in range(1, 13):
    amt = 28_000.00 if m <= 2 else 42_100.00
    lines = [(1100, amt, 0, "CORP", "Invoices issued to tenants")]
    lines.append((4010, 0, 24_000.00, "RIV", "Rivergate office rent (3 tenants)"))
    if m >= 3:
        lines.append((4020, 0, 13_000.00, "MAP", "Maple Street retail rent (3 tenants)"))
    lines.append((4030, 0, 2_000.00, "RIV", "CAM / tax recoveries"))
    if m >= 3:
        lines.append((4030, 0, 1_100.00, "MAP", "CAM recoveries"))
    lines.append((4040, 0, 1_500.00, "RIV", "Parking income"))
    lines.append((4050, 0, 500.00, "RIV", "Roof antenna lease"))
    T(dt.date(2025, m, 1), "Revenue cycle", "Tenants - Rivergate & Maple Street", "Both",
      "Rent roll billed for " + MONTHS[m - 1] + " 2025", 0, lines,
      "Accrual: revenue recognised when billed / earned, not when collected.")

# ---- 4b. tenant collections (12 entries) - one month in arrears
COLLECTIONS = [
    (1, 15, 24_000.00, 0.00, "RIV", "December 2024 rent"),
    (2, 15, 28_000.00, 0.00, "RIV", "January rent"),
    (3, 15, 28_000.00, 1_250.00, "RIV", "February rent + late fees (Vertex Legal Group)"),
    (4, 15, 42_100.00, 0.00, "RIV", "March rent"),
    (5, 15, 42_100.00, 0.00, "RIV", "April rent"),
    (6, 15, 42_100.00, 0.00, "RIV", "May rent"),
    (7, 15, 42_100.00, 0.00, "RIV", "June rent"),
    (8, 15, 42_100.00, 1_000.00, "MAP", "July rent + late fee (Noodle House LLC)"),
    (9, 15, 42_100.00, 0.00, "RIV", "August rent"),
    (10, 15, 42_100.00, 0.00, "RIV", "September rent"),
    (11, 15, 42_100.00, 0.00, "RIV", "October rent"),
    (12, 29, 42_100.00, 0.00, "RIV", "November rent (deposited 31-Dec after bank cut-off)"),
]
for m, d, ar, late, lprop, label in COLLECTIONS:
    lines = [(1000, ar + late, 0, "CORP", "Receipts banked")]
    lines.append((1100, 0, ar, "CORP", "Applied against tenant invoices"))
    if late:
        lines.append((4060, 0, late, lprop, "Late fee income"))
    T(dt.date(2025, m, d), "Revenue cycle", "Tenants - Rivergate & Maple Street", "Both",
      "Tenant collections - " + label, ar + late, lines,
      "December 2025 invoices remain unpaid at year end (tenant receivables).")

# ---- 4c. monthly mortgage payments (12 entries)
for m in range(1, 13):
    a = SCHED_A[m - 1]
    b_idx = m - 4
    b = SCHED_B[b_idx] if b_idx >= 0 else None
    lines = []
    if m == 1:
        lines.append((2020, a["interest"], 0, "RIV",
                      "December 2024 interest accrued at 31-Dec-2024 - clearing opening accrual"))
    else:
        lines.append((5080, a["interest"], 0, "RIV", MONTHS[m - 2] + " interest - Rivergate"))
    lines.append((2100, a["principal"], 0, "RIV", "Principal reduction - Rivergate"))
    cash = a["payment"]
    if b:
        lines.append((5080, b["interest"], 0, "MAP", MONTHS[m - 2] + " interest - Maple Street"))
        lines.append((2110, b["principal"], 0, "MAP", "Principal reduction - Maple Street"))
        cash = r2(cash + b["payment"])
    lines.append((1000, 0, cash, "CORP", "Debit to First Harbor Bank"))
    T(dt.date(2025, m, 1), "Debt service", "First Harbor Bank", "Both",
      "Mortgage instalment - " + MONTHS[m - 1] + " 2025" +
      (" (Rivergate + Maple Street)" if b else " (Rivergate only)"), -cash, lines,
      "Payment covers the prior month's interest: a one-month interest lag.")

# ---- 4d. property tax, insurance, deposits, capital, acquisitions
T(dt.date(2025, 1, 22), "Treasury", "Suite 120 former tenant", "RIV",
  "Security deposit refunded - Suite 120 tenant vacated 31-Dec-2024", -6_000.00,
  [(2030, 6_000.00, 0, "RIV", "Deposit liability released"),
   (1000, 0, 6_000.00, "CORP", "Cheque #2201")],
  "A deposit is a liability, never revenue.")

T(dt.date(2025, 3, 1), "Financing", "Member / Owner", "CORP",
  "Owner capital contribution to fund Maple Street acquisition", 700_000.00,
  [(1000, 700_000.00, 0, "CORP", "Wire received"),
   (3010, 0, 700_000.00, "CORP", "Additional paid-in capital")],
  "Equity injection - a financing inflow, not income.")

T(dt.date(2025, 3, 1), "Investing", "Pinnacle Retail Partners (vendor)", "MAP",
  "Acquisition of Maple Street Retail Shops - land, building and capitalised costs", -740_000.00,
  [(1500, 620_000.00, 0, "MAP", "Land allocation (30% of price)"),
   (1510, 1_480_000.00, 0, "MAP", "Building allocation (70% of price)"),
   (1510, 40_000.00, 0, "MAP", "Legal, title, survey and due diligence - capitalised"),
   (2110, 0, 1_400_000.00, "MAP", "New mortgage - 6.25%, 25-year amortisation"),
   (1000, 0, 740_000.00, "CORP", "Cash at closing (700,000 down + 40,000 costs)")],
  "Investing outflow; purchase price allocated between land and building.")

T(dt.date(2025, 3, 1), "Treasury", "Bloom & Co., Harbor Cuts, Noodle House", "MAP",
  "Security deposits collected from Maple Street tenants on lease commencement", 36_000.00,
  [(1000, 36_000.00, 0, "CORP", "Three deposits received"),
   (2030, 0, 36_000.00, "MAP", "Refundable deposit liability")],
  "Held in trust; refundable at lease expiry.")

T(dt.date(2025, 4, 10), "Property operations", "Multnomah County Tax Collector", "RIV",
  "Rivergate property tax - first half 2025 instalment", -21_000.00,
  [(5010, 21_000.00, 0, "RIV", "Jan-Jun 2025 assessment"),
   (1000, 0, 21_000.00, "CORP", "Cheque #2216")])

T(dt.date(2025, 10, 10), "Property operations", "Multnomah County Tax Collector", "RIV",
  "Rivergate property tax - second half 2025 instalment", -21_000.00,
  [(5010, 21_000.00, 0, "RIV", "Jul-Dec 2025 assessment"),
   (1000, 0, 21_000.00, "CORP", "Cheque #2248")])

T(dt.date(2025, 10, 1), "Property operations", "Cascade Mutual Insurance", "Both",
  "Annual property insurance premium paid - covers 1-Oct-2025 to 30-Sep-2026", -54_000.00,
  [(1200, 30_000.00, 0, "RIV", "Rivergate share - 12 months cover (asset)"),
   (1200, 24_000.00, 0, "MAP", "Maple Street share - 12 months cover (asset)"),
   (1000, 0, 54_000.00, "CORP", "Cheque #2244")],
  "Prepaid: 9 months of cover (40,500) remains an asset at 31-Dec-2025.")

T(dt.date(2025, 9, 8), "Property operations", "Mesabi Roofing Co.", "RIV",
  "Emergency roof membrane repair - Suite 415 water ingress", -6_200.00,
  [(5040, 6_200.00, 0, "RIV", "Repair and make-good"),
   (1000, 0, 6_200.00, "CORP", "Cheque #2241")],
  "Revenue repair - expensed, not capitalised.")

# ---- 4e. quarterly property operating cost settlements (3 paid in 2025)
SETTLEMENTS = [
    (dt.date(2025, 4, 15), "Q1 2025", 26_000.00, [
        (5030, 9_900.00, 0, "RIV", "Electricity, gas and water"),
        (5040, 6_800.00, 0, "RIV", "HVAC compressor and general repairs"),
        (5050, 7_200.00, 0, "RIV", "Janitorial and grounds"),
        (5060, 2_100.00, 0, "RIV", "Management fee - 700/month"),
    ], "Cheque #2219"),
    (dt.date(2025, 7, 15), "Q2 2025", 39_420.00, [
        (5030, 9_900.00, 0, "RIV", "Electricity, gas and water"),
        (5030, 2_400.00, 0, "MAP", "Common area utilities"),
        (5040, 13_600.00, 0, "RIV", "Parking lot sealcoating and restriping"),
        (5040, 2_000.00, 0, "MAP", "Storefront and sign repairs"),
        (5050, 5_800.00, 0, "RIV", "Janitorial and grounds"),
        (5050, 2_000.00, 0, "MAP", "Landscaping and lot sweeping"),
        (5060, 2_100.00, 0, "RIV", "Management fee - 700/month"),
        (5060, 1_620.00, 0, "MAP", "Management fee - 540/month"),
    ], "Cheque #2233"),
    (dt.date(2025, 10, 15), "Q3 2025", 32_620.00, [
        (5030, 10_400.00, 0, "RIV", "Electricity, gas and water"),
        (5030, 3_000.00, 0, "MAP", "Common area utilities"),
        (5040, 6_400.00, 0, "RIV", "Elevator service and general repairs"),
        (5040, 1_000.00, 0, "MAP", "General repairs"),
        (5050, 5_900.00, 0, "RIV", "Janitorial and grounds"),
        (5050, 2_200.00, 0, "MAP", "Landscaping and lot sweeping"),
        (5060, 2_100.00, 0, "RIV", "Management fee - 700/month"),
        (5060, 1_620.00, 0, "MAP", "Management fee - 540/month"),
    ], "Cheque #2246"),
]
for date, q, total, lines, cheque in SETTLEMENTS:
    T(date, "Property operations", "Sightline Property Management", "Both",
      "Property operating cost settlement - " + q + " (manager re-bills vendors quarterly)",
      -total, lines + [(1000, 0, total, "CORP", cheque)],
      "Q4 settlement is accrued at 31-Dec-2025 (see AJ-05).")

# ---- 4f. equity distributions and other income
T(dt.date(2025, 6, 30), "Financing", "Member / Owner", "CORP",
  "Owner distribution - H1 2025", -30_000.00,
  [(3020, 30_000.00, 0, "CORP", "Distribution of equity"),
   (1000, 0, 30_000.00, "CORP", "Wire out")],
  "A distribution of equity - NOT an operating expense.")

T(dt.date(2025, 12, 31), "Financing", "Member / Owner", "CORP",
  "Owner distribution - H2 2025", -30_000.00,
  [(3020, 30_000.00, 0, "CORP", "Distribution of equity"),
   (1000, 0, 30_000.00, "CORP", "Cheque #2255 - outstanding at year end")],
  "A distribution of equity - NOT an operating expense.")

T(dt.date(2025, 12, 31), "Treasury", "First Harbor Bank", "CORP",
  "Bank interest credited on operating account for 2025", 1_190.00,
  [(1000, 1_190.00, 0, "CORP", "Interest credited"),
   (4900, 0, 1_190.00, "CORP", "Interest income")])

TX.sort(key=lambda x: x["date"])
assert len(TX) == 50, len(TX)

# ----------------------------------------------------------------------------
# 5. adjusting entries (8) - all dated 31 December 2025
# ----------------------------------------------------------------------------
AJ = []


def A(no, aj_type, desc, lines, basis):
    AJ.append({"no": no, "type": aj_type, "desc": desc,
               "lines": [(a, r2(d), r2(c), p, m) for a, d, c, p, m in lines],
               "basis": basis, "date": dt.date(2025, 12, 31)})


A("AJ-01", "Non-cash", "Depreciation - buildings and improvements for 2025", [
    (5070, 60_000.00, 0, "RIV", "Rivergate: 2,400,000 / 40 years"),
    (5070, 31_666.67, 0, "MAP", "Maple Street: 1,520,000 / 40 years x 10/12"),
    (1515, 0, 91_666.67, "CORP", "Accumulated depreciation"),
], "Straight line, 40-year life, no residual value.")

A("AJ-02", "Non-cash", "Depreciation - furniture, fixtures & equipment", [
    (5070, 9_285.71, 0, "RIV", "65,000 / 7 years"),
    (1525, 0, 9_285.71, "CORP", "Accumulated depreciation"),
], "Straight line, 7-year life, full year charge.")

A("AJ-03", "Prepaid expense", "Insurance expired during 2025 (prepaid amortisation)", [
    (5020, 16_500.00, 0, "RIV", "Jan-Mar 9,000 + Oct-Dec 7,500"),
    (5020, 6_000.00, 0, "MAP", "Oct-Dec 2025 (24,000 x 3/12)"),
    (1200, 0, 22_500.00, "CORP", "Prepaid insurance released to expense"),
], "Opening prepaid 9,000 fully expired + 3 of 12 months on the new policy. "
   "Prepaid remaining: 40,500.")

A("AJ-04", "Accrued expense", "Property tax accrued - Maple Street (Mar-Dec 2025)", [
    (5010, 12_000.00, 0, "MAP", "10 months at 1,200/month"),
    (2020, 0, 12_000.00, "CORP", "Tax liability - billed 2026"),
], "Assessment billed in arrears; expense belongs to 2025.")

A("AJ-05", "Accrued expense", "Q4 2025 operating cost settlement accrued (invoiced 15-Jan-2026)", [
    (5030, 9_000.00, 0, "RIV", "Electricity, gas and water"),
    (5030, 2_700.00, 0, "MAP", "Common area utilities"),
    (5040, 3_000.00, 0, "RIV", "General repairs"),
    (5040, 2_700.00, 0, "MAP", "General repairs"),
    (5050, 5_400.00, 0, "RIV", "Janitorial and grounds"),
    (5050, 2_100.00, 0, "MAP", "Landscaping and lot sweeping"),
    (5060, 2_100.00, 0, "RIV", "Management fee"),
    (5060, 2_160.00, 0, "MAP", "Management fee incl. March catch-up"),
    (2010, 0, 29_160.00, "CORP", "Payable to Sightline Property Management"),
], "Manager settles quarterly in arrears - Q4 costs are 2025 expenses.")

A("AJ-06", "Accrued expense", "Mortgage interest accrued - December 2025", [
    (5080, INT_A_DEC, 0, "RIV", "Rivergate: 5.85% on " + format(SCHED_A[-1]["balance"], ",.2f")),
    (5080, INT_B_DEC, 0, "MAP", "Maple Street: 6.25% on " + format(SCHED_B[-1]["balance"], ",.2f")),
    (2020, 0, r2(INT_A_DEC + INT_B_DEC), "CORP", "Interest payable 1-Jan-2026"),
], "Interest lags the payment by one month; December is unpaid at year end.")

A("AJ-07", "Bank reconciliation", "Bank service charges for 2025 not yet recorded", [
    (5090, 185.00, 0, "CORP", "Account analysis and wire fees"),
    (1000, 0, 185.00, "CORP", "Charged by First Harbor Bank"),
], "On the bank statement, not in the cash book - recorded here (see Bank Rec).")

A("AJ-08", "Bank reconciliation", "NSF cheque - Bloom & Co. Florist returned unpaid", [
    (1100, 3_600.00, 0, "MAP", "Receivable reinstated - tenant still owes"),
    (1000, 0, 3_600.00, "CORP", "Reversal of deposit"),
], "On the bank statement, not in the cash book - recorded here (see Bank Rec).")

# ----------------------------------------------------------------------------
# 6. model validation (data level)
# ----------------------------------------------------------------------------
for e in TX:
    dr = round(sum(l[1] for l in e["lines"]), 2)
    cr = round(sum(l[2] for l in e["lines"]), 2)
    assert dr == cr, ("unbalanced txn", e["desc"], dr, cr)
for e in AJ:
    dr = round(sum(l[1] for l in e["lines"]), 2)
    cr = round(sum(l[2] for l in e["lines"]), 2)
    assert dr == cr, ("unbalanced AJ", e["no"], dr, cr)

BAL = dict(OPEN)
MOV = {a[0]: {"dr": 0.0, "cr": 0.0} for a in ACCOUNTS}
MOV_AJ = {a[0]: {"dr": 0.0, "cr": 0.0} for a in ACCOUNTS}
for e in TX:
    for acct, d, c, p, m in e["lines"]:
        BAL[acct] = round(BAL.get(acct, 0.0) + d - c, 2)
        MOV[acct]["dr"] = round(MOV[acct]["dr"] + d, 2)
        MOV[acct]["cr"] = round(MOV[acct]["cr"] + c, 2)
for e in AJ:
    for acct, d, c, p, m in e["lines"]:
        BAL[acct] = round(BAL.get(acct, 0.0) + d - c, 2)
        MOV_AJ[acct]["dr"] = round(MOV_AJ[acct]["dr"] + d, 2)
        MOV_AJ[acct]["cr"] = round(MOV_AJ[acct]["cr"] + c, 2)

for _a in ACCOUNTS:
    BAL.setdefault(_a[0], 0.0)
assert round(sum(BAL.values()), 2) == 0.0, ("trial balance does not balance", sum(BAL.values()))

REV = sum(-BAL[a[0]] for a in ACCOUNTS if a[2] == "Revenue")
EXP = sum(BAL[a[0]] for a in ACCOUNTS if a[2] == "Expense")
INT_INC = r2(-BAL[4900])
REV_PROP = r2(REV - INT_INC)
DEP_PY = r2(BAL[5070])
INT_PY = r2(BAL[5080])
OPEX_PY = r2(sum(BAL[a[0]] for a in ACCOUNTS
                 if a[2] == "Expense" and a[0] not in (5070, 5080, 5090)))
NOI_PY = r2(REV_PROP - OPEX_PY)
DS_PY = r2(sum(x["payment"] for x in SCHED_A) + sum(x["payment"] for x in SCHED_B))
NI = r2(REV - EXP)
ASSETS = sum(BAL[a[0]] for a in ACCOUNTS if a[2] == "Asset")
LIAB = sum(-BAL[a[0]] for a in ACCOUNTS if a[2] == "Liability")
EQUITY = sum(-BAL[a[0]] for a in ACCOUNTS if a[2] == "Equity")
CASH = BAL[1000]
assert round(ASSETS - LIAB - EQUITY - NI, 2) == 0.0, (ASSETS, LIAB, EQUITY, NI)

RECEIPTS = round(sum(l[1] - l[2] for e in TX for l in e["lines"] if l[0] == 1000
                     and l[1] > l[2]), 2)
PAYMENTS = round(sum(l[2] - l[1] for e in TX for l in e["lines"] if l[0] == 1000
                     and l[2] > l[1]), 2)
CASH_JE = round(RECEIPTS - PAYMENTS, 2)
CASH_AJ = round(sum(l[1] - l[2] for e in AJ for l in e["lines"] if l[0] == 1000), 2)
assert round(OPEN[1000] + CASH_JE + CASH_AJ, 2) == CASH, (CASH_JE, CASH_AJ, CASH)

# bank reconciliation assumptions
DIT = 42_100.00          # 29-Dec receipts deposited 31-Dec, after the bank cut-off
OS_CHECKS = [("2255", dt.date(2025, 12, 31), "Member / Owner - distribution", 30_000.00)]
OS_TOTAL = round(sum(c[3] for c in OS_CHECKS), 2)
CASH_OUT_AJ = r2(sum(l[2] - l[1] for e in AJ for l in e["lines"]
                     if l[0] == 1000 and l[2] > l[1]))
BANK_CLOSING = r2(OPEN[1000] + (RECEIPTS - DIT) - (PAYMENTS + CASH_OUT_AJ - OS_TOTAL))
assert round(BANK_CLOSING + DIT - OS_TOTAL, 2) == CASH

# ----------------------------------------------------------------------------
# 7. workbook styling helpers
# ----------------------------------------------------------------------------
NAVY, BLUE, PALE, GREY = "1F3864", "2E75B6", "DCE6F1", "F2F2F2"
GREEN, RED, AMBER = "C6EFCE", "FFC7CE", "FFEB9C"
MONEY = '#,##0.00_);(#,##0.00);"-"_)'
MONEY0 = '#,##0_);(#,##0);"-"_)'
PCT = '0.0%'
DATE = 'dd-mmm-yyyy'

thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
TOPRULE = Border(top=Side(style="thin", color="1F3864"))
DBLRULE = Border(top=Side(style="thin", color="1F3864"),
                 bottom=Side(style="double", color="1F3864"))

F_TITLE = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_SUB = Font(name="Calibri", size=9.5, italic=True, color="595959")
F_SEC = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
F_HDR = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
F_BODY = Font(name="Calibri", size=10)
F_BOLD = Font(name="Calibri", size=10, bold=True)
F_SMALL = Font(name="Calibri", size=8.5, color="595959")
F_ITAL = Font(name="Calibri", size=9, italic=True, color="595959")

FILL_SEC = PatternFill("solid", fgColor=NAVY)
FILL_HDR = PatternFill("solid", fgColor=BLUE)
FILL_BAND = PatternFill("solid", fgColor=GREY)
FILL_TOTAL = PatternFill("solid", fgColor=PALE)
FILL_OK = PatternFill("solid", fgColor=GREEN)
FILL_BAD = PatternFill("solid", fgColor=RED)


def sheet_title(ws, col_last, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_last)
    c = ws.cell(row=1, column=1, value=title)
    c.font = F_TITLE
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 24
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=col_last)
    c = ws.cell(row=2, column=2 - 1, value=subtitle)
    c.font = F_SUB
    ws.row_dimensions[2].height = 14


def section(ws, row, col_first, col_last, text):
    ws.merge_cells(start_row=row, start_column=col_first, end_row=row, end_column=col_last)
    c = ws.cell(row=row, column=col_first, value=text)
    c.fill = FILL_SEC
    c.font = F_SEC
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 19


def headers(ws, row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.fill = FILL_HDR
        c.font = F_HDR
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def put(ws, row, col, value, fmt=None, font=None, align=None, fill=None, border=None,
        indent=0):
    c = ws.cell(row=row, column=col, value=value)
    if fmt:
        c.number_format = fmt
    c.font = font or F_BODY
    if align or indent:
        c.alignment = Alignment(horizontal=align or "general", vertical="center",
                                indent=indent)
    if fill:
        c.fill = fill
    if border:
        c.border = border
    return c


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def page(ws, repeat_rows=None):
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    if repeat_rows:
        ws.print_title_rows = repeat_rows


def check_cell(ws, row, col, formula, label_ok="OK", width_label=None):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = Font(name="Calibri", size=9.5, bold=True)
    c.alignment = Alignment(horizontal="center")
    c.fill = FILL_OK
    c.border = BOX
    return c


# ----------------------------------------------------------------------------
# 8. build
# ----------------------------------------------------------------------------
wb = Workbook()
wb.calculation = CalcProperties(fullCalcOnLoad=True)

S_PROF = "Company Profile & COA"
S_REG = "Transaction Register"
S_JE = "Journal Entries"
S_GL = "General Ledger"
S_TB = "Trial Balance"
S_AJ = "Adjusting Entries"
S_BR = "Bank Reconciliation"
S_PL = "Profit & Loss Statement"
S_BS = "Balance Sheet"
S_CF = "Cash Flow Statement"
S_PKG = "Management Financial Package"

ws1 = wb.active
ws1.title = S_PROF
ws2 = wb.create_sheet(S_REG)
ws3 = wb.create_sheet(S_JE)
ws4 = wb.create_sheet(S_GL)
ws5 = wb.create_sheet(S_TB)
ws6 = wb.create_sheet(S_AJ)
ws7 = wb.create_sheet(S_BR)
ws8 = wb.create_sheet(S_PL)
ws9 = wb.create_sheet(S_BS)
ws10 = wb.create_sheet(S_CF)
ws11 = wb.create_sheet(S_PKG)

for ws, colour in ((ws1, NAVY), (ws2, BLUE), (ws3, BLUE), (ws4, BLUE), (ws5, BLUE),
                   (ws6, "C55A11"), (ws7, "C55A11"), (ws8, "548235"), (ws9, "548235"),
                   (ws10, "548235"), (ws11, NAVY)):
    ws.sheet_properties.tabColor = colour

# =====================================================================
# SHEET 1 - Company profile, policies, rent roll, chart of accounts
# =====================================================================
ws = ws1
widths(ws, {"A": 6, "B": 34, "C": 30, "D": 16, "E": 14, "F": 16, "G": 16, "H": 30})
sheet_title(ws, 8, "HARBORVIEW COMMERCIAL PROPERTIES, LLC",
            "Commercial Property Accounting Workbook  |  Fiscal year 1 January - 31 December 2025  |  "
            "Accrual basis (US GAAP)  |  All amounts in US dollars")

r = 4
section(ws, r, 1, 8, "1.  COMPANY PROFILE")
r += 1
profile = [
    ("Legal entity", "Harborview Commercial Properties, LLC - Oregon limited liability company (pass-through)"),
    ("Membership", "Two members: A. Whitfield (60%), M. Okafor (40%)"),
    ("Principal place of business", "400 Rivergate Plaza, Suite 500, Portland, Oregon 97204"),
    ("Business activity", "Owns, leases and manages commercial real estate: one multi-tenant office building and "
                          "one multi-tenant retail strip"),
    ("Fiscal year", "1 January - 31 December 2025 (comparatives at 31 December 2024)"),
    ("Basis of accounting", "Accrual basis, US GAAP. Revenue is recognised when earned, expenses when incurred - "
                            "not when cash moves"),
    ("Financial reporting currency", "US dollar (USD)"),
    ("Property manager", "Sightline Property Management (vendor invoices re-billed and settled quarterly in arrears)"),
    ("Banking", "First Harbor Bank - operating account no. xx4417 (single bank account; deposits held in it)"),
    ("Income tax", "Pass-through entity: no entity-level income tax; results flow to the members' returns"),
    ("Prepared by", "Financial Controller - 15 February 2026"),
]
for label, val in profile:
    put(ws, r, 2, label, font=F_BOLD)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    c = put(ws, r, 3, val, font=F_BODY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 26 if len(val) > 95 else 15
    r += 1

r += 1
section(ws, r, 1, 8, "2.  PROPERTY PORTFOLIO")
r += 1
headers(ws, r, ["", "Property (code)", "Location / type", "Acquired", "Cost basis",
                "Rentable sq ft", "Occupancy", "Tenants"])
r += 1
props = [
    ("Rivergate Office Center (RIV)", "400 Rivergate Plaza, Portland OR - Class B office", dt.date(2023, 3, 1),
     3_300_000.00, 28_500, 0.93, "3 of 4 suites let"),
    ("Maple Street Retail Shops (MAP)", "1180-1190 Maple Street, Beaverton OR - neighbourhood retail",
     dt.date(2025, 3, 1), 2_140_000.00, 11_200, 0.87, "3 of 4 units let"),
]
for p in props:
    put(ws, r, 2, p[0], font=F_BOLD)
    put(ws, r, 3, p[1], font=F_BODY).alignment = Alignment(wrap_text=True, vertical="center")
    put(ws, r, 4, p[2], fmt=DATE, align="center")
    put(ws, r, 5, p[3], fmt=MONEY, align="right")
    put(ws, r, 6, p[4], fmt="#,##0", align="right")
    put(ws, r, 7, p[5], fmt="0%", align="right")
    put(ws, r, 8, p[6], font=F_BODY)
    ws.row_dimensions[r].height = 26
    r += 1

r += 1
section(ws, r, 1, 8, "3.  RENT ROLL AT 31 DECEMBER 2025")
r += 1
headers(ws, r, ["", "Property", "Unit / suite", "Tenant", "Monthly rent",
                "Annual rent", "Lease expiry", "Status"])
r += 1
rentroll = [
    ("RIV", "Suite 200", "Northwind Consulting LLC", 8_000.00, dt.date(2027, 12, 31), "Let"),
    ("RIV", "Suite 310", "Cascade Design Studio", 6_500.00, dt.date(2026, 6, 30), "Let"),
    ("RIV", "Suite 415", "Vertex Legal Group", 9_500.00, dt.date(2028, 2, 28), "Let"),
    ("RIV", "Suite 120", "- vacant -", 0.00, None, "Vacant"),
    ("MAP", "Unit 1", "Bloom & Co. Florist", 4_200.00, dt.date(2028, 2, 28), "Let"),
    ("MAP", "Unit 2", "Harbor Cuts Barbershop", 3_300.00, dt.date(2027, 2, 28), "Let"),
    ("MAP", "Unit 3", "Noodle House LLC", 5_500.00, dt.date(2030, 2, 28), "Let"),
    ("MAP", "Unit 4", "- vacant -", 0.00, None, "Vacant"),
]
rr_first = r
for code, unit, tenant, rent, exp, status in rentroll:
    put(ws, r, 2, code, align="center")
    put(ws, r, 3, unit)
    put(ws, r, 4, tenant, font=F_BOLD if rent else F_ITAL)
    put(ws, r, 5, rent, fmt=MONEY, align="right")
    put(ws, r, 6, "=E%d*12" % r, fmt=MONEY, align="right")
    put(ws, r, 7, exp if exp else "-", fmt=DATE, align="center")
    put(ws, r, 8, status, align="center",
        fill=FILL_BAD if status == "Vacant" else None)
    r += 1
put(ws, r, 4, "Total contracted rent", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 5, "=SUM(E%d:E%d)" % (rr_first, r - 1), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 6, "=SUM(F%d:F%d)" % (rr_first, r - 1), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 7, "Vacancy: 7% RIV / 13% MAP of area", font=F_ITAL)
r += 2

section(ws, r, 1, 8, "4.  ACCOUNTING POLICIES (the accrual rules this workbook applies)")
r += 1
policies = [
    ("Revenue recognition", "Rent is earned evenly over the lease term and recognised when billed for the month it "
                           "covers. December 2025 invoices unpaid at year end are tenant receivables - the revenue "
                           "is still 2025 revenue. Late fees are recognised when chargeable; interest income when "
                           "credited."),
    ("Expense recognition", "Expenses are recognised in the period the benefit is consumed, regardless of payment "
                            "date. Costs incurred but not invoiced are accrued (AJ-04, AJ-05, AJ-06)."),
    ("Prepaid expenses", "Payments covering future periods are recorded as assets and released to expense over the "
                         "coverage period (AJ-03)."),
    ("Property, plant & equipment", "Land is not depreciated. Buildings are depreciated over 40 years and FF&E over "
                                    "7 years, straight line, no residual value. Acquisition costs are capitalised."),
    ("Tenant security deposits", "Cash received is a refundable liability (2030), not income, and is released only "
                                 "when the tenant vacates (Jan-2025 refund)."),
    ("Owner distributions", "Distributions to members are charges against equity (3020), never operating expenses, "
                            "and are excluded from the income statement."),
    ("Bad debts", "Specific write-off; the NSF cheque is reinstated as a receivable (AJ-08) because the tenant still "
                  "owes the money."),
    ("Cash flow statement", "Indirect method for operating activities; interest paid and interest received are "
                            "treated as operating."),
]
for label, text in policies:
    put(ws, r, 2, label, font=F_BOLD)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    c = put(ws, r, 3, text, font=F_BODY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 34
    r += 1

r += 1
section(ws, r, 1, 8, "5.  CHART OF ACCOUNTS")
r += 1
headers(ws, r, ["Code", "Account name", "Type", "Normal balance", "Statement", "", "", ""])
r += 1
coa_first = r
for code, name, typ, norm, stmt in ACCOUNTS:
    put(ws, r, 1, code, fmt="0", align="center")
    put(ws, r, 2, name, border=BOX)
    put(ws, r, 3, typ, border=BOX)
    put(ws, r, 4, norm, align="center", border=BOX)
    put(ws, r, 5, stmt, border=BOX)
    if typ == "Revenue":
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="E2EFDA")
    elif typ == "Expense":
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FCE4D6")
    elif typ in ("Liability", "Equity"):
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FFF2CC")
    else:
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="DDEBF7")
    r += 1
coa_last = r - 1
COA = "'%s'!$A$%d:$C$%d" % (S_PROF, coa_first, coa_last)
put(ws, r, 2, "Property tags used in the journal: RIV = Rivergate, MAP = Maple Street, "
              "CORP = corporate / balance-sheet only.", font=F_ITAL)
r += 2

section(ws, r, 1, 8, "6.  KEY ASSUMPTIONS OF THE SIMULATION")
r += 1
assumptions = [
    "1. Rents are billed on the 1st for the month they cover and collected in the following month; the December "
    "2025 invoices are unpaid at 31 December (tenant receivables 45,700 including the returned cheque).",
    "2. Rivergate mortgage: 1,150,000 at 5.85%, 240-month amortisation, instalments on the 1st covering the "
    "previous month's interest.",
    "3. Maple Street mortgage: 1,400,000 at 6.25%, 300-month amortisation, drawn 1 March 2025, first instalment "
    "1 April 2025.",
    "4. The property manager pays vendors and settles quarterly in arrears; Q4 is accrued (AJ-05) and paid "
    "15 January 2026.",
    "5. Insurance is billed annually in advance; the policy renewed 1 October 2025 covers 12 months.",
    "6. Maple Street property tax for March-December 2025 is assessed and billed in 2026, so it is accrued.",
    "7. One bank account. Tenant receipts of 42,100 deposited on 31 December cleared after the bank cut-off "
    "(deposit in transit); the H2 distribution cheque (30,000) was not presented by year end.",
    "8. No entity-level income tax, no sales tax, no payroll - the company has no employees.",
]
for a in assumptions:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    c = put(ws, r, 2, a, font=F_BODY)
    c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    ws.row_dimensions[r].height = 26 if len(a) > 120 else 15
    r += 1

r += 1
section(ws, r, 1, 8, "7.  HOW THE WORKBOOK FLOWS")
r += 1
flow = ("Transaction Register (50 transactions)  ->  Journal Entries (double entry)  ->  General Ledger "
        "(account-wise)  ->  Trial Balance (opening + operating + adjusting)  ->  Adjusting Entries (8 accruals / "
        "prepaids / bank items)  ->  Bank Reconciliation  ->  Profit & Loss, Balance Sheet, Cash Flow Statement  ->  "
        "Management Financial Package.")
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
c = put(ws, r, 2, flow, font=F_BODY)
c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
ws.row_dimensions[r].height = 46
page(ws)

# =====================================================================
# SHEET 2 - Transaction register
# =====================================================================
ws = ws2
widths(ws, {"A": 9, "B": 13, "C": 17, "D": 60, "E": 11, "F": 30, "G": 15, "H": 11, "I": 46})
sheet_title(ws, 9, "TRANSACTION REGISTER - 2025",
            "Every economic event of the year, in date order, with the cash effect and the journal it produced")
HDR_REG = 5
headers(ws, HDR_REG, ["Txn", "Date", "Cycle", "Transaction description", "Property",
                      "Counterparty / payee", "Cash movement (+/-)", "Journal", "Note"])
r = HDR_REG + 1
reg_first = r
for i, e in enumerate(TX, start=1):
    band = FILL_BAND if i % 2 == 0 else None
    put(ws, r, 1, "TXN-%03d" % i, align="center", fill=band)
    put(ws, r, 2, e["date"], fmt=DATE, align="center", fill=band)
    put(ws, r, 3, e["kind"], fill=band)
    put(ws, r, 4, e["desc"], fill=band)
    put(ws, r, 5, e["scope"], align="center", fill=band)
    put(ws, r, 6, e["party"], fill=band)
    put(ws, r, 7, e["cash"], fmt=MONEY, align="right", fill=band)
    put(ws, r, 8, "JE-%03d" % i, align="center", fill=band)
    c = put(ws, r, 9, e["note"], font=F_SMALL, fill=band)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 24 if e["note"] else 15
    r += 1
reg_last = r - 1
put(ws, r, 4, "Net cash movement - operating transactions", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 7, "=SUM(G%d:G%d)" % (reg_first, reg_last), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
r += 1
put(ws, r, 4, "Add: net cash effect of adjusting entries (AJ-07, AJ-08)", font=F_BODY)
put(ws, r, 7, "=-3785", fmt=MONEY, align="right")
put(ws, r, 9, "Bank charges 185 + NSF cheque 3,600", font=F_SMALL,
    align="right")
r += 1
put(ws, r, 4, "Net increase in cash for 2025", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 7, "=G%d+G%d" % (r - 2, r - 1), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
r += 1
put(ws, r, 4, "Cash at 1 January 2025", font=F_BODY)
put(ws, r, 7, OPEN[1000], fmt=MONEY, align="right")
r += 1
put(ws, r, 4, "Cash at 31 December 2025 - agrees to General Ledger account 1000", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 7, "=G%d+G%d" % (r - 1, r - 2), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 9, "=IF(ROUND(G%d-'%s'!G%d,2)=0,\"OK\",\"CHECK\")" % (r, S_GL, 6 + 0), font=F_BOLD,
    align="center")
r += 2
put(ws, r, 1, "Cycle key:", font=F_BOLD)
put(ws, r, 2, "Revenue cycle = billing and collection; Debt service = mortgage instalments; Property operations = "
              "taxes, insurance, utilities, repairs, management; Investing = acquisition; Financing = capital, "
              "distributions, debt; Treasury = deposits, refunds, bank interest.",
    font=F_SMALL)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
page(ws, "5:5")
ws.freeze_panes = "A6"

# =====================================================================
# SHEET 3 - Journal entries
# =====================================================================
ws = ws3
widths(ws, {"A": 10, "B": 12, "C": 17, "D": 8, "E": 40, "F": 52, "G": 8, "H": 14, "I": 14, "J": 9})
sheet_title(ws, 10, "JOURNAL ENTRIES - 2025",
            "50 transactions in double-entry form. Every entry is checked: debits must equal credits")

_n_je = 5 + sum(len(e["lines"]) for e in TX)
put(ws, 3, 6, "Totals ->", font=F_BOLD, align="right")
put(ws, 3, 8, "=SUM(H6:H%d)" % _n_je, fmt=MONEY, font=F_BOLD, align="right")
put(ws, 3, 9, "=SUM(I6:I%d)" % _n_je, fmt=MONEY, font=F_BOLD, align="right")
put(ws, 3, 10, '=IF(ROUND(H3-I3,2)=0,"OK","CHECK")', font=F_BOLD, align="center", fill=FILL_OK, border=BOX)
put(ws, 3, 5, "Journal - debits must equal credits", font=F_ITAL)

HDR_JE = 5
headers(ws, HDR_JE, ["JE no.", "Date", "Cycle", "Acct", "Account name", "Description / memo",
                     "Property", "Debit", "Credit", "Check"])
r = HDR_JE + 1
JE_FIRST = r
je_rows = []          # (acct, sheet_row, date, desc, prop, dr, cr, je_no)
for i, e in enumerate(TX, start=1):
    jno = "JE-%03d" % i
    band = FILL_BAND if i % 2 == 0 else None
    for acct, dr, cr, prop, memo in e["lines"]:
        put(ws, r, 1, jno, align="center", fill=band)
        put(ws, r, 2, e["date"], fmt=DATE, align="center", fill=band)
        put(ws, r, 3, e["kind"], font=F_SMALL, fill=band)
        put(ws, r, 4, acct, fmt="0", align="center", fill=band)
        put(ws, r, 5, "=IFERROR(VLOOKUP($D%d,%s,2,FALSE),\"?\")" % (r, COA), fill=band)
        put(ws, r, 6, memo, fill=band)
        put(ws, r, 7, prop, align="center", fill=band)
        put(ws, r, 8, dr if dr else None, fmt=MONEY, align="right", fill=band)
        put(ws, r, 9, cr if cr else None, fmt=MONEY, align="right", fill=band)
        put(ws, r, 10, '=IF(ROUND(SUMIF($A$6:$A$999,$A%d,$H$6:$H$999)'
                       '-SUMIF($A$6:$A$999,$A%d,$I$6:$I$999),2)=0,"OK","ERR")' % (r, r),
            font=F_SMALL, align="center", fill=band)
        je_rows.append((acct, r, e["date"], memo, prop, dr, cr, jno))
        r += 1
JE_LAST = r - 1
put(ws, r, 6, "Total journals", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 8, "=SUM(H%d:H%d)" % (JE_FIRST, JE_LAST), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 9, "=SUM(I%d:I%d)" % (JE_FIRST, JE_LAST), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 10, '=IF(ROUND(H%d-I%d,2)=0,"OK","CHECK")' % (r, r), font=F_BOLD,
    align="center", fill=FILL_OK, border=BOX)
r += 2
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
put(ws, r, 1, "Note: adjusting entries AJ-01 to AJ-08 are on the 'Adjusting Entries' sheet and are added to the "
              "ledger and trial balance in the adjusting column. Entries dated 1 January settle December 2024 "
              "items carried in opening balances.", font=F_ITAL)
page(ws, "5:5")
ws.freeze_panes = "A6"

# =====================================================================
# SHEET 4 - General ledger (account summary + full detail with running balance)
# =====================================================================
ws = ws4
widths(ws, {"A": 12, "B": 52, "C": 12, "D": 14, "E": 14, "F": 14, "G": 15})
sheet_title(ws, 7, "GENERAL LEDGER - 2025",
            "Account-wise record of every posting, showing opening balance, debits, credits, closing balance and a "
            "running balance")
section(ws, 4, 1, 7, "PART A - ACCOUNT ACTIVITY SUMMARY")
headers(ws, 5, ["Acct", "Account", "Type", "Opening balance", "Debits", "Credits", "Closing balance"])
r = 6
GL_FIRST = r
gl_row = {}
# The Adjusting Entries sheet layout is fixed: headers on row 5, data from row 6.
AJ_FIRST = 6
n_aj_lines = sum(len(a["lines"]) for a in AJ)
AJ_LAST = AJ_FIRST + n_aj_lines - 1


def sumifs(sheet, col, acct):
    return "SUMIFS('%s'!$%s$%d:$%s$%d,'%s'!$D$%d:$D$%d,%d)" % (
        sheet, col, (JE_FIRST if sheet == S_JE else AJ_FIRST),
        col, (JE_LAST if sheet == S_JE else AJ_LAST),
        sheet, (JE_FIRST if sheet == S_JE else AJ_FIRST),
        (JE_LAST if sheet == S_JE else AJ_LAST), acct)


for code, name, typ, norm, stmt in ACCOUNTS:
    open_amt = OPEN.get(code, 0.0)
    put(ws, r, 1, code, fmt="0", align="center", border=BOX)
    put(ws, r, 2, name, border=BOX)
    put(ws, r, 3, typ, border=BOX)
    put(ws, r, 4, open_amt if open_amt else None, fmt=MONEY, align="right", border=BOX)
    put(ws, r, 5, "=" + sumifs(S_JE, "H", code) + "+" + sumifs(S_AJ, "H", code),
        fmt=MONEY, align="right", border=BOX)
    put(ws, r, 6, "=" + sumifs(S_JE, "I", code) + "+" + sumifs(S_AJ, "I", code),
        fmt=MONEY, align="right", border=BOX)
    put(ws, r, 7, "=D%d+E%d-F%d" % (r, r, r), fmt=MONEY, align="right", font=F_BOLD, border=BOX)
    gl_row[code] = r
    r += 1
GL_LAST = r - 1
put(ws, r, 2, "Totals - debits and credits must agree", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 4, "=SUM(D%d:D%d)" % (GL_FIRST, GL_LAST), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL,
    align="right")
put(ws, r, 5, "=SUM(E%d:E%d)" % (GL_FIRST, GL_LAST), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL,
    align="right")
put(ws, r, 6, "=SUM(F%d:F%d)" % (GL_FIRST, GL_LAST), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL,
    align="right")
put(ws, r, 7, '=IF(ROUND(D%d,2)=0,"BALANCED","CHECK")' % r, font=F_BOLD, align="center",
    fill=FILL_OK, border=BOX)
r += 2

section(ws, r, 1, 7, "PART B - LEDGER DETAIL (debit balances positive; credit balances in brackets)")
put(ws, r + 1, 1, "Running balance = opening balance + debits - credits for the account.", font=F_ITAL)
r += 2
DET_HDR = ["Date", "Journal", "Description", "Property", "Debit", "Credit", "Balance"]
detail_map = []
for code, name, typ, norm, stmt in ACCOUNTS:
    lines = [x for x in je_rows if x[0] == code]
    aj_lines = []
    for a in AJ:
        for acct, dr, cr, prop, memo in a["lines"]:
            if acct == code:
                aj_lines.append((acct, None, a["date"], memo, prop, dr, cr, a["no"]))
    if not lines and not aj_lines and not OPEN.get(code):
        continue
    # header for the account
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c = put(ws, r, 1, "%d   %s      (%s - %s balance, opening %s)" % (
        code, name, typ, norm, format(OPEN.get(code, 0.0), ",.2f")), font=Font(
        name="Calibri", size=10, bold=True, color=NAVY))
    c.fill = PatternFill("solid", fgColor=PALE)
    c.border = BOX
    r += 1
    hdr_row = r
    headers(ws, r, DET_HDR)
    r += 1
    first_line_row = r
    prev_bal_cell = None
    for acct, jrow, date, memo, prop, dr, cr, jno in (lines + aj_lines):
        put(ws, r, 1, date, fmt=DATE, align="center")
        put(ws, r, 2, jno, align="center")
        put(ws, r, 3, memo)
        put(ws, r, 4, prop, align="center")
        if jrow is None:
            put(ws, r, 5, dr if dr else None, fmt=MONEY, align="right")
            put(ws, r, 6, cr if cr else None, fmt=MONEY, align="right")
        else:
            put(ws, r, 5, "='%s'!H%d" % (S_JE, jrow), fmt=MONEY, align="right")
            put(ws, r, 6, "='%s'!I%d" % (S_JE, jrow), fmt=MONEY, align="right")
        if r == first_line_row:
            put(ws, r, 7, "=$D$%d+E%d-F%d" % (gl_row[code], r, r), fmt=MONEY, align="right", font=F_BOLD)
        else:
            put(ws, r, 7, "=G%d+E%d-F%d" % (r - 1, r, r), fmt=MONEY, align="right", font=F_BOLD)
        detail_map.append((code, r))
        r += 1
    # account total
    put(ws, r, 1, "", fill=FILL_TOTAL)
    put(ws, r, 3, "Account total / closing balance", font=F_BOLD, fill=FILL_TOTAL)
    put(ws, r, 5, "=SUM(E%d:E%d)" % (first_line_row, r - 1), fmt=MONEY, font=F_BOLD,
        align="right", fill=FILL_TOTAL)
    put(ws, r, 6, "=SUM(F%d:F%d)" % (first_line_row, r - 1), fmt=MONEY, font=F_BOLD,
        align="right", fill=FILL_TOTAL)
    put(ws, r, 7, "=G%d" % (r - 1), fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
    r += 2
page(ws, "5:5")

# =====================================================================
# SHEET 5 - Trial balance
# =====================================================================
ws = ws5
widths(ws, {"A": 7, "B": 50, "C": 12, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14,
            "I": 14, "J": 14, "K": 14})
sheet_title(ws, 11, "TRIAL BALANCE - AS AT 31 DECEMBER 2025",
            "Opening balances + 50 operating transactions + 8 adjusting entries = closing balances. "
            "Total debits must equal total credits")
section(ws, 4, 1, 11, "TRIAL BALANCE")
headers(ws, 5, ["Acct", "Account", "Type", "Opening\nDr", "Opening\nCr",
                "Operating\nDr", "Operating\nCr", "Adjusting\nDr", "Adjusting\nCr",
                "Closing\nDr", "Closing\nCr"])
r = 6
TB_FIRST = r
tb_row = {}
for code, name, typ, norm, stmt in ACCOUNTS:
    g = gl_row[code]
    put(ws, r, 1, code, fmt="0", align="center", border=BOX)
    put(ws, r, 2, name, border=BOX)
    put(ws, r, 3, typ, border=BOX)
    put(ws, r, 4, "=IF('%s'!D%d>0,'%s'!D%d,0)" % (S_GL, g, S_GL, g), fmt=MONEY, align="right", border=BOX)
    put(ws, r, 5, "=IF('%s'!D%d<0,-'%s'!D%d,0)" % (S_GL, g, S_GL, g), fmt=MONEY, align="right", border=BOX)
    put(ws, r, 6, "=" + sumifs(S_JE, "H", code), fmt=MONEY, align="right", border=BOX)
    put(ws, r, 7, "=" + sumifs(S_JE, "I", code), fmt=MONEY, align="right", border=BOX)
    put(ws, r, 8, "=" + sumifs(S_AJ, "H", code), fmt=MONEY, align="right", border=BOX)
    put(ws, r, 9, "=" + sumifs(S_AJ, "I", code), fmt=MONEY, align="right", border=BOX)
    put(ws, r, 10, "=IF('%s'!G%d>0.004,'%s'!G%d,0)" % (S_GL, g, S_GL, g), fmt=MONEY,
        align="right", font=F_BOLD, border=BOX)
    put(ws, r, 11, "=IF('%s'!G%d<-0.004,-'%s'!G%d,0)" % (S_GL, g, S_GL, g), fmt=MONEY,
        align="right", font=F_BOLD, border=BOX)
    tb_row[code] = r
    r += 1
TB_LAST = r - 1
put(ws, r, 2, "TOTALS", font=F_BOLD, fill=FILL_TOTAL)
for col in range(4, 12):
    L = get_column_letter(col)
    put(ws, r, col, "=SUM(%s%d:%s%d)" % (L, TB_FIRST, L, TB_LAST), fmt=MONEY, font=F_BOLD,
        align="right", fill=FILL_TOTAL)
TBTOT = r
r += 1
put(ws, r, 2, "Debit / credit difference (must be zero)", font=F_BOLD)
put(ws, r, 10, "=ROUND(J%d-K%d,2)" % (TBTOT, TBTOT), fmt=MONEY, font=F_BOLD, align="right")
put(ws, r, 11, '=IF(ROUND(J%d-K%d,2)=0,"BALANCED","CHECK")' % (TBTOT, TBTOT), font=F_BOLD,
    align="center", fill=FILL_OK, border=BOX)
r += 2
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
put(ws, r, 1, "Reading the trial balance: revenue accounts carry credit balances in the Closing Cr column; "
              "expenses and assets carry debit balances in Closing Dr. Equity distributions sit in Closing Dr "
              "because they reduce equity (contra-equity). The credit balances on revenue and expense accounts are "
              "cleared only by the year-end closing entries, which this workbook deliberately leaves unposted so the "
              "trial balance can be reviewed first.", font=F_ITAL)
ws.row_dimensions[r].height = 40
page(ws, "5:5")
ws.freeze_panes = "A6"

# =====================================================================
# SHEET 6 - Adjusting entries
# =====================================================================
ws = ws6
widths(ws, {"A": 9, "B": 12, "C": 20, "D": 8, "E": 44, "F": 50, "G": 8, "H": 14, "I": 14, "J": 9})
sheet_title(ws, 10, "ADJUSTING ENTRIES - 31 DECEMBER 2025",
            "Eight year-end adjustments that convert cash records to accrual accounting: prepaids, accruals, "
            "depreciation and bank items")
put(ws, 3, 6, "Totals ->", font=F_BOLD, align="right")
put(ws, 3, 8, "=SUM(H%d:H%d)" % (AJ_FIRST, AJ_LAST), fmt=MONEY, font=F_BOLD, align="right")
put(ws, 3, 9, "=SUM(I%d:I%d)" % (AJ_FIRST, AJ_LAST), fmt=MONEY, font=F_BOLD, align="right")
put(ws, 3, 10, '=IF(ROUND(H3-I3,2)=0,"OK","CHECK")', font=F_BOLD, align="center",
    fill=FILL_OK, border=BOX)
put(ws, 3, 5, "Adjustments - debits must equal credits", font=F_ITAL)

HDR = 5
headers(ws, HDR, ["AJ no.", "Date", "Category", "Acct", "Account name", "Adjustment and basis",
                  "Property", "Debit", "Credit", "Check"])
r = 6
assert r == AJ_FIRST
for i, a in enumerate(AJ, start=1):
    band = FILL_BAND if i % 2 == 0 else None
    for acct, dr, cr, prop, memo in a["lines"]:
        put(ws, r, 1, a["no"], align="center", fill=band)
        put(ws, r, 2, a["date"], fmt=DATE, align="center", fill=band)
        put(ws, r, 3, a["type"], font=F_SMALL, fill=band)
        put(ws, r, 4, acct, fmt="0", align="center", fill=band)
        put(ws, r, 5, "=IFERROR(VLOOKUP($D%d,%s,2,FALSE),\"?\")" % (r, COA), fill=band)
        put(ws, r, 6, memo, fill=band)
        put(ws, r, 7, prop, align="center", fill=band)
        put(ws, r, 8, dr if dr else None, fmt=MONEY, align="right", fill=band)
        put(ws, r, 9, cr if cr else None, fmt=MONEY, align="right", fill=band)
        put(ws, r, 10, '=IF(ROUND(SUMIF($A$6:$A$999,$A%d,$H$6:$H$999)'
                       '-SUMIF($A$6:$A$999,$A%d,$I$6:$I$999),2)=0,"OK","ERR")' % (r, r),
            font=F_SMALL, align="center", fill=band)
        r += 1
assert r - 1 == AJ_LAST, (r - 1, AJ_LAST)
put(ws, r, 6, "Total adjusting entries", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 8, "=SUM(H%d:H%d)" % (AJ_FIRST, AJ_LAST), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 9, "=SUM(I%d:I%d)" % (AJ_FIRST, AJ_LAST), fmt=MONEY, font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 10, '=IF(ROUND(H%d-I%d,2)=0,"OK","CHECK")' % (r, r), font=F_BOLD,
    align="center", fill=FILL_OK, border=BOX)
r += 2

# explanation block
section(ws, r, 1, 10, "WHY EACH ADJUSTMENT EXISTS")
r += 1
expl = [
    ("AJ-01 / AJ-02", "Depreciation is a non-cash allocation of the buildings (40 years) and FF&E (7 years) over "
                      "the periods that consume them. Nothing was paid in December - the expense still belongs to 2025."),
    ("AJ-03", "Prepaid insurance: 54,000 was paid on 1 October for 12 months of cover. Only 3 months (13,500) "
              "expired in 2025, plus 9,000 carried in from the prior policy, so 22,500 is expensed and 40,500 stays "
              "as an asset for 2026."),
    ("AJ-04 / AJ-05 / AJ-06", "Accrued expenses: property tax on the Maple Street shops, the Q4 operating cost "
                              "settlement and December mortgage interest were all incurred in 2025 but invoiced or "
                              "paid in 2026. Recording them now is the core of accrual accounting."),
    ("AJ-07 / AJ-08", "Bank items known only from the statement: service charges of 185 and the returned NSF "
                      "cheque of 3,600. Both are booked here and then appear as reconciling items on the Bank "
                      "Reconciliation sheet."),
]
for a, b in expl:
    put(ws, r, 1, a, font=F_BOLD)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
    c = put(ws, r, 2, b, font=F_BODY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 32
    r += 1

# loan schedule support
r += 1
section(ws, r, 1, 10, "SUPPORTING SCHEDULE - MORTGAGE INTEREST (supports AJ-06 and the interest expense)")
r += 1
headers(ws, r, ["Loan", "", "Opening balance 1-Jan-25", "Rate", "Term (mths)", "Monthly payment",
                "Payments in 2025", "Interest paid 2025", "Principal paid 2025", "Closing balance 31-Dec-25"])
r += 1
loan_first = r
for spec, sch in ((LOAN_A, SCHED_A), (LOAN_B, SCHED_B)):
    put(ws, r, 1, spec["name"], font=F_BOLD)
    put(ws, r, 2, "")
    put(ws, r, 3, spec["principal"], fmt=MONEY, align="right")
    put(ws, r, 4, spec["rate"], fmt="0.00%", align="right")
    put(ws, r, 5, spec["term"], fmt="0", align="right")
    put(ws, r, 6, sch[0]["payment"], fmt=MONEY, align="right")
    put(ws, r, 7, len(sch), fmt="0", align="right")
    put(ws, r, 8, round(sum(x["interest"] for x in sch), 2), fmt=MONEY, align="right")
    put(ws, r, 9, round(sum(x["principal"] for x in sch), 2), fmt=MONEY, align="right")
    put(ws, r, 10, sch[-1]["balance"], fmt=MONEY, align="right")
    r += 1
put(ws, r, 1, "Accrued December 2025 interest (AJ-06)", font=F_BOLD, fill=FILL_TOTAL)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
put(ws, r, 8, r2(INT_A_DEC + INT_B_DEC), fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
acc_row = r
r += 1
put(ws, r, 1, "Total 2025 interest expense = interest paid less the Dec-2024 element of the 1-Jan "
              "instalment plus the December 2025 accrual", font=F_BOLD, fill=FILL_TOTAL)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
put(ws, r, 7, "Total 2025", font=F_BOLD, align="right", fill=FILL_TOTAL)
put(ws, r, 8, "=SUM(H%d:H%d)-C%d*D%d/12+H%d" % (loan_first, loan_first + 1, loan_first,
                                                loan_first, acc_row),
    fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
put(ws, r, 9, '=IF(ROUND(H%d-\'%s\'!J%d,2)=0,"OK","CHECK")' % (r, S_TB, tb_row[5080]),
    font=F_BOLD, align="center", fill=FILL_OK, border=BOX)
put(ws, r, 10, "agrees to P&L", font=F_SMALL)
r += 2
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
put(ws, r, 1, "The 1-Jan-2025 instalment settled December 2024 interest already accrued in opening balances "
              "(5,606.25), which is why 2025 interest expense equals 12 months of Rivergate interest and 10 months "
              "of Maple Street interest even though only 12 and 9 instalments were paid.", font=F_ITAL)
ws.row_dimensions[r].height = 30
page(ws, "5:5")
ws.freeze_panes = "A6"

# =====================================================================
# SHEET 7 - Bank reconciliation
# =====================================================================
ws = ws7
widths(ws, {"A": 4, "B": 46, "C": 16, "D": 14, "E": 16, "F": 46})
sheet_title(ws, 6, "BANK RECONCILIATION - 31 DECEMBER 2025",
            "First Harbor Bank operating account no. xx4417 against the cash book (General Ledger account 1000)")

r = 4
section(ws, r, 2, 6, "STEP 1 - BALANCE PER BANK STATEMENT (as reported by First Harbor Bank)")
r += 1
bs_rows = [
    ("Opening balance, 1 January 2025", OPEN[1000]),
    ("Add: deposits and other credits cleared", r2(RECEIPTS - DIT)),
    ("Less: cheques and other debits cleared", r2(-(PAYMENTS + CASH_OUT_AJ - OS_TOTAL))),
]
bstart = r
for label, val in bs_rows:
    put(ws, r, 2, label)
    put(ws, r, 3, val, fmt=MONEY, align="right")
    r += 1
put(ws, r, 2, "Closing balance per bank statement, 31 December 2025", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 3, "=SUM(C%d:C%d)" % (bstart, r - 1), fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
BANK_CELL = "C%d" % r
r += 2

section(ws, r, 2, 6, "STEP 2 - ITEMS IN THE CASH BOOK NOT YET ON THE STATEMENT")
r += 1
headers(ws, r, ["", "Item", "Date", "Reference", "Amount", "Explanation"])
r += 1
put(ws, r, 2, "Deposit in transit - tenant receipts banked 31 December")
put(ws, r, 3, dt.date(2025, 12, 31), fmt=DATE, align="center")
put(ws, r, 4, "Deposit slip 4471", align="center")
put(ws, r, 5, DIT, fmt=MONEY, align="right")
put(ws, r, 6, "After the bank's cut-off; credited 2 January 2026.", font=F_SMALL)
DIT_CELL = "E%d" % r
r += 1
put(ws, r, 2, "Outstanding cheque - H2 owner distribution")
put(ws, r, 3, dt.date(2025, 12, 31), fmt=DATE, align="center")
put(ws, r, 4, "Cheque 2255", align="center")
put(ws, r, 5, OS_TOTAL, fmt=MONEY, align="right")
put(ws, r, 6, "Issued 31 December, presented January 2026.", font=F_SMALL)
OS_CELL = "E%d" % r
r += 1
put(ws, r, 2, "Total outstanding cheques", font=F_BOLD)
put(ws, r, 5, "=SUM(E%d:E%d)" % (r - 1, r - 1), fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
OS_TOT_CELL = "E%d" % r
r += 2

section(ws, r, 2, 6, "STEP 3 - RECONCILIATION")
r += 1
headers(ws, r, ["", "Reconciling item", "", "", "Amount", "Working"])
r += 1
start = r
put(ws, r, 2, "Balance per bank statement", font=F_BOLD)
put(ws, r, 5, "=" + BANK_CELL, fmt=MONEY, align="right")
r += 1
put(ws, r, 2, "Add: deposit in transit")
put(ws, r, 5, "=" + DIT_CELL, fmt=MONEY, align="right")
put(ws, r, 6, "Cash book already includes it", font=F_SMALL)
r += 1
put(ws, r, 2, "Less: outstanding cheques")
put(ws, r, 5, "=-" + OS_TOT_CELL, fmt=MONEY, align="right")
put(ws, r, 6, "Cash book already deducts it", font=F_SMALL)
r += 1
put(ws, r, 2, "ADJUSTED BANK BALANCE", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 5, "=SUM(E%d:E%d)" % (start, r - 1), fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
ADJ_BANK = "E%d" % r
r += 2

put(ws, r, 2, "Balance per cash book (General Ledger 1000)", font=F_BOLD)
put(ws, r, 5, "='%s'!G%d" % (S_GL, gl_row[1000]), fmt=MONEY, align="right", font=F_BOLD)
BOOK_CELL = "E%d" % r
r += 1
put(ws, r, 2, "Add: interest credited by bank not yet recorded", font=F_BODY)
put(ws, r, 5, 0, fmt=MONEY, align="right")
put(ws, r, 6, "None - interest recorded 31 December", font=F_SMALL)
r += 1
put(ws, r, 2, "Less: bank charges and NSF cheque not yet recorded", font=F_BODY)
put(ws, r, 5, 0, fmt=MONEY, align="right")
put(ws, r, 6, "Booked via AJ-07 and AJ-08, so already in the cash book", font=F_SMALL)
r += 1
put(ws, r, 2, "ADJUSTED CASH BOOK BALANCE", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 5, "=SUM(E%d:E%d)" % (int(BOOK_CELL[1:]), r - 1), fmt=MONEY, font=F_BOLD,
    align="right", fill=FILL_TOTAL)
ADJ_BOOK = "E%d" % r
r += 2
put(ws, r, 2, "Difference - must be zero", font=F_BOLD)
put(ws, r, 5, "=ROUND(%s-%s,2)" % (ADJ_BANK, ADJ_BOOK), fmt=MONEY, font=F_BOLD, align="right")
put(ws, r, 6, '=IF(ROUND(%s-%s,2)=0,"RECONCILED","INVESTIGATE")' % (ADJ_BANK, ADJ_BOOK),
    font=F_BOLD, align="center", fill=FILL_OK, border=BOX)
r += 2

section(ws, r, 2, 6, "MEMO - ITEMS ALREADY RECORDED THROUGH THE ADJUSTING ENTRIES")
r += 1
memo = [
    ("AJ-07", "Bank service charges 185 - on the statement, not in the cash book before adjustment."),
    ("AJ-08", "NSF cheque 3,600 from Bloom & Co. - deposit reversed by the bank; the receivable is reinstated "
              "because the tenant still owes the rent."),
]
for a, b in memo:
    put(ws, r, 2, a, font=F_BOLD)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    c = put(ws, r, 3, b, font=F_BODY)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 28
    r += 1
r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
put(ws, r, 2, "Both sides are proved to the same figure. The bank statement and the cash book differ only by timing, "
              "which is exactly what a reconciliation proves.", font=F_ITAL)
page(ws)

# =====================================================================
# SHEET 8 - Profit & loss
# =====================================================================
ws = ws8
widths(ws, {"A": 4, "B": 48, "C": 16, "D": 12, "E": 46})
sheet_title(ws, 5, "PROFIT & LOSS STATEMENT - YEAR ENDED 31 DECEMBER 2025",
            "Accrual basis. Property revenue and operating expenses, then net operating income, then other items")
r = 4
put(ws, r, 2, "Harborview Commercial Properties, LLC", font=Font(name="Calibri", size=11, bold=True, color=NAVY))
put(ws, r, 5, "All amounts in USD", font=F_ITAL, align="right")
r += 2
headers(ws, r, ["", "Account", "Amount", "% of revenue", "Notes"])
r += 1
pl_start = r


def pl_line(label, code, note="", bold=False, fill=None, top=False, indent=0, sign=1):
    global r
    c = put(ws, r, 2, ("    " * indent) + label, font=F_BOLD if bold else F_BODY, fill=fill)
    if top:
        c.border = TOPRULE
    put(ws, r, 3, "=" + ("-" if sign < 0 else "") + "'%s'!%s%d"
        % (S_TB, "K" if TYPE[code] == "Revenue" else "J", tb_row[code]),
        fmt=MONEY, align="right", font=F_BOLD if bold else F_BODY, fill=fill)
    if note:
        put(ws, r, 5, note, font=F_SMALL, fill=fill)
    r += 1
    return r - 1


def pl_sub(label, formula, note="", bold=True, fill=FILL_TOTAL, top=False, double=False):
    global r
    c = put(ws, r, 2, label, font=F_BOLD if bold else F_BODY, fill=fill)
    if top:
        c.border = TOPRULE
    cc = put(ws, r, 3, formula, fmt=MONEY, align="right", font=F_BOLD if bold else F_BODY, fill=fill)
    if double:
        cc.border = DBLRULE
    if note:
        put(ws, r, 5, note, font=F_SMALL, fill=fill)
    r += 1
    return r - 1


put(ws, r, 2, "PROPERTY REVENUE", font=F_BOLD, fill=FILL_SEC).font = F_SEC
ws.cell(row=r, column=3).fill = FILL_SEC
ws.cell(row=r, column=4).fill = FILL_SEC
ws.cell(row=r, column=5).fill = FILL_SEC
r += 1
rev_first = r
pl_line("Rental revenue - offices (Rivergate)", 4010, "3 tenants, 24,000/month")
pl_line("Rental revenue - retail (Maple Street)", 4020, "3 tenants, 13,000/month from March")
pl_line("Expense recoveries (CAM / NNN)", 4030, "Recoverable operating costs re-billed to tenants")
pl_line("Parking revenue", 4040)
pl_line("Other property income (antenna lease)", 4050)
pl_line("Late fee income", 4060, "2,250 charged on overdue rent")
rev_last = r - 1
TOTREV = pl_sub("Total property revenue", "=SUM(C%d:C%d)" % (rev_first, rev_last), "", fill=FILL_TOTAL)
r += 1

put(ws, r, 2, "PROPERTY OPERATING EXPENSES", font=F_BOLD, fill=FILL_SEC).font = F_SEC
for col in (3, 4, 5):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
opex_first = r
pl_line("Property taxes", 5010, "Rivergate 42,000 paid; Maple Street 12,000 accrued (AJ-04)")
pl_line("Property insurance", 5020, "22,500 expired of the 54,000 annual premium (AJ-03)")
pl_line("Utilities", 5030)
pl_line("Repairs & maintenance", 5040, "Includes the 6,200 roof repair")
pl_line("Cleaning & landscaping", 5050)
pl_line("Property management fees", 5060, "700/month Rivergate + 540/month Maple Street")
pl_line("Bad debt expense", 5100, "Nil - the NSF cheque was reinstated as a receivable, not written off")
opex_last = r - 1
TOTOPEX = pl_sub("Total property operating expenses", "=SUM(C%d:C%d)" % (opex_first, opex_last))
NOI = pl_sub("NET OPERATING INCOME (NOI)", "=C%d-C%d" % (TOTREV, TOTOPEX),
             "Property-level earnings before depreciation, financing and corporate items", double=True)
r += 1

put(ws, r, 2, "OTHER INCOME AND EXPENSES", font=F_BOLD, fill=FILL_SEC).font = F_SEC
for col in (3, 4, 5):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
oth_first = r
pl_line("Interest income", 4900, "Bank interest credited on the operating account", bold=False)
pl_line("Bank charges", 5090, "Account analysis and wire fees (AJ-07)", sign=-1)
pl_line("Depreciation expense", 5070, "Buildings 40 years, FF&E 7 years (AJ-01, AJ-02)", sign=-1)
pl_line("Interest expense - mortgages", 5080, "Rivergate 5.85%, Maple Street 6.25% (incl. AJ-06)",
        sign=-1)
oth_last = r - 1
TOTOTH = pl_sub("Total other income and expenses", "=SUM(C%d:C%d)" % (oth_first, oth_last))
r += 1
NIPL = pl_sub("NET INCOME", "=C%d+C%d" % (NOI, TOTOTH),
              "Pass-through LLC: no entity-level income tax", double=True)

# percent column + total revenue percent
for row in range(rev_first, r):
    if ws.cell(row=row, column=3).value is not None:
        ws.cell(row=row, column=4).value = "=IF($C$%d=0,\"\",C%d/$C$%d)" % (TOTREV, row, TOTREV)
        ws.cell(row=row, column=4).number_format = PCT
        ws.cell(row=row, column=4).font = F_BODY
        ws.cell(row=row, column=4).alignment = Alignment(horizontal="right")

r += 1
put(ws, r, 2, "Key ratios", font=F_BOLD, fill=FILL_SEC).font = F_SEC
for col in (3, 4, 5):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
put(ws, r, 2, "NOI margin (NOI / property revenue)")
put(ws, r, 3, "=C%d/C%d" % (NOI, TOTREV), fmt=PCT, align="right")
r += 1
put(ws, r, 2, "Operating expense ratio (opex / property revenue)")
put(ws, r, 3, "=C%d/C%d" % (TOTOPEX, TOTREV), fmt=PCT, align="right")
r += 1
put(ws, r, 2, "Interest cover (NOI / interest expense)")
put(ws, r, 3, "=C%d/'%s'!J%d" % (NOI, S_TB, tb_row[5080]), fmt="0.00x", align="right")
r += 1
put(ws, r, 2, "Distributions to owners (shown in equity, not here)")
put(ws, r, 3, "='%s'!J%d" % (S_TB, tb_row[3020]), fmt=MONEY, align="right")
put(ws, r, 5, "A distribution of equity - never an operating expense", font=F_SMALL)
r += 2
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
put(ws, r, 2, "Accrual note: December 2025 rent (42,100) was earned but not collected at year end and is included "
              "in revenue above. Conversely the 54,000 insurance premium paid in October is not all expensed - "
              "only the 13,500 consumed in 2025 is, with 40,500 carried forward as an asset.", font=F_ITAL)
ws.row_dimensions[r].height = 30
PL_NI_CELL = "C%d" % NIPL
PL_NOI_CELL = "C%d" % NOI
PL_REV_CELL = "C%d" % TOTREV
PL_OPEX_CELL = "C%d" % TOTOPEX
page(ws)

# =====================================================================
# SHEET 9 - Balance sheet
# =====================================================================
ws = ws9
widths(ws, {"A": 4, "B": 50, "C": 16, "D": 16, "E": 6, "F": 40})
sheet_title(ws, 4, "BALANCE SHEET - AS AT 31 DECEMBER 2025",
            "Current year from the trial balance; comparative column from the audited opening balances at "
            "31 December 2024")
r = 4
put(ws, r, 2, "Harborview Commercial Properties, LLC", font=Font(name="Calibri", size=11, bold=True, color=NAVY))
r += 2
headers(ws, r, ["", "Account", "31-Dec-2025", "31-Dec-2024", "", "Note"])
r += 1
bs_start = r


def bs_line(label, code, note="", bold=False, fill=None, indent=0, top=False):
    global r
    c = put(ws, r, 2, ("    " * indent) + label, font=F_BOLD if bold else F_BODY, fill=fill)
    if top:
        c.border = TOPRULE
    if code in (1515, 1525):        # contra-asset: credit balance, shown as a deduction
        cur, prior, mult = "K", "E", -1
    elif code == 3020:              # contra-equity
        cur, prior, mult = "J", "D", -1
    elif TYPE[code] == "Asset":
        cur, prior, mult = "J", "D", 1
    else:
        cur, prior, mult = "K", "E", 1
    put(ws, r, 3, "=%s'%s'!%s%d" % ("-" if mult < 0 else "", S_TB, cur, tb_row[code]),
        fmt=MONEY, align="right", font=F_BOLD if bold else F_BODY, fill=fill)
    put(ws, r, 4, "=%s'%s'!%s%d" % ("-" if mult < 0 else "", S_TB, prior, tb_row[code]),
        fmt=MONEY, align="right", font=F_BODY, fill=fill)
    if note:
        put(ws, r, 6, note, font=F_SMALL, fill=fill)
    r += 1
    return r - 1


def bs_sub(label, formula, note="", fill=FILL_TOTAL, top=True, double=False):
    global r
    c = put(ws, r, 2, label, font=F_BOLD, fill=fill)
    if top:
        c.border = TOPRULE
    cc = put(ws, r, 3, formula, fmt=MONEY, align="right", font=F_BOLD, fill=fill)
    if double:
        cc.border = DBLRULE
    put(ws, r, 4, re.sub(r"C(\d+)", r"D\1", formula), fmt=MONEY, align="right",
        font=F_BOLD, fill=fill)
    if note:
        put(ws, r, 6, note, font=F_SMALL, fill=fill)
    r += 1
    return r - 1


put(ws, r, 2, "ASSETS", font=F_SEC, fill=FILL_SEC)
for col in (3, 4, 6):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
put(ws, r, 2, "Current assets", font=F_BOLD)
r += 1
ca_first = r
bs_line("Cash - operating bank account", 1000, indent=1, note="Reconciled to First Harbor Bank")
bs_line("Tenant receivables", 1100, indent=1, note="December 2025 rent + returned NSF cheque")
bs_line("Prepaid insurance", 1200, indent=1, note="9 months of cover to 30-Sep-2026")
TCA = bs_sub("Total current assets", "=SUM(C%d:C%d)" % (ca_first, r - 1), top=True)
r += 1
put(ws, r, 2, "Property and equipment", font=F_BOLD)
r += 1
pe_first = r
bs_line("Land", 1500, indent=1)
bs_line("Building & improvements", 1510, indent=1, note="Includes 40,000 capitalised acquisition costs")
bs_line("Less: accumulated depreciation - building", 1515, indent=1)
bs_line("Furniture, fixtures & equipment", 1520, indent=1)
bs_line("Less: accumulated depreciation - FF&E", 1525, indent=1)
TPE = bs_sub("Net property and equipment", "=SUM(C%d:C%d)" % (pe_first, r - 1), top=True)
TASSET = bs_sub("TOTAL ASSETS", "=C%d+C%d" % (TCA, TPE), double=True)
r += 1

put(ws, r, 2, "LIABILITIES AND EQUITY", font=F_SEC, fill=FILL_SEC)
for col in (3, 4, 6):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
put(ws, r, 2, "Current liabilities", font=F_BOLD)
r += 1
cl_first = r
bs_line("Accounts payable", 2010, indent=1, note="Q4 operating settlement invoiced 15-Jan-2026")
bs_line("Accrued expenses", 2020, indent=1, note="Property tax 12,000 + December interest")
bs_line("Tenant security deposits", 2030, indent=1, note="Refundable deposits held in trust")
TCL = bs_sub("Total current liabilities", "=SUM(C%d:C%d)" % (cl_first, r - 1), top=True)
r += 1
put(ws, r, 2, "Long-term liabilities", font=F_BOLD)
r += 1
lt_first = r
bs_line("Mortgage payable - Rivergate Office Center", 2100, indent=1, note="5.85%, 240-month amortisation")
bs_line("Mortgage payable - Maple Street Retail Shops", 2110, indent=1, note="6.25%, 300-month amortisation")
TLT = bs_sub("Total long-term liabilities", "=SUM(C%d:C%d)" % (lt_first, r - 1), top=True)
TLIAB = bs_sub("TOTAL LIABILITIES", "=C%d+C%d" % (TCL, TLT))
r += 1

put(ws, r, 2, "Equity", font=F_BOLD)
r += 1
eq_first = r
bs_line("Owner's capital", 3010, indent=1, note="Opening 1,890,000 + 700,000 contribution (Mar-2025)")
put(ws, r, 2, "    Current year net income", font=F_BODY)
put(ws, r, 3, "='%s'!%s" % (S_PL, PL_NI_CELL), fmt=MONEY, align="right")
put(ws, r, 4, 0, fmt=MONEY, align="right")
put(ws, r, 6, "From the Profit & Loss statement", font=F_SMALL)
NIBS = r
r += 1
bs_line("Less: owner distributions", 3020, indent=1, note="Distributions of equity - not expenses")
TEQ = bs_sub("TOTAL EQUITY", "=C%d+C%d+C%d" % (eq_first, NIBS, NIBS + 1), top=True)
r += 1
TLE = bs_sub("TOTAL LIABILITIES AND EQUITY", "=C%d+C%d" % (TLIAB, TEQ), double=True)
r += 1
put(ws, r, 2, "Check: total assets less (liabilities + equity)", font=F_BOLD)
put(ws, r, 3, "=ROUND(C%d-C%d,2)" % (TASSET, TLE), fmt=MONEY, font=F_BOLD, align="right")
put(ws, r, 4, "=ROUND(D%d-D%d,2)" % (TASSET, TLE), fmt=MONEY, align="right", font=F_BOLD)
put(ws, r, 6, '=IF(AND(ROUND(C%d,2)=0,ROUND(D%d,2)=0),"BALANCED","CHECK")' % (r, r),
    font=F_BOLD, align="center", fill=FILL_OK, border=BOX)
r += 2
put(ws, r, 2, "Equity roll-forward", font=F_BOLD)
r += 1
put(ws, r, 2, "Opening equity 1 January 2025")
put(ws, r, 3, -OPEN[3010], fmt=MONEY, align="right")
r += 1
put(ws, r, 2, "Add: owner capital contribution")
put(ws, r, 3, "='%s'!F%d" % (S_TB, tb_row[3010]), fmt=MONEY, align="right")
ws.cell(row=r, column=3).value = "='%s'!G%d-'%s'!D%d" % (S_TB, tb_row[3010], S_TB, tb_row[3010])
r += 1
put(ws, r, 2, "Add: net income for 2025")
put(ws, r, 3, "='%s'!%s" % (S_PL, PL_NI_CELL), fmt=MONEY, align="right")
r += 1
put(ws, r, 2, "Less: owner distributions")
put(ws, r, 3, "='%s'!J%d" % (S_TB, tb_row[3020]), fmt=MONEY, align="right")
ws.cell(row=r, column=3).value = "=-'%s'!J%d" % (S_TB, tb_row[3020])
r += 1
put(ws, r, 2, "Closing equity 31 December 2025", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 3, "=SUM(C%d:C%d)" % (r - 4, r - 1), fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
put(ws, r, 6, '=IF(ROUND(C%d-C%d,2)=0,"TIES TO BALANCE SHEET","CHECK")' % (r, TEQ),
    font=F_BOLD, align="center", fill=FILL_OK, border=BOX)
r += 2
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
put(ws, r, 2, "Note: owner distributions of 60,000 appear as a deduction from equity, not as an expense in the "
              "income statement, so net income is unaffected by them.", font=F_ITAL)
page(ws)

# =====================================================================
# SHEET 10 - Cash flow statement
# =====================================================================
ws = ws10
widths(ws, {"A": 4, "B": 58, "C": 16, "D": 6, "E": 48})
sheet_title(ws, 5, "CASH FLOW STATEMENT - YEAR ENDED 31 DECEMBER 2025",
            "Indirect method. Operating, investing and financing flows, reconciled to the movement in cash")
r = 4
put(ws, r, 2, "Harborview Commercial Properties, LLC", font=Font(name="Calibri", size=11, bold=True, color=NAVY))
r += 2
headers(ws, r, ["", "Description", "Amount", "", "Source / working"])
r += 1
cf_start = r


def cf_line(label, formula, note="", bold=False, fill=None, indent=0, top=False):
    global r
    c = put(ws, r, 2, ("    " * indent) + label, font=F_BOLD if bold else F_BODY, fill=fill)
    if top:
        c.border = TOPRULE
    put(ws, r, 3, formula, fmt=MONEY, align="right", font=F_BOLD if bold else F_BODY, fill=fill,
        border=DBLRULE if bold and indent == 0 and fill == FILL_TOTAL else None)
    if note:
        put(ws, r, 5, note, font=F_SMALL, fill=fill)
    r += 1
    return r - 1


put(ws, r, 2, "CASH FLOWS FROM OPERATING ACTIVITIES", font=F_SEC, fill=FILL_SEC)
for col in (3, 5):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
cf_line("Net income", "='%s'!%s" % (S_PL, PL_NI_CELL), "From the Profit & Loss statement")
r += 1
put(ws, r, 2, "Adjustments for non-cash items", font=F_BOLD, indent=1)
r += 1
DEP = cf_line("Depreciation - buildings and FF&E", "='%s'!J%d" % (S_TB, tb_row[5070]),
              "Added back: no cash outflow", indent=2)
r += 1
put(ws, r, 2, "Changes in working capital", font=F_BOLD, indent=1)
r += 1
d_ar = cf_line("Increase in tenant receivables", "=-('{s}'!J{a}-'{s}'!D{a})".format(s=S_TB, a=tb_row[1100]),
               "Accrued December rent and the returned NSF cheque", indent=2)
d_pre = cf_line("Increase in prepaid insurance", "=-('{s}'!J{p}-'{s}'!D{p})".format(s=S_TB, p=tb_row[1200]),
                "9 months of cover carried into 2026", indent=2)
d_ap = cf_line("Increase in accounts payable", "='{s}'!K{a}-'{s}'!E{a}".format(s=S_TB, a=tb_row[2010]),
               "Q4 operating settlement", indent=2)
d_ae = cf_line("Increase in accrued expenses", "='{s}'!K{a}-'{s}'!E{a}".format(s=S_TB, a=tb_row[2020]),
               "Property tax, December interest", indent=2)
r += 1
CFO = cf_line("NET CASH FROM OPERATING ACTIVITIES",
              "=SUM(C%d:C%d)" % (cf_start, r - 1), bold=True, fill=FILL_TOTAL)
r += 1

put(ws, r, 2, "CASH FLOWS FROM INVESTING ACTIVITIES", font=F_SEC, fill=FILL_SEC)
for col in (3, 5):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
i_first = r
cf_line("Acquisition of Maple Street Retail Shops - land",
        "=-('{s}'!J{a}-'{s}'!D{a})".format(s=S_TB, a=tb_row[1500]), indent=2)
cf_line("Acquisition - building and capitalised acquisition costs",
        "=-('{s}'!J{a}-'{s}'!D{a})".format(s=S_TB, a=tb_row[1510]), indent=2)
cf_line("Furniture, fixtures & equipment purchased", "=0", "None in 2025 - opening balance only", indent=2)
r += 1
CFI = cf_line("NET CASH USED IN INVESTING ACTIVITIES", "=SUM(C%d:C%d)" % (i_first, r - 1),
              bold=True, fill=FILL_TOTAL)
r += 1

put(ws, r, 2, "CASH FLOWS FROM FINANCING ACTIVITIES", font=F_SEC, fill=FILL_SEC)
for col in (3, 5):
    ws.cell(row=r, column=col).fill = FILL_SEC
r += 1
f_first = r
cf_line("Owner capital contribution", "='{s}'!K{a}-'{s}'!E{a}".format(s=S_TB, a=tb_row[3010]),
        "March 2025, to fund the acquisition", indent=2)
proceeds_row = cf_line("Proceeds from Maple Street mortgage draw", "=1400000",
                       "1,400,000 drawn at closing on 1 March 2025", indent=2)
cf_line("Principal repayments - Maple Street mortgage",
        "='{s}'!K{a}-C{p}".format(s=S_TB, a=tb_row[2110], p=proceeds_row),
        "9 instalments Apr-Dec 2025", indent=2)
cf_line("Principal repayments - Rivergate mortgage",
        "='{s}'!K{a}-'{s}'!E{a}".format(s=S_TB, a=tb_row[2100]),
        "12 instalments Jan-Dec 2025", indent=2)
cf_line("Owner distributions", "=-'{s}'!J{a}".format(s=S_TB, a=tb_row[3020]),
        "Charged to equity, not to profit", indent=2)
cf_line("Tenant security deposits received, net of refunds",
        "='{s}'!K{a}-'{s}'!E{a}".format(s=S_TB, a=tb_row[2030]),
        "36,000 received less 6,000 refunded", indent=2)
r += 1
CFF = cf_line("NET CASH FROM FINANCING ACTIVITIES", "=SUM(C%d:C%d)" % (f_first, r - 1),
              bold=True, fill=FILL_TOTAL)
r += 2

put(ws, r, 2, "NET INCREASE / (DECREASE) IN CASH", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 3, "=C%d+C%d+C%d" % (CFO, CFI, CFF), fmt=MONEY, font=F_BOLD, align="right",
    fill=FILL_TOTAL)
NETCHG = r
r += 1
put(ws, r, 2, "Cash at 1 January 2025")
put(ws, r, 3, OPEN[1000], fmt=MONEY, align="right")
CASH0 = r
r += 1
put(ws, r, 2, "CASH AT 31 DECEMBER 2025", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 3, "=C%d+C%d" % (NETCHG, CASH0), fmt=MONEY, font=F_BOLD, align="right",
    fill=FILL_TOTAL)
CASH1 = r
put(ws, r, 5, '=IF(ROUND(C%d-\'%s\'!J%d,2)=0,"AGREES TO BALANCE SHEET","CHECK")'
    % (r, S_TB, tb_row[1000]), font=F_BOLD, align="left", fill=FILL_OK, border=BOX)
r += 2

section(ws, r, 2, 5, "SUPPLEMENTARY DISCLOSURES")
r += 1
def je_cash(acct):
    return ("=SUMIFS('{j}'!$H${f}:$H${l},'{j}'!$D${f}:$D${l},{a})"
            .format(j=S_JE, f=JE_FIRST, l=JE_LAST, a=acct))


sup = [
    ("Interest paid in cash (2025 expense paid, excluding December accrual)", je_cash(5080)),
    ("Cash paid for property taxes", je_cash(5010)),
    ("Cash paid for the annual insurance premium", je_cash(1200)),
    ("Cash paid to owners as distributions", "='%s'!J%d" % (S_TB, tb_row[3020])),
]
for label, f in sup:
    put(ws, r, 2, "    " + label)
    put(ws, r, 3, f, fmt=MONEY, align="right")
    r += 1
put(ws, r, 2, "    Interest paid excludes the 12,000 accrued for December 2025", font=F_ITAL)
r += 2
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
put(ws, r, 2, "Why operating cash flow exceeds net income: depreciation of 100,953 is a non-cash charge, and "
              "30,000 of accruals (Q4 settlement, property tax, December interest) are 2025 expenses that had not "
              "been paid at year end. Debt service, the acquisition and distributions sit in financing and investing, "
              "so they never flatter operating cash flow.", font=F_ITAL)
ws.row_dimensions[r].height = 44
page(ws)

# =====================================================================
# SHEET 11 - Management financial package
# =====================================================================
ws = ws11
widths(ws, {"A": 4, "B": 42, "C": 16, "D": 16, "E": 14, "F": 16, "G": 17})
sheet_title(ws, 7, "MANAGEMENT FINANCIAL PACKAGE - FY2025",
            "Consolidated property-level reporting pack for the members, prepared 15 February 2026")
r = 4
put(ws, r, 2, "To:  A. Whitfield and M. Okafor, Members", font=F_BOLD)
r += 1
put(ws, r, 2, "From:  Financial Controller", font=F_BODY)
r += 1
put(ws, r, 2, "Subject:  Results for the year ended 31 December 2025", font=F_BOLD)
r += 2
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
memo_txt = (
    "Harborview's first full year owning both properties produced property revenue of "
    + format(REV_PROP, ",.0f") + " and net operating income of " + format(NOI_PY, ",.0f")
    + ", a NOI margin of " + format(NOI_PY / REV_PROP, ".1%")
    + ". After depreciation of " + format(DEP_PY, ",.0f") + " and mortgage interest of "
    + format(INT_PY, ",.0f") + ", net income was " + format(NI, ",.0f")
    + ". Distributions of " + format(-BAL[3020], ",.0f")
    + " were paid out of equity and are excluded from profit. December rent, the Q4 operating settlement and "
      "the Maple Street property tax are accrued, so the result reflects what was earned, not what was banked.")
c = put(ws, r, 2, memo_txt, font=F_BODY)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[r].height = 74
r += 2

section(ws, r, 2, 7, "1.  KEY FINANCIAL INDICATORS")
r += 1
headers(ws, r, ["", "Indicator", "FY2025", "Reference / budget", "", "Comment", ""])
r += 1


def kpi(label, formula, fmt, ref="", comment=""):
    global r
    put(ws, r, 2, label, font=F_BOLD)
    put(ws, r, 3, formula, fmt=fmt, align="right")
    if ref:
        put(ws, r, 4, ref, font=F_SMALL, align="center")
    if comment:
        put(ws, r, 6, comment, font=F_SMALL)
    rr = r
    r += 1
    return rr


k_rev = kpi("Total property revenue", "='%s'!%s" % (S_PL, PL_REV_CELL), MONEY,
            "479,250 budget", "Maple Street contributed 10 months")
k_noi = kpi("Net operating income", "='%s'!%s" % (S_PL, PL_NOI_CELL), MONEY,
            "on budget", "Recoveries and parking held the margin")
k_marg = kpi("NOI margin", "=C%d/C%d" % (k_noi, k_rev), PCT, "55.0% target",
             "Strong for a two-property portfolio")
k_ni = kpi("Net income", "='%s'!%s" % (S_PL, PL_NI_CELL), MONEY, "-",
           "After depreciation and mortgage interest")
k_cfo = kpi("Cash from operations", "='%s'!C%d" % (S_CF, CFO), MONEY, "-",
            "Funds debt service and distributions")
k_dist = kpi("Owner distributions", "='%s'!J%d" % (S_TB, tb_row[3020]), MONEY, "60,000",
             "Charged to equity, never to profit")
k_ta = kpi("Total assets", "='%s'!C%d" % (S_BS, TASSET), MONEY, "-", "See balance sheet")
k_td = kpi("Total mortgage debt",
           "='{t}'!K{a}+'{t}'!K{b}".format(t=S_TB, a=tb_row[2100], b=tb_row[2110]),
           MONEY, "-", "Two facilities, both with First Harbor Bank")
k_lev = kpi("Debt to total assets", "=C%d/C%d" % (k_td, k_ta), PCT, "39% at 1 Jan 2025",
            "Leverage rose with the Maple Street mortgage")
k_ds = kpi("Annual debt service (12 instalments)", DS_PY, MONEY, "amortisation schedule",
           "Interest plus principal actually paid")
k_dscr = kpi("Debt service coverage (NOI / debt service)", "=C%d/C%d" % (k_noi, k_ds),
             "0.00x", "1.20x covenant", "Comfortable headroom")
r += 1

section(ws, r, 2, 7, "2.  PROPERTY-LEVEL OPERATING RESULTS")
r += 1
headers(ws, r, ["", "USD", "Rivergate (RIV)", "Maple Street (MAP)", "Corporate",
                "Total", "% of revenue"])
r += 1


def sar(sheet, f, l, acct, prop):
    dr = ("SUMIFS('{s}'!$H${f}:$H${l},'{s}'!$D${f}:$D${l},{a},'{s}'!$G${f}:$G${l},\"{p}\")"
          .format(s=sheet, f=f, l=l, a=acct, p=prop))
    cr = ("SUMIFS('{s}'!$I${f}:$I${l},'{s}'!$D${f}:$D${l},{a},'{s}'!$G${f}:$G${l},\"{p}\")"
          .format(s=sheet, f=f, l=l, a=acct, p=prop))
    return dr, cr


def prop_formula(acct, prop, debit_positive):
    parts = []
    for sheet, f, l in ((S_JE, JE_FIRST, JE_LAST), (S_AJ, AJ_FIRST, AJ_LAST)):
        dr, cr = sar(sheet, f, l, acct, prop)
        parts.append("(%s)-(%s)" % (dr, cr) if debit_positive else "(%s)-(%s)" % (cr, dr))
    return "=" + "+".join(parts)


pct_rows = []


def pr_row(label, code, kind, bold=False, indent=2, fill=None, top=False):
    global r
    c = put(ws, r, 2, ("    " * indent) + label, font=F_BOLD if bold else F_BODY, fill=fill)
    if top:
        c.border = TOPRULE
    for col, prop in ((3, "RIV"), (4, "MAP"), (5, "CORP")):
        put(ws, r, col, prop_formula(code, prop, kind == "exp"), fmt=MONEY, align="right",
            font=F_BOLD if bold else F_BODY, fill=fill)
    put(ws, r, 6, "=SUM(C%d:E%d)" % (r, r), fmt=MONEY, align="right",
        font=F_BOLD if bold else F_BODY, fill=fill)
    put(ws, r, 7, None, fmt=PCT, align="right", font=F_SMALL, fill=fill)
    pct_rows.append(r)
    rr = r
    r += 1
    return rr


def pr_total(label, first, last):
    global r
    put(ws, r, 2, label, font=F_BOLD, fill=FILL_TOTAL)
    for col in range(3, 7):
        L = get_column_letter(col)
        put(ws, r, col, "=SUM(%s%d:%s%d)" % (L, first, L, last), fmt=MONEY, font=F_BOLD,
            align="right", fill=FILL_TOTAL)
    put(ws, r, 7, None, fmt=PCT, align="right", font=F_BOLD, fill=FILL_TOTAL)
    rr = r
    r += 1
    return rr


put(ws, r, 2, "Property revenue", font=F_BOLD)
r += 1
rr_first = r
pr_row("Rental revenue - offices", 4010, "rev")
pr_row("Rental revenue - retail", 4020, "rev")
pr_row("Expense recoveries", 4030, "rev")
pr_row("Parking", 4040, "rev")
pr_row("Other property income", 4050, "rev")
pr_row("Late fees", 4060, "rev")
rr_last = r - 1
pr_tot_rev = pr_total("        Total property revenue", rr_first, rr_last)
r += 1

put(ws, r, 2, "Property operating expenses", font=F_BOLD)
r += 1
px_first = r
pr_row("Property taxes", 5010, "exp")
pr_row("Property insurance", 5020, "exp")
pr_row("Utilities", 5030, "exp")
pr_row("Repairs & maintenance", 5040, "exp")
pr_row("Cleaning & landscaping", 5050, "exp")
pr_row("Management fees", 5060, "exp")
pr_row("Bad debt", 5100, "exp")
px_last = r - 1
pr_tot_opex = pr_total("        Total operating expenses", px_first, px_last)
r += 1

put(ws, r, 2, "NET OPERATING INCOME", font=F_BOLD, fill=FILL_TOTAL)
for col in range(3, 7):
    L = get_column_letter(col)
    put(ws, r, col, "=%s%d-%s%d" % (L, pr_tot_rev, L, pr_tot_opex), fmt=MONEY, font=F_BOLD,
        align="right", fill=FILL_TOTAL)
pr_noi = r
put(ws, r, 7, None, fmt=PCT, font=F_BOLD, fill=FILL_TOTAL, align="right")
r += 1
put(ws, r, 2, "        NOI margin", font=F_ITAL)
for col in (3, 4, 5, 6):
    L = get_column_letter(col)
    put(ws, r, col, "=IF(%s%d=0,\"\",%s%d/%s%d)" % (L, pr_tot_rev, L, pr_noi, L, pr_tot_rev),
        fmt=PCT, align="right", font=F_ITAL)
r += 1

put(ws, r, 2, "Other income and expenses", font=F_BOLD)
r += 1
oi_inc = pr_row("Interest income", 4900, "rev")
oi_bank = pr_row("Bank charges", 5090, "exp")
oi_dep = pr_row("Depreciation", 5070, "exp")
oi_int = pr_row("Interest expense", 5080, "exp")
put(ws, r, 2, "        NET INCOME FOR THE YEAR", font=F_BOLD, fill=FILL_TOTAL)
for col in range(3, 7):
    L = get_column_letter(col)
    put(ws, r, col, "=%s%d+%s%d-%s%d-%s%d-%s%d" % (L, pr_noi, L, oi_inc, L, oi_bank,
                                                   L, oi_dep, L, oi_int),
        fmt=MONEY, font=F_BOLD, align="right", fill=FILL_TOTAL)
pr_ni = r
put(ws, r, 7, None, fmt=PCT, font=F_BOLD, fill=FILL_TOTAL, align="right")
r += 1

put(ws, r, 2, "Check - agrees to the consolidated Profit & Loss statement", font=F_BOLD)
put(ws, r, 6, "=ROUND(F%d-'%s'!%s,2)" % (pr_ni, S_PL, PL_NI_CELL), fmt=MONEY, font=F_BOLD,
    align="right")
put(ws, r, 7, '=IF(ROUND(F%d,2)=0,"OK","CHECK")' % r, font=F_BOLD, align="center",
    fill=FILL_OK, border=BOX)
r += 1

# percent-of-revenue column, now that the revenue base row is known
for row in pct_rows:
    ws.cell(row=row, column=7).value = "=IF($F$%d=0,\"\",F%d/$F$%d)" % (pr_tot_rev, row, pr_tot_rev)
ws.cell(row=pr_tot_rev, column=7).value = "=IF(F%d=0,\"\",1)" % pr_tot_rev
ws.cell(row=pr_noi, column=7).value = "=IF(F%d=0,\"\",F%d/$F$%d)" % (pr_tot_rev, pr_noi, pr_tot_rev)
ws.cell(row=pr_ni, column=7).value = "=IF(F%d=0,\"\",F%d/$F$%d)" % (pr_tot_rev, pr_ni, pr_tot_rev)
r += 1

section(ws, r, 2, 7, "3.  DEBT AND CAPITAL POSITION AT 31 DECEMBER 2025")
r += 1
headers(ws, r, ["", "Facility", "Lender", "Balance", "Rate", "Monthly payment",
                "Amortisation / maturity"])
r += 1
d_first = r
put(ws, r, 2, "Rivergate Office Center")
put(ws, r, 3, "First Harbor Bank")
put(ws, r, 4, "='%s'!K%d" % (S_TB, tb_row[2100]), fmt=MONEY, align="right")
put(ws, r, 5, LOAN_A["rate"], fmt="0.00%", align="right")
put(ws, r, 6, SCHED_A[0]["payment"], fmt=MONEY, align="right")
put(ws, r, 7, "240 months, maturing 2029", font=F_SMALL)
r += 1
put(ws, r, 2, "Maple Street Retail Shops")
put(ws, r, 3, "First Harbor Bank")
put(ws, r, 4, "='%s'!K%d" % (S_TB, tb_row[2110]), fmt=MONEY, align="right")
put(ws, r, 5, LOAN_B["rate"], fmt="0.00%", align="right")
put(ws, r, 6, SCHED_B[0]["payment"], fmt=MONEY, align="right")
put(ws, r, 7, "300 months, maturing 2032", font=F_SMALL)
r += 1
put(ws, r, 2, "Total debt", font=F_BOLD, fill=FILL_TOTAL)
put(ws, r, 4, "=SUM(D%d:D%d)" % (d_first, r - 1), fmt=MONEY, font=F_BOLD, align="right",
    fill=FILL_TOTAL)
put(ws, r, 6, "=SUM(F%d:F%d)" % (d_first, r - 1), fmt=MONEY, font=F_BOLD, align="right",
    fill=FILL_TOTAL)
put(ws, r, 7, '=IF(ROUND(D%d-C%d,2)=0,"OK","CHECK")' % (r, k_td), font=F_BOLD,
    align="center", fill=FILL_OK, border=BOX)
r += 2

section(ws, r, 2, 7, "4.  COMMENTARY")
r += 1
notes = [
    "1. Two-property platform: Maple Street was acquired on 1 March 2025 for 2,140,000 (land 620,000, building "
    "1,480,000, capitalised costs 40,000), funded by a 1,400,000 mortgage and a 700,000 member contribution.",
    "2. Revenue quality: " + format(REV_PROP, ",.0f") + " of property revenue, of which rent of "
    + format(-BAL[4010] - BAL[4020], ",.0f") + " and recoveries of " + format(-BAL[4030], ",.0f")
    + " are contracted; only " + format(-BAL[4060], ",.0f") + " came from late fees. Vacancy is limited to "
    "Suite 120 (Rivergate) and Unit 4 (Maple Street).",
    "3. Accrual discipline: December rent of 42,100 is a receivable at year end, the Q4 operating settlement of "
    "29,160 and the Maple Street property tax of 12,000 are accrued, and 40,500 of insurance cover is carried "
    "forward. These are period allocations, not cash timing.",
    "4. Distributions: " + format(-BAL[3020], ",.0f") + " was distributed in two instalments. It is a return of "
    "equity, so it reduces equity and cash but never net income or NOI.",
    "5. Liquidity and covenant: cash of " + format(CASH, ",.0f") + " plus operating cash flow covers the next "
    "quarter's debt service; a DSCR of " + format(NOI_PY / DS_PY, ".2f") + "x sits above the 1.20x covenant.",
    "6. Next actions: re-let Suite 120, complete the 2026 CAM reconciliation, and open refinancing discussions "
    "on the Rivergate loan ahead of its 2029 maturity.",
]
for n in notes:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = put(ws, r, 2, n, font=F_BODY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 32
    r += 1
page(ws)

# ----------------------------------------------------------------------------
# 9. save + report
# ----------------------------------------------------------------------------
OUT = "Commercial_Property_Accounting_Workbook_FY2025.xlsx"
wb.save(OUT)

print("=" * 78)
print("WORKBOOK BUILT:", OUT)
print("=" * 78)
print("Transactions .............. %d" % len(TX))
print("Journal lines ............. %d" % len(je_rows))
print("Adjusting entries ......... %d (%d lines)" % (len(AJ), n_aj_lines))
print("Accounts .................. %d" % len(ACCOUNTS))
print("-" * 78)
print("Property revenue .......... %14s" % format(REV_PROP, ",.2f"))
print("Interest income ........... %14s" % format(-BAL[4900], ",.2f"))
print("Total revenue ............. %14s" % format(REV, ",.2f"))
print("Total expenses ............ %14s" % format(EXP, ",.2f"))
print("Net income ................ %14s" % format(NI, ",.2f"))
print("NOI ....................... %14s" % format(NOI_PY, ",.2f"))
print("-" * 78)
print("Total assets .............. %14s" % format(ASSETS, ",.2f"))
print("Total liabilities ......... %14s" % format(LIAB, ",.2f"))
print("Total equity .............. %14s" % format(EQUITY, ",.2f"))
print("  A - L - E - NI .......... %14s" % format(ASSETS - LIAB - EQUITY - NI, ",.2f"))
print("-" * 78)
print("Cash 1 Jan 2025 ........... %14s" % format(OPEN[1000], ",.2f"))
print("Cash 31 Dec 2025 .......... %14s" % format(CASH, ",.2f"))
print("Receipts / payments ....... %14s / %s" % (format(RECEIPTS, ",.2f"), format(PAYMENTS, ",.2f")))
print("Bank statement closing .... %14s" % format(BANK_CLOSING, ",.2f"))
print("  + DIT %s - OS %s = %s" % (format(DIT, ",.2f"), format(OS_TOTAL, ",.2f"),
                                   format(BANK_CLOSING + DIT - OS_TOTAL, ",.2f")))
print("-" * 78)
print("Trial balance debits ...... %14s" % format(sum(MOV[a[0]]["dr"] for a in ACCOUNTS)
                                                  + sum(MOV_AJ[a[0]]["dr"] for a in ACCOUNTS)
                                                  + sum(v for v in OPEN.values() if v > 0), ",.2f"))
print("Trial balance credits ..... %14s" % format(sum(MOV[a[0]]["cr"] for a in ACCOUNTS)
                                                  + sum(MOV_AJ[a[0]]["cr"] for a in ACCOUNTS)
                                                  + sum(-v for v in OPEN.values() if v < 0), ",.2f"))
print("=" * 78)

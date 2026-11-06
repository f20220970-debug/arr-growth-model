"""Builds ARR_Growth_Model.xlsx: Inputs (with a Bear/Base/Bull switch), a 24-month driver model in live formulas,
unit economics, and a Risk sheet with the Python Monte Carlo and sensitivity results (static, labelled as such)."""
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
import model as M

F = "Arial"
BLUE, BLACK, GREEN = Font(name=F, color="0000FF"), Font(name=F, color="000000"), Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, size=14)
HEAD = Font(name=F, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
YELLOW = PatternFill("solid", fgColor="FFFF00")
THIN = Border(bottom=Side(style="thin", color="999999"))
USD = '$#,##0;($#,##0);-'
NUM = '#,##0;(#,##0);-'
PCT = '0.0%;(0.0%);-'

wb = Workbook()

# ---------------- Inputs ----------------
ws = wb.active
ws.title = "Inputs"
ws["A1"] = "ARR Growth Model: commission-revenue insurance brokerage (illustrative inputs)"; ws["A1"].font = TITLE
ws["A2"] = "Blue = input you can change. Yellow = scenario switch. Black = formula. Green = link from another sheet."
ws["A2"].font = Font(name=F, italic=True, color="666666")
ws["A4"] = "Scenario in use (type Bear, Base or Bull)"; ws["A4"].font = BOLD
ws["B4"] = "Base"; ws["B4"].font = BLUE; ws["B4"].fill = YELLOW
ws["B4"].comment = Comment("Switch the whole model between the Bear, Base and Bull columns below.", "model")
for c, h in zip("ABCDEF", ["Driver", "Live value", "Bear", "Base", "Bull", "Note"]):
    ws[f"{c}6"] = h; ws[f"{c}6"].font = HEAD; ws[f"{c}6"].fill = HEAD_FILL
NOTES = {
    "lead_growth": "Monthly growth in inbound leads",
    "annual_renewal": "Share of policies renewing each year; monthly churn = 1 - renewal^(1/12)",
    "leads_per_agent": "Agent capacity; agents = ROUNDUP(leads / capacity)",
    "start_policies": "Book of active policies at month 0",
    "target_arr": "Milestone used for 'month target ARR is hit'",
}
fmt = {"lead_growth": PCT, "lead_to_bind": PCT, "commission_rate": PCT, "annual_renewal": PCT}
INPUT_ROW = {}
for i, (k, vals) in enumerate(M.SCENARIOS.items()):
    r = 7 + i
    INPUT_ROW[k] = r
    ws[f"A{r}"] = M.LABELS[k]
    ws[f"B{r}"] = f'=INDEX(C{r}:E{r},MATCH($B$4,$C$6:$E$6,0))'
    ws[f"B{r}"].font = BLACK
    for c, v in zip("CDE", vals):
        ws[f"{c}{r}"] = v; ws[f"{c}{r}"].font = BLUE
    ws[f"F{r}"] = NOTES.get(k, "Illustrative assumption, not company data")
    ws[f"F{r}"].font = Font(name=F, color="666666")
    f = fmt.get(k, USD if any(w in k for w in ["premium", "cost", "opex", "cash", "arr"]) else NUM)
    for c in "BCDE":
        ws[f"{c}{r}"].number_format = f
    for c in "ABCDEF":
        ws[f"{c}{r}"].border = THIN
ws.column_dimensions["A"].width = 42
for c in "BCDE":
    ws.column_dimensions[c].width = 14
ws.column_dimensions["F"].width = 62


def inp(k):
    return f"Inputs!$B${INPUT_ROW[k]}"


# ---------------- Model ----------------
m = wb.create_sheet("Model")
m["A1"] = "24-month driver model (all formulas; change Inputs to update)"; m["A1"].font = TITLE
m["A3"] = "Month"; m["A3"].font = HEAD; m["A3"].fill = HEAD_FILL
LINES = [
    ("Leads", NUM), ("New policies bound", NUM), ("Policies churned", NUM), ("Active policies (end of month)", NUM),
    ("ARR ($, annualised commission)", USD), ("Revenue in month ($)", USD), ("Marketing spend ($)", USD),
    ("Agents needed", NUM), ("Agent cost ($)", USD), ("Fixed opex ($)", USD), ("Total cost ($)", USD),
    ("Net burn ($)", USD), ("Cash at end of month ($)", USD),
]
R = {name: 4 + i for i, (name, _) in enumerate(LINES)}
for name, f in LINES:
    m[f"A{R[name]}"] = name
    m[f"A{R[name]}"].font = BOLD if name.startswith(("ARR", "Cash")) else Font(name=F)
churn = f"(1-{inp('annual_renewal')}^(1/12))"
for j in range(1, M.MONTHS + 1):
    c = get_column_letter(j + 1); p = get_column_letter(j)
    m[f"{c}3"] = j; m[f"{c}3"].font = HEAD; m[f"{c}3"].fill = HEAD_FILL
    first = j == 1
    cell = {
        "Leads": f"={inp('start_leads')}" if first else f"={p}{R['Leads']}*(1+{inp('lead_growth')})",
        "New policies bound": f"={c}{R['Leads']}*{inp('lead_to_bind')}",
        "Policies churned": (f"={inp('start_policies')}*{churn}" if first else f"={p}{R['Active policies (end of month)']}*{churn}"),
        "Active policies (end of month)": (f"={inp('start_policies')}+{c}{R['New policies bound']}-{c}{R['Policies churned']}" if first
                                           else f"={p}{R['Active policies (end of month)']}+{c}{R['New policies bound']}-{c}{R['Policies churned']}"),
        "ARR ($, annualised commission)": f"={c}{R['Active policies (end of month)']}*{inp('avg_premium')}*{inp('commission_rate')}",
        "Revenue in month ($)": f"={c}{R['ARR ($, annualised commission)']}/12",
        "Marketing spend ($)": f"={c}{R['Leads']}*{inp('cost_per_lead')}",
        "Agents needed": f"=ROUNDUP({c}{R['Leads']}/{inp('leads_per_agent')}-0.000000001,0)",
        "Agent cost ($)": f"={c}{R['Agents needed']}*{inp('agent_cost')}",
        "Fixed opex ($)": f"={inp('fixed_opex')}",
        "Total cost ($)": f"={c}{R['Marketing spend ($)']}+{c}{R['Agent cost ($)']}+{c}{R['Fixed opex ($)']}",
        "Net burn ($)": f"={c}{R['Revenue in month ($)']}-{c}{R['Total cost ($)']}",
        "Cash at end of month ($)": (f"={inp('start_cash')}+{c}{R['Net burn ($)']}" if first
                                     else f"={p}{R['Cash at end of month ($)']}+{c}{R['Net burn ($)']}"),
    }
    for name, f in LINES:
        x = m[f"{c}{R[name]}"]
        x.value = cell[name]; x.number_format = f
        x.font = GREEN if cell[name].count("Inputs!") and cell[name].startswith("=Inputs") else BLACK
    m.column_dimensions[c].width = 13
m.column_dimensions["A"].width = 34
m.freeze_panes = "B4"
last = get_column_letter(M.MONTHS + 1)
arr_row, cash_row = R["ARR ($, annualised commission)"], R["Cash at end of month ($)"]

# ---------------- Summary ----------------
s = wb.create_sheet("Summary", 1)
s["A1"] = "Summary (live formulas)"; s["A1"].font = TITLE
rows = [
    ("Scenario", "=Inputs!B4", None),
    ("ARR at month 12 ($)", f"=Model!M{arr_row}", USD),
    ("ARR at month 24 ($)", f"=Model!{last}{arr_row}", USD),
    ("Month target ARR is first hit (25 = not within 24 months)", f'=IF(MAX(Model!B{arr_row}:{last}{arr_row})>={inp("target_arr")},COUNTIF(Model!B{arr_row}:{last}{arr_row},"<"&{inp("target_arr")})+1,25)', NUM),
    ("Lowest cash balance ($)", f"=MIN(Model!B{cash_row}:{last}{cash_row})", USD),
    ("Extra cash needed to never go below zero ($)", f"=MAX(0,-B6)", USD),
    ("CAC per bound policy ($)", f"={inp('cost_per_lead')}/{inp('lead_to_bind')}", USD),
    ("Commission per policy per year ($)", f"={inp('avg_premium')}*{inp('commission_rate')}", USD),
    ("Commission LTV ($) = annual commission / (1 - renewal)", f"=B9/(1-{inp('annual_renewal')})", USD),
    ("LTV / CAC", "=B10/B8", '0.0"x"'),
    ("CAC payback (months)", "=12*B8/B9", '0.0'),
]
for i, (label, f, nf) in enumerate(rows):
    r = 2 + i
    s[f"A{r}"] = label; s[f"B{r}"] = f
    s[f"B{r}"].font = GREEN if f.startswith(("=Inputs", "=Model")) else BLACK
    if nf:
        s[f"B{r}"].number_format = nf
s.column_dimensions["A"].width = 58; s.column_dimensions["B"].width = 18

# ---------------- Risk (static results from model.py) ----------------
res = json.load(open("results.json"))
k = wb.create_sheet("Risk")
k["A1"] = "Risk view: static results from model.py (Python mirror of this workbook); rerun model.py to refresh"
k["A1"].font = TITLE
k["A3"] = "Monte Carlo, 5,000 runs around Base"; k["A3"].font = BOLD
mc = res["monte_carlo"]
mrows = [("ARR at month 24, P10 ($)", mc["arr_m24_p10_p50_p90"][0], USD), ("ARR at month 24, P50 ($)", mc["arr_m24_p10_p50_p90"][1], USD),
         ("ARR at month 24, P90 ($)", mc["arr_m24_p10_p50_p90"][2], USD), ("Chance target hit by month 18", mc["p_target_by_m18"] / 100, PCT),
         ("Chance target hit by month 24", mc["p_target_by_m24"] / 100, PCT), ("Chance cash goes below zero", mc["p_cash_below_zero"] / 100, PCT)]
for i, (a, v, nf) in enumerate(mrows):
    k[f"A{4+i}"] = a; k[f"B{4+i}"] = v; k[f"B{4+i}"].number_format = nf
k["A11"] = "Drawn: lead growth N(7%, 2%), conversion N(15%, 2%), renewal N(81%, 4%), cost per lead N($40, $7), premium N($2,300, $200)"
k["A11"].font = Font(name=F, italic=True, color="666666")
k["A13"] = "Sensitivity: ARR at month 18 when one driver moves -20% / +20%"; k["A13"].font = BOLD
for c, h in zip("ABC", ["Driver", "-20% ($)", "+20% ($)"]):
    k[f"{c}14"] = h; k[f"{c}14"].font = HEAD; k[f"{c}14"].fill = HEAD_FILL
for i, d in enumerate(res["tornado_arr_m18"]["drivers"]):
    k[f"A{15+i}"] = d["driver"]; k[f"B{15+i}"] = d["low"]; k[f"C{15+i}"] = d["high"]
    k[f"B{15+i}"].number_format = USD; k[f"C{15+i}"].number_format = USD
k.column_dimensions["A"].width = 52; k.column_dimensions["B"].width = 16; k.column_dimensions["C"].width = 16

for sh in wb.worksheets:
    for row in sh.iter_rows():
        for x in row:
            if x.font is None or x.font.name != F:
                x.font = Font(name=F, bold=x.font.bold if x.font else False, color=x.font.color if x.font else None,
                              italic=x.font.italic if x.font else False, size=x.font.size if x.font else 11)
wb.save("ARR_Growth_Model.xlsx")
print("saved")

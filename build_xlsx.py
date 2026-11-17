"""Build MTN_Pricer.xlsx: curve bootstrap, MTN pricer, cash flows, blotter - all in native Excel formulas."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

BLUE = "1F4E8C"
H = Font(bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor=BLUE)
INP = PatternFill("solid", fgColor="FFF2CC")      # inputs in yellow
OUT = PatternFill("solid", fgColor="E2EFDA")      # key outputs in green
B = Font(bold=True)
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

# illustrative inputs - NOT market data, replace with the desk curve
TENORS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 20, 25, 30]
PAR = [2.55, 2.58, 2.62, 2.66, 2.71, 2.76, 2.81, 2.86, 2.91, 2.96, 3.04, 3.12, 3.16, 3.12, 3.06]
SPREAD = [18, 25, 32, 38, 44, 50, 55, 60, 64, 68, 74, 80, 86, 90, 94]

wb = Workbook()

# ---------------------------------------------------------------- Market ---
m = wb.active
m.title = "Market"
m["A1"] = "Données de marché – courbe swap EUR et spread de financement émetteur"
m["A1"].font = Font(bold=True, size=13, color=BLUE)
m["A2"] = ("Valeurs ILLUSTRATIVES, à remplacer par la courbe du desk (ex. Bloomberg EUSA) "
           "et la grille de spreads de l'émetteur. Cellules jaunes = saisies.")
m["A2"].font = Font(italic=True, color="C00000")
m["A3"], m["B3"] = "Date de valorisation", dt.date(2026, 11, 18)
m["B3"].number_format = "dd/mm/yyyy"
m["A4"], m["B4"] = "Euribor 6M (fixing, %)", 2.60
for c in ("B3", "B4"):
    m[c].fill = INP
for j, h in enumerate(["Maturité (ans)", "Taux swap par (%)", "Spread de financement (bp)"], 1):
    c = m.cell(6, j, h); c.font, c.fill = H, HF
for i, (t, r, s) in enumerate(zip(TENORS, PAR, SPREAD)):
    m.cell(7 + i, 1, t)
    for j, v in ((2, r), (3, s)):
        c = m.cell(7 + i, j, v); c.fill = INP
for col, w in zip("ABC", (26, 20, 28)):
    m.column_dimensions[col].width = w
LAST = 6 + len(TENORS)          # last row of the tenor table
TEN = f"Market!$A$7:$A${LAST}"
PARR = f"Market!$B$7:$B${LAST}"
SPR = f"Market!$C$7:$C${LAST}"

# ----------------------------------------------------------------- Curve ---
c = wb.create_sheet("Curve")
c["A1"] = "Bootstrap annuel : facteurs d'actualisation sans risque et émetteur"
c["A1"].font = Font(bold=True, size=13, color=BLUE)
hdr = ["t (ans)", "Taux par interpolé", "Spread interpolé (bp)", "DF swap",
       "Taux zéro (cont.)", "Forward 1 an", "DF émetteur", "Annuité émetteur cumulée",
       "Σ forward × DF émetteur"]
for j, h in enumerate(hdr, 1):
    x = c.cell(3, j, h); x.font, x.fill = H, HF
    c.column_dimensions[chr(64 + j)].width = 22


def interp(col_rng, t):
    # linear interpolation on the tenor table, flat beyond the ends
    lo = f"MATCH({t},{TEN},1)"
    return (f"=IF({t}<=MIN({TEN}),INDEX({col_rng},1),IF({t}>=MAX({TEN}),INDEX({col_rng},ROWS({TEN})),"
            f"INDEX({col_rng},{lo})+({t}-INDEX({TEN},{lo}))/(INDEX({TEN},{lo}+1)-INDEX({TEN},{lo}))"
            f"*(INDEX({col_rng},{lo}+1)-INDEX({col_rng},{lo}))))")


for k in range(1, 31):
    r = 3 + k
    c.cell(r, 1, k)
    c.cell(r, 2, interp(PARR, f"A{r}"))
    c.cell(r, 3, interp(SPR, f"A{r}"))
    # par-swap bootstrap, annual fixed leg: DF_n = (1 - S_n * sum DF_1..n-1) / (1 + S_n)
    prev = "0" if k == 1 else f"SUM($D$4:D{r-1})"
    c.cell(r, 4, f"=(1-B{r}/100*{prev})/(1+B{r}/100)")
    c.cell(r, 5, f"=-LN(D{r})/A{r}")
    c.cell(r, 6, "=1/D4-1" if k == 1 else f"=D{r-1}/D{r}-1")
    c.cell(r, 7, f"=D{r}*EXP(-C{r}/10000*A{r})")
    c.cell(r, 8, f"=SUM($G$4:G{r})")
    c.cell(r, 9, f"=SUMPRODUCT($F$4:F{r},$G$4:G{r})")
    for j in range(2, 10):
        c.cell(r, j).number_format = "0.000000" if j in (4, 7, 8) else "0.0000"

# ---------------------------------------------------------------- Pricer ---
p = wb.create_sheet("Pricer")
p["A1"] = "Pricer MTN – émission à taux fixe et FRN"
p["A1"].font = Font(bold=True, size=13, color=BLUE)
p["A2"] = "Coupons annuels, maturités entières de 1 à 30 ans. Cellules jaunes = saisies."
p["A2"].font = Font(italic=True)
rows = [
    ("Émetteur", "Émetteur (exemple)", None),
    ("Devise", "EUR", None),
    ("Nominal", 50_000_000, "#,##0"),
    ("Maturité (ans, entier)", 7, "0"),
    ("Coupon fixe proposé (%)", 3.35, "0.000"),
    ("", None, None),
    ("Taux swap à maturité (%)", "=INDEX(Curve!$B$4:$B$33,B7)", "0.000"),
    ("Spread de financement (bp)", "=INDEX(Curve!$C$4:$C$33,B7)", "0.0"),
    ("Coupon fixe au pair (%)", "=(1-INDEX(Curve!$G$4:$G$33,B7))/INDEX(Curve!$H$4:$H$33,B7)*100", "0.0000"),
    ("Prix du MTN au coupon proposé (%)", "=(B8/100*INDEX(Curve!$H$4:$H$33,B7)+INDEX(Curve!$G$4:$G$33,B7))*100", "0.0000"),
    ("Écart coupon proposé vs pair (bp)", "=(B8-B12)*100", "0.0"),
    ("Coût tout compris swappé (Euribor + bp)", "=(B8-B10)*100", "0.0"),
    ("Marge FRN au pair (Euribor + bp)",
     "=(1-INDEX(Curve!$G$4:$G$33,B7)-INDEX(Curve!$I$4:$I$33,B7))/INDEX(Curve!$H$4:$H$33,B7)*10000", "0.0"),
    ("Duration (ans)", "=SUMPRODUCT(Cashflows!$A$4:$A$33,Cashflows!$E$4:$E$33)/SUMPRODUCT(Cashflows!$E$4:$E$33)", "0.000"),
    ("DV01 (EUR par bp)", "=B6*B13/100*B17/10000", "#,##0"),
    ("Valeur actuelle (EUR)", "=B6*B13/100", "#,##0"),
]
for i, (lab, val, fmt) in enumerate(rows):
    r = 4 + i
    p.cell(r, 1, lab).font = B if lab else Font()
    if val is not None:
        x = p.cell(r, 2, val)
        if fmt:
            x.number_format = fmt
        x.fill = INP if r <= 8 else OUT
        x.border = BOX
p.column_dimensions["A"].width = 40
p.column_dimensions["B"].width = 22
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="30", showErrorMessage=True,
                    error="Maturité entière entre 1 et 30 ans")
p.add_data_validation(dv)
dv.add("B7")
p["D4"] = "Lecture"
p["D4"].font = B
notes = [
    "Coupon au pair : coupon qui fait coter le MTN à 100 sur la courbe de l'émetteur (swap + spread).",
    "Coût tout compris : coupon fixe swappé contre Euribor, soit Euribor + (coupon − taux swap).",
    "Marge FRN : marge sur Euribor qui fait coter un FRN de même maturité à 100.",
    "DV01 : variation de valeur pour 1 bp de hausse du taux zéro émetteur (duration) ; le VBA la recalcule par bump de la courbe par.",
    "Avec les macros (MTN_Pricer.xlsm) : contrôles VBA indépendants, blotter et export de booking.",
]
for i, n in enumerate(notes):
    p.cell(5 + i, 4, n)
p.column_dimensions["D"].width = 95

# ------------------------------------------------------------- Cashflows ---
f = wb.create_sheet("Cashflows")
f["A1"] = "Échéancier du MTN à taux fixe (coupon proposé)"
f["A1"].font = Font(bold=True, size=13, color=BLUE)
for j, h in enumerate(["t (ans)", "Date de paiement", "Flux (% nominal)", "DF émetteur", "VA (% nominal)", "VA (EUR)"], 1):
    x = f.cell(3, j, h); x.font, x.fill = H, HF
    f.column_dimensions[chr(64 + j)].width = 20
for k in range(1, 31):
    r = 3 + k
    f.cell(r, 1, k)
    f.cell(r, 2, f"=IF(A{r}<=Pricer!$B$7,EDATE(Market!$B$3,12*A{r}),\"\")").number_format = "dd/mm/yyyy"
    f.cell(r, 3, f"=IF(A{r}<Pricer!$B$7,Pricer!$B$8,IF(A{r}=Pricer!$B$7,100+Pricer!$B$8,0))").number_format = "0.000"
    f.cell(r, 4, f"=Curve!G{r}").number_format = "0.000000"
    f.cell(r, 5, f"=C{r}*D{r}").number_format = "0.0000"
    f.cell(r, 6, f"=E{r}/100*Pricer!$B$6").number_format = "#,##0"
f["A35"], f["E35"], f["F35"] = "Total", "=SUM(E4:E33)", "=SUM(F4:F33)"
f["A35"].font = B
f["E35"].number_format, f["F35"].number_format = "0.0000", "#,##0"

# --------------------------------------------------------------- Blotter ---
b = wb.create_sheet("Blotter")
b["A1"] = "Blotter des émissions MTN – contrôles avant booking"
b["A1"].font = Font(bold=True, size=13, color=BLUE)
b["A2"] = "Macros : AjouterAuBlotter (depuis le pricer), ValiderBlotter, ExporterBooking (CSV)."
b["A2"].font = Font(italic=True)
bh = ["ID trade", "Date", "Émetteur", "Devise", "Nominal", "Maturité", "Coupon (%)", "Coupon au pair (%)",
      "Écart (bp)", "Tout compris (E+bp)", "Statut", "Commentaire contrôle"]
for j, h in enumerate(bh, 1):
    x = b.cell(4, j, h); x.font, x.fill = H, HF
    b.column_dimensions[chr(64 + j)].width = 16 if j < 12 else 45
# red fill when a row fails validation
b.conditional_formatting.add("A5:L500", FormulaRule(formula=['$K5="Erreur"'], fill=PatternFill("solid", fgColor="F8CBAD")))
b.conditional_formatting.add("A5:L500", FormulaRule(formula=['$K5="Validé"'], fill=PatternFill("solid", fgColor="E2EFDA")))

# --------------------------------------------------------------- Controls (VBA) ---
k = wb.create_sheet("Controles_VBA")
k["A1"] = "Contrôles : formules Excel vs fonctions VBA indépendantes"
k["A1"].font = Font(bold=True, size=13, color=BLUE)
k["A2"] = "Ces cellules utilisent les fonctions VBA du module modPricing (classeur .xlsm, macros activées)."
k["A2"].font = Font(italic=True)
for j, h in enumerate(["Mesure", "Formules Excel", "VBA", "Écart", "OK ?"], 1):
    x = k.cell(4, j, h); x.font, x.fill = H, HF
checks = [
    ("Coupon au pair (%)", "=Pricer!B12", "=MTN_COUPON_PAIR(Pricer!B7)", 1e-6),
    ("Prix au coupon proposé (%)", "=Pricer!B13", "=MTN_PRIX(Pricer!B8,Pricer!B7)", 1e-6),
    ("Marge FRN (bp)", "=Pricer!B16", "=MTN_MARGE_FRN(Pricer!B7)", 1e-4),
    ("Tout compris (bp)", "=Pricer!B15", "=MTN_TOUT_COMPRIS(Pricer!B8,Pricer!B7)", 1e-6),
    ("DV01 (EUR/bp) – duration vs bump de courbe", "=Pricer!B18", "=MTN_DV01(Pricer!B6,Pricer!B8,Pricer!B7)", None),
]
for i, (lab, fx, vba, tol) in enumerate(checks):
    r = 5 + i
    k.cell(r, 1, lab); k.cell(r, 2, fx); k.cell(r, 3, vba)
    k.cell(r, 4, f"=C{r}-B{r}")
    k.cell(r, 5, f'=IF(ABS(D{r})<={tol},"OK","ÉCART")' if tol else f'=IF(ABS(D{r}/B{r})<0.05,"OK (<5 %)","ÉCART")')
    for j in (2, 3, 4):
        k.cell(r, j).number_format = "0.000000"
k.column_dimensions["A"].width = 36
for col in "BCDE":
    k.column_dimensions[col].width = 18

wb.move_sheet("Pricer", offset=-2)
wb.active = 0
wb.save("MTN_Pricer.xlsx")
print("saved")

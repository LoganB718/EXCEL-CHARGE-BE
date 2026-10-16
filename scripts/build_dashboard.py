# -*- coding: utf-8 -*-
"""Génère le classeur « Plan de charge BE » : onglets Charge, Data, Dashboard et Calculs.

Usage : python scripts/build_dashboard.py [chemin_sortie.xlsx]
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.data_source import NumFmt, StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.marker import Marker
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break

OUT = sys.argv[1] if len(sys.argv) > 1 else "Plan_de_charge_BE_Dashboard.xlsx"

FONT = "Arial"
NAVY = "1F2A44"
INK = "1F2937"
GREY = "6B7280"
MUTED = "9CA3AF"
LINE = "D6DCE4"
BG = "EEF1F5"
WHITE = "FFFFFF"
ACCENT = "2A78D6"
AMBER = "EDA100"
INPUT_FILL = "EAF2FD"
INPUT_INK = "1F4E9A"
# Palette catégorielle validée (ordre fixe) + gris pour « Autres projets »
SLOT_COLORS = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948", "9CA3AF"]
N_SLOTS = len(SLOT_COLORS)

FIRST, LAST = 4, 503            # lignes de données de l'onglet Charge (500 affaires max)
N_MONTHS = 12
N_TABLE = 100                   # lignes du tableau des affaires sur le Dashboard

PERIM_ALL = "Tous les projets"
PERIM_RUN = "Projets en cours"
REP_SMOOTH = "Lissée jusqu'à la date de besoin"
REP_START = "Tout sur le mois de début"


def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def side(color=LINE, style="thin"):
    return Side(style=style, color=color)


def rng(col, r1=FIRST, r2=LAST, sheet=None):
    ref = f"${col}${r1}:${col}${r2}"
    return f"{sheet}!{ref}" if sheet else ref


wb = Workbook()

# ---------------------------------------------------------------- Data
ws_data = wb.active
ws_data.title = "Data"
ws_charge = wb.create_sheet("Charge", 0)
ws_dash = wb.create_sheet("Dashboard")
ws_calc = wb.create_sheet("Calculs")

LISTS = [
    ("Chefs de projet", "lst_ChefsProjet",
     ["Julie Martin", "Thomas Bernard", "Sophie Lefèvre", "Nicolas Moreau", "Camille Laurent"]),
    ("Noms de projet", "lst_Projets",
     ["ATLAS", "BORÉAL", "CASSIOPÉE", "DELTA", "ORION", "PÉGASE", "VEGA", "TITAN"]),
    ("Types", "lst_Types",
     ["Étude de conception", "Modification", "Mise à jour de plans", "Industrialisation",
      "Chiffrage / avant-projet"]),
    ("Coordinateurs", "lst_Coordinateurs", ["Marc Dubois", "Isabelle Roux"]),
    ("Dessinateurs projeteurs", "lst_Dessinateurs",
     ["Lucas Garnier", "Emma Faure", "Hugo Mercier", "Léa Chevalier", "Antoine Blanc"]),
    ("Prestataires (ressources externes)", "lst_Prestataires",
     ["Alpha Ingénierie", "CAO Services Ouest", "Technidessin"]),
]
DATA_LAST = 203

ws_data["A1"] = "LISTES DE RÉFÉRENCE"
ws_data["A1"].font = font(14, True, NAVY)
ws_data["A2"] = ("Saisir une valeur par cellule, sans ligne vide, sous chaque en-tête : les listes déroulantes "
                 "de l'onglet Charge se mettent à jour automatiquement (200 valeurs max par liste). "
                 "Les valeurs ci-dessous sont des exemples à remplacer.")
ws_data["A2"].font = font(9, color=GREY, italic=True)
ws_data.row_dimensions[3].height = 30
for i, (title, name, values) in enumerate(LISTS):
    col = L(i + 1)
    ws_data.column_dimensions[col].width = 30
    h = ws_data[f"{col}3"]
    h.value = title
    h.font = font(10, True, WHITE)
    h.fill = fill(NAVY)
    h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(4, DATA_LAST + 1):
        c = ws_data[f"{col}{r}"]
        c.font = font(10)
        c.border = Border(bottom=side("E5E7EB"), left=side("E5E7EB"), right=side("E5E7EB"))
    for j, v in enumerate(values):
        ws_data[f"{col}{4 + j}"] = v
    wb.defined_names[name] = DefinedName(
        name, attr_text=f"OFFSET(Data!${col}$4,0,0,MAX(1,COUNTA(Data!${col}$4:${col}${DATA_LAST})),1)")
ws_data.freeze_panes = "A4"
ws_data.sheet_properties.tabColor = "6B7280"

# ---------------------------------------------------------------- Charge
CHARGE_COLS = [
    # (lettre, en-tête, largeur, liste, format)
    ("A", "Chef de projet", 20, "lst_ChefsProjet", None),
    ("B", "Nom du projet", 18, "lst_Projets", None),
    ("C", "Sous-projet", 22, None, None),
    ("D", "Produit", 22, None, None),
    ("E", "Type", 20, "lst_Types", None),
    ("F", "Date de besoin", 13, None, "DD/MM/YYYY"),
    ("G", "Description du projet", 38, None, None),
    ("H", "Contraintes, spécifications, délais, budget", 34, None, None),
    ("I", "Autres informations pertinentes / suggestions d'optimisation des ressources", 34, None, None),
    ("J", "Coordinateur", 18, "lst_Coordinateurs", None),
    ("K", "Dessinateur projeteur", 20, "lst_Dessinateurs", None),
    ("L", "Ressources externes (prestataire)", 22, "lst_Prestataires", None),
    ("M", "Charge de travail estimée (h)", 13, None, None),
    ("N", "Date de début planifiée", 13, None, "DD/MM/YYYY"),
    ("O", "Taux d'avancement", 13, None, "0%"),
]

ws = ws_charge
ws["A1"] = "PLAN DE CHARGE — BUREAU D'ÉTUDES"
ws["A1"].font = font(14, True, NAVY)
ws["A2"] = ("Une ligne = une affaire.  Colonnes A à I : informations projet  ·  Colonnes J à O : affectation et "
            "planification.  Listes déroulantes alimentées par l'onglet Data.  "
            "Les lignes 4 à 21 sont des EXEMPLES FICTIFS à remplacer par vos données.")
ws["A2"].font = font(9, color=GREY, italic=True)
ws.row_dimensions[1].height = 24
ws.row_dimensions[3].height = 48
thin = side("D9DEE5")
for col, title, width, _, fmt in CHARGE_COLS:
    ws.column_dimensions[col].width = width
    h = ws[f"{col}3"]
    h.value = title
    h.font = font(10, True, WHITE)
    h.fill = fill(NAVY if col <= "I" else "2F5C8F")
    h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    h.border = Border(left=side(WHITE), right=side(WHITE))
    wrap = col in "GHI"
    for r in range(FIRST, LAST + 1):
        c = ws[f"{col}{r}"]
        c.font = font(10)
        c.border = Border(bottom=thin, left=thin, right=thin)
        c.alignment = Alignment(vertical="top", wrap_text=wrap,
                                horizontal="center" if col in "FMNO" else None)
        if fmt:
            c.number_format = fmt

SAMPLE = [
    ("Julie Martin", "ATLAS", "Lot 1 – Structure", "Châssis principal", "Étude de conception", date(2027, 3, 31),
     "Conception du châssis mécano-soudé de la nouvelle gamme", "Masse < 450 kg ; budget étude 60 k€",
     "Mutualiser les calculs RDM avec le lot 2", "Marc Dubois", "Lucas Garnier", "Alpha Ingénierie", 420,
     date(2026, 11, 2), 0.15),
    ("Julie Martin", "ATLAS", "Lot 2 – Carénage", "Capots latéraux", "Étude de conception", date(2027, 5, 28),
     "Dessin des capots et de leurs fixations", "Matière ABS, outillage existant", None, "Marc Dubois",
     "Emma Faure", None, 260, date(2027, 1, 11), 0),
    ("Julie Martin", "ATLAS", "Lot 3 – Intégration", "Faisceau électrique", "Industrialisation", date(2027, 7, 30),
     "Routage des faisceaux et supports", "Normes NF C 15-100", "Prévoir un prestataire si pic en juin",
     "Marc Dubois", "Hugo Mercier", "CAO Services Ouest", 180, date(2027, 4, 5), 0),
    ("Thomas Bernard", "BORÉAL", "Phase APD", "Groupe froid", "Étude de conception", date(2027, 1, 29),
     "Avant-projet détaillé du groupe froid", "Délai client ferme", None, "Isabelle Roux", "Léa Chevalier",
     None, 310, date(2026, 10, 5), 0.35),
    ("Thomas Bernard", "BORÉAL", "Phase EXE", "Groupe froid", "Mise à jour de plans", date(2027, 4, 30),
     "Plans d'exécution", "Budget 25 k€", None, "Isabelle Roux", "Léa Chevalier", None, 220,
     date(2027, 2, 1), 0),
    ("Thomas Bernard", "BORÉAL", "Phase EXE", "Armoire de commande", "Modification", date(2027, 3, 15),
     "Reprise de l'armoire existante", None, None, "Isabelle Roux", "Antoine Blanc", "Technidessin", 90,
     date(2027, 1, 18), 0),
    ("Sophie Lefèvre", "CASSIOPÉE", "Rétrofit ligne 2", "Convoyeur", "Modification", date(2026, 12, 18),
     "Adaptation du convoyeur au nouveau format", "Arrêt de ligne limité à 2 semaines",
     "Réutiliser les plans de la ligne 1", "Marc Dubois", "Hugo Mercier", None, 150, date(2026, 10, 12), 0.5),
    ("Sophie Lefèvre", "CASSIOPÉE", "Rétrofit ligne 2", "Poste de contrôle", "Étude de conception",
     date(2027, 2, 26), "Nouveau poste de contrôle vision", None, None, "Marc Dubois", "Emma Faure",
     "Alpha Ingénierie", 240, date(2026, 12, 1), 0.05),
    ("Sophie Lefèvre", "DELTA", "Lot unique", "Bâti machine", "Industrialisation", date(2027, 6, 30),
     "Industrialisation du bâti", "Cible coût -10 %", None, "Isabelle Roux", "Antoine Blanc", None, 360,
     date(2027, 2, 15), 0),
    ("Nicolas Moreau", "ORION", "Tranche 1", "Support moteur", "Étude de conception", date(2026, 11, 27),
     "Conception du support moteur", None, None, "Isabelle Roux", "Lucas Garnier", None, 120,
     date(2026, 9, 14), 0.8),
    ("Nicolas Moreau", "ORION", "Tranche 1", "Carter de protection", "Mise à jour de plans", date(2026, 10, 9),
     "Mise à jour indice C", None, None, "Isabelle Roux", "Lucas Garnier", None, 40, date(2026, 9, 21), 1),
    ("Nicolas Moreau", "ORION", "Tranche 2", "Support moteur", "Modification", date(2027, 5, 14),
     "Évolution pour moteur 30 kW", None, None, "Isabelle Roux", "Emma Faure", None, 160, date(2027, 3, 1), 0),
    ("Nicolas Moreau", "PÉGASE", "Avant-projet", "Ensemble complet", "Chiffrage / avant-projet",
     date(2026, 12, 11), "Chiffrage et esquisses pour l'appel d'offres", "Remise d'offre le 15/12", None,
     "Marc Dubois", "Léa Chevalier", None, 70, date(2026, 11, 16), 0.1),
    ("Camille Laurent", "PÉGASE", "Lot 1", "Ensemble complet", "Étude de conception", date(2027, 9, 30),
     "Étude complète si commande", "Sous réserve de commande", "Prévoir un renfort externe", "Marc Dubois",
     "Hugo Mercier", "Technidessin", 480, date(2027, 3, 15), 0),
    ("Camille Laurent", "VEGA", "Mise en conformité", "Garde-corps", "Modification", date(2027, 1, 15),
     "Mise en conformité des garde-corps", "Norme NF E 85-015", None, "Isabelle Roux", "Antoine Blanc", None,
     85, date(2026, 11, 23), 0),
    ("Camille Laurent", "VEGA", "Mise en conformité", "Passerelle", "Mise à jour de plans", date(2027, 3, 31),
     "Plans de recollement", None, None, "Isabelle Roux", "Antoine Blanc", None, 60, date(2027, 2, 8), 0),
    ("Julie Martin", "TITAN", "Lot 1", "Poutre de levage", "Étude de conception", date(2026, 9, 30),
     "Note de calcul et plans", None, None, "Marc Dubois", "Lucas Garnier", None, 95, date(2026, 8, 24), 1),
    ("Thomas Bernard", "DELTA", "Lot unique", "Habillage", "Étude de conception", date(2027, 8, 27),
     "Habillage design", None, None, "Isabelle Roux", "Emma Faure", "CAO Services Ouest", 200,
     date(2027, 5, 3), 0),
]
for i, row in enumerate(SAMPLE):
    r = FIRST + i
    for (col, *_), v in zip(CHARGE_COLS, row):
        if v is not None:
            ws[f"{col}{r}"] = v

ws.freeze_panes = "C4"
ws.auto_filter.ref = f"A3:O{LAST}"
ws.sheet_view.zoomScale = 90

for col, _, _, lst, _ in CHARGE_COLS:
    if lst:
        dv = DataValidation(type="list", formula1=lst, allow_blank=True, errorStyle="warning",
                            showErrorMessage=True, errorTitle="Valeur hors liste",
                            error="Cette valeur n'est pas dans la liste de l'onglet Data. Continuer quand même ?")
        dv.add(f"{col}{FIRST}:{col}{LAST}")
        ws.add_data_validation(dv)
for col in "FN":
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                        showErrorMessage=True, errorTitle="Date attendue", error="Saisir une date (jj/mm/aaaa).")
    dv.add(f"{col}{FIRST}:{col}{LAST}")
    ws.add_data_validation(dv)
dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                    showErrorMessage=True, errorTitle="Nombre d'heures",
                    error="Saisir un nombre d'heures positif.")
dv.add(f"M{FIRST}:M{LAST}")
ws.add_data_validation(dv)
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                    showErrorMessage=True, errorTitle="Taux d'avancement",
                    error="Saisir un pourcentage entre 0 % et 100 %.")
dv.add(f"O{FIRST}:O{LAST}")
ws.add_data_validation(dv)

# Contrôles visuels : charge ou date de début manquante, début après la date de besoin
miss = PatternFill("solid", start_color="FDE2E1", end_color="FDE2E1", bgColor="FDE2E1")
warn = PatternFill("solid", start_color="FFF1CC", end_color="FFF1CC", bgColor="FFF1CC")
ws.conditional_formatting.add(f"M{FIRST}:N{LAST}",
                              FormulaRule(formula=[f'AND($B{FIRST}<>"",M{FIRST}="")'], fill=miss))
ws.conditional_formatting.add(f"N{FIRST}:N{LAST}",
                              FormulaRule(formula=[f"AND(ISNUMBER($N{FIRST}),ISNUMBER($F{FIRST}),"
                                                   f"$N{FIRST}>$F{FIRST})"], fill=warn))
ws.conditional_formatting.add(f"O{FIRST}:O{LAST}",
                              DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                          color="86B6EF", showValue=True))

# ---------------------------------------------------------------- Dashboard : paramètres (nommés)
GRID0 = 2          # colonne Excel de la 1re colonne de la grille (B)
NG = 59            # nombre de colonnes de la grille


def gc(i):
    return GRID0 + i


def gl(i):
    return L(gc(i))


P_ROW = 5
PARAMS = {  # nom: (col début input, col fin input)
    "p_Perimetre": (4, 12),
    "p_Repartition": (21, 32),
    "p_PremierMois": (40, 44),
    "p_HeuresETP": (54, 58),
}
for name, (c1, _) in PARAMS.items():
    wb.defined_names[name] = DefinedName(name, attr_text=f"Dashboard!${gl(c1)}${P_ROW}")

# ---------------------------------------------------------------- Calculs
ws = ws_calc
ws.sheet_properties.tabColor = "9CA3AF"
ws["A1"] = "CALCULS INTERMÉDIAIRES — alimentent le Dashboard, ne pas modifier"
ws["A1"].font = font(12, True, NAVY)

LINE_COLS = {
    "A": "Ligne (onglet Charge)", "B": "Affaire valide (projet renseigné)", "C": "Avancement (0-1)",
    "D": "Incluse dans le périmètre", "E": "Charge incluse (h)", "F": "Chef de projet", "G": "Projet",
    "H": "Type", "I": "Début planifié", "J": "Fin de répartition", "K": "Durée (jours)",
    "L": "Rang affaire (charge décroissante)",
    "M": "1re occurrence chef", "N": "Charge du chef", "O": "Rang chef",
    "P": "1re occurrence projet", "Q": "Charge du projet", "R": "Rang projet",
    "S": "1re occurrence type", "T": "Charge du type", "U": "Rang type",
    "V": "Groupe couleur (projet)",
}
MONTH_COL0 = 23  # W
MCOLS = [L(MONTH_COL0 + m) for m in range(N_MONTHS)]  # W..AH
SUM_COL = L(MONTH_COL0 + N_MONTHS)                    # AI

hdr_font = font(9, True, WHITE)
for col, title in list(LINE_COLS.items()) + [(SUM_COL, "Charge sur la période (h)")]:
    c = ws[f"{col}3"]
    c.value = title
    c.font = hdr_font
    c.fill = fill(NAVY)
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.column_dimensions[col].width = 12
ws.row_dimensions[3].height = 54
ws["V2"] = "Fin de mois →"
ws["V2"].font = font(9, color=GREY, italic=True)

R1, R2 = FIRST, LAST
for r in range(R1, R2 + 1):
    f = {
        "A": f"=ROW()",
        "B": f'=IF(TRIM(Charge!B{r})<>"",1,0)',
        "C": f"=IF(B{r}=1,MIN(1,MAX(0,N(Charge!O{r}))),0)",
        "D": f'=IF(B{r}=1,IF(OR(p_Perimetre="{PERIM_ALL}",C{r}<1),1,0),0)',
        "E": f"=IF(D{r}=1,MAX(0,N(Charge!M{r})),0)",
        "F": f'=IF(D{r}=1,IF(TRIM(Charge!A{r})="","(non renseigné)",TRIM(Charge!A{r})),"")',
        "G": f'=IF(D{r}=1,TRIM(Charge!B{r}),"")',
        "H": f'=IF(D{r}=1,IF(TRIM(Charge!E{r})="","(non renseigné)",TRIM(Charge!E{r})),"")',
        "I": f'=IF(AND(D{r}=1,ISNUMBER(Charge!N{r})),Charge!N{r},"")',
        "J": f'=IF(I{r}="","",IF(AND(p_Repartition="{REP_SMOOTH}",ISNUMBER(Charge!F{r})),'
             f'IF(Charge!F{r}>=I{r},Charge!F{r},I{r}),I{r}))',
        "K": f'=IF(I{r}="","",J{r}-I{r}+1)',
        "L": f'=IF(D{r}=1,COUNTIFS($D${R1}:$D${R2},1,$E${R1}:$E${R2},">"&E{r})'
             f'+COUNTIFS($D${R1}:$D{r},1,$E${R1}:$E{r},E{r}),"")',
        "V": f'=IF(D{r}=1,MIN(INDEX($R${R1}:$R${R2},MATCH(G{r},$G${R1}:$G${R2},0)),'
             f'IF($AL$7>{N_SLOTS},{N_SLOTS},999)),"")',
    }
    # chefs (F), projets (G), types (H) : 1re occurrence, charge cumulée, rang décroissant
    for key, first, tot, rank in (("F", "M", "N", "O"), ("G", "P", "Q", "R"), ("H", "S", "T", "U")):
        f[first] = f'=IF({key}{r}="",0,IF(COUNTIF(${key}${R1}:{key}{r},{key}{r})=1,1,0))'
        f[tot] = f'=IF({first}{r}=1,SUMIF(${key}${R1}:${key}${R2},{key}{r},$E${R1}:$E${R2}),"")'
        f[rank] = (f'=IF({first}{r}=1,COUNTIFS(${first}${R1}:${first}${R2},1,${tot}${R1}:${tot}${R2},">"&{tot}{r})'
                   f'+COUNTIFS(${first}${R1}:{first}{r},1,${tot}${R1}:{tot}{r},{tot}{r}),"")')
    for col, formula in f.items():
        ws[f"{col}{r}"] = formula
    for mc in MCOLS:
        ws[f"{mc}{r}"] = (f'=IF(OR($I{r}="",$E{r}=0),0,IF(OR($J{r}<{mc}$3,$I{r}>{mc}$2),0,'
                          f'$E{r}*(MIN($J{r},{mc}$2)-MAX($I{r},{mc}$3)+1)/$K{r}))')
        ws[f"{mc}{r}"].number_format = "0.0"
    ws[f"{SUM_COL}{r}"] = f"=SUM({MCOLS[0]}{r}:{MCOLS[-1]}{r})"
    ws[f"{SUM_COL}{r}"].number_format = "0.0"
    for col in "IJ":
        ws[f"{col}{r}"].number_format = "DD/MM/YYYY"

# en-têtes de mois (début en ligne 3, fin en ligne 2)
for m, mc in enumerate(MCOLS):
    ws[f"{mc}3"] = "=$AL$14" if m == 0 else f"=EDATE({MCOLS[m - 1]}3,1)"
    ws[f"{mc}3"].number_format = "MMM YYYY"
    ws[f"{mc}2"] = f"=EOMONTH({mc}3,0)"
    ws[f"{mc}2"].number_format = "DD/MM/YYYY"
    ws[f"{mc}2"].font = font(9, color=GREY)

# Indicateurs (AK libellé / AL valeur)
K = {}
INDICATORS = [
    ("total", "Charge totale identifiée (h)", f"=SUM(E{R1}:E{R2})", "#,##0.0"),
    ("nbAff", "Nombre d'affaires", f"=SUM(D{R1}:D{R2})", "0"),
    ("nbChefs", "Nombre de chefs de projet", f"=SUM(M{R1}:M{R2})", "0"),
    ("nbProj", "Nombre de projets", f"=SUM(P{R1}:P{R2})", "0"),
    ("nbTypes", "Nombre de types", f"=SUM(S{R1}:S{R2})", "0"),
    ("realise", "Charge réalisée (selon avancement)", f"=SUMPRODUCT(E{R1}:E{R2},C{R1}:C{R2})", "#,##0.0"),
    ("reste", "Reste à faire (h)", "=AL4-AL9", "#,##0.0"),
    ("avgSimple", "Avancement moyen (simple)",
     f"=IF(AL5=0,0,SUMPRODUCT(D{R1}:D{R2},C{R1}:C{R2})/AL5)", "0.0%"),
    ("avgPond", "Avancement moyen pondéré par la charge", "=IF(AL4=0,0,AL9/AL4)", "0.0%"),
    ("autoStart", "1er mois automatique (1er début planifié)",
     f"=IF(COUNT(I{R1}:I{R2})=0,DATE(YEAR(TODAY()),MONTH(TODAY()),1),"
     f"DATE(YEAR(MIN(I{R1}:I{R2})),MONTH(MIN(I{R1}:I{R2})),1))", "MMM YYYY"),
    ("start", "1er mois affiché",
     "=IF(ISNUMBER(p_PremierMois),DATE(YEAR(p_PremierMois),MONTH(p_PremierMois),1),AL13)", "MMM YYYY"),
    ("periode", "Charge répartie sur la période affichée (h)", f"=SUM(AP4:AP{3 + N_MONTHS})", "#,##0.0"),
    ("hors", "Charge hors période ou sans date de début (h)", "=MAX(0,ROUND(AL4-AL15,1))", "#,##0.0"),
    ("pic", "Pic de charge mensuel (h)", f"=MAX(AP4:AP{3 + N_MONTHS})", "#,##0.0"),
    ("picIdx", "Rang du mois de pic", f"=MATCH(AL17,AP4:AP{3 + N_MONTHS},0)", "0"),
    ("picMois", "Mois le plus chargé", f"=INDEX(AO4:AO{3 + N_MONTHS},AL18)", "MMM YYYY"),
    ("picActives", "Affaires actives sur le mois de pic", f"=INDEX(AQ4:AQ{3 + N_MONTHS},AL18)", "0"),
    ("picETP", "ETP nécessaires au pic", "=IF(N(p_HeuresETP)<=0,0,AL17/p_HeuresETP)", "0.0"),
    ("totalETP", "Charge totale en ETP-mois", "=IF(N(p_HeuresETP)<=0,0,AL4/p_HeuresETP)", "0.0"),
    ("bigIdx", "Position de la plus grosse affaire", "=BX4", "0"),
    ("nbTermines", "Affaires terminées (avancement 100 %)", f"=COUNTIFS(B{R1}:B{R2},1,C{R1}:C{R2},\">=1\")", "0"),
    ("finPeriode", "Dernier mois affiché", f"=EDATE(AL14,{N_MONTHS - 1})", "MMM YYYY"),
    ("picPart", "Part du mois de pic dans la charge de la période", "=IF(AL15=0,0,AL17/AL15)", "0%"),
]
ws["AK3"], ws["AL3"] = "Indicateur", "Valeur"
for c in ("AK3", "AL3"):
    ws[c].font = hdr_font
    ws[c].fill = fill(NAVY)
ws.column_dimensions["AK"].width = 44
ws.column_dimensions["AL"].width = 14
for i, (key, label, formula, fmt) in enumerate(INDICATORS):
    r = 4 + i
    K[key] = f"Calculs!$AL${r}"
    ws[f"AK{r}"] = label
    ws[f"AL{r}"] = formula
    ws[f"AL{r}"].number_format = fmt
# contrôle de cohérence des références codées en dur (AL4, AL5, AL7, AL9, AL13...)
expected = {"total": 4, "nbAff": 5, "nbProj": 7, "realise": 9, "autoStart": 13, "start": 14,
            "periode": 15, "pic": 17, "picIdx": 18}
for k, row in expected.items():
    assert K[k] == f"Calculs!$AL${row}", (k, K[k])

# Charge par mois (AN..AQ)
for c, t in zip(("AN", "AO", "AP", "AQ"), ("N° mois", "Mois", "Charge (h)", "Affaires actives")):
    ws[f"{c}3"] = t
    ws[f"{c}3"].font = hdr_font
    ws[f"{c}3"].fill = fill(NAVY)
for m, mc in enumerate(MCOLS):
    r = 4 + m
    ws[f"AN{r}"] = m + 1
    ws[f"AO{r}"] = f"={mc}$3"
    ws[f"AO{r}"].number_format = "MMM YYYY"
    ws[f"AP{r}"] = f"=SUM({mc}{R1}:{mc}{R2})"
    ws[f"AP{r}"].number_format = "#,##0.0"
    ws[f"AQ{r}"] = f'=COUNTIF({mc}{R1}:{mc}{R2},">0")'

# Matrice des groupes de couleur (AS..BH) : séries des graphiques
SM = [L(49 + m) for m in range(N_MONTHS)]  # AW..BH
for c, t in zip(("AS", "AT", "AU", "AV"), ("Groupe", "Projet", "Charge (h)", "Part")):
    ws[f"{c}3"] = t
    ws[f"{c}3"].font = hdr_font
    ws[f"{c}3"].fill = fill(NAVY)
ws.column_dimensions["AT"].width = 22
for m, c in enumerate(SM):
    ws[f"{c}3"] = f"={MCOLS[m]}$3"
    ws[f"{c}3"].number_format = "MMM YY"
    ws[f"{c}3"].font = hdr_font
    ws[f"{c}3"].fill = fill(NAVY)
SLOT_R1 = 4
SLOT_R2 = SLOT_R1 + N_SLOTS - 1
TOT_ROW = SLOT_R2 + 1
for k in range(1, N_SLOTS + 1):
    r = SLOT_R1 + k - 1
    ws[f"AS{r}"] = k
    if k == N_SLOTS:
        ws[f"AT{r}"] = (f'=IF(AS{r}>$AL$7,"",IF($AL$7>{N_SLOTS},"Autres projets ("&($AL$7-{N_SLOTS - 1})&")",'
                        f'INDEX($G${R1}:$G${R2},MATCH(AS{r},$R${R1}:$R${R2},0))))')
    else:
        ws[f"AT{r}"] = f'=IF(AS{r}>$AL$7,"",INDEX($G${R1}:$G${R2},MATCH(AS{r},$R${R1}:$R${R2},0)))'
    ws[f"AU{r}"] = f"=SUMIF($V${R1}:$V${R2},AS{r},$E${R1}:$E${R2})"
    ws[f"AV{r}"] = f"=IF($AL$4=0,0,AU{r}/$AL$4)"
    ws[f"AV{r}"].number_format = "0.0%"
    for m, c in enumerate(SM):
        ws[f"{c}{r}"] = f"=SUMIF($V${R1}:$V${R2},$AS{r},{MCOLS[m]}${R1}:{MCOLS[m]}${R2})"
        ws[f"{c}{r}"].number_format = "0.0"
ws[f"AT{TOT_ROW}"] = "Total"
ws[f"AU{TOT_ROW}"] = f"=SUM(AU{SLOT_R1}:AU{SLOT_R2})"
for c in SM:
    ws[f"{c}{TOT_ROW}"] = f"=ROUND(SUM({c}{SLOT_R1}:{c}{SLOT_R2}),0)"
    ws[f"{c}{TOT_ROW}"].number_format = "#,##0"


def top_list(c_k, c_name, c_h, n, count_ref, key_col, tot_col, rank_col, other_label, title):
    """Liste des n premiers éléments par charge décroissante ; la dernière ligne regroupe le reste."""
    for c, t in zip((c_k, c_name, c_h), ("Rang", title, "Charge (h)")):
        ws[f"{c}3"] = t
        ws[f"{c}3"].font = hdr_font
        ws[f"{c}3"].fill = fill(NAVY)
    ws.column_dimensions[c_name].width = 22
    for k in range(1, n + 1):
        r = 3 + k
        ws[f"{c_k}{r}"] = k
        look = f"MATCH({c_k}{r},${rank_col}${R1}:${rank_col}${R2},0)"
        if k == n:
            ws[f"{c_name}{r}"] = (f'=IF({c_k}{r}>{count_ref},"",IF({count_ref}>{n},'
                                  f'"+ "&({count_ref}-{n - 1})&" {other_label}",'
                                  f'INDEX(${key_col}${R1}:${key_col}${R2},{look})))')
            ws[f"{c_h}{r}"] = (f'=IF({c_name}{r}="","",IF({count_ref}>{n},$AL$4-SUM({c_h}4:{c_h}{r - 1}),'
                               f'INDEX(${tot_col}${R1}:${tot_col}${R2},{look})))')
        else:
            ws[f"{c_name}{r}"] = f'=IF({c_k}{r}>{count_ref},"",INDEX(${key_col}${R1}:${key_col}${R2},{look}))'
            ws[f"{c_h}{r}"] = f'=IF({c_name}{r}="","",INDEX(${tot_col}${R1}:${tot_col}${R2},{look}))'


top_list("BK", "BL", "BM", 8, "$AL$6", "F", "N", "O", "autres", "Chef de projet")
top_list("BO", "BP", "BQ", 8, "$AL$7", "G", "Q", "R", "autres projets", "Projet")
top_list("BS", "BT", "BU", 4, "$AL$8", "H", "T", "U", "autres types", "Type")

# Tableau des affaires : position (dans Charge) de l'affaire de rang k
for c, t in zip(("BW", "BX"), ("Rang", "Position dans Charge")):
    ws[f"{c}3"] = t
    ws[f"{c}3"].font = hdr_font
    ws[f"{c}3"].fill = fill(NAVY)
for k in range(1, N_TABLE + 1):
    r = 3 + k
    ws[f"BW{r}"] = k
    ws[f"BX{r}"] = f'=IFERROR(MATCH(BW{r},$L${R1}:$L${R2},0),"")'

ws.freeze_panes = "B4"

# ---------------------------------------------------------------- Dashboard
ws = ws_dash
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 90
ws.column_dimensions["A"].width = 2
for i in range(NG):
    ws.column_dimensions[gl(i)].width = 2.71
ws.column_dimensions[L(gc(NG))].width = 2

ROW_H = {1: 6, 2: 30, 3: 18, 4: 8, 5: 22, 6: 14, 7: 10, 8: 20, 9: 6,
         10: 6, 11: 18, 12: 34, 13: 15, 14: 6, 23: 8, 24: 14, 25: 20, 26: 6,
         27: 6, 28: 18, 29: 34, 30: 15, 31: 15, 32: 8, 33: 14, 34: 20, 35: 6,
         36: 6, 37: 18, 38: 15, 55: 8, 56: 14, 57: 20, 58: 6, 59: 30}
for r in range(15, 23):
    ROW_H[r] = 16
for r in range(39, 55):
    ROW_H[r] = 15
TABLE_R1 = 60
TABLE_R2 = TABLE_R1 + N_TABLE - 1
for r in range(TABLE_R1, TABLE_R2 + 1):
    ROW_H[r] = 16.5
LAST_ROW = TABLE_R2 + 7
for r, h in ROW_H.items():
    ws.row_dimensions[r].height = h

# fond de page
for r in range(1, LAST_ROW + 1):
    for c in range(1, gc(NG) + 1):
        ws.cell(r, c).fill = fill(BG)


def box(r1, r2, c1, c2, bg=WHITE, border=LINE, top=None):
    """Carte : fond + bordure extérieure (c1, c2 = indices de grille)."""
    for r in range(r1, r2 + 1):
        for i in range(c1, c2 + 1):
            cell = ws.cell(r, gc(i))
            cell.fill = fill(bg)
            cell.border = Border(
                left=side(border) if i == c1 else None,
                right=side(border) if i == c2 else None,
                top=(top or side(border)) if r == r1 else None,
                bottom=side(border) if r == r2 else None)


def put(r, c1, c2, value, f=None, h="left", v="center", fmt=None, r2=None, bg=None, wrap=False, indent=0):
    r2 = r2 or r
    if c2 > c1 or r2 > r:
        ws.merge_cells(start_row=r, start_column=gc(c1), end_row=r2, end_column=gc(c2))
    cell = ws.cell(r, gc(c1))
    cell.value = value
    cell.font = f or font()
    cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)
    if fmt:
        cell.number_format = fmt
    if bg:
        cell.fill = fill(bg)
    return cell


def section(r, text):
    for i in range(NG):
        ws.cell(r, gc(i)).border = Border(bottom=side("B8C2CF"))
    c = ws.cell(r, gc(0))
    c.value = text
    c.font = font(11, True, NAVY)
    c.alignment = Alignment(vertical="bottom")


# Bandeau titre
for r in (2, 3):
    for i in range(NG):
        ws.cell(r, gc(i)).fill = fill(NAVY)
put(2, 1, 40, "TABLEAU DE BORD — PLAN DE CHARGE DU BUREAU D'ÉTUDES", font(17, True, WHITE), bg=NAVY)
put(3, 1, 44, "Synthèse automatique de l'onglet Charge : mettre à jour Charge et Data, "
               "le tableau de bord se recalcule seul.", font(9, color="C9D3E3"), bg=NAVY)
put(2, 44, 57, "=TODAY()", font(10, color="C9D3E3"), h="right", fmt='"Édité le "DD/MM/YYYY', bg=NAVY)

# Barre de paramètres
box(5, 6, 0, NG - 1)
labels = {"p_Perimetre": (0, 3, "Périmètre"), "p_Repartition": (14, 20, "Répartition mensuelle"),
          "p_PremierMois": (34, 39, "1er mois affiché"), "p_HeuresETP": (46, 53, "Heures / ETP / mois")}
defaults = {"p_Perimetre": PERIM_RUN, "p_Repartition": REP_SMOOTH, "p_PremierMois": None, "p_HeuresETP": 140}
hints = {"p_Perimetre": ("En cours = avancement < 100 %", None),
         "p_Repartition": (f'=IF(p_Repartition="{REP_START}","Toute la charge sur le mois de début planifié",'
                           f'"Prorata des jours entre début planifié et date de besoin")', None),
         "p_PremierMois": ('=IF(ISNUMBER(p_PremierMois),"Saisi manuellement",Calculs!$AL$13)',
                           '"Vide = auto ("MMM YYYY")"'),
         "p_HeuresETP": ("Conversion en ETP", None)}
inp_border = side("8FB3E6")
for name, (c1, c2) in PARAMS.items():
    lc1, lc2, text = labels[name]
    put(P_ROW, lc1, lc2, text, font(9, True, GREY), h="right", indent=1)
    for i in range(c1, c2 + 1):
        cell = ws.cell(P_ROW, gc(i))
        cell.fill = fill(INPUT_FILL)
        cell.border = Border(top=inp_border, bottom=inp_border,
                             left=inp_border if i == c1 else None, right=inp_border if i == c2 else None)
    cell = put(P_ROW, c1, c2, defaults[name], font(10, True, INPUT_INK), h="center", bg=INPUT_FILL,
               fmt="MMMM YYYY" if name == "p_PremierMois" else None)
    put(6, c1, c2, hints[name][0], font(7, color=MUTED, italic=True), h="center", v="top", fmt=hints[name][1])

dv = DataValidation(type="list", formula1=f'"{PERIM_RUN},{PERIM_ALL}"', allow_blank=False)
dv.add(f"{gl(PARAMS['p_Perimetre'][0])}{P_ROW}")
ws.add_data_validation(dv)
dv = DataValidation(type="list", formula1=f'"{REP_SMOOTH},{REP_START}"', allow_blank=False)
dv.add(f"{gl(PARAMS['p_Repartition'][0])}{P_ROW}")
ws.add_data_validation(dv)
dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                    showErrorMessage=True, error="Saisir une date (le 1er du mois par exemple) ou laisser vide.",
                    showInputMessage=True, promptTitle="1er mois affiché",
                    prompt="Laisser vide pour démarrer au premier mois de début planifié.")
dv.add(f"{gl(PARAMS['p_PremierMois'][0])}{P_ROW}")
ws.add_data_validation(dv)
dv = DataValidation(type="decimal", operator="between", formula1="1", formula2="744", allow_blank=False,
                    showErrorMessage=True, error="Saisir un nombre d'heures par mois (ex. 140).")
dv.add(f"{gl(PARAMS['p_HeuresETP'][0])}{P_ROW}")
ws.add_data_validation(dv)

# ---- Section 1 : projets en cours
section(8, f'=IF(p_Perimetre="{PERIM_ALL}","TOUS LES PROJETS","PROJETS EN COURS")')
CARDS = [(0, 13), (15, 28), (30, 43), (45, 58)]
H_FMT = '#,##0" h"'
title_f = font(9, True, GREY)
big_f = font(24, True, NAVY)
sub_f = font(8.5, color=GREY)
item_f = font(9.5, color=INK)
item_h = font(9, color=GREY)
for idx, (c1, c2) in enumerate(CARDS):
    if idx == 3:
        box(10, 23, c1, c2, bg=NAVY, border=NAVY, top=side(AMBER, "thick"))
    else:
        box(10, 23, c1, c2)
    if idx < 3:
        for i in range(c1 + 1, c2):
            ws.cell(14, gc(i)).border = Border(bottom=side("E5E7EB"))

# Carte 1 : chefs de projet
c1, c2 = CARDS[0]
put(11, c1 + 1, c2 - 1, "CHEFS DE PROJET", title_f)
put(12, c1 + 1, c2 - 1, f"={K['nbChefs']}", big_f, fmt="0")
put(13, c1 + 1, c2 - 1, f'=IF(p_Perimetre="{PERIM_ALL}","sur l\'ensemble des projets","sur les projets en cours")',
    sub_f)
for k in range(8):
    r = 15 + k
    put(r, c1 + 1, c2 - 4, f"=Calculs!BL{4 + k}", item_f)
    put(r, c2 - 3, c2 - 1, f"=Calculs!BM{4 + k}", item_h, h="right", fmt=H_FMT)

# Carte 2 : programmes (projets) suivis
c1, c2 = CARDS[1]
put(11, c1 + 1, c2 - 1, "PROGRAMMES SUIVIS", title_f)
put(12, c1 + 1, c2 - 1, f"={K['nbProj']}", big_f, fmt="0")
put(13, c1 + 1, c2 - 1, "projets distincts · classés par charge", sub_f)
for k in range(8):
    r = 15 + k
    sw = ws.cell(r, gc(c1 + 1))
    sw.fill = fill(SLOT_COLORS[k])
    sw.border = Border(top=side(WHITE, "thick"), bottom=side(WHITE, "thick"), right=side(WHITE, "medium"))
    put(r, c1 + 2, c2 - 4, f"=Calculs!BP{4 + k}", item_f, indent=0)
    put(r, c2 - 3, c2 - 1, f"=Calculs!BQ{4 + k}", item_h, h="right", fmt=H_FMT)
    name_ref = f"{gl(c1 + 2)}{r}"
    ws.conditional_formatting.add(
        f"{gl(c1 + 1)}{r}",
        FormulaRule(formula=[f'OR({name_ref}="",LEFT({name_ref},1)="+")'],
                    fill=PatternFill("solid", start_color=WHITE, end_color=WHITE, bgColor=WHITE)))

# Carte 3 : charge identifiée
c1, c2 = CARDS[2]
put(11, c1 + 1, c2 - 1, "CHARGE IDENTIFIÉE", title_f)
put(12, c1 + 1, c2 - 1, f"={K['total']}", big_f, fmt=H_FMT)
put(13, c1 + 1, c2 - 1, f"={K['totalETP']}", sub_f, fmt='"soit ≈ "0.0" ETP-mois"')
put(15, c1 + 1, c2 - 4, "Réalisé (selon avancement)", item_f)
put(15, c2 - 3, c2 - 1, f"={K['realise']}", font(9.5, True, INK), h="right", fmt=H_FMT)
put(16, c1 + 1, c2 - 4, "Reste à faire", item_f)
put(16, c2 - 3, c2 - 1, f"={K['reste']}", font(9.5, True, INK), h="right", fmt=H_FMT)
put(18, c1 + 1, c2 - 1, "PAR TYPE D'AFFAIRE", font(8, True, MUTED))
for k in range(4):
    r = 19 + k
    put(r, c1 + 1, c2 - 4, f"=Calculs!BT{4 + k}", item_f)
    put(r, c2 - 3, c2 - 1, f"=Calculs!BU{4 + k}", item_h, h="right", fmt=H_FMT)

# Carte 4 : plus grosse affaire (mise en avant)
c1, c2 = CARDS[3]
IDX = K["bigIdx"]
pop_lab = font(8.5, color="AEB9CC")
pop_val = font(9.5, True, WHITE)
put(11, c1 + 1, c2 - 1, "★  PLUS GROSSE AFFAIRE", font(9, True, AMBER), bg=NAVY)
put(12, c1 + 1, c2 - 1, f'=IF({IDX}="",0,INDEX(Calculs!$E${R1}:$E${R2},{IDX}))', font(24, True, WHITE),
    fmt=H_FMT, bg=NAVY)
put(13, c1 + 1, c2 - 1, f'=IF(OR({IDX}="",{K["total"]}=0),0,INDEX(Calculs!$E${R1}:$E${R2},{IDX})/{K["total"]})',
    font(8.5, color="C9D3E3"), fmt='"soit "0%" de la charge identifiée"', bg=NAVY)
for i in range(c1 + 1, c2):
    ws.cell(14, gc(i)).border = Border(bottom=side("3B4A6B"))
POP = [
    ("Projet", "B", "text"), ("Sous-projet", "C", "text"), ("Produit", "D", "text"),
    ("Chef de projet", "A", "text"), ("Type", "E", "text"), ("Début planifié", "N", "date"),
    ("Date de besoin", "F", "date"), ("Avancement", "O", "pct"),
]
for k, (lab, col, kind) in enumerate(POP):
    r = 15 + k
    put(r, c1 + 1, c1 + 5, lab, pop_lab, bg=NAVY)
    src = f"INDEX(Charge!${col}${R1}:${col}${R2},{IDX})"
    if kind == "text":
        formula, fmt = f'=IF({IDX}="","—",{src}&"")', None
    elif kind == "date":
        formula, fmt = f'=IF({IDX}="","",IF(ISNUMBER({src}),{src},"—"))', "DD/MM/YYYY"
    else:
        formula, fmt = f'=IF({IDX}="","",INDEX(Calculs!$C${R1}:$C${R2},{IDX}))', "0%"
    put(r, c1 + 6, c2 - 1, formula, pop_val, fmt=fmt, bg=NAVY)

# ---- Section 2 : indicateurs clés
section(25, "INDICATEURS CLÉS")
TILES = [(i * 12, i * 12 + 10) for i in range(5)]
kpi_val = font(22, True, NAVY)
KPIS = [
    ("CHARGE TOTALE IDENTIFIÉE", f"={K['total']}", H_FMT,
     (f"={K['reste']}", '"Reste à faire : "#,##0" h"'),
     (f"={K['nbAff']}", '"sur "0" affaires"')),
    ("PIC DE CHARGE", f"={K['pic']}", H_FMT,
     (f"={K['picMois']}", '"en "MMMM YYYY'),
     (f"={K['picETP']}", '"≈ "0.0" ETP sur le mois"')),
    ("MOIS LE PLUS CHARGÉ", f'=IF({K["pic"]}=0,"—",{K["picMois"]})', "MMMM YYYY",
     (f"={K['picActives']}", '0" affaire(s) active(s) ce mois-ci"'),
     (f"={K['picPart']}", '0%" de la charge sur 12 mois"')),
    ("NOMBRE D'AFFAIRES", f"={K['nbAff']}", "0",
     (f"={K['nbProj']}", '"réparties sur "0" projet(s)"'),
     (f'=IF(p_Perimetre="{PERIM_ALL}","dont "&{K["nbTermines"]}&" terminée(s)",'
      f'{K["nbTermines"]}&" terminée(s) non comptée(s)")', None)),
    ("TAUX D'AVANCEMENT MOYEN", f"={K['avgSimple']}", "0%",
     (f"={K['avgPond']}", '"Pondéré par la charge : "0%'),
     ("moyenne simple par affaire", None)),
]
for (c1, c2), (title, formula, fmt, s1, s2) in zip(TILES, KPIS):
    box(27, 32, c1, c2, top=side(ACCENT, "thick"))
    put(28, c1 + 1, c2 - 1, title, title_f)
    put(29, c1 + 1, c2 - 1, formula, kpi_val, fmt=fmt)
    put(30, c1 + 1, c2 - 1, s1[0], sub_f, fmt=s1[1])
    put(31, c1 + 1, c2 - 1, s2[0], sub_f, fmt=s2[1])

# ---- Section 3 : graphiques
section(34, "RÉPARTITION DE LA CHARGE")
LP = (0, 30)
RP = (32, 58)
box(36, 55, *LP)
box(36, 55, *RP)
put(37, LP[0] + 1, LP[1] - 1, "CHARGE EN HEURES PAR MOIS ET PAR PROJET", title_f)
put(38, LP[0] + 1, LP[0] + 14, f"={K['start']}", sub_f, fmt='"12 mois à partir de "MMMM YYYY')
put(38, LP[0] + 15, LP[1] - 1, f"={K['hors']}", sub_f, h="right",
    fmt='"Hors période ou sans date de début : "#,##0" h";;"Toute la charge est visible sur la période"')
put(37, RP[0] + 1, RP[1] - 1, "RÉPARTITION DE LA CHARGE PAR PROJET", title_f)
put(38, RP[0] + 1, RP[1] - 1, "Total au centre · une couleur par projet, commune aux deux graphiques", sub_f)

CH_R1, CH_R2 = 39, 54
DN = (33, 45)       # colonnes de grille du donut
# total au centre du donut (le graphique est transparent)
put(45, DN[0], DN[1], f"={K['total']}", font(18, True, NAVY), h="center", v="bottom", fmt='#,##0" h"', r2=46)
put(47, DN[0], DN[1], "au total", font(9, color=GREY), h="center", v="top", r2=48)

# légende commune
LG = {"sw": 47, "name": (48, 52), "h": (53, 55), "p": (56, 57)}
put(42, *LG["name"], "Projet", font(8, True, MUTED))
put(42, *LG["h"], "Heures", font(8, True, MUTED), h="right")
put(42, *LG["p"], "Part", font(8, True, MUTED), h="right")
for k in range(N_SLOTS):
    r = 43 + k
    sw = ws.cell(r, gc(LG["sw"]))
    sw.fill = fill(SLOT_COLORS[k])
    sw.border = Border(top=side(WHITE, "thick"), bottom=side(WHITE, "thick"), right=side(WHITE, "medium"))
    put(r, *LG["name"], f"=Calculs!AT{SLOT_R1 + k}", font(9, color=INK))
    put(r, *LG["h"], f'=IF(Calculs!AT{SLOT_R1 + k}="","",Calculs!AU{SLOT_R1 + k})', font(9, color=GREY),
        h="right", fmt="#,##0")
    put(r, *LG["p"], f'=IF(Calculs!AT{SLOT_R1 + k}="","",Calculs!AV{SLOT_R1 + k})', font(9, color=GREY),
        h="right", fmt="0%")
    ws.conditional_formatting.add(
        f"{gl(LG['sw'])}{r}",
        FormulaRule(formula=[f'{gl(LG["name"][0])}{r}=""'],
                    fill=PatternFill("solid", start_color=WHITE, end_color=WHITE, bgColor=WHITE)))


def text_props(size=800, color="595959", bold=False):
    cp = CharacterProperties(latin=DFont(typeface=FONT), sz=size, b=bold, solidFill=color)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def no_fill():
    return GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))


def anchor(c1, r1, c2, r2):
    a = TwoCellAnchor()
    a._from = AnchorMarker(col=gc(c1) - 1, row=r1 - 1, colOff=0, rowOff=0)
    a.to = AnchorMarker(col=gc(c2), row=r2, colOff=0, rowOff=0)
    return a


# Histogramme empilé : heures par mois et par projet + total au-dessus de chaque colonne
bar = BarChart()
bar.type = "col"
bar.grouping = "stacked"
bar.overlap = 100
bar.gapWidth = 55
for k in range(N_SLOTS):
    r = SLOT_R1 + k
    bar.add_data(Reference(ws_calc, min_col=49, max_col=48 + N_MONTHS, min_row=r, max_row=r),
                 from_rows=True, titles_from_data=False)
    s = bar.series[-1]
    s.tx = SeriesLabel(strRef=StrRef(f"Calculs!$AT${r}"))
    s.graphicalProperties = GraphicalProperties(solidFill=SLOT_COLORS[k])
    s.graphicalProperties.line = LineProperties(solidFill=WHITE, w=9525)
cats = Reference(ws_calc, min_col=49, max_col=48 + N_MONTHS, min_row=3, max_row=3)
bar.set_categories(cats)

line = LineChart()
line.add_data(Reference(ws_calc, min_col=49, max_col=48 + N_MONTHS, min_row=TOT_ROW, max_row=TOT_ROW),
              from_rows=True, titles_from_data=False)
ls = line.series[0]
ls.tx = SeriesLabel(strRef=StrRef(f"Calculs!$AT${TOT_ROW}"))
ls.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
ls.marker = Marker(symbol="none")
ls.smooth = False
ls.dLbls = DataLabelList(numFmt='#,##0;;;', dLblPos="t", showVal=True, showLegendKey=False,
                         showCatName=False, showSerName=False, showPercent=False, showBubbleSize=False,
                         showLeaderLines=False)
ls.dLbls.txPr = text_props(800, "374151", True)
line.set_categories(cats)
bar += line

bar.legend = None
bar.title = None
bar.roundedCorners = False
bar.graphical_properties = no_fill()
bar.plot_area.graphicalProperties = no_fill()
bar.y_axis.delete = False
bar.x_axis.delete = False
bar.y_axis.title = None
bar.y_axis.numFmt = NumFmt(formatCode='#,##0" h"', sourceLinked=False)
bar.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E5E7EB", w=6350))
bar.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
bar.y_axis.txPr = text_props(800, "6B7280")
bar.y_axis.majorTickMark = "none"
bar.x_axis.numFmt = NumFmt(formatCode="mmm yy", sourceLinked=False)
bar.x_axis.auto = False
bar.x_axis.txPr = text_props(800, "6B7280")
bar.x_axis.majorTickMark = "none"
bar.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="BFC6D1", w=9525))
bar.y_axis.scaling.min = 0
bar.anchor = anchor(LP[0] + 1, CH_R1, LP[1] - 1, CH_R2)
ws.add_chart(bar)

# Anneau : répartition par projet (sans légende ni titre ; total affiché dans la cellule sous le trou)
donut = DoughnutChart(holeSize=62)
donut.add_data(Reference(ws_calc, min_col=47, min_row=SLOT_R1, max_row=SLOT_R2), titles_from_data=False)
donut.set_categories(Reference(ws_calc, min_col=46, min_row=SLOT_R1, max_row=SLOT_R2))
ds = donut.series[0]
ds.tx = SeriesLabel(v="Charge (h)")
for k, color in enumerate(SLOT_COLORS):
    pt = DataPoint(idx=k)
    pt.graphicalProperties = GraphicalProperties(solidFill=color)
    pt.graphicalProperties.line = LineProperties(solidFill=WHITE, w=19050)
    ds.dPt.append(pt)
donut.firstSliceAng = 0
donut.varyColors = True
donut.legend = None
donut.title = None
donut.roundedCorners = False
donut.graphical_properties = no_fill()
donut.plot_area.graphicalProperties = no_fill()
donut.layout = Layout(manualLayout=ManualLayout(layoutTarget="inner", xMode="edge", yMode="edge",
                                                x=0.04, y=0.04, w=0.92, h=0.92))
donut.anchor = anchor(DN[0], CH_R1, DN[1], CH_R2)
ws.add_chart(donut)

# ---- Section 4 : liste des affaires
section(57, "LISTE DES AFFAIRES (par charge décroissante)")
TCOLS = [  # (titre, nb colonnes de grille, alignement, format)
    ("N°", 2, "center", "0"),
    ("Projet", 9, "left", None),
    ("Sous-projet", 9, "left", None),
    ("Produit", 9, "left", None),
    ("Chef de projet", 8, "left", None),
    ("Charge (h)", 5, "right", "#,##0"),
    ("% de la charge totale", 6, "right", "0.0%"),
    ("Avancement", 6, "right", "0%"),
    ("Date de début planifiée", 5, "center", "DD/MM/YYYY"),
]
spans = []
c = 0
for title, w, al, fmt in TCOLS:
    spans.append((c, c + w - 1))
    put(59, c, c + w - 1, title, font(9, True, WHITE), h="center", bg=NAVY, wrap=True)
    for i in range(c, c + w):
        ws.cell(59, gc(i)).fill = fill(NAVY)
    c += w
assert c == NG

row_line = side("E5E7EB")
for k in range(1, N_TABLE + 1):
    r = TABLE_R1 + k - 1
    idx = f"Calculs!$BX${3 + k}"
    for i in range(NG):
        cell = ws.cell(r, gc(i))
        cell.fill = fill(BG)
    vals = [
        f'=IF({idx}="","",Calculs!$BW${3 + k})',
        f'=IF({idx}="","",INDEX(Charge!$B${R1}:$B${R2},{idx})&"")',
        f'=IF({idx}="","",INDEX(Charge!$C${R1}:$C${R2},{idx})&"")',
        f'=IF({idx}="","",INDEX(Charge!$D${R1}:$D${R2},{idx})&"")',
        f'=IF({idx}="","",INDEX(Charge!$A${R1}:$A${R2},{idx})&"")',
        f'=IF({idx}="","",INDEX(Calculs!$E${R1}:$E${R2},{idx}))',
        f'=IF({idx}="","",IF({K["total"]}=0,0,INDEX(Calculs!$E${R1}:$E${R2},{idx})/{K["total"]}))',
        f'=IF({idx}="","",INDEX(Calculs!$C${R1}:$C${R2},{idx}))',
        f'=IF({idx}="","",IF(ISNUMBER(INDEX(Charge!$N${R1}:$N${R2},{idx})),'
        f'INDEX(Charge!$N${R1}:$N${R2},{idx}),"à planifier"))',
    ]
    for (title, w, al, fmt), (s1, s2), formula in zip(TCOLS, spans, vals):
        bold = title in ("Projet",)
        put(r, s1, s2, formula, font(9, bold, INK if title != "N°" else GREY), h=al, fmt=fmt, indent=1 if al != "center" else 0)

first_cell = f"{gl(0)}{TABLE_R1}"
tbl_range = f"{gl(0)}{TABLE_R1}:{gl(NG - 1)}{TABLE_R2}"
ws.conditional_formatting.add(
    tbl_range,
    FormulaRule(formula=[f'AND(${gl(0)}{TABLE_R1}<>"",MOD(ROW(),2)=0)'], border=Border(bottom=row_line),
                fill=PatternFill("solid", start_color="F5F7FA", end_color="F5F7FA", bgColor="F5F7FA")))
ws.conditional_formatting.add(
    tbl_range,
    FormulaRule(formula=[f'${gl(0)}{TABLE_R1}<>""'], border=Border(bottom=row_line),
                fill=PatternFill("solid", start_color=WHITE, end_color=WHITE, bgColor=WHITE)))
pct_col = gl(spans[6][0])
adv_col = gl(spans[7][0])
ws.conditional_formatting.add(f"{pct_col}{TABLE_R1}:{pct_col}{TABLE_R2}",
                              DataBarRule(start_type="num", start_value=0, end_type="max",
                                          color="86B6EF", showValue=True))
ws.conditional_formatting.add(f"{adv_col}{TABLE_R1}:{adv_col}{TABLE_R2}",
                              DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                          color="7FD1AE", showValue=True))

NOTE_R = TABLE_R2 + 2
put(NOTE_R, 0, NG - 1,
    f'=IF({K["nbAff"]}>{N_TABLE},"Attention : seules les {N_TABLE} premières affaires (par charge) sont '
    f'listées sur "&{K["nbAff"]}&".","")', font(9, True, "B45309"))
NOTES = [
    "Affaire = une ligne de l'onglet Charge dont le nom de projet (colonne B) est renseigné.  "
    "Projet en cours = affaire dont le taux d'avancement (colonne O) est inférieur à 100 %.",
    "Charge par mois : la charge estimée (M) est répartie au prorata des jours calendaires entre la date de "
    "début planifiée (N) et la date de besoin (F). Sans date de besoin, ou si elle précède le début, "
    "toute la charge est placée sur le mois de début.",
    "Au-delà de 9 projets, les plus petits sont regroupés en « Autres projets » (gris) dans les graphiques.  "
    "ETP = heures ÷ heures productives par ETP et par mois (paramètre en haut de page, 140 h par défaut).",
]
for i, t in enumerate(NOTES):
    put(NOTE_R + 1 + i, 0, NG - 1, t, font(8, color=GREY, italic=True))

ws.print_area = f"A1:{L(gc(NG))}{LAST_ROW}"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.scale = 90
ws.row_breaks.append(Break(id=56))
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.4
ws.sheet_properties.tabColor = ACCENT

wb.active = wb.sheetnames.index("Dashboard")
for sheet in wb.worksheets:
    sheet.sheet_view.tabSelected = sheet.title == "Dashboard"
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("OK ->", OUT)

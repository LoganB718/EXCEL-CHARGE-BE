# -*- coding: utf-8 -*-
"""Contrôles statiques du modèle sémantique (model_def.build_model()).

0. Témoins : les deux analyseurs rejettent bien des expressions fausses (M et DAX) et acceptent des formes valides.
1. Expressions M (partitions + expressions partagées) analysées par @microsoft/powerquery-parser (check_m.js) ;
   références entre requêtes (existence, pas de cycle, pas de variable locale masquant une requête) ; chaque
   colonne déclarée est produite par sa requête.
2. Mesures DAX analysées par pbi_parsers (fin de texte atteinte) ; fonctions connues avec un nombre d'arguments
   valide, aucun argument vide (l'analyseur les tolère) ; chaque 'Table'[Colonne], Table[Colonne], [Mesure] et table
   référencée existe (arbre syntaxique + contrôle textuel) ; variables déclarées avant usage.
3. Modèle : noms uniques ; aucune mesure homonyme d'une colonne ou d'une table du modèle ; relations (colonnes,
   types, un seul chemin de filtre entre deux tables) ; colonnes de tri ; clés et numéros non additionnés ;
   lineageTag uniques ; annotations usuelles de Power BI Desktop ; descriptions des mesures ; textes en français.

Usage : <venv>/bin/python -I check_model.py [dossier_de_sortie] [--modele autre_model_def.py]
(le dossier reçoit model.bim et les expressions M extraites ; zéro erreur exigé, code de sortie 1 sinon ;
--modele sert aux essais d'injection de fautes de fautes_check_model.py)
"""
import importlib.util
import json
import os
import re
import subprocess
import sys
import tempfile

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
from pbi_parsers import dax  # noqa: E402
from pbi_parsers.dax import exprs as dax_exprs  # noqa: E402

ARGS = sys.argv[1:]
if "--modele" in ARGS:
    i = ARGS.index("--modele")
    spec = importlib.util.spec_from_file_location("model_def", ARGS[i + 1])
    model_def = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(model_def)
    del ARGS[i:i + 2]
else:
    import model_def  # noqa: E402

SORTIE = ARGS[0] if ARGS else tempfile.mkdtemp(prefix="check_model_")
os.makedirs(SORTIE, exist_ok=True)
CHECK_M = os.path.join(ICI, "check_m.js")
erreurs = []


def erreur(msg):
    erreurs.append(msg)
    print("ERREUR", msg)


def analyse_m(expressions, globaux=None):
    """Analyse {nom: M} par check_m.js (identifiants résolus si `globaux` = noms des requêtes) ;
    renvoie (code de sortie, lignes de sortie)."""
    fichiers = []
    for contenu in (expressions, globaux):
        if contenu is None:
            continue
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump(contenu, f, ensure_ascii=False)
            fichiers.append(f.name)
    try:
        res = subprocess.run(["node", CHECK_M] + fichiers, capture_output=True, text=True)
    finally:
        for chemin in fichiers:
            os.unlink(chemin)
    return res.returncode, (res.stdout + res.stderr).strip().splitlines()


# fonctions et constantes de la bibliothèque standard M utilisées (référence Power Query M de Microsoft)
BIBLIOTHEQUE_M = {
    "Date.AddDays", "Date.AddMonths", "Date.Day", "Date.DayOfWeek", "Date.EndOfMonth", "Date.From", "Date.Month",
    "Date.StartOfMonth", "Date.Year", "DateTimeZone.RemoveZone", "DateTimeZone.UtcNow", "Day.Monday", "Day.Sunday",
    "Duration.Days", "Duration.TotalDays", "Error.Record", "Excel.Workbook", "Int64.Type", "JoinKind.Inner",
    "JoinKind.LeftOuter", "List.Average", "List.Buffer", "List.Contains", "List.Count", "List.Distinct",
    "List.IsEmpty", "List.Max", "List.Min", "List.Numbers", "List.PositionOf", "List.RemoveNulls", "List.Repeat",
    "List.Select", "List.Sort", "List.Sum", "List.Transform", "List.Zip", "MissingField.UseNull",
    "Number.IntegerDivide", "Number.Mod", "Number.Power", "Number.Round", "Number.ToText", "Order.Ascending",
    "Order.Descending", "Record.Field", "RoundingMode.AwayFromZero", "Splitter.SplitByNothing", "Table.AddColumn",
    "Table.AddIndexColumn", "Table.Buffer", "Table.Column", "Table.ColumnCount", "Table.ColumnNames", "Table.Combine",
    "Table.ExpandListColumn", "Table.ExpandRecordColumn", "Table.ExpandTableColumn", "Table.FirstN",
    "Table.FirstValue", "Table.FromColumns", "Table.FromList", "Table.FromRecords", "Table.Group", "Table.IsEmpty",
    "Table.Min", "Table.NestedJoin", "Table.RemoveColumns", "Table.RenameColumns", "Table.ReplaceErrorValues",
    "Table.RowCount", "Table.SelectColumns", "Table.SelectRows", "Table.Sort", "Table.ToRecords",
    "Table.TransformColumnTypes", "Text.Combine", "Text.From", "Text.Lower", "Text.PadStart", "Text.Replace",
    "Text.Split", "Text.Start", "Time.Hour", "Time.Minute", "Value.Type", "Web.Contents"}


def analyse_dax(code):
    """Arbre de pbi_parsers ; exception si le texte n'est pas entièrement reconnu."""
    ast = dax.to_ast(code)
    if ast is None:
        raise ValueError("aucune expression reconnue")
    return ast


# ------------------------------------------------------------------------------------------------ 0. témoins
TEMOINS_M = {
    "faux_let_incomplet": "let a = 1 in",
    "faux_parenthese": "let a = (1 + 2 in a",
    "faux_operateur": 'let a = Table.AddColumn(t, "x", each [b] +) in a',
    "faux_virgule": "let a = 1, in a",
    "juste_each": 'let t = #table({"b c"}, {{1}}), a = Table.AddColumn(t, "x", each [#"b c"] * 2, type number) in a',
    "juste_meta": '"x" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]',
    "faux_identifiant": "let a = 1, b = c + 1 in b",
    "faux_portee": "let f = (x) => x + 1, g = x in g",
    "juste_portees": 'let f = (x) => [a = x, b = a + 1][b], t = #table({"c"}, {{1}}) in List.Transform({1}, each _ + f(1))',
}
code_m, sortie_m = analyse_m(TEMOINS_M, [])
for nom in TEMOINS_M:
    ligne = next((x for x in sortie_m if f" {nom} " in x + " " or x.endswith(f" {nom}")), "")
    attendu_ok = nom.startswith("juste")
    if ligne.startswith("OK") != attendu_ok:
        erreur(f"témoin M {nom} : {'accepté' if not attendu_ok else 'rejeté'} à tort ({ligne})")
print(f"témoins M : {len(TEMOINS_M)} expressions, analyseur discriminant : "
      f"{'oui' if not any(e.startswith('témoin M') for e in erreurs) else 'non'}")
TEMOINS_DAX = {"faux_parenthese": "SUM ( Lignes[Heures (h)] ", "faux_return": "VAR _a = 1 RETURN",
               "faux_operateur": "1 +", "faux_arguments": "IF ( 1, 2 3 )",
               "juste_var": "VAR _a = 1 RETURN _a & \"x\"", "juste_in": "x IN { \"a\", \"b\" }",
               "juste_calculate_var": "CALCULATE ( VAR _t = 1 RETURN _t, 'Périmètre'[Périmètre] = \"x\" )"}
for nom, code in TEMOINS_DAX.items():
    try:
        analyse_dax(code)
        ok = True
    except Exception:  # noqa: BLE001 - toute erreur de l'analyseur = expression rejetée
        ok = False
    if ok != nom.startswith("juste"):
        erreur(f"témoin DAX {nom} : {'accepté' if ok else 'rejeté'} à tort")
print(f"témoins DAX : {len(TEMOINS_DAX)} expressions, analyseur discriminant : "
      f"{'oui' if not any(e.startswith('témoin DAX') for e in erreurs) else 'non'}")

# ------------------------------------------------------------------------------------------------ modèle
bim = model_def.build_model()
model = bim["model"]
with open(os.path.join(SORTIE, "model.bim"), "w", encoding="utf-8") as f:
    json.dump(bim, f, ensure_ascii=False, indent=2)

# ------------------------------------------------------------------------------------------------ 1. M
m = model_def.m_expressions(model)
with open(os.path.join(SORTIE, "expressions_m.json"), "w", encoding="utf-8") as f:
    json.dump(m, f, ensure_ascii=False, indent=1)
code_m, sortie_m = analyse_m(m, sorted(m))
print("\n".join(x for x in sortie_m if not x.startswith(("OK ", "BIBLIOTHÈQUE "))))
biblio = {x.split(" ", 1)[1] for x in sortie_m if x.startswith("BIBLIOTHÈQUE ")}
for x in sorted(biblio - BIBLIOTHEQUE_M):
    erreur(f"M : fonction ou constante de bibliothèque inconnue {x}")
print(f"{len(biblio)} fonctions et constantes de la bibliothèque M, toutes connues : "
      f"{'oui' if biblio <= BIBLIOTHEQUE_M else 'non'}")
if code_m != 0:
    erreur(f"analyse M : code {code_m}")
print(f"{sum(1 for x in sortie_m if x.startswith('OK '))} expressions M reconnues par @microsoft/powerquery-parser")

noms_requetes = set(m)
dependances = {}
for nom, code in m.items():
    sans_textes = re.sub(r'"(?:[^"]|"")*"', '""', code)
    ids = set(re.findall(r'#"((?:[^"]|"")*)"', code)) | set(re.findall(r"(?<![\w.#\[])([^\W\d]\w*)", sans_textes))
    dependances[nom] = {i for i in ids if i in noms_requetes and i != nom}
for nom, code in m.items():
    locales = set(re.findall(r'^\s*([^\W\d]\w*)\s*=(?!=)', code, re.M))
    locales |= set(re.findall(r'[(,]\s*([^\W\d]\w*)\s+as\s', code))
    for v in sorted(locales & noms_requetes):
        erreur(f"requête M {nom} : la variable locale {v} masque la requête du même nom")
        dependances[nom].discard(v)
vus, en_cours = set(), set()


def visite(n, chemin):
    if n in en_cours:
        erreur("cycle de requêtes M : " + " -> ".join(chemin + [n]))
        return
    if n in vus:
        return
    en_cours.add(n)
    for dep in sorted(dependances[n]):
        visite(dep, chemin + [n])
    en_cours.discard(n)
    vus.add(n)


for n in sorted(m):
    visite(n, [])
print(f"{len(m)} requêtes M, dépendances sans cycle : {'oui' if not any('cycle' in e for e in erreurs) else 'non'}")
# une seule source de données (pas de Formula.Firewall) : seule la requête Classeur appelle Web.Contents
acces = [(n, x) for n, code in m.items() for x in re.findall(
    r"\b(Web\.\w+|File\.\w+|Folder\.\w+|SharePoint\.\w+|Sql\.\w+|OData\.\w+|Odbc\.\w+|Csv\.\w+|Json\.\w+)\s*\(", code)]
if acces != [("Classeur", "Web.Contents")]:
    erreur(f"accès aux données : un seul Web.Contents attendu dans Classeur, trouvé {acces}")
if 'Excel.Workbook(Web.Contents(UrlFichier), null, true)' not in m["Classeur"]:
    erreur("Classeur : lecture attendue Excel.Workbook(Web.Contents(UrlFichier), null, true)")


def amont(n, acc=None):
    acc = set() if acc is None else acc
    if n not in acc:
        acc.add(n)
        for dep in dependances.get(n, ()):
            amont(dep, acc)
    return acc


nb_cols = 0
for t in model_def.TABLES:
    corps = {n: m[n] for n in amont(t["name"])}
    corps[t["name"]] = corps[t["name"]].split("    Sélection = Table.SelectColumns(")[0]
    texte = "\n".join(corps.values())
    for c in t["cols"]:
        nb_cols += 1
        motifs = ['"%s"' % c.name, '#"%s"' % c.name]
        if not any(x in texte for x in motifs) and not re.search(r"[\[,]\s*%s\s*=" % re.escape(c.name), texte):
            erreur(f"colonne {t['name']}[{c.name}] non produite par la requête M")
    if not m[t["name"]].rstrip().endswith("in\n    Typée"):
        erreur(f"table {t['name']} : la requête ne se termine pas par l'étape Typée")
print(f"{nb_cols} colonnes déclarées, toutes produites par leur requête M et typées : "
      f"{'oui' if not any('non produite' in e or 'Typée' in e for e in erreurs) else 'non'}")

# ------------------------------------------------------------------------------------------------ 2. DAX
colonnes = {t["name"]: {c["name"]: c for c in t["columns"]} for t in model["tables"]}
mesures = model_def.dax_expressions(model)
noms_tables = set(colonnes)
MOTS_CLES = {"VAR", "RETURN", "IN", "NOT", "AND", "OR", "TRUE", "FALSE", "ASC", "DESC", "DEFINE", "EVALUATE"}
# fonctions DAX autorisées : (nombre minimal, maximal d'arguments) ; None = sans maximum
ARITE = {"ABS": (1, 1), "ALLSELECTED": (0, None), "AVERAGEX": (2, 2), "BLANK": (0, 0), "CALCULATE": (1, None),
         "COALESCE": (2, None), "CONCATENATEX": (2, None), "COUNTROWS": (0, 1), "DISTINCTCOUNT": (1, 1),
         "DIVIDE": (2, 3), "FALSE": (0, 0), "FILTER": (2, 2), "FORMAT": (2, 3), "HASONEVALUE": (1, 1), "IF": (2, 3), "INT": (1, 1),
         "ISBLANK": (1, 1), "ISEMPTY": (1, 1), "KEEPFILTERS": (1, 1), "LOOKUPVALUE": (3, None), "MAX": (1, 2), "MAXX": (2, 2),
         "MIN": (1, 2), "MINX": (2, 2), "NOT": (1, 1), "REMOVEFILTERS": (0, None), "ROUND": (2, 2),
         "SELECTEDVALUE": (1, 2), "SUM": (1, 1), "SUMX": (2, 2), "SWITCH": (3, None), "TOPN": (2, None),
         "TRUE": (0, 0), "UPPER": (1, 1), "VALUES": (1, 1)}


def nom_table(texte):
    return texte[1:-1].replace("''", "'") if texte.startswith("'") else texte


def parcourir(e, out):
    out.append(e)
    for c in e.children():
        parcourir(c, out)


def sans_textes_ni_commentaires(code):
    code = re.sub(r'"(?:[^"]|"")*"', '""', code)
    code = re.sub(r"/\*.*?\*/", " ", code, flags=re.S)
    return re.sub(r"(--|//)[^\n]*", " ", code)


nb_refs = 0
fonctions_vues = set()
for nom, code in mesures.items():
    try:
        ast = analyse_dax(code)
    except Exception as ex:  # noqa: BLE001
        erreur(f"DAX [{nom}] : analyse impossible ({type(ex).__name__} {ex})")
        continue
    noeuds = []
    parcourir(ast, noeuds)
    variables = {n.var_name.text for n in noeuds if isinstance(n, dax_exprs.VariableExpression)}
    for n in noeuds:
        if isinstance(n, dax_exprs.ColumnExpression):
            t, c = nom_table(n.table.text), n.column.text[1:-1]
            nb_refs += 1
            if t not in colonnes or c not in colonnes[t]:
                erreur(f"DAX [{nom}] : colonne inconnue {t}[{c}]")
        elif isinstance(n, dax_exprs.MeasureExpression):
            x = n.name.text[1:-1]
            nb_refs += 1
            if x not in mesures:
                erreur(f"DAX [{nom}] : mesure inconnue [{x}]")
        elif isinstance(n, dax_exprs.TableExpression):
            t = nom_table(n.name.text)
            if t not in noms_tables and t not in variables:
                erreur(f"DAX [{nom}] : table ou variable inconnue {t}")
        elif isinstance(n, dax_exprs.FunctionExpression):
            f_nom = n.function_name().upper()
            fonctions_vues.add(f_nom)
            if f_nom not in ARITE:
                erreur(f"DAX [{nom}] : fonction inconnue ou non autorisée {f_nom}")
            else:
                mini, maxi = ARITE[f_nom]
                if len(n.args) < mini or (maxi is not None and len(n.args) > maxi):
                    erreur(f"DAX [{nom}] : {f_nom} avec {len(n.args)} argument(s)")
    texte = sans_textes_ni_commentaires(code)
    # arguments vides (tolérés par pbi_parsers, refusés par Power BI)
    if re.search(r",\s*\)|\(\s*,|,\s*,", texte):
        erreur(f"DAX [{nom}] : argument vide")
    if texte.count("(") != texte.count(")"):
        erreur(f"DAX [{nom}] : parenthèses déséquilibrées")
    # contrôle textuel : Table[Colonne] (espaces permis), 'Table'[Colonne], [Mesure]
    for m_ in re.finditer(r"('(?:[^']|'')+'|\b[^\W\d][\w]*)?(\s*)\[([^\]]+)\]", texte):
        prefixe, ref = m_.group(1), m_.group(3)
        if prefixe and prefixe.upper() not in MOTS_CLES and prefixe not in variables:
            t = nom_table(prefixe)
            if t not in colonnes or ref not in colonnes[t]:
                erreur(f"DAX [{nom}] : référence {t}[{ref}] introuvable (contrôle textuel)")
        elif ref not in mesures:
            erreur(f"DAX [{nom}] : mesure [{ref}] introuvable (contrôle textuel)")
    # variables : déclarées une fois, chacune utilisée
    decl = re.findall(r"\bVAR\s+([^\W\d]\w*)", texte)
    for v in set(decl):
        if decl.count(v) > 1 and texte.count("RETURN") < 2:
            erreur(f"DAX [{nom}] : variable {v} déclarée plusieurs fois")
        if len(re.findall(r"(?<![\w\[])%s(?![\w\]])" % re.escape(v), texte)) < 2:
            erreur(f"DAX [{nom}] : variable {v} jamais utilisée")
print(f"{len(mesures)} mesures DAX analysées par pbi_parsers, {nb_refs} références contrôlées, "
      f"fonctions : {', '.join(sorted(fonctions_vues))}")

# ------------------------------------------------------------------------------------------------ 3. modèle
toutes_mesures = [mm["name"] for t in model["tables"] for mm in t.get("measures", [])]
if len(toutes_mesures) != len(set(toutes_mesures)):
    erreur("noms de mesures en double")
toutes_colonnes = {c: t for t, cols in colonnes.items() for c in cols}
for t in model["tables"]:
    if len({c["name"].casefold() for c in t["columns"]}) != len(t["columns"]):
        erreur(f"{t['name']} : noms de colonnes en double (casse ignorée)")
    for mm in t.get("measures", []):
        homonymes = sorted(tb for tb, cols in colonnes.items()
                           if any(c.casefold() == mm["name"].casefold() for c in cols))
        if homonymes:
            erreur(f"mesure [{mm['name']}] homonyme d'une colonne de {homonymes}")
        if mm["name"].casefold() in {x.casefold() for x in noms_tables}:
            erreur(f"mesure [{mm['name']}] homonyme d'une table")
        if not mm.get("description"):
            erreur(f"mesure {mm['name']} sans description")
        if not mm.get("displayFolder"):
            erreur(f"mesure {mm['name']} sans dossier d'affichage")
    for c in t["columns"]:
        if c.get("sortByColumn") and c["sortByColumn"] not in colonnes[t["name"]]:
            erreur(f"{t['name']}[{c['name']}] : colonne de tri inconnue {c['sortByColumn']}")
        technique = re.match(r"(Clé|N°|Rang|Ordre|Position)\b", c["name"]) or c["name"] in ("Retard", "Année")
        if technique and c["dataType"] in ("int64", "double") and c.get("summarizeBy") != "none":
            erreur(f"{t['name']}[{c['name']}] : clé ou numéro additionné (summarizeBy {c.get('summarizeBy')})")
        if c["dataType"] in ("int64", "double") and not c.get("isHidden") and not c.get("formatString"):
            erreur(f"{t['name']}[{c['name']}] : colonne numérique visible sans format")
        if c["dataType"] == "dateTime" and not c.get("formatString"):
            erreur(f"{t['name']}[{c['name']}] : date sans format")
    p = t["partitions"]
    if len(p) != 1 or p[0]["source"]["type"] != "m" or p[0].get("mode") != "import":
        erreur(f"{t['name']} : partition M en importation attendue")
for r in model["relationships"]:
    for tb, cl in ((r["fromTable"], r["fromColumn"]), (r["toTable"], r["toColumn"])):
        if tb not in colonnes or cl not in colonnes[tb]:
            erreur(f"relation {r['name']} : {tb}[{cl}] inconnue")
    if colonnes[r["fromTable"]][r["fromColumn"]]["dataType"] != colonnes[r["toTable"]][r["toColumn"]]["dataType"]:
        erreur(f"relation {r['fromTable']}->{r['toTable']} : types différents")
tags = [t["lineageTag"] for t in model["tables"]] + [c["lineageTag"] for t in model["tables"] for c in t["columns"]] \
    + [mm["lineageTag"] for t in model["tables"] for mm in t.get("measures", [])] \
    + [e["lineageTag"] for e in model["expressions"]]
if len(tags) != len(set(tags)):
    erreur("lineageTag en double")
if len({r["name"] for r in model["relationships"]}) != len(model["relationships"]):
    erreur("noms de relations en double")

aretes = {}
for r in model["relationships"]:
    aretes.setdefault(r["toTable"], []).append(r["fromTable"])          # le côté un filtre le côté plusieurs
    if r.get("crossFilteringBehavior") == "bothDirections":
        aretes.setdefault(r["fromTable"], []).append(r["toTable"])


def chemins(src, dst, chemin):
    if src == dst:
        return [chemin]
    out = []
    for nxt in aretes.get(src, []):
        if nxt not in chemin:
            out += chemins(nxt, dst, chemin + [nxt])
    return out


for a in colonnes:
    for b in colonnes:
        if a != b:
            ch = chemins(a, b, [a])
            if len(ch) > 1:
                erreur(f"plusieurs chemins de filtre {a} -> {b} : {ch}")

# annotations et propriétés usuelles de Power BI Desktop
if bim.get("compatibilityLevel", 0) < 1550:
    erreur("compatibilityLevel < 1550")
for k, v in (("culture", "fr-FR"), ("sourceQueryCulture", "fr-FR"), ("defaultPowerBIDataSourceVersion", "powerBI_V3")):
    if model.get(k) != v:
        erreur(f"model.{k} = {model.get(k)!r}, attendu {v!r}")
ann = {a["name"]: a["value"] for a in model.get("annotations", [])}
for k in ("PBI_QueryOrder", "__PBI_TimeIntelligenceEnabled"):
    if k not in ann:
        erreur(f"annotation de modèle {k} absente")
ordre = json.loads(ann.get("PBI_QueryOrder", "[]"))
if sorted(ordre) != sorted(m):
    erreur("PBI_QueryOrder ne liste pas exactement les requêtes du modèle")
for e in model["expressions"]:
    if not any(a["name"] == "PBI_ResultType" for a in e.get("annotations", [])):
        erreur(f"expression {e['name']} sans PBI_ResultType")
    if e.get("queryGroup") and e["queryGroup"] not in {g["folder"] for g in model.get("queryGroups", [])}:
        erreur(f"expression {e['name']} : groupe de requêtes inconnu {e['queryGroup']}")
url = next(e for e in model["expressions"] if e["name"] == "UrlFichier")
if "IsParameterQuery = true" not in json.dumps(url["expression"], ensure_ascii=False) or \
        "IsParameterQueryRequired = true" not in json.dumps(url["expression"], ensure_ascii=False):
    erreur("UrlFichier n'est pas un paramètre requis")

# textes vus par l'utilisateur : français accentué, pas de mot anglais courant ni de faute d'accent fréquente
visibles = [t["name"] for t in model["tables"]] + [c["name"] for t in model["tables"] for c in t["columns"]
                                                   if not c.get("isHidden")] + toutes_mesures
visibles += [mm.get("description", "") for t in model["tables"] for mm in t.get("measures", [])]
for x in visibles:
    if re.search(r"\b(Total|Load|Hours|Month|Project|Resource|Capacity|Progress|Delay)\b", x) or \
            re.search(r"\b(Periode|Repartition|Perimetre|Debut|Realise|Affecte|Equipe|Controle|Ecart)\b", x):
        erreur(f"texte visible non français ou sans accents : {x!r}")
print(f"{len(model['tables'])} tables, {len(model['relationships'])} relations, {len(model['expressions'])} "
      f"expressions partagées, {len(toutes_mesures)} mesures")
print("model.bim :", os.path.abspath(os.path.join(SORTIE, "model.bim")))
print("CONTRÔLES STATIQUES : " + ("0 erreur" if not erreurs else f"{len(erreurs)} erreur(s)"))
sys.exit(1 if erreurs else 0)

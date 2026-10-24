#!/usr/bin/env python3
"""Contrôles structurels d'un modèle sémantique TMSL (model.bim) destiné à Power BI Desktop.

Usage : <venv>/bin/python validate_bim.py <model.bim> [--donnees echantillon.json] [--sans-m] [--sans-dax]
                                          [--strict] [--quiet]

Contrôles : niveau de compatibilité, culture, version de source de données, tables (partitions M ou DAX),
colonnes (nom, dataType TOM, sourceColumn, formatString, summarizeBy, sortByColumn), mesures (noms uniques dans
tout le modèle, pas de collision avec une colonne de la même table, références DAX existantes, syntaxe DAX par
pbi_parsers si disponible), relations (colonnes existantes, types identiques, cardinalités, côté « un »),
expressions partagées (paramètres M « meta [IsParameterQuery=true] »), annotations usuelles (PBI_QueryOrder,
PBI_ResultType, __PBI_TimeIntelligenceEnabled), syntaxe M par @microsoft/powerquery-parser si disponible,
cohérence heuristique des types M / TMSL.
--donnees : JSON {"Table": [{"Colonne": valeur, ...}, ...]} (par ex. produit par le miroir Python) pour vérifier
la cardinalité des colonnes de tri, l'unicité du côté « un » des relations et l'intégrité référentielle.
Code retour : 0 si aucune erreur, 1 sinon.
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from collections import defaultdict

ICI = os.path.dirname(os.path.abspath(__file__))
# node_modules contenant @microsoft/powerquery-parser : variable PBI_NODE_MODULES, sinon node_modules/ à côté du
# script, sinon ../tools/mparser/node_modules (dossier de travail)
NODE_MODULES_M = os.environ.get("PBI_NODE_MODULES") or next(
    (d for d in (os.path.join(ICI, "node_modules"), os.path.normpath(os.path.join(ICI, "..", "tools", "mparser",
                                                                                  "node_modules")))
     if os.path.isdir(d)), os.path.join(ICI, "node_modules"))
SCRIPT_M = os.path.join(ICI, "outils", "analyse_m.js")

TYPES_TOM = {"string", "int64", "double", "dateTime", "decimal", "boolean", "binary", "variant", "unknown",
             "automatic"}
TYPES_COLONNE_OK = {"string", "int64", "double", "dateTime", "decimal", "boolean", "binary"}
SUMMARIZE = {"default", "none", "sum", "min", "max", "count", "average", "distinctCount"}
TYPES_COLONNE = {"data", "calculated", "calculatedTableColumn", "rowNumber"}
MODES = {"import", "directQuery", "dual", "default", "push", "directLake"}
SOURCES = {"m", "calculated", "query", "entity", "policyRange", "calculationGroup", "parquet", "inferred"}
CARDINALITES = {"none", "one", "many"}
FILTRAGE = {"oneDirection", "bothDirections", "automatic"}
CLES_MODELE = {"name", "description", "culture", "collation", "dataAccessOptions", "defaultPowerBIDataSourceVersion",
               "sourceQueryCulture", "defaultMode", "defaultDataView", "discourageImplicitMeasures", "tables",
               "relationships", "roles", "perspectives", "cultures", "expressions", "dataSources", "annotations",
               "queryGroups", "functions", "forceUniqueNames", "disableAutoExists", "maxParallelismPerQuery",
               "valueFilterBehavior", "storageLocation", "automaticAggregationOptions", "extendedProperties",
               "analyticsAIMetadata", "dataSourceDefaultMaxConnections", "mAttributes", "discourageCompositeModels",
               "selectionExpressionBehavior", "dataSourceVariablesOverrideBehavior", "directLakeBehavior",
               "valueFilterBehavior", "disableSystemDefaultExpression", "bindingInfoCollection"}
# correspondance types M -> dataType TOM (heuristique de cohérence)
TYPES_M = {"text": "string", "number": "double", "Int64.Type": "int64", "Int32.Type": "int64",
           "Int16.Type": "int64", "Int8.Type": "int64", "date": "dateTime", "datetime": "dateTime",
           "datetimezone": "dateTime", "logical": "boolean", "Currency.Type": "decimal",
           "Percentage.Type": "double", "Decimal.Type": "double", "Double.Type": "double",
           "Single.Type": "double", "Date.Type": "dateTime", "DateTime.Type": "dateTime", "Text.Type": "string",
           "Logical.Type": "boolean", "Number.Type": "double"}


class Constats:
    def __init__(self):
        self.erreurs, self.avert, self.infos = [], [], []

    def err(self, ou, msg):
        self.erreurs.append((ou, msg))

    def warn(self, ou, msg):
        self.avert.append((ou, msg))

    def info(self, ou, msg):
        self.infos.append((ou, msg))


def _texte(expr):
    """Expression TMSL : chaîne ou liste de lignes."""
    if isinstance(expr, list):
        if not all(isinstance(x, str) for x in expr):
            return None
        return "\n".join(expr)
    return expr if isinstance(expr, str) else None


def _nom_valide(nom, quoi, c):
    if not isinstance(nom, str) or not nom:
        c.err(quoi, "nom absent ou vide")
        return False
    if nom != nom.strip():
        c.err(quoi, "espaces en début ou fin de nom : %r" % nom)
    if re.search(r"[\x00-\x1f]", nom):
        c.err(quoi, "caractère de contrôle dans le nom %r" % nom)
    return True


def _annotations(obj, quoi, c):
    vus = {}
    for a in obj.get("annotations", []) or []:
        n = a.get("name")
        if n in vus:
            c.err(quoi, "annotation en double : %s" % n)
        v = a.get("value")
        if not isinstance(v, str):
            c.err(quoi, "annotation %s : valeur non textuelle" % n)
        vus[n] = v
    return vus


# ---------------------------------------------------------------------------
# DAX : lexique minimal (chaînes, commentaires, références)
# ---------------------------------------------------------------------------
def _sans_commentaires_ni_chaines(dax, quoi, c):
    """Remplace chaînes et commentaires par des espaces ; vérifie l'équilibre des délimiteurs."""
    out, i, n = [], 0, len(dax)
    pile = []
    while i < n:
        ch = dax[i]
        if dax.startswith("//", i) or dax.startswith("--", i):
            j = dax.find("\n", i)
            j = n if j < 0 else j
            out.append(" " * (j - i))
            i = j
        elif dax.startswith("/*", i):
            j = dax.find("*/", i + 2)
            if j < 0:
                c.err(quoi, "commentaire /* non fermé")
                break
            out.append(" " * (j + 2 - i))
            i = j + 2
        elif ch == '"':
            j = i + 1
            while j < n:
                if dax[j] == '"':
                    if j + 1 < n and dax[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            if j >= n:
                c.err(quoi, "chaîne non fermée")
                break
            out.append('"' + " " * (j - i - 1) + '"')
            i = j + 1
        elif ch == "'":
            j = i + 1
            while j < n:
                if dax[j] == "'":
                    if j + 1 < n and dax[j + 1] == "'":
                        j += 2
                        continue
                    break
                j += 1
            if j >= n:
                c.err(quoi, "nom de table entre apostrophes non fermé")
                break
            out.append(dax[i:j + 1])
            i = j + 1
        elif ch == "[":
            j = i + 1
            while j < n:
                if dax[j] == "]":
                    if j + 1 < n and dax[j + 1] == "]":
                        j += 2
                        continue
                    break
                j += 1
            if j >= n:
                c.err(quoi, "crochet [ non fermé")
                break
            out.append(dax[i:j + 1])
            i = j + 1
        else:
            if ch in "({":
                pile.append(ch)
            elif ch in ")}":
                if not pile or {"(": ")", "{": "}"}[pile.pop()] != ch:
                    c.err(quoi, "parenthèse ou accolade fermante %r sans ouvrante correspondante" % ch)
                    return "".join(out)
            out.append(ch)
            i += 1
    if pile:
        c.err(quoi, "parenthèses ou accolades non fermées : %s" % "".join(pile))
    return "".join(out)


RE_REF_QUALIFIEE = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z_À-ɏ][\wÀ-ɏ.]*))\s*\[((?:[^\]]|\]\])+)\]")
RE_REF_SIMPLE = re.compile(r"\[((?:[^\]]|\]\])+)\]")
# mots-clés DAX qui peuvent précéder une référence [Nom] sans être un nom de table
MOTS_CLES_DAX = {"RETURN", "VAR", "IN", "NOT", "AND", "OR", "ASC", "DESC", "DEFINE", "EVALUATE", "MEASURE",
                 "ORDER", "BY", "START", "AT", "COLUMN", "TABLE", "FUNCTION", "WITH", "VISUAL", "SHAPE"}


def references_dax(dax, quoi, c):
    """(table|None, nom) des références de colonnes/mesures d'une expression DAX."""
    propre = _sans_commentaires_ni_chaines(dax, quoi, c)
    refs = []
    pris = set()
    for m in RE_REF_QUALIFIEE.finditer(propre):
        if m.group(2) and m.group(2).upper() in MOTS_CLES_DAX:
            continue  # « RETURN [Mesure] », « IN [Liste] »... : référence non qualifiée
        table = (m.group(1) or "").replace("''", "'") or m.group(2)
        refs.append((table, m.group(3).replace("]]", "]")))
        pris.add(m.start(3) - 1)
    for m in RE_REF_SIMPLE.finditer(propre):
        if m.start() in pris:
            continue
        refs.append((None, m.group(1).replace("]]", "]")))
    variables = set(re.findall(r"\bVAR\s+([A-Za-z_][\w]*)", propre, flags=re.I))
    return refs, variables


def analyse_dax_externe(expressions):
    """Analyse syntaxique par pbi_parsers : {clé: None si correct, sinon message}. None si indisponible."""
    try:
        from pbi_parsers.dax import to_ast
    except Exception:
        return None
    res = {}
    for cle, texte in expressions.items():
        try:
            r = to_ast(texte)
            res[cle] = None if r is not None else "expression non reconnue"
        except AssertionError:
            res[cle] = "jetons non consommés en fin d'expression"
        except Exception as e:
            res[cle] = "%s : %s" % (type(e).__name__, str(e)[:160])
    return res


def analyse_m_externe(expressions):
    """Analyse syntaxique M par @microsoft/powerquery-parser (node). None si indisponible."""
    node = shutil.which("node")
    if not node or not os.path.isdir(os.path.join(NODE_MODULES_M, "@microsoft", "powerquery-parser")):
        return None
    entree = json.dumps([{"nom": k, "texte": v} for k, v in expressions.items()])
    try:
        p = subprocess.run([node, SCRIPT_M, NODE_MODULES_M], input=entree, capture_output=True, text=True,
                           timeout=300, cwd=ICI)
    except (OSError, subprocess.TimeoutExpired):
        return None
    if p.returncode != 0:
        return None
    return {r["nom"]: (None if r["ok"] else r.get("erreur", "erreur")) for r in json.loads(p.stdout)}


# ---------------------------------------------------------------------------
# M : heuristiques (types déclarés, noms de colonnes)
# ---------------------------------------------------------------------------
RE_TYPE_TABLE = re.compile(r"type\s+table\s*\[(.*?)\]", re.S)
RE_TRANSFORM = re.compile(r"Table\.TransformColumnTypes\s*\(", re.S)


def _types_m(m):
    """{colonne: type M} relevés dans « type table [...] » et Table.TransformColumnTypes (dernier gagne)."""
    types = {}
    for bloc in RE_TYPE_TABLE.finditer(m):
        for nom, t in re.findall(r'(#"(?:[^"]|"")*"|[A-Za-z_][\w.]*)\s*=\s*(?:nullable\s+)?([A-Za-z][\w.]*)', bloc.group(1)):
            if nom.startswith('#"'):
                nom = nom[2:-1].replace('""', '"')
            types[nom] = t
    for mt in RE_TRANSFORM.finditer(m):
        seg = m[mt.end():mt.end() + 20000]
        for nom, t in re.findall(r'\{\s*"((?:[^"]|"")*)"\s*,\s*(?:type\s+)?(?:nullable\s+)?([A-Za-z][\w.]*)\s*\}', seg):
            types[nom.replace('""', '"')] = t
    return types


def _apparait_dans_m(nom, m):
    e = re.escape(nom.replace('"', '""'))
    return bool(re.search(r'"' + e + r'"', m) or re.search(r"\[" + re.escape(nom) + r"\]", m)
                or re.search(r"(?<![\w.])" + re.escape(nom) + r"\s*=", m))


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def valider(bim, donnees=None, analyse_m=True, analyse_dax=True):
    c = Constats()
    if not isinstance(bim, dict):
        c.err("racine", "objet JSON attendu")
        return c
    for k in bim:
        if k not in ("compatibilityLevel", "model", "name", "id", "$schema"):
            c.warn("racine", "clé inattendue : %s" % k)
    cl = bim.get("compatibilityLevel")
    if not isinstance(cl, int):
        c.err("racine", "compatibilityLevel entier requis")
    elif not 1500 <= cl <= 1800:
        c.err("racine", "compatibilityLevel %s hors de la plage Power BI (1500-1800)" % cl)
    elif cl < 1550:
        c.warn("racine", "compatibilityLevel %s ancien (Desktop 2025 écrit 1550 ou plus)" % cl)
    else:
        c.info("racine", "compatibilityLevel %s" % cl)
    m = bim.get("model")
    if not isinstance(m, dict):
        c.err("racine", "objet model requis")
        return c
    for k in m:
        if k not in CLES_MODELE:
            c.warn("model", "propriété inconnue : %s" % k)
    cult = m.get("culture")
    if not isinstance(cult, str) or not re.match(r"^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$", cult or ""):
        c.err("model", "culture absente ou invalide : %r" % cult)
    else:
        c.info("model", "culture %s" % cult)
    if m.get("defaultPowerBIDataSourceVersion") != "powerBI_V3":
        c.err("model", "defaultPowerBIDataSourceVersion doit valoir powerBI_V3 (métadonnées améliorées), trouvé %r"
              % m.get("defaultPowerBIDataSourceVersion"))
    if "sourceQueryCulture" not in m:
        c.warn("model", "sourceQueryCulture absente (culture des requêtes M ; Desktop l'écrit)")
    elif m["sourceQueryCulture"] != cult:
        c.info("model", "sourceQueryCulture %s différente de culture %s" % (m["sourceQueryCulture"], cult))
    dao = m.get("dataAccessOptions")
    if not dao:
        c.warn("model", "dataAccessOptions absent (Desktop écrit legacyRedirects et returnErrorValuesAsNull)")
    if m.get("dataSources"):
        c.warn("model", "dataSources hérités présents : un modèle Power Query n'en a normalement pas")
    ann_modele = _annotations(m, "model", c)

    tables = m.get("tables") or []
    if not tables:
        c.err("model", "aucune table")
    noms_tables = {}
    index = {}
    noms_mesures = {}
    tags = defaultdict(list)
    requetes_m = []
    textes_dax = {}
    textes_m = {}
    for t in tables:
        tn = t.get("name")
        quoi = "table %s" % tn
        if not _nom_valide(tn, quoi, c):
            continue
        if tn.lower() in noms_tables:
            c.err(quoi, "nom de table en double (casse ignorée) : %s" % noms_tables[tn.lower()])
        noms_tables[tn.lower()] = tn
        if t.get("lineageTag"):
            tags[t["lineageTag"]].append(quoi)
        ann_t = _annotations(t, quoi, c)
        cols = {}
        for col in t.get("columns") or []:
            cn = col.get("name")
            cq = "%s[%s]" % (tn, cn)
            if not _nom_valide(cn, cq, c):
                continue
            if cn.lower() in {k.lower() for k in cols}:
                c.err(cq, "nom de colonne en double dans la table (casse ignorée)")
            cols[cn] = col
            _annotations(col, cq, c)
            if col.get("lineageTag"):
                tags[col["lineageTag"]].append(cq)
            typ = col.get("type", "data")
            if typ not in TYPES_COLONNE:
                c.err(cq, "type de colonne inconnu : %s" % typ)
            if typ == "rowNumber":
                c.warn(cq, "colonne rowNumber explicite (gérée par le moteur)")
            dt = col.get("dataType")
            if dt not in TYPES_TOM:
                c.err(cq, "dataType invalide : %r (valeurs TOM : %s)" % (dt, ", ".join(sorted(TYPES_TOM))))
            elif typ == "data" and dt not in TYPES_COLONNE_OK:
                c.err(cq, "dataType %s non permis pour une colonne de données" % dt)
            if typ == "data":
                if not col.get("sourceColumn"):
                    c.err(cq, "sourceColumn requis pour une colonne de données")
                if "expression" in col:
                    c.err(cq, "expression sur une colonne de données")
            elif typ == "calculated":
                if not _texte(col.get("expression")):
                    c.err(cq, "colonne calculée sans expression DAX")
                else:
                    textes_dax[cq] = _texte(col["expression"])
            if "formatString" in col and (not isinstance(col["formatString"], str) or not col["formatString"]):
                c.err(cq, "formatString vide ou non textuel")
            sb = col.get("summarizeBy")
            if sb is not None and sb not in SUMMARIZE:
                c.err(cq, "summarizeBy invalide : %r" % sb)
            if dt in ("string", "dateTime", "boolean") and sb not in (None, "none", "count", "distinctCount", "min", "max", "default"):
                c.err(cq, "summarizeBy %s incompatible avec le type %s" % (sb, dt))
            if dt in ("string",) and sb in (None, "default"):
                pass
            if dt in ("int64", "double", "decimal") and sb is None:
                c.info(cq, "summarizeBy non précisé sur une colonne numérique (somme par défaut)")
            if dt == "dateTime":
                ud = {a.get("name"): a.get("value") for a in col.get("annotations", []) or []}
                if "UnderlyingDateTimeDataType" not in ud and "formatString" not in col:
                    c.info(cq, "colonne date sans formatString ni annotation UnderlyingDateTimeDataType")
        index[tn] = {"colonnes": cols, "mesures": {}, "table": t}
        # partitions
        parts = t.get("partitions") or []
        if not parts:
            c.err(quoi, "aucune partition")
        noms_p = set()
        types_src = set()
        for p in parts:
            pn = p.get("name")
            pq = "%s / partition %s" % (quoi, pn)
            _nom_valide(pn, pq, c)
            if pn in noms_p:
                c.err(pq, "nom de partition en double")
            noms_p.add(pn)
            if p.get("mode", "import") not in MODES:
                c.err(pq, "mode invalide : %r" % p.get("mode"))
            src = p.get("source") or {}
            st = src.get("type")
            types_src.add(st)
            if st not in SOURCES:
                c.err(pq, "source.type invalide : %r" % st)
            if st in ("m", "calculated"):
                tx = _texte(src.get("expression"))
                if not tx or not tx.strip():
                    c.err(pq, "source.expression vide")
                elif st == "m":
                    textes_m["table " + tn] = tx
                else:
                    textes_dax[pq] = tx
        if len(types_src) > 1:
            c.err(quoi, "partitions de types différents : %s" % sorted(map(str, types_src)))
        if "m" in types_src:
            requetes_m.append(tn)
            if ann_t.get("PBI_ResultType") != "Table":
                c.warn(quoi, "annotation PBI_ResultType = Table attendue sur une table issue de Power Query")
            if "PBI_NavigationStepName" not in ann_t:
                c.info(quoi, "annotation PBI_NavigationStepName absente (Desktop écrit « Navigation »)")
            # heuristiques M <-> colonnes
            mtx = textes_m.get("table " + tn, "")
            tm = _types_m(mtx)
            for cn, col in cols.items():
                if col.get("type", "data") != "data":
                    continue
                sc = col.get("sourceColumn", "")
                if sc and not _apparait_dans_m(sc, mtx):
                    c.warn("%s[%s]" % (tn, cn), "sourceColumn %r introuvable dans l'expression M (vérifier le nom)" % sc)
                if sc in tm and tm[sc] in TYPES_M and col.get("dataType") in TYPES_TOM:
                    attendu = TYPES_M[tm[sc]]
                    if attendu != col["dataType"] and not (attendu == "double" and col["dataType"] in ("int64", "decimal")):
                        c.warn("%s[%s]" % (tn, cn), "type M %s -> %s attendu, dataType %s déclaré"
                               % (tm[sc], attendu, col["dataType"]))
        # hiérarchies
        for h in t.get("hierarchies") or []:
            for lv in h.get("levels") or []:
                if lv.get("column") not in cols:
                    c.err("%s / hiérarchie %s" % (quoi, h.get("name")), "colonne de niveau inconnue : %s" % lv.get("column"))
        # clés
        cles = [cn for cn, col in cols.items() if col.get("isKey")]
        if len(cles) > 1:
            c.err(quoi, "plusieurs colonnes isKey : %s" % ", ".join(cles))
        # tri par colonne
        for cn, col in cols.items():
            sbc = col.get("sortByColumn")
            if not sbc:
                continue
            cq = "%s[%s]" % (tn, cn)
            if sbc == cn:
                c.err(cq, "sortByColumn sur elle-même")
            elif sbc not in cols:
                c.err(cq, "sortByColumn %r absente de la table" % sbc)
            else:
                if cols[sbc].get("sortByColumn"):
                    c.err(cq, "sortByColumn %r est elle-même triée par une autre colonne" % sbc)
                if cols[sbc].get("dataType") == "string":
                    c.warn(cq, "colonne de tri %r de type texte (tri alphabétique)" % sbc)
                if donnees and tn in donnees:
                    _verifier_tri(donnees[tn], cn, sbc, cq, c)
    # mesures (après toutes les tables pour les collisions)
    for t in tables:
        tn = t.get("name")
        if tn not in index:
            continue
        for me in t.get("measures") or []:
            mn = me.get("name")
            mq = "mesure %s[%s]" % (tn, mn)
            if not _nom_valide(mn, mq, c):
                continue
            if mn.lower() in noms_mesures:
                c.err(mq, "nom de mesure en double dans le modèle (avec %s)" % noms_mesures[mn.lower()])
            noms_mesures[mn.lower()] = "%s[%s]" % (tn, mn)
            if mn.lower() in {k.lower() for k in index[tn]["colonnes"]}:
                c.err(mq, "même nom qu'une colonne de la table")
            for t2, ix in index.items():
                if t2 != tn and mn.lower() in {k.lower() for k in ix["colonnes"]}:
                    c.info(mq, "même nom qu'une colonne de la table %s" % t2)
            index[tn]["mesures"][mn] = me
            _annotations(me, mq, c)
            if me.get("lineageTag"):
                tags[me["lineageTag"]].append(mq)
            tx = _texte(me.get("expression"))
            if not tx or not tx.strip():
                c.err(mq, "expression vide")
            else:
                textes_dax[mq] = tx
            if "formatString" in me and (not isinstance(me["formatString"], str) or not me["formatString"]):
                c.err(mq, "formatString vide ou non textuel")
    for tag, ou in tags.items():
        if len(ou) > 1:
            c.err("lineageTag", "lineageTag %s en double : %s" % (tag, ", ".join(ou)))
    # références DAX
    toutes_mesures = {k for ix in index.values() for k in ix["mesures"]}
    tables_par_nom = {k: v for k, v in index.items()}
    for quoi, tx in textes_dax.items():
        refs, variables = references_dax(tx, quoi, c)
        table_locale = quoi.split(" ", 1)[1].split("[")[0] if quoi.startswith("mesure ") else quoi.split("[")[0]
        for table, nom in refs:
            if table is not None:
                if table in variables:
                    continue
                if table not in tables_par_nom:
                    c.err(quoi, "table %r inconnue dans %s[%s]" % (table, table, nom))
                    continue
                ix = tables_par_nom[table]
                if nom not in ix["colonnes"] and nom not in ix["mesures"]:
                    c.err(quoi, "%s[%s] : ni colonne ni mesure de la table %s" % (table, nom, table))
            else:
                ix = tables_par_nom.get(table_locale, {"colonnes": {}})
                if nom not in toutes_mesures and nom not in ix["colonnes"]:
                    if any(nom in v["colonnes"] for v in tables_par_nom.values()):
                        c.info(quoi, "[%s] non qualifiée : colonne d'une autre table (contexte de ligne ?)" % nom)
                    else:
                        c.warn(quoi, "[%s] : mesure inconnue (ou colonne créée par ADDCOLUMNS/SELECTCOLUMNS)" % nom)
    if analyse_dax and textes_dax:
        res = analyse_dax_externe(textes_dax)
        if res is None:
            c.info("DAX", "pbi_parsers indisponible : syntaxe DAX non analysée")
        else:
            nb = 0
            for quoi, e in res.items():
                if e:
                    c.warn(quoi, "pbi_parsers n'a pas pu analyser l'expression DAX (%s)" % e)
                else:
                    nb += 1
            c.info("DAX", "%d/%d expressions DAX analysées par pbi_parsers" % (nb, len(res)))
    # relations
    rels = m.get("relationships") or []
    noms_r = set()
    paires_actives = defaultdict(list)
    for r in rels:
        rn = r.get("name")
        rq = "relation %s" % rn
        if not rn:
            c.err(rq, "nom absent")
        elif rn in noms_r:
            c.err(rq, "nom de relation en double")
        noms_r.add(rn)
        ft, fc, tt, tc = r.get("fromTable"), r.get("fromColumn"), r.get("toTable"), r.get("toColumn")
        ok = True
        for tab, colname, cote in ((ft, fc, "from"), (tt, tc, "to")):
            if tab not in index:
                c.err(rq, "%sTable %r inconnue" % (cote, tab))
                ok = False
            elif colname not in index[tab]["colonnes"]:
                c.err(rq, "%sColumn %s[%s] inconnue" % (cote, tab, colname))
                ok = False
        if ft == tt:
            c.err(rq, "relation d'une table sur elle-même")
        fcard, tcard = r.get("fromCardinality", "many"), r.get("toCardinality", "one")
        for v in (fcard, tcard):
            if v not in CARDINALITES:
                c.err(rq, "cardinalité invalide : %r" % v)
        if r.get("crossFilteringBehavior", "oneDirection") not in FILTRAGE:
            c.err(rq, "crossFilteringBehavior invalide : %r" % r.get("crossFilteringBehavior"))
        if r.get("crossFilteringBehavior") == "bothDirections":
            c.warn(rq, "filtrage croisé bidirectionnel (ambiguïtés possibles)")
        if r.get("joinOnDateBehavior", "dateAndTime") not in ("dateAndTime", "datePartOnly"):
            c.err(rq, "joinOnDateBehavior invalide")
        if r.get("securityFilteringBehavior", "oneDirection") not in ("oneDirection", "bothDirections", "none"):
            c.err(rq, "securityFilteringBehavior invalide")
        if fcard == "many" and tcard == "many":
            c.warn(rq, "relation plusieurs-à-plusieurs")
        if fcard == "one" and tcard == "many":
            c.warn(rq, "côté « un » en from : TMSL attend from = plusieurs, to = un")
        if not ok:
            continue
        dtf = index[ft]["colonnes"][fc].get("dataType")
        dtt = index[tt]["colonnes"][tc].get("dataType")
        if dtf != dtt:
            c.err(rq, "types différents : %s[%s] %s / %s[%s] %s" % (ft, fc, dtf, tt, tc, dtt))
        if dtt in ("double", "decimal"):
            c.warn(rq, "colonne clé %s[%s] de type %s (clé entière ou texte recommandée)" % (tt, tc, dtt))
        if r.get("isActive", True):
            paires_actives[frozenset((ft, tt))].append(rn)
        for tab, colname, card in ((tt, tc, tcard), (ft, fc, fcard)):
            col = index[tab]["colonnes"][colname]
            if card == "one" and col.get("type", "data") == "calculated":
                c.warn(rq, "côté « un » sur la colonne calculée %s[%s] (dépendances circulaires possibles)" % (tab, colname))
            if card == "one" and not col.get("isKey") and not (donnees and tab in donnees):
                c.info(rq, "côté « un » %s[%s] : unicité non vérifiable sans --donnees" % (tab, colname))
        if donnees:
            _verifier_relation(donnees, ft, fc, tt, tc, tcard, fcard, rq, c)
    for paire, ns in paires_actives.items():
        if len(ns) > 1:
            c.err("relations", "plusieurs relations actives entre %s : %s" % (" et ".join(sorted(paire)), ", ".join(ns)))
    _cycles(paires_actives, c)
    # expressions partagées
    noms_expr = set()
    for e in m.get("expressions") or []:
        en = e.get("name")
        eq = "expression %s" % en
        if not _nom_valide(en, eq, c):
            continue
        if en.lower() in noms_expr:
            c.err(eq, "nom d'expression en double")
        noms_expr.add(en.lower())
        if en.lower() in noms_tables:
            c.err(eq, "même nom qu'une table (espace de noms Power Query commun)")
        if e.get("kind", "m") != "m":
            c.err(eq, "kind %r (m attendu)" % e.get("kind"))
        tx = _texte(e.get("expression"))
        if not tx or not tx.strip():
            c.err(eq, "expression vide")
            continue
        textes_m[eq] = tx
        requetes_m.append(en)
        ann = _annotations(e, eq, c)
        if re.search(r"\bmeta\s*\[", tx) and re.search(r"IsParameterQuery\s*=\s*true", tx):
            typ = re.search(r'\bType\s*=\s*"([^"]+)"', tx)
            if not typ:
                c.err(eq, "paramètre sans Type=\"...\" dans meta")
            if not re.search(r"IsParameterQueryRequired\s*=\s*(true|false)", tx):
                c.warn(eq, "paramètre sans IsParameterQueryRequired")
            if typ:
                attendu = {"Text": "Text", "Number": "Number", "Date": "Date", "DateTime": "DateTime",
                           "Logical": "Logical", "Any": "Any"}.get(typ.group(1))
                if ann.get("PBI_ResultType") not in (None, attendu):
                    c.err(eq, "PBI_ResultType %r incohérent avec Type=%s" % (ann.get("PBI_ResultType"), typ.group(1)))
                if "PBI_ResultType" not in ann:
                    c.warn(eq, "annotation PBI_ResultType absente (Desktop écrit %s)" % attendu)
                valeur = tx.split("meta")[0].strip()
                if typ.group(1) == "Text" and not re.match(r'^"(?:[^"]|"")*"$', valeur):
                    c.err(eq, "valeur du paramètre texte non littérale : %s" % valeur[:80])
            c.info(eq, "paramètre Power Query")
    # PBI_QueryOrder
    qo = ann_modele.get("PBI_QueryOrder")
    if qo is None:
        if requetes_m:
            c.warn("model", "annotation PBI_QueryOrder absente (ordre des requêtes dans l'éditeur Power Query)")
    else:
        try:
            ordre = json.loads(qo)
            if not isinstance(ordre, list) or not all(isinstance(x, str) for x in ordre):
                raise ValueError
        except ValueError:
            c.err("model", "PBI_QueryOrder n'est pas une liste JSON de noms")
            ordre = None
        if ordre is not None:
            if len(set(ordre)) != len(ordre):
                c.err("model", "PBI_QueryOrder contient des doublons")
            manquantes = [q for q in requetes_m if q not in ordre]
            en_trop = [q for q in ordre if q not in requetes_m]
            if manquantes:
                c.warn("model", "PBI_QueryOrder ne liste pas : %s" % ", ".join(manquantes))
            if en_trop:
                c.warn("model", "PBI_QueryOrder liste des requêtes inexistantes : %s" % ", ".join(en_trop))
    if ann_modele.get("__PBI_TimeIntelligenceEnabled") != "0":
        c.warn("model", "__PBI_TimeIntelligenceEnabled != 0 : Desktop créera des tables de dates automatiques")
    # syntaxe M
    if analyse_m and textes_m:
        res = analyse_m_externe(textes_m)
        if res is None:
            c.info("M", "@microsoft/powerquery-parser indisponible : syntaxe M non analysée")
        else:
            for quoi, e in res.items():
                if e:
                    c.err(quoi, "syntaxe M : %s" % e)
            c.info("M", "%d expression(s) M analysée(s) par @microsoft/powerquery-parser, %d en erreur"
                   % (len(res), sum(1 for e in res.values() if e)))
    # cultures
    for cu in m.get("cultures") or []:
        if not cu.get("name"):
            c.err("cultures", "culture sans nom")
    return c


def _verifier_tri(lignes, col, tri, quoi, c):
    vu = {}
    for li in lignes:
        if col not in li or tri not in li:
            c.warn(quoi, "échantillon sans colonnes %s / %s" % (col, tri))
            return
        a, b = li[col], li[tri]
        if a in vu and vu[a] != b:
            c.err(quoi, "cardinalité incohérente : %r trié par %r et %r (une seule valeur de tri par valeur)" % (a, vu[a], b))
            return
        vu[a] = b
    c.info(quoi, "tri par %s vérifié sur %d valeur(s)" % (tri, len(vu)))


def _verifier_relation(donnees, ft, fc, tt, tc, tcard, fcard, quoi, c):
    if fcard == "one" and ft in donnees:
        vals = [li.get(fc) for li in donnees[ft]]
        dbl = {v for v in vals if vals.count(v) > 1}
        if dbl:
            c.err(quoi, "côté « un » %s[%s] (relation un-à-un) non unique : %s" % (ft, fc, sorted(map(str, dbl))[:5]))
    if tt in donnees:
        vals = [li.get(tc) for li in donnees[tt]]
        if tcard == "one":
            dbl = {v for v in vals if vals.count(v) > 1}
            if dbl:
                c.err(quoi, "côté « un » %s[%s] non unique : %s" % (tt, tc, sorted(map(str, dbl))[:5]))
            if any(v is None or v == "" for v in vals):
                c.err(quoi, "côté « un » %s[%s] contient des valeurs vides" % (tt, tc))
        if ft in donnees:
            ens = set(vals)
            orph = {li.get(fc) for li in donnees[ft] if li.get(fc) not in ens and li.get(fc) not in (None, "")}
            if orph:
                c.warn(quoi, "valeurs de %s[%s] sans correspondance dans %s[%s] : %s"
                       % (ft, fc, tt, tc, sorted(map(str, orph))[:5]))


def _cycles(paires, c):
    """Cycle dans le graphe des relations actives (chemins de filtrage ambigus possibles)."""
    parent = {}

    def racine(x):
        while parent.setdefault(x, x) != x:
            x = parent[x]
        return x
    for paire, ns in paires.items():
        if len(paire) != 2:
            continue
        a, b = sorted(paire)
        ra, rb = racine(a), racine(b)
        if ra == rb:
            c.warn("relations", "cycle de relations actives passant par %s - %s (%s)" % (a, b, ", ".join(ns)))
        else:
            parent[ra] = rb


def valider_fichier(chemin, donnees=None, analyse_m=True, analyse_dax=True):
    with open(chemin, encoding="utf-8") as f:
        texte = f.read()
    c = Constats()
    if texte.startswith("﻿"):
        c.warn("fichier", "BOM en tête de fichier")
        texte = texte[1:]
    try:
        bim = json.loads(texte)
    except json.JSONDecodeError as e:
        c.err("fichier", "JSON invalide : %s" % e)
        return c
    r = valider(bim, donnees, analyse_m, analyse_dax)
    r.erreurs[:0], r.avert[:0] = c.erreurs, c.avert
    return r


def main(argv=None):
    ap = argparse.ArgumentParser(description="Contrôles structurels d'un model.bim (TMSL).")
    ap.add_argument("bim")
    ap.add_argument("--donnees", help="échantillon JSON {table: [lignes]} pour les contrôles de cardinalité")
    ap.add_argument("--sans-m", action="store_true", help="ne pas analyser la syntaxe M")
    ap.add_argument("--sans-dax", action="store_true", help="ne pas analyser la syntaxe DAX (pbi_parsers)")
    ap.add_argument("--strict", action="store_true")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)
    donnees = None
    if a.donnees:
        with open(a.donnees, encoding="utf-8") as f:
            donnees = json.load(f)
    c = valider_fichier(a.bim, donnees, not a.sans_m, not a.sans_dax)
    for ou, msg in c.erreurs:
        print("ERREUR  %s : %s" % (ou, msg))
    if not a.quiet:
        for ou, msg in c.avert:
            print("AVERT.  %s : %s" % (ou, msg))
        for ou, msg in c.infos:
            print("INFO    %s : %s" % (ou, msg))
    print("Bilan model.bim : %d erreur(s), %d avertissement(s), %d info(s)" % (len(c.erreurs), len(c.avert), len(c.infos)))
    return 1 if c.erreurs or (a.strict and c.avert) else 0


if __name__ == "__main__":
    sys.exit(main())

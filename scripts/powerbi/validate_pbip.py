#!/usr/bin/env python3
"""Validation d'un projet Power BI PBIP dont le rapport est au format PBIR.

Usage : <venv>/bin/python validate_pbip.py <dossier_pbip | dossier .Report> [--bim] [--strict] [--quiet]

Contrôles :
  1. chaque fichier JSON (rapport, .pbip, .platform, definition.pbir/.pbism) contre le schéma officiel désigné par
     son $schema (copies locales dans schemas/, résolution des $ref hors ligne) ; thèmes contre le schéma
     de thème de Power BI Desktop ;
  2. cohérence structurelle : noms uniques, dossiers = noms, pages listées / existantes, visuels dans la page,
     chevauchements non voulus (avertissement), ressources de thème présentes et déclarées, liens de navigation,
     extraction (drillthrough), groupes de synchronisation, interactions ;
  3. syntaxe des littéraux PBIR et des objets de mise en forme (noms, types et valeurs permises d'après le schéma
     de thème officiel), rôles de requête par type de visuel ;
  4. références croisées avec <modèle>.SemanticModel/model.bim : chaque Entity/Property existe avec le bon genre
     (colonne / mesure), types des littéraux de filtres, segments de dates sur des colonnes de type date.
Code retour : 0 si aucune erreur, 1 sinon (--strict : les avertissements comptent comme erreurs).
"""
import argparse
import json
import os
import re
import sys
from collections import defaultdict

ICI = os.path.dirname(os.path.abspath(__file__))
if ICI not in sys.path:
    sys.path.insert(0, ICI)

# copies locales des schémas : dossier schemas/ à côté du script (dépôt), sinon ../schemas (dossier de travail)
DOSSIER_SCHEMAS = next((d for d in (os.path.join(ICI, "schemas"), os.path.normpath(os.path.join(ICI, "..", "schemas")))
                        if os.path.isdir(d)), os.path.join(ICI, "schemas"))
BASE_URL = "https://developer.microsoft.com/json-schemas/"
SCHEMA_THEME = os.path.join(DOSSIER_SCHEMAS, "powerbi-desktop-samples", "Report Theme JSON Schema",
                            "reportThemeSchema-2.144.json")

# Versions maximales compatibles avec Power BI Desktop de juin 2025 (2.144) : au-delà, avertissement.
VERSIONS_MAX = {
    "report": "2.1.0", "page": "2.0.0", "pagesMetadata": "1.0.0", "versionMetadata": "1.0.0",
    "visualContainer": "2.1.0", "visualContainerMobileState": "2.1.0", "bookmark": "1.4.0",
    "bookmarksMetadata": "1.0.0", "reportExtension": "1.0.0", "definitionProperties": "1.0.0",
    "platformProperties": "2.0.0", "pbipProperties": "1.0.0",
}

# Rôles de requête admis par type de visuel (relevés sur des rapports réels et la documentation)
ROLES = {
    "card": {"Values"}, "multiRowCard": {"Values"}, "slicer": {"Values"}, "tableEx": {"Values"},
    "pivotTable": {"Rows", "Columns", "Values"},
    "lineStackedColumnComboChart": {"Category", "Series", "Y", "Y2", "Tooltips", "Rows"},
    "lineClusteredColumnComboChart": {"Category", "Series", "Y", "Y2", "Tooltips", "Rows"},
    "lineChart": {"Category", "Series", "Y", "Y2", "Tooltips", "Rows"},
    "clusteredBarChart": {"Category", "Series", "Y", "Tooltips", "Rows"},
    "barChart": {"Category", "Series", "Y", "Tooltips", "Rows"},
    "columnChart": {"Category", "Series", "Y", "Tooltips", "Rows"},
    "clusteredColumnChart": {"Category", "Series", "Y", "Tooltips", "Rows"},
    "hundredPercentStackedBarChart": {"Category", "Series", "Y", "Tooltips", "Rows"},
    "hundredPercentStackedColumnChart": {"Category", "Series", "Y", "Tooltips", "Rows"},
    "donutChart": {"Category", "Series", "Y", "Tooltips"}, "pieChart": {"Category", "Series", "Y", "Tooltips"},
    "textbox": set(), "shape": set(), "actionButton": set(), "image": set(), "basicShape": set(),
}
ROLES_MESURES = {"Y", "Y2", "Tooltips"}           # mesures ou agrégations
ROLES_COLONNES = {"Category", "Series", "Rows", "Columns"}
TYPES_DECOR = {"shape", "textbox", "image", "basicShape"}
TYPES_SUPERPOSABLES = {"card", "textbox", "shape", "actionButton", "image", "multiRowCard"}
PROPRIETES_NON_THEMATIQUES = {("slicer", "general", "filter"), ("advancedSlicerVisual", "general", "filter"),
                              ("listSlicer", "general", "filter")}

RE_TEXTE = re.compile(r"^'(?:[^']|'')*'$")
RE_NOMBRE = re.compile(r"^-?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?[DLM]$")
RE_DATETIME = re.compile(r"^datetime'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{1,7})?'$")
RE_COULEUR = re.compile(r"^'#(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8}|[0-9A-Fa-f]{3})'$")
RE_GUID = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
RE_NOM = re.compile(r"^[A-Za-z0-9_-]{1,50}$")


class Constats:
    def __init__(self, racine):
        self.racine = racine
        self.erreurs, self.avert, self.infos = [], [], []

    def _rel(self, f):
        try:
            return os.path.relpath(f, self.racine) if f else "-"
        except ValueError:
            return f

    def err(self, f, msg):
        self.erreurs.append((self._rel(f), msg))

    def warn(self, f, msg):
        self.avert.append((self._rel(f), msg))

    def info(self, f, msg):
        self.infos.append((self._rel(f), msg))


def lire_json(chemin, c):
    try:
        with open(chemin, encoding="utf-8") as f:
            texte = f.read()
    except OSError as e:
        c.err(chemin, "lecture impossible : %s" % e)
        return None
    if texte.startswith("﻿"):
        c.warn(chemin, "marque d'ordre d'octets (BOM) en tête de fichier")
        texte = texte[1:]
    try:
        return json.loads(texte)
    except json.JSONDecodeError as e:
        c.err(chemin, "JSON invalide : %s" % e)
        return None


# ---------------------------------------------------------------------------
# Schémas officiels (hors ligne)
# ---------------------------------------------------------------------------
class Schemas:
    def __init__(self):
        from referencing import Registry, Resource
        from referencing.jsonschema import DRAFT7
        ressources = []
        self.docs = {}
        for dp, _d, fs in os.walk(os.path.join(DOSSIER_SCHEMAS, "fabric")):
            for f in fs:
                if not f.endswith(".json"):
                    continue
                p = os.path.join(dp, f)
                with open(p, encoding="utf-8") as fh:
                    doc = json.load(fh)
                url = BASE_URL + os.path.relpath(p, DOSSIER_SCHEMAS).replace(os.sep, "/")
                self.docs[url] = doc
                res = Resource(contents=doc, specification=DRAFT7)
                ressources.append((url, res))
                if isinstance(doc, dict) and doc.get("$id") and doc["$id"] != url:
                    ressources.append((doc["$id"], res))
        self.registre = Registry().with_resources(ressources)
        self._validateurs = {}

    def validateur(self, url):
        if url not in self._validateurs:
            from jsonschema import Draft7Validator
            self._validateurs[url] = Draft7Validator(self.docs[url], registry=self.registre)
        return self._validateurs[url]

    def valider(self, doc, url, fichier, c, schema_impose=False):
        if url not in self.docs:
            c.err(fichier, "schéma non disponible localement : %s" % url)
            return False
        if schema_impose and "$schema" not in doc:
            doc = dict(doc)
            doc["$schema"] = url
            c.info(fichier, "pas de $schema : validé contre %s" % url.replace(BASE_URL, ""))
        ok = True
        for e in sorted(self.validateur(url).iter_errors(doc), key=lambda e: list(e.absolute_path)):
            chemin = "/".join(str(x) for x in e.absolute_path) or "(racine)"
            c.err(fichier, "schéma %s : %s : %s" % (url.split("/definition/")[-1].split("/item/")[-1],
                                                   chemin, _court(e.message)))
            ok = False
        return ok


def _court(s, n=300):
    return s if len(s) <= n else s[:n] + "..."


def _version(url):
    m = re.search(r"/([A-Za-z]+)/(\d+\.\d+\.\d+)/schema(?:-embedded)?\.json$", url or "")
    return (m.group(1), m.group(2)) if m else (None, None)


def _vtuple(v):
    return tuple(int(x) for x in v.split("."))


# ---------------------------------------------------------------------------
# Catalogue des objets de mise en forme (schéma de thème officiel)
# ---------------------------------------------------------------------------
class Catalogue:
    def __init__(self, chemin):
        with open(chemin, encoding="utf-8") as f:
            self.schema = json.load(f)
        self.defs = self.schema["definitions"]
        self.types = set(self.schema["properties"]["visualStyles"]["properties"].keys())
        self._cache = {}
        self.communs = self._objets({"$ref": "#/definitions/commonCards"})

    def resoudre(self, s):
        while isinstance(s, dict) and "$ref" in s:
            s = self.defs[s["$ref"].split("/")[-1]]
        return s

    def _objets(self, defn):
        out = {}
        s = self.resoudre(defn)
        for p in s.get("allOf", [s]):
            p = self.resoudre(p)
            for on, ov in p.get("properties", {}).items():
                if on == "*":
                    continue
                it = self.resoudre(ov.get("items", {}))
                out.setdefault(on, {}).update(it.get("properties", {}))
        return out

    def objets(self, type_visuel):
        if type_visuel not in self._cache:
            nom = "visual-" + type_visuel
            self._cache[type_visuel] = self._objets(self.defs[nom]) if nom in self.defs else None
        return self._cache[type_visuel]


def _enum(spec):
    vals = []
    for k in ("oneOf", "anyOf"):
        for o in spec.get(k, []):
            if "const" in o:
                vals.append(o["const"])
    if "enum" in spec:
        vals.extend(spec["enum"])
    return vals


def verifier_valeur_propriete(cat, spec, valeur, ou, fichier, c):
    """Contrôle de forme d'une valeur de propriété d'après sa définition dans le schéma de thème."""
    ref = spec.get("$ref", "").split("/")[-1]
    spec = cat.resoudre(spec)
    est_fill = ref == "fill" or any(isinstance(o, dict) and "solid" in o.get("required", [])
                                    for o in spec.get("oneOf", []))
    if est_fill:
        if not (isinstance(valeur, dict) and "solid" in valeur and isinstance(valeur["solid"], dict)
                and "color" in valeur["solid"] and isinstance(valeur["solid"]["color"], dict)
                and "expr" in valeur["solid"]["color"]):
            c.err(fichier, "%s : couleur attendue sous la forme {\"solid\": {\"color\": {\"expr\": ...}}}" % ou)
            return
        e = valeur["solid"]["color"]["expr"]
        if "Literal" in e:
            _verifier_couleur(e["Literal"]["Value"], ou, fichier, c)
        return
    if ref in ("paragraphs",):
        if not isinstance(valeur, list):
            c.err(fichier, "%s : liste de paragraphes attendue" % ou)
        return
    if ref in ("image", "icon", "dataBars", "fillRule", "itemLocation", "annotationTemplate", "themeIcon"):
        return
    if not isinstance(valeur, dict) or "expr" not in valeur:
        if isinstance(valeur, dict) and ("solid" in valeur or "filter" in valeur):
            c.err(fichier, "%s : valeur structurée inattendue pour une propriété simple" % ou)
        return
    e = valeur["expr"]
    if "Literal" not in e:
        return  # expression (mesure, règle, couleur de thème) : mise en forme conditionnelle
    v = e["Literal"]["Value"]
    t = spec.get("type")
    if isinstance(t, list):
        t = t[0]
    vals = _enum(spec)
    if t == "boolean":
        if v not in ("true", "false"):
            c.err(fichier, "%s : booléen attendu (true/false), trouvé %s" % (ou, v))
    elif t in ("number", "integer") or ref in ("fontSize", "labelDisplayUnits", "displayUnitsWithoutAuto"):
        if re.match(r"^'-?\d+(\.\d+)?'$", v):
            c.warn(fichier, "%s : nombre écrit comme texte %s (toléré ; forme attendue 12D)" % (ou, v))
        elif not RE_NOMBRE.match(v):
            c.err(fichier, "%s : nombre attendu (ex. 12D, 3L), trouvé %s" % (ou, v))
        elif vals and float(v[:-1]) not in [float(x) for x in vals]:
            c.warn(fichier, "%s : valeur %s hors des valeurs du schéma de thème %s" % (ou, v, vals))
        else:
            x = float(v[:-1])
            if ("minimum" in spec and x < spec["minimum"]) or ("maximum" in spec and x > spec["maximum"]):
                c.warn(fichier, "%s : valeur %s hors de la plage du schéma de thème [%s ; %s] (Desktop la borne)"
                       % (ou, v, spec.get("minimum", "-"), spec.get("maximum", "-")))
    elif t == "string":
        if RE_DATETIME.match(v) and re.search(r"(Date|date)$", ou):
            return  # dates des segments (startDate, endDate, anchorDate) : chaîne dans le thème, datetime'...' dans le visuel
        if not RE_TEXTE.match(v):
            c.err(fichier, "%s : texte entre apostrophes attendu, trouvé %s" % (ou, v))
        elif vals:
            s = v[1:-1].replace("''", "'")
            if s not in vals:
                c.warn(fichier, "%s : valeur '%s' hors des valeurs du schéma de thème (%s)" % (ou, s, ", ".join(map(str, vals))))


# ---------------------------------------------------------------------------
# Modèle sémantique (model.bim)
# ---------------------------------------------------------------------------
class Modele:
    def __init__(self, bim):
        self.tables = {}
        for t in bim.get("model", {}).get("tables", []):
            cols = {cc["name"]: cc.get("dataType", "unknown" if cc.get("type") != "calculated" else "variant")
                    for cc in t.get("columns", [])}
            mesures = {m["name"]: m for m in t.get("measures", [])}
            hier = {h["name"]: [lv.get("name") for lv in h.get("levels", [])] for h in t.get("hierarchies", [])}
            self.tables[t["name"]] = {"colonnes": cols, "mesures": mesures, "hierarchies": hier}

    def table(self, nom):
        if nom in self.tables:
            return nom, None
        for t in self.tables:
            if t.lower() == nom.lower():
                return t, "casse différente (%s)" % t
        return None, None


def _lit_compatible(type_col, valeur):
    if valeur == "null":
        return True
    if type_col in ("string",):
        return bool(RE_TEXTE.match(valeur))
    if type_col in ("int64", "double", "decimal"):
        return bool(RE_NOMBRE.match(valeur))
    if type_col == "dateTime":
        return bool(RE_DATETIME.match(valeur))
    if type_col == "boolean":
        return valeur in ("true", "false")
    return True


# ---------------------------------------------------------------------------
# Parcours des références de champs
# ---------------------------------------------------------------------------
def references(noeud, alias=None, chemin=""):
    """Rend (genre, table|None, alias|None, propriété, chemin, noeud) pour chaque référence de champ."""
    alias = alias or {}
    if isinstance(noeud, dict):
        if isinstance(noeud.get("From"), list) and ("Where" in noeud or "Select" in noeud):
            alias = dict(alias)
            for s in noeud["From"]:
                if isinstance(s, dict) and "Name" in s:
                    alias[s["Name"]] = s.get("Entity")
        for genre in ("Column", "Measure"):
            r = noeud.get(genre)
            if isinstance(r, dict) and "Property" in r and "Expression" in r:
                sr = r["Expression"].get("SourceRef", {}) if isinstance(r["Expression"], dict) else {}
                if "Entity" in sr:
                    yield (genre, sr["Entity"], None, r["Property"], chemin + "/" + genre, sr.get("Schema"))
                elif "Source" in sr:
                    yield (genre, alias.get(sr["Source"]), sr["Source"], r["Property"], chemin + "/" + genre,
                           None)
        for k, v in noeud.items():
            yield from references(v, alias, chemin + "/" + k)
    elif isinstance(noeud, list):
        for i, v in enumerate(noeud):
            yield from references(v, alias, chemin + "[%d]" % i)


def litteraux(noeud, chemin=""):
    if isinstance(noeud, dict):
        lit = noeud.get("Literal")
        if isinstance(lit, dict) and "Value" in lit:
            yield chemin + "/Literal", lit["Value"]
        for k, v in noeud.items():
            yield from litteraux(v, chemin + "/" + k)
    elif isinstance(noeud, list):
        for i, v in enumerate(noeud):
            yield from litteraux(v, chemin + "[%d]" % i)


def comparaisons(noeud, alias=None):
    """(champ_expr, valeurs_littérales, alias) des conditions In et Comparison d'une définition de filtre."""
    alias = alias or {}
    if isinstance(noeud, dict):
        if isinstance(noeud.get("From"), list) and "Where" in noeud:
            alias = dict(alias)
            for s in noeud["From"]:
                alias[s.get("Name")] = s.get("Entity")
        if "In" in noeud and isinstance(noeud["In"], dict):
            exps = noeud["In"].get("Expressions", [])
            for ligne in noeud["In"].get("Values", []):
                for e, val in zip(exps, ligne):
                    if "Literal" in val:
                        yield e, val["Literal"]["Value"], alias
        if "Comparison" in noeud and isinstance(noeud["Comparison"], dict):
            cp = noeud["Comparison"]
            if "Literal" in cp.get("Right", {}):
                yield cp.get("Left"), cp["Right"]["Literal"]["Value"], alias
        for v in noeud.values():
            yield from comparaisons(v, alias)
    elif isinstance(noeud, list):
        for v in noeud:
            yield from comparaisons(v, alias)


def _champ_cle(expr_champ, alias=None):
    """Clé texte d'un champ (genre, table, propriété) pour comparer des expressions."""
    alias = alias or {}
    if not isinstance(expr_champ, dict):
        return None
    if "Aggregation" in expr_champ:
        return ("Aggregation", _champ_cle(expr_champ["Aggregation"].get("Expression"), alias),
                expr_champ["Aggregation"].get("Function"))
    for g in ("Column", "Measure"):
        if g in expr_champ:
            sr = expr_champ[g]["Expression"].get("SourceRef", {})
            t = sr.get("Entity") or alias.get(sr.get("Source"))
            return (g, t, expr_champ[g]["Property"])
    return None


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
class Validateur:
    def __init__(self, racine, strict=False):
        self.racine = os.path.abspath(racine)
        self.c = Constats(self.racine)
        self.strict = strict
        self.schemas = Schemas()
        self.cat = Catalogue(SCHEMA_THEME)
        self.modele = None
        self.fichier_modele = None
        self.pages = {}        # nom -> (page_json, chemin)
        self.visuels = []      # (page_nom, visual_json, chemin)
        self.sync = defaultdict(list)

    # -- point d'entrée
    def executer(self):
        r = self.racine
        reports = []
        if r.endswith(".Report") and os.path.isdir(r):
            reports = [r]
        else:
            pbips = sorted(f for f in os.listdir(r) if f.endswith(".pbip"))
            if not pbips:
                self.c.err(r, "aucun fichier .pbip ni dossier .Report")
            for p in pbips:
                fp = os.path.join(r, p)
                d = lire_json(fp, self.c)
                if d is None:
                    continue
                self.schemas.valider(d, d.get("$schema") or BASE_URL + "fabric/pbip/pbipProperties/1.0.0/schema.json",
                                     fp, self.c, schema_impose=True)
                for a in d.get("artifacts", []):
                    chemin = a.get("report", {}).get("path")
                    cible = os.path.normpath(os.path.join(r, chemin or ""))
                    if not chemin or not os.path.isdir(cible):
                        self.c.err(fp, "artefact de rapport introuvable : %s" % chemin)
                    else:
                        reports.append(cible)
            gi = os.path.join(r, ".gitignore")
            if not os.path.isfile(gi):
                self.c.info(r, "pas de .gitignore (Desktop en écrit un avec .pbi/localSettings.json et cache.abf)")
        for rep in reports:
            self.valider_rapport(rep)
        return self.c

    # -- fichiers d'artefact
    def valider_platform(self, dossier, type_attendu):
        fp = os.path.join(dossier, ".platform")
        if not os.path.isfile(fp):
            self.c.warn(dossier, "fichier .platform absent")
            return
        d = lire_json(fp, self.c)
        if d is None:
            return
        self._version_ok(d.get("$schema"), fp)
        self.schemas.valider(d, d.get("$schema") or BASE_URL + "fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                             fp, self.c, schema_impose=True)
        md = d.get("metadata", {})
        if md.get("type") != type_attendu:
            self.c.err(fp, "metadata.type = %r, attendu %r" % (md.get("type"), type_attendu))
        nom = os.path.basename(dossier).rsplit(".", 1)[0]
        if md.get("displayName") != nom:
            self.c.warn(fp, "displayName %r différent du nom du dossier %r" % (md.get("displayName"), nom))
        if not RE_GUID.match(d.get("config", {}).get("logicalId", "")):
            self.c.err(fp, "logicalId n'est pas un GUID")

    def _version_ok(self, url, fichier):
        fam, ver = _version(url)
        if fam in VERSIONS_MAX and ver and _vtuple(ver) > _vtuple(VERSIONS_MAX[fam]):
            self.c.warn(fichier, "schéma %s %s plus récent que la version de référence %s (Desktop juin 2025)"
                        % (fam, ver, VERSIONS_MAX[fam]))

    def valider_modele(self, dossier_rapport, pbir):
        ref = pbir.get("datasetReference", {})
        if bool(ref.get("byPath")) == bool(ref.get("byConnection")):
            self.c.err(dossier_rapport, "definition.pbir : exactement une référence byPath ou byConnection attendue")
            return
        if not ref.get("byPath"):
            self.c.info(dossier_rapport, "rapport connecté (byConnection) : références au modèle non vérifiées")
            return
        dm = os.path.normpath(os.path.join(dossier_rapport, ref["byPath"]["path"]))
        if not os.path.isdir(dm):
            self.c.err(os.path.join(dossier_rapport, "definition.pbir"), "modèle sémantique introuvable : %s" % dm)
            return
        self.valider_platform(dm, "SemanticModel")
        fp = os.path.join(dm, "definition.pbism")
        if not os.path.isfile(fp):
            self.c.err(dm, "definition.pbism absent")
        else:
            d = lire_json(fp, self.c)
            if d is not None:
                self.schemas.valider(d, d.get("$schema") or BASE_URL + "fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
                                     fp, self.c, schema_impose=True)
        bim = os.path.join(dm, "model.bim")
        if os.path.isfile(bim):
            d = lire_json(bim, self.c)
            if d is not None:
                self.modele = Modele(d)
                self.fichier_modele = bim
        elif os.path.isdir(os.path.join(dm, "definition")):
            self.c.warn(dm, "modèle au format TMDL : références croisées non vérifiées (model.bim attendu)")
        else:
            self.c.err(dm, "ni model.bim ni dossier definition/ dans le modèle sémantique")

    # -- rapport
    def valider_rapport(self, rep):
        c = self.c
        self.valider_platform(rep, "Report")
        fpbir = os.path.join(rep, "definition.pbir")
        pbir = lire_json(fpbir, c) if os.path.isfile(fpbir) else None
        if pbir is None:
            c.err(rep, "definition.pbir absent ou illisible")
            return
        url = pbir.get("$schema") or BASE_URL + "fabric/item/report/definitionProperties/1.0.0/schema.json"
        self._version_ok(pbir.get("$schema"), fpbir)
        self.schemas.valider(pbir, url, fpbir, c, schema_impose=True)
        ddef = os.path.join(rep, "definition")
        if not os.path.isdir(ddef):
            c.err(rep, "dossier definition/ absent : le rapport n'est pas au format PBIR")
            return
        try:
            if float(pbir.get("version", "0")) < 4.0:
                c.err(fpbir, "version %r : le format PBIR (dossier definition/) exige la version 4.0 ou plus"
                      % pbir.get("version"))
        except ValueError:
            c.err(fpbir, "version illisible : %r" % pbir.get("version"))
        self.valider_modele(rep, pbir)

        # validation par schéma de tous les fichiers JSON de definition/
        docs = {}
        for dp, _d, fs in os.walk(ddef):
            for f in sorted(fs):
                p = os.path.join(dp, f)
                if not f.endswith(".json"):
                    c.warn(p, "fichier inattendu dans definition/")
                    continue
                d = lire_json(p, c)
                if d is None:
                    continue
                docs[p] = d
                s = d.get("$schema") if isinstance(d, dict) else None
                if not s:
                    c.err(p, "$schema absent")
                    continue
                self._version_ok(s, p)
                self.schemas.valider(d, s, p, c)
        attendus = {"version.json": "versionMetadata", "report.json": "report", os.path.join("pages", "pages.json"): "pagesMetadata"}
        for rel, fam in attendus.items():
            p = os.path.join(ddef, rel)
            if p not in docs:
                c.err(p, "fichier obligatoire absent")
            elif _version(docs[p].get("$schema"))[0] != fam:
                c.err(p, "$schema inattendu pour ce fichier : %s" % docs[p].get("$schema"))
        report = docs.get(os.path.join(ddef, "report.json"))
        pages_meta = docs.get(os.path.join(ddef, "pages", "pages.json"))
        if report:
            self.valider_ressources(rep, report)
            self.valider_noeud_champs(report, os.path.join(ddef, "report.json"))
            self.valider_filtres(report.get("filterConfig"), os.path.join(ddef, "report.json"))
        self.valider_pages(ddef, docs, pages_meta)

    def valider_ressources(self, rep, report):
        c = self.c
        fr = os.path.join(rep, "definition", "report.json")
        declares = {}
        for paquet in report.get("resourcePackages", []):
            for it in paquet.get("items", []):
                p = os.path.join(rep, "StaticResources", paquet["name"], it["path"])
                declares[os.path.normpath(p)] = (paquet["name"], it)
                if not os.path.isfile(p):
                    c.err(fr, "ressource déclarée introuvable : StaticResources/%s/%s" % (paquet["name"], it["path"]))
        tc = report.get("themeCollection", {})
        for cle, type_item in (("baseTheme", "BaseTheme"), ("customTheme", "CustomTheme")):
            th = tc.get(cle)
            if not th:
                continue
            meme_nom = [v for v in declares.values() if v[0] == th.get("type") and v[1]["name"] == th["name"]]
            trouve = [v for v in meme_nom if v[1]["type"] == type_item]
            if meme_nom and not trouve:
                c.warn(fr, "themeCollection.%s %r déclaré avec le type %r (attendu %s)"
                       % (cle, th.get("name"), meme_nom[0][1]["type"], type_item))
            elif not trouve:
                c.err(fr, "themeCollection.%s %r non déclaré dans resourcePackages (%s, type %s)"
                      % (cle, th.get("name"), th.get("type"), type_item))
        # fichiers présents mais non déclarés + validation des thèmes
        ds = os.path.join(rep, "StaticResources")
        with open(SCHEMA_THEME, encoding="utf-8") as f:
            sch_theme = json.load(f)
        from jsonschema import Draft7Validator
        vt = Draft7Validator(sch_theme)
        for dp, _d, fs in os.walk(ds):
            for f in fs:
                p = os.path.normpath(os.path.join(dp, f))
                if p not in declares:
                    c.warn(p, "ressource non déclarée dans report.json")
                    continue
                if declares[p][1]["type"] == "CustomTheme":
                    d = lire_json(p, c)
                    if d is None:
                        continue
                    for e in vt.iter_errors(d):
                        chemin = "/".join(str(x) for x in e.absolute_path) or "(racine)"
                        c.err(p, "thème : %s : %s" % (chemin, _court(e.message)))
                    # Desktop nomme le fichier d'après le thème, espaces remplacées par « _ » (« Accessible Tidal »
                    # -> Accessible_Tidal9104699003566352.json dans le corpus)
                    prefixes = (d.get("name") or "", (d.get("name") or "").replace(" ", "_"))
                    if d.get("name") and declares[p][1]["name"] not in (d["name"], d["name"] + ".json") \
                            and not declares[p][1]["name"].startswith(prefixes):
                        c.warn(p, "nom du thème %r différent de l'élément déclaré %r" % (d["name"], declares[p][1]["name"]))
                elif declares[p][1]["type"] == "BaseTheme":
                    lire_json(p, c)

    # -- pages et visuels
    def valider_pages(self, ddef, docs, meta):
        c = self.c
        dpages = os.path.join(ddef, "pages")
        presents = {}
        for d in sorted(os.listdir(dpages)) if os.path.isdir(dpages) else []:
            p = os.path.join(dpages, d, "page.json")
            if os.path.isdir(os.path.join(dpages, d)):
                if p not in docs:
                    c.err(os.path.join(dpages, d), "dossier de page sans page.json valide")
                    continue
                pj = docs[p]
                if pj.get("name") != d:
                    c.warn(p, "name %r différent du dossier %r (toléré par Desktop)" % (pj.get("name"), d))
                if not RE_NOM.match(pj.get("name", "")):
                    c.warn(p, "nom de page hors [A-Za-z0-9_-]{1,50} : %r" % pj.get("name"))
                presents[pj.get("name")] = (pj, p)
        self.pages = presents
        if not presents:
            c.err(dpages, "le rapport doit contenir au moins une page")
        if meta:
            ordre = meta.get("pageOrder", [])
            fm = os.path.join(dpages, "pages.json")
            if len(set(ordre)) != len(ordre):
                c.err(fm, "pageOrder contient des doublons")
            for n in ordre:
                if n not in presents:
                    c.err(fm, "page listée mais absente : %s" % n)
            for n in presents:
                if n not in ordre:
                    c.warn(fm, "page présente mais absente de pageOrder : %s" % n)
            if meta.get("activePageName") and meta["activePageName"] not in presents:
                c.err(fm, "activePageName inconnue : %s" % meta["activePageName"])
            if meta.get("activePageName") and presents.get(meta["activePageName"], ({},))[0].get("visibility") == "HiddenInViewMode":
                c.warn(fm, "la page active est masquée")
        noms_aff = defaultdict(list)
        for n, (pj, p) in presents.items():
            noms_aff[pj.get("displayName", "").strip().lower()].append(n)
            if not pj.get("displayName", "").strip():
                c.err(p, "displayName vide")
        for k, v in noms_aff.items():
            if len(v) > 1:
                c.warn(dpages, "nom affiché de page en double : %s (%s)" % (k, ", ".join(v)))
        tous_noms_visuels = defaultdict(list)
        for n, (pj, p) in presents.items():
            self.valider_page(n, pj, p, docs, tous_noms_visuels)
        for nv, ou in tous_noms_visuels.items():
            if len(ou) > 1:
                c.info(dpages, "nom de visuel présent sur plusieurs pages : %s (%s)" % (nv, ", ".join(ou)))
        # noms des liaisons de page et de leurs paramètres : uniques dans tout le rapport (schéma page)
        liaisons = defaultdict(list)
        for n, (pj, p) in presents.items():
            pb = pj.get("pageBinding") or {}
            if pb.get("name"):
                liaisons[("liaison", pb["name"])].append(n)
            for prm in pb.get("parameters", []):
                liaisons[("paramètre", prm.get("name"))].append(n)
        for (quoi, nom), ou in liaisons.items():
            if len(ou) > 1:
                c.warn(dpages, "nom de %s de page %r utilisé plusieurs fois (%s) : unique dans le rapport attendu"
                       % (quoi, nom, ", ".join(ou)))
        self.valider_sync()

    def valider_page(self, nom, pj, fichier, docs, tous_noms_visuels):
        c = self.c
        W, H = pj.get("width", 0), pj.get("height", 0)
        if W <= 0 or H <= 0:
            c.err(fichier, "dimensions de page invalides")
        self.valider_noeud_champs(pj, fichier)
        noms_filtres = self.valider_filtres(pj.get("filterConfig"), fichier)
        # extraction (drillthrough)
        pb = pj.get("pageBinding")
        filtres = {f.get("name"): f for f in (pj.get("filterConfig") or {}).get("filters", [])}
        if pb and pb.get("type") in ("Drillthrough", "Tooltip"):
            if pj.get("type") != pb.get("type"):
                c.warn(fichier, "pageBinding de type %s mais page.type = %r" % (pb.get("type"), pj.get("type")))
            if not pb.get("parameters") and pb.get("type") == "Drillthrough":
                c.err(fichier, "page d'extraction sans paramètre")
            for prm in pb.get("parameters", []):
                bf = prm.get("boundFilter")
                if bf not in filtres:
                    c.err(fichier, "paramètre d'extraction %r : filtre lié %r absent du filterConfig" % (prm.get("name"), bf))
                    continue
                f = filtres[bf]
                if f.get("howCreated") != "Drillthrough":
                    c.warn(fichier, "filtre %r lié à l'extraction : howCreated %r (Drillthrough attendu)" % (bf, f.get("howCreated")))
                if _champ_cle(f.get("field")) != _champ_cle(prm.get("fieldExpr")):
                    c.err(fichier, "paramètre %r : fieldExpr différent du champ du filtre lié" % prm.get("name"))
        elif pj.get("type") == "Drillthrough":
            c.warn(fichier, "page de type Drillthrough sans pageBinding (aucun champ d'extraction)")
        for f in filtres.values():
            if f.get("howCreated") == "Drillthrough" and not (pb and pb.get("type") in ("Drillthrough", "Tooltip")):
                c.err(fichier, "filtre d'extraction %r sans pageBinding" % f.get("name"))
        del noms_filtres
        # visuels
        dv = os.path.join(os.path.dirname(fichier), "visuals")
        visuels = {}
        for d in sorted(os.listdir(dv)) if os.path.isdir(dv) else []:
            p = os.path.join(dv, d, "visual.json")
            if not os.path.isdir(os.path.join(dv, d)):
                c.warn(os.path.join(dv, d), "fichier inattendu dans visuals/")
                continue
            if p not in docs:
                c.err(os.path.join(dv, d), "dossier de visuel sans visual.json valide")
                continue
            vj = docs[p]
            if vj.get("name") != d:
                c.warn(p, "name %r différent du dossier %r (toléré par Desktop)" % (vj.get("name"), d))
            if not RE_NOM.match(vj.get("name", "")):
                c.warn(p, "nom de visuel hors [A-Za-z0-9_-]{1,50} : %r" % vj.get("name"))
            visuels[vj.get("name")] = (vj, p)
            tous_noms_visuels[vj.get("name")].append(nom)
        groupes = {n for n, (vj, _p) in visuels.items() if "visualGroup" in vj}
        for n, (vj, p) in visuels.items():
            pg = vj.get("parentGroupName")
            if pg and pg not in groupes:
                c.err(p, "parentGroupName %r inconnu" % pg)
            self.valider_visuel(nom, pj, vj, p)
        self.valider_disposition(pj, fichier, visuels)
        for it in pj.get("visualInteractions", []):
            for cle in ("source", "target"):
                if it.get(cle) not in visuels:
                    c.err(fichier, "visualInteractions : %s %r inconnu sur la page" % (cle, it.get(cle)))

    def valider_disposition(self, pj, fichier, visuels):
        c = self.c
        W, H = pj.get("width", 0), pj.get("height", 0)
        rects = []
        zs, tabs = defaultdict(list), defaultdict(list)
        for n, (vj, p) in visuels.items():
            pos = vj.get("position", {})
            x, y, w, h = pos.get("x", 0), pos.get("y", 0), pos.get("width", 0), pos.get("height", 0)
            if w <= 0 or h <= 0:
                c.err(p, "taille nulle ou négative")
            if x < -0.5 or y < -0.5 or x + w > W + 0.5 or y + h > H + 0.5:
                c.err(p, "visuel hors de la page (%g,%g %gx%g ; page %gx%g)" % (x, y, w, h, W, H))
            if vj.get("parentGroupName") or vj.get("isHidden") or "visualGroup" in vj:
                continue
            zs[pos.get("z")].append(n)
            if "tabOrder" in pos:
                tabs[pos["tabOrder"]].append(n)
            rects.append((n, x, y, w, h, pos.get("z", 0), vj.get("visual", {}).get("visualType")))
        for z, ns in zs.items():
            if len(ns) > 1:
                c.warn(fichier, "même z (%s) pour plusieurs visuels : %s" % (z, ", ".join(ns)))
        for t, ns in tabs.items():
            if len(ns) > 1:
                c.warn(fichier, "même tabOrder (%s) pour plusieurs visuels : %s" % (t, ", ".join(ns)))
        for i in range(len(rects)):
            for j in range(i + 1, len(rects)):
                a, b = rects[i], rects[j]
                ix = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
                iy = min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2])
                if ix <= 1 or iy <= 1:
                    continue
                dessous, dessus = (a, b) if a[5] <= b[5] else (b, a)
                contient = (dessous[1] - 0.5 <= dessus[1] and dessous[2] - 0.5 <= dessus[2]
                            and dessus[1] + dessus[3] <= dessous[1] + dessous[3] + 0.5
                            and dessus[2] + dessus[4] <= dessous[2] + dessous[4] + 0.5)
                if contient and (dessous[6] in TYPES_DECOR or dessus[6] in TYPES_SUPERPOSABLES):
                    continue  # décor en arrière-plan ou superposition volontaire (ex. total au centre d'un anneau)
                c.warn(fichier, "chevauchement de %s (%s) et %s (%s) sur %gx%g px"
                       % (a[0], a[6], b[0], b[6], ix, iy))

    def valider_visuel(self, page_nom, pj, vj, fichier):
        c = self.c
        v = vj.get("visual")
        if not v:
            return
        t = v.get("visualType")
        q = v.get("query") or {}
        qs = q.get("queryState") or {}
        refs = []
        roles_ok = ROLES.get(t)
        for role, st in qs.items():
            if roles_ok is not None and role not in roles_ok:
                c.err(fichier, "rôle %r non admis pour %s (admis : %s)" % (role, t, ", ".join(sorted(roles_ok)) or "aucun"))
            for pr in st.get("projections", []):
                refs.append(pr.get("queryRef"))
                genre = next(iter(pr.get("field", {})), None)
                if role in ROLES_MESURES and genre == "Column":
                    c.err(fichier, "rôle %s : mesure ou agrégation attendue, colonne %s trouvée" % (role, pr.get("queryRef")))
                if role in ROLES_COLONNES and genre in ("Measure", "Aggregation") and t not in ("tableEx",):
                    c.err(fichier, "rôle %s : colonne attendue, %s trouvée (%s)" % (role, genre, pr.get("queryRef")))
                if t == "slicer" and genre != "Column":
                    c.err(fichier, "segment : colonne attendue, %s trouvée" % genre)
                if t == "card" and genre == "Column":
                    c.warn(fichier, "carte sur une colonne (affiche la première valeur) : %s" % pr.get("queryRef"))
                if t == "pivotTable" and role == "Values" and genre == "Column":
                    c.err(fichier, "matrice : valeurs = mesures ou agrégations (%s)" % pr.get("queryRef"))
        if roles_ok and t not in ("textbox", "shape", "actionButton", "image") and not qs:
            c.err(fichier, "visuel de données sans requête")
        if roles_ok == set() and qs:
            c.err(fichier, "%s ne doit pas avoir de requête" % t)
        for role, st in qs.items():
            rr = [pr.get("queryRef") for pr in st.get("projections", [])]
            doublons = {r for r in rr if rr.count(r) > 1}
            if doublons:
                c.err(fichier, "queryRef en double dans le rôle %s : %s" % (role, ", ".join(sorted(doublons))))
        if t == "slicer":
            vals = qs.get("Values", {}).get("projections", [])
            if len(vals) != 1:
                c.warn(fichier, "segment avec %d champs (1 attendu)" % len(vals))
        cles_proj = {_champ_cle(pr.get("field")) for st in qs.values() for pr in st.get("projections", [])}
        for s in (q.get("sortDefinition") or {}).get("sort", []):
            if _champ_cle(s.get("field")) not in cles_proj:
                c.warn(fichier, "tri sur un champ absent des projections : %s" % (_champ_cle(s.get("field")),))
        # objets de mise en forme
        self.valider_objets(t, v.get("objects") or {}, fichier, set(refs), conteneur=False)
        self.valider_objets(t, v.get("visualContainerObjects") or {}, fichier, set(refs), conteneur=True)
        # liens et info-bulles
        for e in (v.get("visualContainerObjects") or {}).get("visualLink", []):
            pr = e.get("properties", {})
            typ = _lit_texte(pr.get("type"))
            for cle in ("navigationSection", "drillthroughSection"):
                cible = _lit_texte(pr.get(cle))
                if cible == "":
                    c.warn(fichier, "visualLink.%s vide (bouton sans destination)" % cle)
                elif cible is not None and cible not in self.pages:
                    c.warn(fichier, "visualLink.%s : page %r inconnue (lien sans effet)" % (cle, cible))
                if cle == "drillthroughSection" and cible in self.pages and self.pages[cible][0].get("type") != "Drillthrough":
                    c.warn(fichier, "visualLink.drillthroughSection : la page %r n'est pas une page d'extraction" % cible)
            if typ == "PageNavigation" and _lit_texte(pr.get("navigationSection")) is None:
                c.warn(fichier, "bouton de navigation sans navigationSection")
        for nom_obj in ("visualTooltip", "visualHeaderTooltip"):
            for e in (v.get("visualContainerObjects") or {}).get(nom_obj, []):
                sec = _lit_texte(e.get("properties", {}).get("section"))
                if sec and sec != "___AUTO___" and sec not in self.pages:
                    c.warn(fichier, "%s.section : page %r inconnue (info-bulle par défaut)" % (nom_obj, sec))
        # segment : sélection enregistrée, dates, synchronisation
        if t == "slicer":
            self.valider_segment(v, fichier)
        self.valider_noeud_champs(vj, fichier)
        self.valider_filtres(vj.get("filterConfig"), fichier)
        if v.get("syncGroup"):
            vals = qs.get("Values", {}).get("projections", [])
            cle = _champ_cle(vals[0]["field"]) if vals else None
            self.sync[v["syncGroup"].get("groupName")].append((page_nom, cle, fichier, _selection_enregistree(v)))

    def valider_segment(self, v, fichier):
        c = self.c
        objs = v.get("objects") or {}
        mode = None
        data = {}
        for e in objs.get("data", []):
            data.update(e.get("properties", {}))
        mode = _lit_texte(data.get("mode"))
        unique = any(_lit_brut(e.get("properties", {}).get("strictSingleSelect")) == "true"
                     for e in objs.get("selection", []))
        vals = (v.get("query") or {}).get("queryState", {}).get("Values", {}).get("projections", [])
        champ = _champ_cle(vals[0]["field"]) if vals else None
        type_col = self._type_colonne(champ)
        dates = "startDate" in data or "endDate" in data or mode in ("Between", "Before", "After", "Relative")
        if dates and (mode in ("Relative",) or "startDate" in data or "endDate" in data) and type_col not in (None, "dateTime"):
            c.err(fichier, "segment de dates sur une colonne de type %s (dateTime attendu)" % type_col)
        if mode in ("Between", "Before", "After") and type_col not in (None, "dateTime", "int64", "double", "decimal"):
            c.err(fichier, "segment de plage (%s) sur une colonne de type %s" % (mode, type_col))
        for cle in ("startDate", "endDate"):
            if cle in data:
                val = _lit_brut(data[cle])
                if not (val and RE_DATETIME.match(val)):
                    c.err(fichier, "data.%s : littéral datetime'...' attendu" % cle)
        for e in objs.get("general", []):
            f = e.get("properties", {}).get("filter")
            if f is None:
                continue
            fd = f.get("filter") if isinstance(f, dict) else None
            if not isinstance(fd, dict) or "Where" not in fd:
                c.err(fichier, "general.filter : structure {\"filter\": {\"Version\", \"From\", \"Where\"}} attendue")
                continue
            alias = {s.get("Name"): s.get("Entity") for s in fd.get("From", [])}
            for ex, val, al in comparaisons(fd):
                if _champ_cle(ex, al) != champ:
                    c.err(fichier, "sélection enregistrée sur un autre champ que celui du segment : %s" % (_champ_cle(ex, al),))
            if unique:
                for w in fd.get("Where", []):
                    cond = w.get("Condition", {})
                    if "In" in cond and len(cond["In"].get("Values", [])) != 1:
                        c.err(fichier, "segment à sélection unique : la sélection enregistrée doit contenir une seule valeur")
            del alias

    def _type_colonne(self, cle):
        if not self.modele or not cle or cle[0] != "Column":
            return None
        t = self.modele.tables.get(cle[1])
        return t["colonnes"].get(cle[2]) if t else None

    def valider_sync(self):
        for g, membres in self.sync.items():
            champs = {m[1] for m in membres}
            if len(champs) > 1:
                self.c.warn(membres[0][2], "groupe de synchronisation %r : champs différents %s (synchronisation par valeurs)"
                            % (g, sorted(map(str, champs))))
            pages = [m[0] for m in membres]
            for p in set(pages):
                if pages.count(p) > 1:
                    self.c.warn(membres[0][2], "groupe de synchronisation %r : plusieurs segments sur la page %s" % (g, p))
            if len(membres) == 1:
                self.c.info(membres[0][2], "groupe de synchronisation %r utilisé sur une seule page" % g)
            selections = {m[3] for m in membres}
            if len(selections) > 1:
                self.c.warn(membres[0][2], "groupe de synchronisation %r : sélections enregistrées différentes selon les "
                            "pages (%s) ; la sélection affichée à l'ouverture dépendra de la page"
                            % (g, ", ".join(sorted(m[0] for m in membres))))

    def valider_filtres(self, fc, fichier):
        noms = set()
        for f in (fc or {}).get("filters", []):
            n = f.get("name")
            if n in noms:
                self.c.err(fichier, "nom de filtre en double : %s" % n)
            noms.add(n)
            fd = f.get("filter")
            if fd:
                for ex, val, al in comparaisons(fd):
                    if _champ_cle(ex, al) != _champ_cle(f.get("field")):
                        self.c.warn(fichier, "filtre %s : condition sur un autre champ que field" % n)
                if f.get("type") == "Categorical" and not any("In" in w.get("Condition", {}) or "Not" in w.get("Condition", {})
                                                              for w in fd.get("Where", [])):
                    self.c.warn(fichier, "filtre %s de type Categorical sans condition In" % n)
            gen = ((f.get("objects") or {}).get("general") or [{}])[0].get("properties", {})
            if _lit_brut(gen.get("requireSingleSelect")) == "true" and fd:
                for w in fd.get("Where", []):
                    if "In" in w.get("Condition", {}) and len(w["Condition"]["In"].get("Values", [])) > 1:
                        self.c.err(fichier, "filtre %s à sélection unique avec plusieurs valeurs" % n)
        return noms

    def valider_objets(self, type_visuel, objets, fichier, query_refs, conteneur):
        c = self.c
        cat = self.cat.communs if conteneur else self.cat.objets(type_visuel)
        for nom, entrees in objets.items():
            if not isinstance(entrees, list):
                continue
            spec_obj = cat.get(nom) if cat is not None else None
            if cat is not None and spec_obj is None:
                (c.err if conteneur else c.warn)(fichier, "objet de mise en forme %r inconnu pour %s" % (nom, type_visuel))
            for i, e in enumerate(entrees):
                sel = e.get("selector") or {}
                md = sel.get("metadata")
                if md is not None and md not in query_refs:
                    c.warn(fichier, "%s[%d] : selector.metadata %r ne correspond à aucun queryRef du visuel" % (nom, i, md))
                for pn, pv in (e.get("properties") or {}).items():
                    ou = "%s.%s" % (nom, pn)
                    if spec_obj is None:
                        continue
                    if pn not in spec_obj:
                        if (type_visuel, nom, pn) in PROPRIETES_NON_THEMATIQUES:
                            continue
                        c.warn(fichier, "propriété %r inconnue dans l'objet %r de %s" % (pn, nom, type_visuel))
                        continue
                    verifier_valeur_propriete(self.cat, spec_obj[pn], pv, ou, fichier, c)

    def valider_noeud_champs(self, doc, fichier):
        """Littéraux, couleurs, références au modèle (genre et existence), types des littéraux de filtres."""
        c = self.c
        for ou, val in litteraux(doc):
            if not isinstance(val, str):
                c.err(fichier, "%s : Value doit être une chaîne JSON" % ou)
            elif not (val in ("true", "false", "null") or RE_TEXTE.match(val) or RE_NOMBRE.match(val)
                      or RE_DATETIME.match(val)):
                c.err(fichier, "%s : littéral PBIR invalide %r" % (ou, val))
        for ou, val in _couleurs_solides(doc):
            _verifier_couleur(val, ou, fichier, c)
        for ou, agg in _agregations(doc):
            if "Measure" in agg.get("Expression", {}):
                c.err(fichier, "%s : agrégation d'une mesure" % ou)
        if not self.modele:
            return
        for genre, table, alias, prop, ou, schema in references(doc):
            if schema == "extension":
                c.warn(fichier, "%s : mesure d'extension de rapport %s[%s] non vérifiée" % (ou, table, prop))
                continue
            if table is None:
                c.err(fichier, "%s : alias de source %r non déclaré dans From" % (ou, alias))
                continue
            tn, remarque = self.modele.table(table)
            if tn is None:
                c.err(fichier, "%s : table %r absente du modèle" % (ou, table))
                continue
            if remarque:
                c.err(fichier, "%s : table %r : %s" % (ou, table, remarque))
            t = self.modele.tables[tn]
            if genre == "Column":
                if prop not in t["colonnes"]:
                    if prop in t["mesures"]:
                        c.err(fichier, "%s : %s[%s] est une mesure, référencée comme colonne" % (ou, tn, prop))
                    else:
                        proches = [x for x in t["colonnes"] if x.lower() == prop.lower()]
                        c.err(fichier, "%s : colonne %s[%s] absente du modèle%s" % (
                            ou, tn, prop, (" (casse : %s)" % proches[0]) if proches else ""))
            else:
                if prop not in t["mesures"]:
                    if prop in t["colonnes"]:
                        c.err(fichier, "%s : %s[%s] est une colonne, référencée comme mesure" % (ou, tn, prop))
                    else:
                        c.err(fichier, "%s : mesure %s[%s] absente du modèle" % (ou, tn, prop))
        for ex, val, al in comparaisons(doc):
            cle = _champ_cle(ex, al)
            tc = self._type_colonne(cle)
            if tc and not _lit_compatible(tc, val):
                c.err(fichier, "littéral %s incompatible avec la colonne %s[%s] de type %s" % (val, cle[1], cle[2], tc))


# noms de couleurs du thème admis dans un littéral de couleur ('minColor', 'foreground'...)
COULEURS_THEME = {"foreground", "background", "good", "neutral", "bad", "maximum", "center", "minimum", "null",
                  "accent", "tableAccent", "minColor", "maxColor", "midColor", "foregroundLight", "foregroundDark",
                  "foregroundNeutralLight", "foregroundNeutralDark", "foregroundNeutralSecondary",
                  "foregroundNeutralTertiary", "backgroundLight", "backgroundNeutral", "backgroundDark",
                  "secondaryBackground", "hyperlink", "visitedHyperlink", "disabledText", "shapeStroke"}


def _verifier_couleur(val, ou, fichier, c):
    if not isinstance(val, str):
        c.err(fichier, "%s : couleur non textuelle" % ou)
    elif RE_COULEUR.match(val):
        return
    elif re.match(r"^'[A-Za-z][A-Za-z0-9]*'$", val):
        if val[1:-1] not in COULEURS_THEME:
            c.warn(fichier, "%s : nom de couleur de thème inconnu %s" % (ou, val))
    else:
        c.err(fichier, "%s : couleur littérale invalide %s (attendu '#RRGGBB')" % (ou, val))


def _selection_enregistree(visuel):
    """Signature de la sélection enregistrée d'un segment (general.filter), alias remplacés par les tables."""
    for e in (visuel.get("objects") or {}).get("general", []):
        f = (e.get("properties") or {}).get("filter")
        fd = f.get("filter") if isinstance(f, dict) else None
        if isinstance(fd, dict):
            alias = {s.get("Name"): s.get("Entity") for s in fd.get("From", [])}

            def remplacer(n):
                if isinstance(n, dict):
                    if set(n) == {"Source"} and n["Source"] in alias:
                        return {"Entity": alias[n["Source"]]}
                    return {k: remplacer(x) for k, x in n.items()}
                if isinstance(n, list):
                    return [remplacer(x) for x in n]
                return n
            return json.dumps(remplacer(fd.get("Where")), sort_keys=True, ensure_ascii=False)
    return None


def _lit_brut(v):
    if isinstance(v, dict) and "expr" in v and "Literal" in v["expr"]:
        return v["expr"]["Literal"]["Value"]
    return None


def _lit_texte(v):
    b = _lit_brut(v)
    if b and RE_TEXTE.match(b):
        return b[1:-1].replace("''", "'")
    return None


def _couleurs_solides(noeud, chemin=""):
    if isinstance(noeud, dict):
        s = noeud.get("solid")
        if isinstance(s, dict) and isinstance(s.get("color"), dict):
            e = s["color"].get("expr", {})
            if isinstance(e, dict) and "Literal" in e:
                yield chemin + "/solid/color", e["Literal"].get("Value")
        for k, v in noeud.items():
            yield from _couleurs_solides(v, chemin + "/" + k)
    elif isinstance(noeud, list):
        for i, v in enumerate(noeud):
            yield from _couleurs_solides(v, chemin + "[%d]" % i)


def _agregations(noeud, chemin=""):
    if isinstance(noeud, dict):
        a = noeud.get("Aggregation")
        if isinstance(a, dict) and "Expression" in a:
            yield chemin + "/Aggregation", a
        for k, v in noeud.items():
            yield from _agregations(v, chemin + "/" + k)
    elif isinstance(noeud, list):
        for i, v in enumerate(noeud):
            yield from _agregations(v, chemin + "[%d]" % i)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Validation d'un projet PBIP (rapport PBIR).")
    ap.add_argument("dossier", help="dossier du projet (contenant le .pbip) ou dossier .Report")
    ap.add_argument("--bim", action="store_true", help="valider aussi model.bim avec validate_bim.py")
    ap.add_argument("--strict", action="store_true", help="les avertissements font échouer la validation")
    ap.add_argument("--quiet", action="store_true", help="n'afficher que le bilan et les erreurs")
    a = ap.parse_args(argv)
    v = Validateur(a.dossier, a.strict)
    c = v.executer()
    nb_bim = 0
    if a.bim and v.fichier_modele:
        import validate_bim
        rb = validate_bim.valider_fichier(v.fichier_modele)
        for f, m in rb.erreurs:
            c.err(v.fichier_modele, "model.bim : " + m)
        for f, m in rb.avert:
            c.warn(v.fichier_modele, "model.bim : " + m)
        nb_bim = 1
    for f, m in c.erreurs:
        print("ERREUR  %s : %s" % (f, m))
    if not a.quiet:
        for f, m in c.avert:
            print("AVERT.  %s : %s" % (f, m))
        for f, m in c.infos:
            print("INFO    %s : %s" % (f, m))
    nb_fichiers = sum(len(fs) for _dp, _d, fs in os.walk(v.racine))
    print("Bilan : %d erreur(s), %d avertissement(s), %d info(s) — %d page(s), modèle %s%s, %d fichiers parcourus"
          % (len(c.erreurs), len(c.avert), len(c.infos), len(v.pages),
             "chargé" if v.modele else "non chargé", " (+ validate_bim)" if nb_bim else "", nb_fichiers))
    echec = bool(c.erreurs) or (a.strict and bool(c.avert))
    return 1 if echec else 0


if __name__ == "__main__":
    sys.exit(main())

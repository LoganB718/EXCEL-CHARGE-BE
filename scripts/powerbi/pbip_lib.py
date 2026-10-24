"""Écriture d'un projet Power BI (PBIP) avec rapport au format PBIR (dossier .Report/definition).

Bibliothèque sans dépendance externe (matplotlib seulement pour preview_pages).
Versions de schémas retenues : jeu publié le 10 juin 2025 (Power BI Desktop 2.144, juin 2025, et suivants),
voir FORMAT_NOTES.md pour la justification.

Usage typique :
    from pbip_lib import *
    r = Rapport("PlanDeCharge_BI", theme=theme_personnalise("PlanDeCharge", [...]))
    p = r.page("Tableau de bord", largeur=1280, hauteur=1500)
    p.ajouter(carte(mes("Mesures", "Charge identifiée"), 20, 80, 200, 110, titre="Charge identifiée (h)"))
    ecrire_projet("/chemin/neuf", "PlanDeCharge_BI", r, model_bim=modele)
"""
import datetime as _dt
import hashlib
import json
import os
import re
import shutil
import uuid
from decimal import Decimal

# ---------------------------------------------------------------------------
# Versions de schémas (URL officielles ; copies locales dans pbi/schemas/)
# ---------------------------------------------------------------------------
BASE_SCHEMAS = "https://developer.microsoft.com/json-schemas/"
SCHEMAS = {
    "report": BASE_SCHEMAS + "fabric/item/report/definition/report/2.1.0/schema.json",
    "page": BASE_SCHEMAS + "fabric/item/report/definition/page/2.0.0/schema.json",
    "pages": BASE_SCHEMAS + "fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
    "version": BASE_SCHEMAS + "fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
    "visual": BASE_SCHEMAS + "fabric/item/report/definition/visualContainer/2.1.0/schema.json",
    # definition.pbir : 1.0.0 (publié le 12/05/2025, avant Desktop de juin 2025) plutôt que 2.0.0 (25/06/2025) ;
    # les deux décrivent byPath à l'identique
    "pbir": BASE_SCHEMAS + "fabric/item/report/definitionProperties/1.0.0/schema.json",
    "pbism": BASE_SCHEMAS + "fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
    "platform": BASE_SCHEMAS + "fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    "pbip": BASE_SCHEMAS + "fabric/pbip/pbipProperties/1.0.0/schema.json",
}
VERSION_DEFINITION = "2.0.0"       # definition/version.json
VERSION_PBIR = "4.0"               # definition.pbir (format PBIR = dossier definition/)
VERSION_PBISM = "4.0"              # definition.pbism (accepte model.bim ou TMDL)
VERSION_THEME_BASE = "5.61"        # reportVersionAtImport du thème de base CY24SU10 (valeur écrite par Desktop)
VERSION_THEME_PERSO = "5.64"       # valeur relevée dans un rapport réel au schéma report 1.3.0 (mai 2025)
THEME_BASE = "CY24SU10"
DOSSIER_RESSOURCES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ressources")

# rôles de requête dont les colonnes sont « actives » (convention de Desktop)
_ROLES_ACTIFS = {"Category", "Rows", "Columns"}
_NOM_OK = re.compile(r"^[A-Za-z0-9_-]{1,50}$")


# ---------------------------------------------------------------------------
# Outils
# ---------------------------------------------------------------------------
def ident(*parties, longueur=20):
    """Identifiant hexadécimal déterministe (comme les noms générés par Desktop)."""
    h = hashlib.sha1("\x1f".join(str(p) for p in parties).encode("utf-8")).hexdigest()
    return h[:longueur]


def guid(*parties):
    """GUID déterministe (logicalId des fichiers .platform)."""
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "pbip:" + "/".join(str(p) for p in parties)))


def ecrire_json(chemin, donnees):
    os.makedirs(os.path.dirname(chemin), exist_ok=True)
    with open(chemin, "w", encoding="utf-8", newline="\n") as f:
        json.dump(donnees, f, ensure_ascii=False, indent=2)
        f.write("\n")


def _nombre(x):
    """Écriture décimale sans notation scientifique ni « .0 » inutile."""
    if isinstance(x, bool):
        raise TypeError("booléen inattendu")
    if isinstance(x, int):
        return str(x)
    d = Decimal(repr(float(x)))
    if d == d.to_integral_value():
        return str(int(d))
    s = format(d, "f")
    return s.rstrip("0").rstrip(".") if "." in s else s


# ---------------------------------------------------------------------------
# Littéraux PBIR (valeurs de « Literal.Value »)
# ---------------------------------------------------------------------------
class Entier(int):
    """Entier à écrire avec le suffixe L (ex. labelPrecision, shapeAngle)."""


def lit_texte(s):
    """'texte' avec apostrophes doublées : "l'axe" -> 'l''axe'."""
    return "'" + str(s).replace("'", "''") + "'"


def lit_nombre(x, suffixe="D"):
    return _nombre(x) + suffixe


def lit_entier(n):
    return str(int(n)) + "L"


def lit_bool(b):
    return "true" if b else "false"


def lit_date(d):
    """datetime'2026-01-31T00:00:00' (date ou datetime Python)."""
    if isinstance(d, _dt.datetime):
        return "datetime'" + d.strftime("%Y-%m-%dT%H:%M:%S") + "'"
    if isinstance(d, _dt.date):
        return "datetime'" + d.strftime("%Y-%m-%d") + "T00:00:00'"
    raise TypeError("date attendue")


def lit_valeur(v):
    """Littéral PBIR d'une valeur Python (pour filtres et formatage)."""
    if v is None:
        return "null"
    if isinstance(v, bool):
        return lit_bool(v)
    if isinstance(v, Entier):
        return lit_entier(v)
    if isinstance(v, int):
        return lit_entier(v)
    if isinstance(v, float):
        return lit_nombre(v)
    if isinstance(v, (_dt.date, _dt.datetime)):
        return lit_date(v)
    if isinstance(v, str):
        return lit_texte(v)
    raise TypeError("valeur non convertible en littéral : %r" % (v,))


def literal(valeur_pbir):
    return {"Literal": {"Value": valeur_pbir}}


def expr(valeur_pbir):
    return {"expr": literal(valeur_pbir)}


# ---------------------------------------------------------------------------
# Couleurs
# ---------------------------------------------------------------------------
_HEX = re.compile(r"^#?([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")


def hex_couleur(c):
    m = _HEX.match(c)
    if not m:
        raise ValueError("couleur hexadécimale invalide : %r" % c)
    return "#" + m.group(1).upper()


class Couleur:
    """Valeur de propriété de type couleur ({"solid": {"color": {"expr": ...}}})."""

    def __init__(self, expression):
        self.expression = expression

    def pbir(self):
        return {"solid": {"color": {"expr": self.expression}}}


def couleur(c):
    """Couleur fixe : couleur('1F2A44') ou couleur('#1F2A44')."""
    return Couleur(literal(lit_texte(hex_couleur(c))))


def couleur_theme(indice, nuance=0):
    return Couleur({"ThemeDataColor": {"ColorId": int(indice), "Percent": float(nuance) if nuance else 0}})


def couleur_mesure(champ):
    """Couleur donnée par une mesure renvoyant un code « #RRGGBB » (mise en forme conditionnelle)."""
    if champ.genre != "Measure":
        raise ValueError("couleur_mesure attend une mesure")
    return Couleur(champ.expr())


# ---------------------------------------------------------------------------
# Champs du modèle
# ---------------------------------------------------------------------------
# QueryAggregateFunction (schéma semanticQuery) : 0 somme, 1 moyenne, 2 nombre distinct, 3 min, 4 max,
# 5 nombre de valeurs non vides ; second élément = préfixe du queryRef écrit par Desktop (« Sum(T.C) »)
_AGREGATS = {"somme": (0, "Sum"), "moyenne": (1, "Avg"), "nb_distinct": (2, "Count"),
             "min": (3, "Min"), "max": (4, "Max"), "nombre": (5, "CountNonNull")}


class Champ:
    """Référence à une colonne, une mesure ou une agrégation de colonne du modèle sémantique."""

    def __init__(self, table, nom, genre="Column", agregation=None):
        if genre not in ("Column", "Measure"):
            raise ValueError(genre)
        self.table, self.nom, self.genre, self.agregation = table, nom, genre, agregation

    def expr_base(self, source=None):
        ref = {"Source": source} if source else {"Entity": self.table}
        return {self.genre: {"Expression": {"SourceRef": ref}, "Property": self.nom}}

    def expr(self, source=None):
        base = self.expr_base(source)
        if self.agregation:
            return {"Aggregation": {"Expression": base, "Function": _AGREGATS[self.agregation][0]}}
        return base

    @property
    def query_ref(self):
        r = "%s.%s" % (self.table, self.nom)
        if self.agregation:
            r = "%s(%s)" % (_AGREGATS[self.agregation][1], r)
        return r

    def __repr__(self):
        return "Champ(%s[%s] %s)" % (self.table, self.nom, self.genre)


def col(table, nom):
    return Champ(table, nom, "Column")


def mes(table, nom):
    return Champ(table, nom, "Measure")


def agr(table, nom, fonction="somme"):
    if fonction not in _AGREGATS:
        raise ValueError(fonction)
    return Champ(table, nom, "Column", fonction)


def projection(champ, role=None, nom_affiche=None, actif=None):
    """Projection d'un rôle de requête : field, queryRef « Table.Champ », nativeQueryRef, displayName éventuel."""
    p = {"field": champ.expr(), "queryRef": champ.query_ref, "nativeQueryRef": champ.nom}
    if nom_affiche:
        p["displayName"] = nom_affiche
    if actif is None:
        actif = role in _ROLES_ACTIFS and champ.genre == "Column" and not champ.agregation
    if actif:
        p["active"] = True
    return p


# ---------------------------------------------------------------------------
# Définitions de filtres (FilterDefinition, version 2)
# ---------------------------------------------------------------------------
class _Sources:
    """Alias des tables (« From ») d'une définition de filtre."""

    def __init__(self):
        self.alias = {}

    def de(self, champ):
        if champ.table not in self.alias:
            base = re.sub(r"[^a-z]", "", champ.table.lower())[:1] or "t"
            a, i = base, 1
            while a in self.alias.values():
                i += 1
                a = "%s%d" % (base, i)
            self.alias[champ.table] = a
        return self.alias[champ.table]

    def from_(self):
        return [{"Name": a, "Entity": t, "Type": 0} for t, a in self.alias.items()]


def _definition(sources, conditions):
    return {"Version": 2, "From": sources.from_(), "Where": [{"Condition": c} for c in conditions]}


def filtre_valeurs(champ, valeurs):
    """Condition « champ dans (valeurs) » (sélection d'un segment, filtre de base)."""
    s = _Sources()
    a = s.de(champ)
    cond = {"In": {"Expressions": [champ.expr(a)], "Values": [[literal(lit_valeur(v))] for v in valeurs]}}
    return _definition(s, [cond])


_OPERATEURS = {"=": 0, ">": 1, ">=": 2, "<": 3, "<=": 4}


def _comparaison(champ, alias, operateur, valeur):
    return {"Comparison": {"ComparisonKind": _OPERATEURS[operateur], "Left": champ.expr(alias),
                           "Right": literal(lit_valeur(valeur))}}


def filtre_comparaison(champ, operateur, valeur):
    s = _Sources()
    a = s.de(champ)
    return _definition(s, [_comparaison(champ, a, operateur, valeur)])


def filtre_non_vide(champ):
    """Champ non vide (filtre avancé « n'est pas vide » d'une mesure ou colonne)."""
    s = _Sources()
    a = s.de(champ)
    return _definition(s, [{"Not": {"Expression": _comparaison(champ, a, "=", None)}}])


def filtre_plage_dates(champ, debut, fin):
    """Plage de dates incluse [debut, fin] telle que l'écrit un segment « Entre » : >= debut et < fin + 1 jour."""
    s = _Sources()
    a = s.de(champ)
    conds = []
    if debut is not None and fin is not None:
        conds.append({"And": {"Left": _comparaison(champ, a, ">=", debut),
                              "Right": _comparaison(champ, a, "<", fin + _dt.timedelta(days=1))}})
    elif debut is not None:
        conds.append(_comparaison(champ, a, ">=", debut))
    elif fin is not None:
        conds.append(_comparaison(champ, a, "<", fin + _dt.timedelta(days=1)))
    return _definition(s, conds)


def conteneur_filtre(champ, definition=None, type_filtre="Categorical", nom=None, cree_par="User",
                     masque=False, verrouille=False, selection_unique=False, ordinal=None):
    """Filtre du volet des filtres (rapport, page ou visuel)."""
    f = {"name": nom or ident("filtre", champ.query_ref, json.dumps(definition, sort_keys=True)),
         "field": champ.expr(), "type": type_filtre}
    if ordinal is not None:
        f["ordinal"] = ordinal
    if definition is not None:
        f["filter"] = definition
    f["howCreated"] = cree_par
    if masque:
        f["isHiddenInViewMode"] = True
    if verrouille:
        f["isLockedInViewMode"] = True
    if selection_unique:
        f["objects"] = {"general": [{"properties": {"requireSingleSelect": expr("true")}}]}
    return f


def filtre_categoriel(champ, valeurs=None, **opts):
    return conteneur_filtre(champ, filtre_valeurs(champ, valeurs) if valeurs is not None else None,
                            "Categorical", **opts)


def filtre_avance(champ, definition, **opts):
    return conteneur_filtre(champ, definition, "Advanced", **opts)


# ---------------------------------------------------------------------------
# Objets de mise en forme
# ---------------------------------------------------------------------------
def valeur_propriete(v):
    """Conversion d'une valeur Python en valeur de propriété PBIR :
    Couleur -> {"solid": ...} ; Champ -> {"expr": champ} ; dict -> tel quel ; bool -> true/false ;
    Entier(n) -> nL ; int/float -> nD ; str -> 'texte' ; date -> datetime'...' ; None -> null."""
    if isinstance(v, Couleur):
        return v.pbir()
    if isinstance(v, Champ):
        return {"expr": v.expr()}
    if isinstance(v, dict):
        return v
    if isinstance(v, bool):
        return expr(lit_bool(v))
    if isinstance(v, Entier):
        return expr(lit_entier(v))
    if isinstance(v, (int, float)):
        return expr(lit_nombre(v))
    return expr(lit_valeur(v))


def objet(proprietes, selecteur=None):
    """Entrée d'objet de mise en forme : {"properties": {...}, "selector": {...}}."""
    e = {"properties": {k: valeur_propriete(v) for k, v in proprietes.items()}}
    if selecteur is not None:
        e["selector"] = selecteur
    return e


def sel_metadata(champ_ou_queryref):
    q = champ_ou_queryref.query_ref if isinstance(champ_ou_queryref, Champ) else champ_ou_queryref
    return {"metadata": q}


def sel_id(i="default"):
    return {"id": i}


def sel_par_point(champ_ou_queryref=None):
    """Évaluation par point de données (mise en forme conditionnelle par mesure)."""
    s = {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}
    if champ_ou_queryref is not None:
        s.update(sel_metadata(champ_ou_queryref))
    return s


def sel_valeur(champ, valeur):
    """Point de données dont la colonne vaut « valeur » (couleur d'une série ou d'une catégorie)."""
    return {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": champ.expr(),
                                                 "Right": literal(lit_valeur(valeur))}}}]}


# ---------------------------------------------------------------------------
# Visuels
# ---------------------------------------------------------------------------
class Visuel:
    """Conteneur de visuel (visual.json). Les fonctions de haut niveau ci-dessous en construisent."""

    def __init__(self, type_visuel, x, y, largeur, hauteur, nom=None, cle=None):
        self.type = type_visuel
        self.x, self.y, self.largeur, self.hauteur = x, y, largeur, hauteur
        self.nom = nom
        self.cle = cle
        self.z = None
        self.ordre_tab = None
        self.roles = {}          # rôle -> liste de projections
        self.tri = []            # [(champ, "Ascending"|"Descending")]
        self.objets = {}         # objets du visuel
        self.conteneur = {}      # visualContainerObjects
        self.filtres = []        # filterConfig.filters
        self.sync = None
        self.masque = False
        self.cree_par = None
        self.drill_autres = True
        self.libelle = None      # texte court pour l'aperçu

    # -- données
    def ajouter_role(self, role, champ, nom_affiche=None, actif=None):
        self.roles.setdefault(role, []).append(projection(champ, role, nom_affiche, actif))
        return self

    def trier(self, champ, sens="Descending"):
        self.tri.append((champ, sens))
        return self

    # -- mise en forme
    def objet(self, nom, proprietes, selecteur=None):
        self.objets.setdefault(nom, []).append(objet(proprietes, selecteur))
        return self

    def objet_conteneur(self, nom, proprietes, selecteur=None):
        self.conteneur.setdefault(nom, []).append(objet(proprietes, selecteur))
        return self

    def query_refs(self):
        return [p["queryRef"] for ps in self.roles.values() for p in ps]

    def vers_json(self, nom):
        pos = {"x": self.x, "y": self.y, "z": self.z if self.z is not None else 0,
               "height": self.hauteur, "width": self.largeur}
        if self.ordre_tab is not None:
            pos["tabOrder"] = self.ordre_tab
        v = {"visualType": self.type}
        if self.roles:
            q = {"queryState": {r: {"projections": ps} for r, ps in self.roles.items()}}
            if self.tri:
                q["sortDefinition"] = {"sort": [{"field": c.expr(), "direction": s} for c, s in self.tri]}
            v["query"] = q
        if self.objets:
            v["objects"] = self.objets
        if self.conteneur:
            v["visualContainerObjects"] = self.conteneur
        if self.sync:
            v["syncGroup"] = self.sync
        v["drillFilterOtherVisuals"] = self.drill_autres
        d = {"$schema": SCHEMAS["visual"], "name": nom, "position": pos, "visual": v}
        if self.filtres:
            d["filterConfig"] = {"filters": self.filtres}
        if self.masque:
            d["isHidden"] = True
        if self.cree_par:
            d["howCreated"] = self.cree_par
        return d


# Couleurs par défaut des titres (bleu nuit du classeur)
STYLE_TITRE = {"fontSize": 11, "bold": True, "fontColor": "1F2A44"}


def _habiller(v, titre=None, sous_titre=None, fond=None, transparence_fond=0, bordure=None, rayon=None,
              ombre=False, entete=None, aide=None, infobulle=True, page_infobulle=None, infobulles=None,
              texte_alt=None, masque=False, filtres=None, z=None, ordre_tab=None, marges=None,
              style_titre=None, nom=None, cle=None):
    """Options communes de conteneur (titre, fond, bordure, en-tête, info-bulles, filtres...)."""
    if nom:
        v.nom = nom
    if cle:
        v.cle = cle
    st = dict(STYLE_TITRE)
    st.update(style_titre or {})
    if titre:
        v.libelle = titre
        v.objet_conteneur("title", {"show": True, "text": titre, "fontSize": st["fontSize"],
                                    "bold": st["bold"], "fontColor": couleur(st["fontColor"]),
                                    "titleWrap": True})
    else:
        v.objet_conteneur("title", {"show": False})
    if sous_titre:
        v.objet_conteneur("subTitle", {"show": True, "text": sous_titre, "fontSize": 9,
                                       "fontColor": couleur("6B7280")})
    if fond is not None:
        if fond is False:
            v.objet_conteneur("background", {"show": False})
        else:
            v.objet_conteneur("background", {"show": True, "color": couleur(fond),
                                             "transparency": float(transparence_fond)})
    if bordure is not None:
        if bordure is False:
            v.objet_conteneur("border", {"show": False})
        else:
            pr = {"show": True, "color": couleur(bordure)}
            if rayon is not None:
                pr["radius"] = float(rayon)
            v.objet_conteneur("border", pr)
    if ombre is not None:
        v.objet_conteneur("dropShadow", {"show": bool(ombre)})
    if marges is not None:
        h, d, b, g = marges
        v.objet_conteneur("padding", {"top": float(h), "right": float(d), "bottom": float(b), "left": float(g)})
    if entete is not None or aide:
        pr = {"show": bool(entete) or bool(aide)}
        if aide:
            pr["showTooltipButton"] = True
        v.objet_conteneur("visualHeader", pr)
    if aide:
        v.objet_conteneur("visualHeaderTooltip", {"type": "Default", "text": aide})
    if page_infobulle:
        v.objet_conteneur("visualTooltip", {"show": True, "type": "Canvas", "section": page_infobulle})
    elif infobulle is False:
        v.objet_conteneur("visualTooltip", {"show": False})
    for c in infobulles or []:
        v.ajouter_role("Tooltips", c)
    if texte_alt:
        v.objet_conteneur("general", {"altText": texte_alt})
    if masque:
        v.masque = True
    for f in filtres or []:
        v.filtres.append(f)
    if z is not None:
        v.z = z
    if ordre_tab is not None:
        v.ordre_tab = ordre_tab
    return v


def _run(texte, gras=False, couleur_texte=None, taille=None, police=None, italique=False, souligne=False,
         url=None):
    st = {}
    if police:
        st["fontFamily"] = police
    if taille:
        st["fontSize"] = "%spt" % _nombre(taille)
    if gras:
        st["fontWeight"] = "bold"
    if italique:
        st["fontStyle"] = "italic"
    if souligne:
        st["textDecoration"] = "underline"
    if couleur_texte:
        st["color"] = hex_couleur(couleur_texte)
    r = {"value": texte}
    if st:
        r["textStyle"] = st
    if url:
        r["url"] = url
    return r


def paragraphe(*morceaux, alignement=None):
    """Paragraphe de zone de texte. Chaque morceau : texte, ou dict(texte=..., gras=, couleur=, taille=,
    police=, italique=, souligne=, url=)."""
    runs = []
    for m in morceaux:
        if isinstance(m, str):
            runs.append(_run(m))
        else:
            m = dict(m)
            runs.append(_run(m.pop("texte"), couleur_texte=m.pop("couleur", None), **m))
    p = {"textRuns": runs or [{"value": ""}]}
    if alignement:
        p["horizontalTextAlignment"] = alignement
    return p


def zone_texte(paragraphes, x, y, largeur, hauteur, fond=False, **opts):
    """Zone de texte riche. paragraphes : liste de paragraphe(...) (ou de chaînes)."""
    paras = [paragraphe(p) if isinstance(p, str) else p for p in paragraphes]
    v = Visuel("textbox", x, y, largeur, hauteur)
    v.objets["general"] = [{"properties": {"paragraphs": paras}}]
    v.libelle = " ".join(r["value"] for p in paras for r in p["textRuns"])[:60]
    opts.setdefault("entete", False)
    opts.setdefault("infobulle", None)
    return _habiller(v, fond=fond, **opts)


def forme_rectangle(x, y, largeur, hauteur, remplissage="1F2A44", transparence=0, arrondi=0, contour=None,
                    **opts):
    """Forme rectangle pleine (bandeau). arrondi > 0 : rectangle à coins arrondis (pixels)."""
    v = Visuel("shape", x, y, largeur, hauteur)
    if arrondi:
        v.objet("shape", {"tileShape": "rectangleRounded", "rectangleRoundedCurve": Entier(arrondi)})
    else:
        v.objet("shape", {"tileShape": "rectangle"})
    v.objet("rotation", {"shapeAngle": Entier(0)})
    v.objet("fill", {"show": True})
    v.objet("fill", {"fillColor": couleur(remplissage), "transparency": float(transparence)}, sel_id())
    if contour:
        v.objet("outline", {"show": True})
        v.objet("outline", {"lineColor": couleur(contour), "weight": 1.0}, sel_id())
    else:
        v.objet("outline", {"show": False})
    v.libelle = "rectangle %s" % hex_couleur(remplissage)
    opts.setdefault("entete", False)
    opts.setdefault("infobulle", None)
    v.cree_par = "InsertVisualButton"
    return _habiller(v, **opts)


def carte(champ, x, y, largeur, hauteur, titre=None, etiquette=None, afficher_etiquette=True,
          taille_valeur=24, couleur_valeur="1F2A44", couleur_etiquette="6B7280", taille_etiquette=9,
          unites=1, decimales=None, gras=True, fond="FFFFFF", **opts):
    """Carte classique (« card ») : une valeur, étiquette de catégorie (sous la valeur) et/ou titre.
    unites : 0 auto, 1 aucune, 1000 milliers... ; decimales : nombre de décimales (None = format du modèle)."""
    v = Visuel("card", x, y, largeur, hauteur)
    v.ajouter_role("Values", champ, nom_affiche=etiquette)
    lab = {"fontSize": float(taille_valeur), "color": couleur(couleur_valeur), "bold": gras,
           "labelDisplayUnits": float(unites)}
    if decimales is not None:
        lab["labelPrecision"] = Entier(decimales)
    v.objet("labels", lab)
    v.objet("categoryLabels", {"show": bool(afficher_etiquette), "color": couleur(couleur_etiquette),
                               "fontSize": float(taille_etiquette)})
    v.objet("wordWrap", {"show": True})
    v.libelle = titre or etiquette or champ.nom
    opts.setdefault("entete", False)
    return _habiller(v, titre=titre, fond=fond, **opts)


def carte_multiligne(champs, x, y, largeur, hauteur, titre=None, noms=None, taille_valeur=14,
                     couleur_valeur="1F2A44", barre=None, fond="FFFFFF", **opts):
    """Carte à plusieurs lignes (multiRowCard) : liste de mesures/colonnes."""
    v = Visuel("multiRowCard", x, y, largeur, hauteur)
    noms = noms or [None] * len(champs)
    for c, n in zip(champs, noms):
        v.ajouter_role("Values", c, nom_affiche=n)
    v.objet("dataLabels", {"fontSize": float(taille_valeur), "color": couleur(couleur_valeur)})
    v.objet("categoryLabels", {"show": True, "fontSize": 9.0, "color": couleur("6B7280")})
    if barre:
        v.objet("card", {"barShow": True, "barColor": couleur(barre)})
    else:
        v.objet("card", {"barShow": False})
    v.libelle = titre or "carte multiligne"
    return _habiller(v, titre=titre, fond=fond, **opts)


def segment(champ, x, y, largeur, hauteur, titre=None, mode="Basic", selection_unique=False,
            valeur_initiale=None, valeurs_initiales=None, groupe_sync=None, orientation=None,
            tout_selectionner=False, taille_texte=10, fond="FFFFFF", **opts):
    """Segment (slicer) de type liste (mode='Basic') ou liste déroulante (mode='Dropdown').
    selection_unique : sélection unique obligatoire ; valeur_initiale(s) : sélection enregistrée dans le visuel ;
    groupe_sync : nom du groupe de synchronisation entre pages (même champ sur toutes les pages)."""
    if mode not in ("Basic", "Dropdown"):
        raise ValueError("mode de segment : Basic ou Dropdown")
    v = Visuel("slicer", x, y, largeur, hauteur)
    v.ajouter_role("Values", champ, actif=True)
    v.objet("data", {"mode": mode})
    gen = {}
    if orientation is not None:
        gen["orientation"] = 1.0 if orientation == "horizontal" else 0.0   # Desktop écrit « 0D » / « 1D »
    valeurs = valeurs_initiales if valeurs_initiales is not None else (
        [valeur_initiale] if valeur_initiale is not None else None)
    if valeurs is not None:
        if selection_unique and len(valeurs) != 1:
            raise ValueError("sélection unique : une seule valeur initiale")
        gen["filter"] = {"filter": filtre_valeurs(champ, valeurs)}
    if gen:
        v.objets["general"] = [objet(gen)]
    sel = {"strictSingleSelect": bool(selection_unique)}
    if not selection_unique:
        sel["selectAllCheckboxEnabled"] = bool(tout_selectionner)
    v.objet("selection", sel)
    if titre:
        v.objet("header", {"show": True, "text": titre, "fontColor": couleur("1F2A44"), "bold": True,
                           "textSize": float(taille_texte)})
    else:
        v.objet("header", {"show": False})
    v.objet("items", {"textSize": float(taille_texte), "fontColor": couleur("1F2937")})
    if groupe_sync:
        v.sync = {"groupName": groupe_sync, "fieldChanges": True, "filterChanges": True}
    v.libelle = titre or champ.nom
    opts.setdefault("entete", False)
    return _habiller(v, fond=fond, **opts)


def segment_dates(champ, x, y, largeur, hauteur, debut=None, fin=None, titre=None, groupe_sync=None,
                  taille_texte=10, fond="FFFFFF", **opts):
    """Segment de plage de dates « Entre » (colonne de type date) avec plage initiale enregistrée."""
    v = Visuel("slicer", x, y, largeur, hauteur)
    v.ajouter_role("Values", champ, actif=True)
    data = {"mode": "Between"}
    if debut is not None:
        data["startDate"] = {"expr": literal(lit_date(debut))}
    if fin is not None:
        data["endDate"] = {"expr": literal(lit_date(fin))}
    v.objet("data", data)
    if debut is not None or fin is not None:
        v.objets["general"] = [{"properties": {"filter": {"filter": filtre_plage_dates(champ, debut, fin)}}}]
    if titre:
        v.objet("header", {"show": True, "text": titre, "fontColor": couleur("1F2A44"), "bold": True,
                           "textSize": float(taille_texte)})
    else:
        v.objet("header", {"show": False})
    v.objet("date", {"textSize": float(taille_texte)})
    if groupe_sync:
        v.sync = {"groupName": groupe_sync, "fieldChanges": True, "filterChanges": True}
    v.libelle = titre or champ.nom
    v.tri.append((champ, "Ascending"))
    opts.setdefault("entete", False)
    return _habiller(v, fond=fond, **opts)


def _axes(v, montrer_valeurs=True, titre_axe_valeurs=None, quadrillage=True, taille=9):
    pa = {"show": True, "showAxisTitle": False, "fontSize": float(taille), "concatenateLabels": False,
          "labelColor": couleur("374151")}
    v.objet("categoryAxis", pa)
    pv = {"show": bool(montrer_valeurs), "fontSize": float(taille), "labelDisplayUnits": 1.0,
          "gridlineShow": bool(quadrillage), "labelColor": couleur("374151")}
    if titre_axe_valeurs:
        pv.update({"showAxisTitle": True, "titleText": titre_axe_valeurs})
    else:
        pv["showAxisTitle"] = False
    return pv


def graphique_combo(categorie, colonnes, x, y, largeur, hauteur, serie=None, lignes=(), titre=None,
                    couleurs_series=None, couleur_colonnes=None, couleurs_lignes=None, pointilles=True,
                    epaisseur_ligne=2, meme_axe=True, etiquettes=False, legende="Top", tri_categorie=True,
                    titre_axe=None, fond="FFFFFF", **opts):
    """Histogramme empilé + courbe (lineStackedColumnComboChart).
    colonnes : mesures empilées (rôle Y) ; serie : colonne de légende (empilement par valeur) ;
    lignes : mesures en courbe (rôle Y2) ; couleurs_series : {valeur de série: couleur} ;
    couleurs_lignes : {mesure ou queryRef: couleur} ; meme_axe : courbe sur l'axe Y principal (axe secondaire masqué)."""
    v = Visuel("lineStackedColumnComboChart", x, y, largeur, hauteur)
    v.ajouter_role("Category", categorie)
    if serie is not None:
        v.ajouter_role("Series", serie)
    for c in colonnes:
        v.ajouter_role("Y", c)
    for c in lignes:
        v.ajouter_role("Y2", c)
    if tri_categorie:
        v.trier(categorie, "Ascending")
    pv = _axes(v, titre_axe_valeurs=titre_axe)
    pv["secShow"] = not meme_axe
    if meme_axe:
        pv["alignZeros"] = True
    v.objet("valueAxis", pv)
    if couleur_colonnes and serie is None:
        for c in colonnes:
            v.objet("dataPoint", {"fill": couleur(couleur_colonnes)}, sel_metadata(c))
    for val, coul in (couleurs_series or {}).items():
        v.objet("dataPoint", {"fill": couleur(coul)}, sel_valeur(serie, val))
    for c in lignes:
        coul = (couleurs_lignes or {}).get(c.nom) or (couleurs_lignes or {}).get(c.query_ref) or "1F2937"
        v.objet("dataPoint", {"fill": couleur(coul)}, sel_metadata(c))
        st = {"strokeWidth": Entier(epaisseur_ligne), "showMarker": False,
              "lineStyle": "dashed" if pointilles else "solid"}
        v.objet("lineStyles", st, sel_metadata(c))
    v.objet("labels", {"show": bool(etiquettes), "labelDisplayUnits": 1.0})
    if legende:
        v.objet("legend", {"show": True, "position": legende, "fontSize": 9.0})
    else:
        v.objet("legend", {"show": False})
    v.libelle = titre or "combo"
    return _habiller(v, titre=titre, fond=fond, **opts)


def graphique_barres(categorie, valeurs, x, y, largeur, hauteur, type_visuel="clusteredBarChart", serie=None,
                     titre=None, couleur_barres=None, couleurs_categories=None, couleurs_series=None,
                     etiquettes=True, decimales=0, legende=None, tri=None, montrer_axe_valeurs=False,
                     fond="FFFFFF", **opts):
    """Barres / histogrammes : clusteredBarChart, barChart (empilé), columnChart (empilé),
    clusteredColumnChart. tri : (champ, 'Ascending'|'Descending'), par défaut la 1re mesure décroissante."""
    if type_visuel not in ("clusteredBarChart", "barChart", "columnChart", "clusteredColumnChart"):
        raise ValueError(type_visuel)
    v = Visuel(type_visuel, x, y, largeur, hauteur)
    v.ajouter_role("Category", categorie)
    if serie is not None:
        v.ajouter_role("Series", serie)
    for c in valeurs:
        v.ajouter_role("Y", c)
    if tri is None:
        v.trier(valeurs[0], "Descending")
    elif tri:
        v.trier(*tri)
    pv = _axes(v, montrer_valeurs=montrer_axe_valeurs, quadrillage=montrer_axe_valeurs)
    v.objet("valueAxis", pv)
    if couleur_barres:
        v.objet("dataPoint", {"fill": couleur(couleur_barres)})
    for val, coul in (couleurs_categories or {}).items():
        v.objet("dataPoint", {"fill": couleur(coul)}, sel_valeur(categorie, val))
    for val, coul in (couleurs_series or {}).items():
        v.objet("dataPoint", {"fill": couleur(coul)}, sel_valeur(serie, val))
    lab = {"show": bool(etiquettes), "labelDisplayUnits": 1.0}
    if etiquettes and decimales is not None:
        lab["labelPrecision"] = Entier(decimales)
    v.objet("labels", lab)
    if legende:
        v.objet("legend", {"show": True, "position": legende, "fontSize": 9.0})
    else:
        v.objet("legend", {"show": False})
    v.libelle = titre or type_visuel
    return _habiller(v, titre=titre, fond=fond, **opts)


def anneau(categorie, valeur, x, y, largeur, hauteur, titre=None, couleurs=None, style_etiquettes="Category, percent of total",
           position_etiquettes="preferOutside", legende=False, total=None, taille_total=18,
           couleur_total="1F2A44", libelle_total=None, fond="FFFFFF", **opts):
    """Graphique en anneau. total : champ affiché au centre par une carte superposée (z supérieur).
    Retourne une liste [anneau] ou [anneau, carte_centre] à ajouter à la page."""
    v = Visuel("donutChart", x, y, largeur, hauteur)
    v.ajouter_role("Category", categorie)
    v.ajouter_role("Y", valeur)
    v.trier(valeur, "Descending")
    for val, coul in (couleurs or {}).items():
        v.objet("dataPoint", {"fill": couleur(coul)}, sel_valeur(categorie, val))
    v.objet("labels", {"show": True, "labelStyle": style_etiquettes, "position": position_etiquettes,
                       "fontSize": 9.0, "labelDisplayUnits": 1.0})
    if legende:
        v.objet("legend", {"show": True, "position": legende, "fontSize": 9.0})
    else:
        v.objet("legend", {"show": False})
    v.libelle = titre or "anneau"
    _habiller(v, titre=titre, fond=fond, **opts)
    if total is None:
        return [v]
    # carte centrée dans le trou de l'anneau (zone de tracé sous le titre)
    haut_titre = 30 if titre else 0
    cw, ch = min(130, largeur * 0.36), 56
    cx = x + (largeur - cw) / 2
    cy = y + haut_titre + (hauteur - haut_titre - ch) / 2
    c = carte(total, round(cx, 1), round(cy, 1), round(cw, 1), ch, etiquette=libelle_total or "Total",
              afficher_etiquette=True, taille_valeur=taille_total, couleur_valeur=couleur_total,
              fond=False, infobulle=False)
    c.cle = (v.cle + "_total") if v.cle else None
    if v.z is not None:
        c.z = v.z + 1
    return [v, c]


class ColonneTableau:
    """Colonne d'un tableau : champ, largeur (px), titre affiché, couleurs conditionnelles (mesures couleur)."""

    def __init__(self, champ, largeur=None, titre=None, couleur_police=None, couleur_fond=None,
                 alignement=None):
        self.champ, self.largeur, self.titre = champ, largeur, titre
        self.couleur_police, self.couleur_fond, self.alignement = couleur_police, couleur_fond, alignement


def tableau(colonnes, x, y, largeur, hauteur, titre=None, tri=None, totaux=False, taille_texte=9,
            entete_fond="1F2A44", entete_texte="FFFFFF", fond="FFFFFF", **opts):
    """Tableau (tableEx). colonnes : liste de ColonneTableau (ou de Champ).
    tri : liste de (champ, sens). Couleur de police / de fond conditionnelle : mesure renvoyant '#RRGGBB'."""
    v = Visuel("tableEx", x, y, largeur, hauteur)
    cols = [c if isinstance(c, ColonneTableau) else ColonneTableau(c) for c in colonnes]
    for c in cols:
        v.ajouter_role("Values", c.champ, nom_affiche=c.titre)
    for champ, sens in (tri or []):
        v.trier(champ, sens)
    for c in cols:
        if c.largeur:
            v.objet("columnWidth", {"value": float(c.largeur)}, sel_metadata(c.champ))
    for c in cols:
        pr = {}
        if c.couleur_police is not None:
            pr["fontColor"] = couleur_mesure(c.couleur_police) if isinstance(c.couleur_police, Champ) \
                else couleur(c.couleur_police)
        if c.couleur_fond is not None:
            pr["backColor"] = couleur_mesure(c.couleur_fond) if isinstance(c.couleur_fond, Champ) \
                else couleur(c.couleur_fond)
        if pr:
            v.objet("values", pr, sel_par_point(c.champ))
        if c.alignement:
            v.objet("columnFormatting", {"alignment": c.alignement}, sel_metadata(c.champ))
    v.objet("values", {"fontSize": float(taille_texte), "wordWrap": False})
    v.objet("columnHeaders", {"fontSize": float(taille_texte), "bold": True, "backColor": couleur(entete_fond),
                              "fontColor": couleur(entete_texte), "wordWrap": True})
    v.objet("total", {"totals": bool(totaux)})
    v.objet("grid", {"rowPadding": 3.0, "gridHorizontal": True, "gridHorizontalColor": couleur("E5E7EB")})
    v.libelle = titre or "tableau"
    return _habiller(v, titre=titre, fond=fond, **opts)


def matrice(lignes, colonnes, valeurs, x, y, largeur, hauteur, titre=None, couleur_fond=None,
            couleur_police=None, totaux_lignes=False, totaux_colonnes=False, largeurs=None, tri=None,
            taille_texte=9, entete_fond="1F2A44", entete_texte="FFFFFF", disposition="Compact",
            fond="FFFFFF", **opts):
    """Matrice (pivotTable) lignes x colonnes. couleur_fond / couleur_police : mesure couleur appliquée
    aux cellules de valeurs (par point). totaux_* : afficher ou non les totaux. largeurs : {champ: px}."""
    v = Visuel("pivotTable", x, y, largeur, hauteur)
    for c in lignes:
        v.ajouter_role("Rows", c)
    for c in colonnes:
        v.ajouter_role("Columns", c)
    for c in valeurs:
        v.ajouter_role("Values", c)
    for champ, sens in (tri or []):
        v.trier(champ, sens)
    for c in valeurs:
        pr = {}
        if couleur_fond is not None:
            pr["backColor"] = couleur_mesure(couleur_fond)
        if couleur_police is not None:
            pr["fontColor"] = couleur_mesure(couleur_police)
        if pr:
            v.objet("values", pr, sel_par_point(c))
    for champ, px in (largeurs or {}).items():
        v.objet("columnWidth", {"value": float(px)}, sel_metadata(champ))
    v.objet("subTotals", {"rowSubtotals": bool(totaux_lignes), "columnSubtotals": bool(totaux_colonnes)})
    v.objet("values", {"fontSize": float(taille_texte)})
    v.objet("columnHeaders", {"fontSize": float(taille_texte), "bold": True, "backColor": couleur(entete_fond),
                              "fontColor": couleur(entete_texte)})
    v.objet("rowHeaders", {"fontSize": float(taille_texte)})
    v.objet("general", {"layout": disposition})
    v.libelle = titre or "matrice"
    return _habiller(v, titre=titre, fond=fond, **opts)


def bouton(texte, x, y, largeur, hauteur, action="PageNavigation", page_cible=None, remplissage="2A78D6",
           couleur_texte="FFFFFF", taille_texte=11, icone="blank", arrondi=6, info=None, **opts):
    """Bouton (actionButton). action : 'PageNavigation' (page_cible = nom de page), 'Back' (retour),
    'Drillthrough' (page_cible = page d'extraction)."""
    if action not in ("PageNavigation", "Back", "Drillthrough"):
        raise ValueError(action)
    v = Visuel("actionButton", x, y, largeur, hauteur)
    v.objet("icon", {"shapeType": icone}, sel_id())
    v.objet("icon", {"show": icone != "blank"})
    v.objet("text", {"show": True})
    v.objet("text", {"text": texte, "fontColor": couleur(couleur_texte), "fontSize": float(taille_texte),
                     "bold": True}, sel_id())
    v.objet("fill", {"show": True})
    v.objet("fill", {"fillColor": couleur(remplissage), "transparency": 0.0}, sel_id())
    v.objet("outline", {"show": False})
    if arrondi:
        v.objet("shape", {"roundEdge": Entier(arrondi)}, sel_id())
    lien = {"show": True, "type": action}
    if action == "PageNavigation":
        if not page_cible:
            raise ValueError("page_cible requise")
        lien["navigationSection"] = page_cible
    elif action == "Drillthrough":
        if not page_cible:
            raise ValueError("page_cible requise")
        lien["drillthroughSection"] = page_cible
    if info:
        lien["tooltip"] = info
    v.objet_conteneur("visualLink", lien)
    v.cree_par = "InsertVisualButton"
    v.libelle = texte
    opts.setdefault("entete", False)
    opts.setdefault("infobulle", None)
    return _habiller(v, **opts)


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
class Page:
    def __init__(self, nom_affiche, nom=None, largeur=1280, hauteur=720, affichage="FitToWidth",
                 fond="FFFFFF", transparence_fond=0, papier_peint=None, masquee=False, filtres=None,
                 extraction=None, garder_tous_filtres=True, type_page=None):
        self.nom_affiche = nom_affiche
        self.nom = nom or ident("page", nom_affiche)
        if not _NOM_OK.match(self.nom):
            raise ValueError("nom de page invalide (lettres, chiffres, _ et -, 50 car. max) : %r" % self.nom)
        if affichage not in ("FitToPage", "FitToWidth", "ActualSize", "ActualSizeTopLeft"):
            raise ValueError(affichage)
        self.largeur, self.hauteur, self.affichage = largeur, hauteur, affichage
        self.fond, self.transparence_fond, self.papier_peint = fond, transparence_fond, papier_peint
        self.masquee = masquee
        self.filtres = list(filtres or [])
        self.extraction = list(extraction or [])
        self.garder_tous_filtres = garder_tous_filtres
        self.type_page = type_page
        self.visuels = []
        self.interactions = []

    def ajouter(self, *visuels):
        """Ajoute un ou plusieurs visuels (ou listes de visuels, ex. retour de anneau())."""
        ajout = []
        for v in visuels:
            if isinstance(v, (list, tuple)):
                ajout.extend(v)
            else:
                ajout.append(v)
        for v in ajout:
            if not isinstance(v, Visuel):
                raise TypeError("Visuel attendu")
            self.visuels.append(v)
        return ajout[-1] if len(ajout) == 1 else ajout

    def interaction(self, source, cible, type_interaction="NoFilter"):
        """Interaction entre deux visuels de la page (objets Visuel ou noms) :
        Default, DataFilter (filtrer), HighlightFilter (mettre en évidence), NoFilter (aucune)."""
        if type_interaction not in ("Default", "DataFilter", "HighlightFilter", "NoFilter"):
            raise ValueError(type_interaction)
        self.interactions.append((source, cible, type_interaction))

    def _noms_visuels(self):
        """Attribue noms, z et ordre de tabulation (lecture haut-gauche -> bas-droite)."""
        noms, vus = {}, set()
        for i, v in enumerate(self.visuels):
            n = v.nom or ident(self.nom, "visuel", v.cle or i, v.type)
            if not _NOM_OK.match(n):
                raise ValueError("nom de visuel invalide : %r" % n)
            if n in vus:
                raise ValueError("nom de visuel en double sur la page %s : %s" % (self.nom_affiche, n))
            vus.add(n)
            noms[id(v)] = n
        for i, v in enumerate(self.visuels):
            if v.z is None:
                v.z = i * 1000
        ordre = sorted(range(len(self.visuels)), key=lambda i: (round(self.visuels[i].y / 10), self.visuels[i].x))
        for rang, i in enumerate(ordre):
            if self.visuels[i].ordre_tab is None:
                self.visuels[i].ordre_tab = rang * 1000
        return noms

    def vers_json(self, noms=None):
        noms = noms or {}
        d = {"$schema": SCHEMAS["page"], "name": self.nom, "displayName": self.nom_affiche,
             "displayOption": self.affichage, "height": self.hauteur, "width": self.largeur}
        filtres = list(self.filtres)
        binding = None
        if self.extraction:
            params = []
            for i, champ in enumerate(self.extraction):
                nf = ident(self.nom, "extraction", i)   # noms uniques dans tout le rapport
                filtres.insert(i, conteneur_filtre(champ, None, "Categorical", nom=nf, cree_par="Drillthrough"))
                params.append({"name": "Param_" + nf, "boundFilter": nf, "fieldExpr": champ.expr()})
            binding = {"name": ident(self.nom, "liaison"), "type": "Drillthrough", "parameters": params}
            if not self.garder_tous_filtres:
                binding["acceptsFilterContext"] = "None"
        if filtres:
            for i, f in enumerate(filtres):
                f.setdefault("ordinal", i)
            d["filterConfig"] = {"filters": filtres, "filterSortOrder": "Custom"}
        if binding:
            d["pageBinding"] = binding
        objs = {}
        if self.fond is not None:
            objs["background"] = [objet({"color": couleur(self.fond), "transparency": float(self.transparence_fond)})]
        if self.papier_peint is not None:
            objs["outspace"] = [objet({"color": couleur(self.papier_peint), "transparency": 0.0})]
        if objs:
            d["objects"] = objs
        if self.extraction:
            d["type"] = "Drillthrough"
        elif self.type_page:
            d["type"] = self.type_page
        if self.masquee:
            d["visibility"] = "HiddenInViewMode"
        if self.interactions:
            def nom_de(v):
                return noms[id(v)] if isinstance(v, Visuel) else v
            d["visualInteractions"] = [{"source": nom_de(s), "target": nom_de(c), "type": t}
                                       for s, c, t in self.interactions]
        return d


# ---------------------------------------------------------------------------
# Thème
# ---------------------------------------------------------------------------
def theme_personnalise(nom, couleurs_donnees, fond="#FFFFFF", texte="#1F2937", accent="#2A78D6",
                       bon="#1E7B34", neutre="#EDA100", mauvais="#B42318", police="Segoe UI",
                       couleur_titres="#1F2A44", fond_page="#F5F7FA", styles_visuels=None):
    """Thème personnalisé (fichier StaticResources/RegisteredResources/<nom>.json)."""
    h = hex_couleur
    t = {
        "name": nom,
        "dataColors": [h(c) for c in couleurs_donnees],
        "foreground": h(texte), "foregroundNeutralSecondary": "#6B7280", "foregroundNeutralTertiary": "#9CA3AF",
        "background": h(fond), "backgroundLight": "#F5F7FA", "backgroundNeutral": "#D6DCE4",
        "tableAccent": h(accent), "good": h(bon), "neutral": h(neutre), "bad": h(mauvais),
        "maximum": h(accent), "center": h(neutre), "minimum": "#EAF2FD", "null": "#9CA3AF",
        "hyperlink": h(accent), "visitedHyperlink": h(accent),
        "textClasses": {
            "callout": {"fontSize": 24, "fontFace": police, "color": h(couleur_titres)},
            "title": {"fontSize": 11, "fontFace": police, "color": h(couleur_titres)},
            "header": {"fontSize": 11, "fontFace": police, "color": h(couleur_titres)},
            "label": {"fontSize": 9, "fontFace": police, "color": h(texte)},
        },
        "visualStyles": {
            "*": {"*": {
                "title": [{"fontColor": {"solid": {"color": h(couleur_titres)}}, "fontSize": 11, "bold": True}],
                "background": [{"show": True, "transparency": 0}],
                "border": [{"show": False}],
                "dropShadow": [{"show": False}],
            }},
            "page": {"*": {
                "background": [{"color": {"solid": {"color": h(fond_page)}}, "transparency": 0}],
                "outspace": [{"color": {"solid": {"color": h(fond_page)}}}],
            }},
        },
    }
    for vt, st in (styles_visuels or {}).items():
        t["visualStyles"].setdefault(vt, {"*": {}})["*"].update(st)
    return t


# ---------------------------------------------------------------------------
# Rapport et projet
# ---------------------------------------------------------------------------
class Rapport:
    def __init__(self, nom, modele=None, theme=None, theme_base=THEME_BASE, volet_filtres_deplie=False,
                 volet_filtres_visible=True, alignement_vertical="Top"):
        """nom : nom du projet (dossiers <nom>.Report / <nom>.SemanticModel) ;
        modele : nom du dossier du modèle sémantique (par défaut <nom>.SemanticModel) ;
        theme : dict de thème personnalisé (theme_personnalise) ou None."""
        self.nom = nom
        self.modele = modele or (nom + ".SemanticModel")
        self.theme = theme
        self.theme_base = theme_base
        self.volet_filtres_deplie = volet_filtres_deplie
        self.volet_filtres_visible = volet_filtres_visible
        self.alignement_vertical = alignement_vertical
        self.pages = []
        self.filtres = []
        self.page_active = None

    def page(self, nom_affiche, **opts):
        p = Page(nom_affiche, **opts)
        if any(q.nom == p.nom for q in self.pages):
            raise ValueError("page en double : %s" % p.nom)
        self.pages.append(p)
        return p

    def _fichier_theme(self):
        return (self.theme["name"] + ".json") if self.theme else None

    def report_json(self):
        tc = {"baseTheme": {"name": self.theme_base, "reportVersionAtImport": VERSION_THEME_BASE,
                            "type": "SharedResources"}}
        paquets = [{"name": "SharedResources", "type": "SharedResources",
                    "items": [{"name": self.theme_base, "path": "BaseThemes/%s.json" % self.theme_base,
                               "type": "BaseTheme"}]}]
        if self.theme:
            ft = self._fichier_theme()
            tc["customTheme"] = {"name": ft, "reportVersionAtImport": VERSION_THEME_PERSO,
                                 "type": "RegisteredResources"}
            paquets.append({"name": "RegisteredResources", "type": "RegisteredResources",
                            "items": [{"name": ft, "path": ft, "type": "CustomTheme"}]})
        d = {"$schema": SCHEMAS["report"], "themeCollection": tc}
        if self.filtres:
            for i, f in enumerate(self.filtres):
                f.setdefault("ordinal", i)
            d["filterConfig"] = {"filters": self.filtres, "filterSortOrder": "Custom"}
        d["objects"] = {
            "section": [objet({"verticalAlignment": self.alignement_vertical})],
            "outspacePane": [objet({"expanded": bool(self.volet_filtres_deplie),
                                    "visible": bool(self.volet_filtres_visible)})],
        }
        d["resourcePackages"] = paquets
        d["settings"] = {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                         "allowChangeFilterTypes": True, "useEnhancedTooltips": True,
                         "useDefaultAggregateDisplayName": True}
        d["slowDataSourceSettings"] = {"isCrossHighlightingDisabled": False,
                                       "isSlicerSelectionsButtonEnabled": False,
                                       "isFilterSelectionsButtonEnabled": False,
                                       "isFieldWellButtonEnabled": False, "isApplyAllButtonEnabled": False}
        return d

    def ecrire(self, dossier_report):
        """Écrit le dossier <nom>.Report complet (dossier neuf ou vide)."""
        if not self.pages:
            raise ValueError("le rapport doit contenir au moins une page")
        _verifier_dossier_neuf(dossier_report)
        df = os.path.join(dossier_report, "definition")
        ecrire_json(os.path.join(dossier_report, ".platform"), platform_json("Report", self.nom))
        ecrire_json(os.path.join(dossier_report, "definition.pbir"),
                    {"$schema": SCHEMAS["pbir"], "version": VERSION_PBIR,
                     "datasetReference": {"byPath": {"path": "../" + self.modele}}})
        ecrire_json(os.path.join(df, "version.json"), {"$schema": SCHEMAS["version"], "version": VERSION_DEFINITION})
        ecrire_json(os.path.join(df, "report.json"), self.report_json())
        active = self.page_active or self.pages[0].nom
        ecrire_json(os.path.join(df, "pages", "pages.json"),
                    {"$schema": SCHEMAS["pages"], "pageOrder": [p.nom for p in self.pages], "activePageName": active})
        for p in self.pages:
            noms = p._noms_visuels()
            ecrire_json(os.path.join(df, "pages", p.nom, "page.json"), p.vers_json(noms))
            for v in p.visuels:
                n = noms[id(v)]
                ecrire_json(os.path.join(df, "pages", p.nom, "visuals", n, "visual.json"), v.vers_json(n))
        # thèmes
        src = os.path.join(DOSSIER_RESSOURCES, self.theme_base + ".json")
        dst = os.path.join(dossier_report, "StaticResources", "SharedResources", "BaseThemes", self.theme_base + ".json")
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copyfile(src, dst)
        if self.theme:
            ecrire_json(os.path.join(dossier_report, "StaticResources", "RegisteredResources", self._fichier_theme()),
                        self.theme)
        return dossier_report


def platform_json(type_element, nom_affiche):
    return {"$schema": SCHEMAS["platform"],
            "metadata": {"type": type_element, "displayName": nom_affiche},
            "config": {"version": "2.0", "logicalId": guid(type_element, nom_affiche)}}


def _verifier_dossier_neuf(d):
    if os.path.exists(d) and os.listdir(d):
        raise FileExistsError("le dossier existe et n'est pas vide (utiliser un dossier neuf) : %s" % d)


def ecrire_modele(dossier_modele, nom, model_bim, version_pbism=VERSION_PBISM):
    """Écrit <nom>.SemanticModel : .platform, definition.pbism, model.bim (dict TMSL)."""
    _verifier_dossier_neuf(dossier_modele)
    ecrire_json(os.path.join(dossier_modele, ".platform"), platform_json("SemanticModel", nom))
    ecrire_json(os.path.join(dossier_modele, "definition.pbism"),
                {"$schema": SCHEMAS["pbism"], "version": version_pbism, "settings": {}})
    ecrire_json(os.path.join(dossier_modele, "model.bim"), model_bim)
    return dossier_modele


def ecrire_projet(dossier, nom, rapport, model_bim=None, version_pbism=VERSION_PBISM, auto_recuperation=True):
    """Écrit un projet PBIP complet dans un dossier neuf :
    <nom>.pbip, .gitignore, <nom>.SemanticModel/ (si model_bim fourni), <nom>.Report/."""
    _verifier_dossier_neuf(dossier)
    os.makedirs(dossier, exist_ok=True)
    if rapport.modele != nom + ".SemanticModel" and model_bim is not None:
        raise ValueError("le rapport doit référencer %s.SemanticModel" % nom)
    ecrire_json(os.path.join(dossier, nom + ".pbip"),
                {"$schema": SCHEMAS["pbip"], "version": "1.0",
                 "artifacts": [{"report": {"path": nom + ".Report"}}],
                 "settings": {"enableAutoRecovery": bool(auto_recuperation)}})
    with open(os.path.join(dossier, ".gitignore"), "w", encoding="utf-8", newline="\n") as f:
        f.write("**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    if model_bim is not None:
        ecrire_modele(os.path.join(dossier, nom + ".SemanticModel"), nom, model_bim, version_pbism)
    rapport.ecrire(os.path.join(dossier, nom + ".Report"))
    return dossier


# ---------------------------------------------------------------------------
# Aperçu schématique (matplotlib)
# ---------------------------------------------------------------------------
_COULEURS_APERCU = {
    "shape": "#1F2A44", "textbox": "#D6DCE4", "card": "#EAF2FD", "multiRowCard": "#EAF2FD",
    "slicer": "#FFF1CC", "lineStackedColumnComboChart": "#86B6EF", "clusteredBarChart": "#8FB3E6",
    "barChart": "#8FB3E6", "columnChart": "#8FB3E6", "clusteredColumnChart": "#8FB3E6",
    "donutChart": "#7FD1AE", "tableEx": "#E3F4E8", "pivotTable": "#FDE2E1", "actionButton": "#2A78D6",
}


def _lire_rapport(dossier):
    """Lit pages et visuels d'un dossier .Report (ou d'un projet contenant un seul .Report)."""
    if not dossier.endswith(".Report"):
        cands = [os.path.join(dossier, d) for d in sorted(os.listdir(dossier)) if d.endswith(".Report")]
        if len(cands) != 1:
            raise ValueError("dossier .Report introuvable ou ambigu dans %s" % dossier)
        dossier = cands[0]
    dp = os.path.join(dossier, "definition", "pages")
    with open(os.path.join(dp, "pages.json"), encoding="utf-8") as f:
        ordre = json.load(f).get("pageOrder", [])
    pages = []
    for pn in ordre:
        with open(os.path.join(dp, pn, "page.json"), encoding="utf-8") as f:
            page = json.load(f)
        vis = []
        dv = os.path.join(dp, pn, "visuals")
        for vn in sorted(os.listdir(dv)) if os.path.isdir(dv) else []:
            with open(os.path.join(dv, vn, "visual.json"), encoding="utf-8") as f:
                vis.append(json.load(f))
        pages.append((page, vis))
    return pages


def _libelle_visuel(v):
    vi = v.get("visual", {})
    t = vi.get("visualType", "groupe")
    titre = None
    for e in vi.get("visualContainerObjects", {}).get("title", []):
        tx = e.get("properties", {}).get("text")
        if tx and "Literal" in tx.get("expr", {}):
            titre = tx["expr"]["Literal"]["Value"].strip("'").replace("''", "'")
    if not titre and t == "textbox":
        paras = vi.get("objects", {}).get("general", [{}])[0].get("properties", {}).get("paragraphs", [])
        titre = " ".join(r.get("value", "") for p in paras for r in p.get("textRuns", []))
    if not titre:
        for e in vi.get("objects", {}).get("header", []) + vi.get("objects", {}).get("text", []):
            tx = e.get("properties", {}).get("text")
            if tx and "Literal" in tx.get("expr", {}):
                titre = tx["expr"]["Literal"]["Value"].strip("'").replace("''", "'")
    if not titre:
        champs = [p.get("nativeQueryRef") or p.get("queryRef")
                  for r in vi.get("query", {}).get("queryState", {}).values() for p in r.get("projections", [])]
        titre = ", ".join(champs[:2])
    return t, (titre or "")[:48]


def preview_pages(dossier, sortie):
    """Aperçu PNG schématique de toutes les pages (rectangles + type et titre de chaque visuel)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle

    pages = _lire_rapport(dossier)
    largeur_fig = 13
    hauteurs = [p["height"] / p["width"] * largeur_fig + 0.6 for p, _ in pages]
    fig, axes = plt.subplots(len(pages), 1, figsize=(largeur_fig, sum(hauteurs)),
                             gridspec_kw={"height_ratios": hauteurs})
    if len(pages) == 1:
        axes = [axes]
    for ax, (page, vis) in zip(axes, pages):
        W, H = page["width"], page["height"]
        ax.add_patch(Rectangle((0, 0), W, H, facecolor="#F5F7FA", edgecolor="#9CA3AF"))
        for v in sorted(vis, key=lambda d: d["position"].get("z", 0)):
            pos = v["position"]
            t, titre = _libelle_visuel(v)
            c = _COULEURS_APERCU.get(t, "#E5E7EB")
            ax.add_patch(Rectangle((pos["x"], pos["y"]), pos["width"], pos["height"], facecolor=c,
                                   edgecolor="#374151", linewidth=0.8, alpha=0.85))
            fonce = t in ("shape", "actionButton")
            ax.text(pos["x"] + 4, pos["y"] + 4, "%s\n%s" % (t, titre), fontsize=6.5, va="top", ha="left",
                    color="white" if fonce else "#111827", clip_on=True)
        ax.set_xlim(-5, W + 5)
        ax.set_ylim(H + 5, -5)
        ax.set_aspect("equal")
        ax.set_xticks([])
        ax.set_yticks([])
        etat = []
        if page.get("type") == "Drillthrough":
            etat.append("extraction")
        if page.get("visibility") == "HiddenInViewMode":
            etat.append("masquée")
        ax.set_title("%s  (%s, %d x %d%s)" % (page["displayName"], page["name"], W, H,
                                               (", " + ", ".join(etat)) if etat else ""), fontsize=10)
    fig.tight_layout()
    os.makedirs(os.path.dirname(os.path.abspath(sortie)), exist_ok=True)
    fig.savefig(sortie, dpi=110)
    plt.close(fig)
    return sortie

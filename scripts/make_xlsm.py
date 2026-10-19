# -*- coding: utf-8 -*-
"""
make_xlsm.py — transforme un classeur .xlsx (produit par openpyxl) en classeur .xlsm
contenant un projet VBA (xl/vbaProject.bin) et des boutons affectés à des macros.

Uniquement la bibliothèque standard de Python (zipfile, xml.parsers.expat, struct, re).

Principe : chirurgie minimale au niveau des octets.
    Les parties XML existantes ne sont JAMAIS ré-sérialisées : un analyseur XML conforme
    (expat) sert seulement à repérer la position exacte (en octets) des éléments, puis on
    insère ou remplace quelques octets à ces positions. Tout le reste du classeur (valeurs
    en cache, validations, mises en forme conditionnelles, noms définis, zones d'impression,
    sauts de page, graphiques...) est donc recopié octet pour octet.

Ce que fait make_xlsm() (références : ECMA-376 1re partie / ISO-IEC 29500-1, sml.xsd,
dml-spreadsheetDrawing.xsd, [MS-OI29500], [MS-OVBA], [MS-CFB], et le comportement de
XlsxWriter dont les sorties sont comparées à des fichiers Excel dans sa suite de tests) :

  1. ajoute la partie xl/vbaProject.bin (octets fournis, copiés tels quels) ;
  2. ajoute la relation  http://schemas.microsoft.com/office/2006/relationships/vbaProject
     dans xl/_rels/workbook.xml.rels (cible relative « vbaProject.bin », comme Excel) ;
  3. [Content_Types].xml : type du classeur -> application/vnd.ms-excel.sheet.macroEnabled.main+xml
     et Default Extension="bin" -> application/vnd.ms-office.vbaProject (Override si « bin »
     est déjà pris par un autre type, p. ex. printerSettings) ;
  4. ajoute codeName="..." à <workbookPr> (créé à la bonne place s'il est absent) et au
     <sheetPr> de chaque feuille (créé en tête de <worksheet> s'il est absent : sml.xsd impose
     sheetPr comme tout premier enfant de CT_Worksheet) ;
  5. insère les boutons : une forme DrawingML <xdr:sp> (rectangle arrondi, texte centré)
     portant l'attribut macro="[0]!NomMacro". La forme est ajoutée au dessin existant de la
     feuille (à côté des graphiques), ou dans un nouveau dessin (xl/drawings/drawingN.xml +
     relation + type de contenu + élément <drawing> placé selon l'ordre de CT_Worksheet).

Pourquoi une forme DrawingML plutôt qu'un bouton de formulaire VML (insert_button de XlsxWriter) ?
  - Un bouton de formulaire (VML, <x:ClientData ObjectType="Button"> + <x:FmlaMacro>) est
    toujours peint par Excel avec la couleur système « buttonFace » (gris) : son fond ne peut
    pas être bleu foncé 1F2A44. Seule une forme (ou un contrôle ActiveX, à proscrire sur un
    poste d'entreprise) accepte un fond et une police au choix.
  - VML est le format hérité (ECMA-376 4e partie) : il faudrait ajouter une 2e partie de dessin
    (vmlDrawingN.vml + <legacyDrawing> + Default « vml ») à côté du dessin DrawingML existant,
    avec ses propres numérotations (o:idmap, _x0000_sNNNN) : davantage de points de rupture.
  - CT_Shape (dml-spreadsheetDrawing.xsd) définit l'attribut macro (xsd:string). Valeur écrite :
    « [0]!NomMacro », où [0] désigne le classeur lui-même :
      * c'est la référence qu'écrit XlsxWriter pour un bouton (vml.py : "[0]!" + macro dans
        <x:FmlaMacro>), format que sa suite de tests compare à des fichiers créés par Excel ;
      * la note [MS-OI29500] 2.1.1609 (Part 1 §21.3.2.22) indique qu'Excel attend pour macro
        la grammaire des références de nom (préfixe de classeur "[n]!" admis), 256 caractères
        au plus ;
      * l'import OOXML de LibreOffice (écrit pour lire les fichiers d'Excel) résout
        « [0]!say_hello » et « [0]!Module1.say_hello » vers la macro du projet VBA (testé).
  - [MS-OI29500] : si textlink est présent, txBody est obligatoire -> toujours présent ici.

Coordonnées des boutons (indices 0-based, exactement la sémantique de <xdr:from>/<xdr:to>) :
    from_col/from_row = cellule du coin haut-gauche, to_col/to_row = cellule où se trouve le
    coin bas-droit avec un décalage nul, c.-à-d. que le bouton couvre les colonnes
    from_col .. to_col-1 et les lignes from_row .. to_row-1. Décalages optionnels en EMU :
    from_col_off, from_row_off, to_col_off, to_row_off (1 pixel = 9525 EMU, 1 pt = 12700 EMU).
    Plus simple : "range": "BL2:BM4" à la place des quatre indices (ou bouton_plage(...)).
    inset_px (facultatif) : marge intérieure en pixels sur les quatre côtés ; le coin bas-droit
    est alors reporté dans la dernière cellule (to_col-1, largeur - marge), d'après les largeurs
    de colonnes et hauteurs de lignes de la feuille (police Normal Calibri 11 : 7 px/chiffre).
Options de style par bouton (facultatives) : fill ("1F2A44"), font_color ("FFFFFF"),
    bold (True), font_size (10, en points), font_name (None = police du thème), name (nom de
    la forme), edit_as ("oneCell" = déplacer sans dimensionner avec les cellules).

Feuilles sans module document dans le projet VBA : chaque feuille reçoit quand même son
    codeName ; à l'ouverture, Excel crée le module document (vide) qui manque — c'est aussi
    ce que fait XlsxWriter, qui donne codeName="SheetN" à toutes les feuilles dès qu'un
    projet VBA est ajouté. make_xlsm() le signale par un avertissement.

Utilisation en ligne de commande :
    python make_xlsm.py source.xlsx cible.xlsm vbaProject.bin config.json
    avec config.json = {"codenames": {...}, "buttons": [...]}
"""

import json
import posixpath
import re
import struct
import sys
import warnings
import xml.parsers.expat
import zipfile
from xml.sax.saxutils import escape as _xml_escape

__all__ = ["make_xlsm", "lire_projet_vba", "bouton_plage", "plage_vers_indices"]

# --------------------------------------------------------------------------------------
# Constantes (espaces de noms, types de contenu, types de relation)
# --------------------------------------------------------------------------------------
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

CT_MAIN_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_MAIN_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
CT_VBA = "application/vnd.ms-office.vbaProject"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"

REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

# Enfants de CT_Worksheet (sml.xsd) qui PRÉCÈDENT <drawing> : un nouvel élément <drawing>
# est inséré juste après le dernier d'entre eux présent dans la feuille (sheetData est
# obligatoire, il y en a donc toujours un). Ce qui suit (legacyDrawing, controls souvent
# enveloppé dans mc:AlternateContent, tableParts, extLst...) reste après.
AVANT_DRAWING = ("sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
                 "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
                 "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
                 "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
                 "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
                 "customProperties", "cellWatches", "ignoredErrors", "smartTags")
# Enfants de CT_Workbook qui précèdent <workbookPr>.
AVANT_WORKBOOKPR = ("fileVersion", "fileSharing")

EMU_PAR_PIXEL = 9525          # 914400 EMU par pouce / 96 pixels par pouce
EMU_PAR_POINT = 12700         # 914400 / 72
LARGEUR_CHIFFRE_MAX = 7       # largeur (px) du chiffre le plus large en Calibri 11 (police
                              # Normal des classeurs openpyxl) : sert aux marges inset_px et
                              # à <a:xfrm> (informatif : Excel place une forme twoCellAnchor
                              # d'après <xdr:from>/<xdr:to>)
MAX_COL, MAX_ROW = 16384, 1048576

# Identifiant VBA : lettre puis lettres/chiffres/_ ; « Module.Macro » accepté pour les macros.
RE_IDENT = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")
RE_MACRO = re.compile(r"^[A-Za-z][A-Za-z0-9_]*(\.[A-Za-z][A-Za-z0-9_]*)?$")
RE_COULEUR = re.compile(r"^[0-9A-Fa-f]{6}$")
RE_CELLULE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


# --------------------------------------------------------------------------------------
# Utilitaires : bouton défini par une plage de cellules
# --------------------------------------------------------------------------------------
def _cellule_vers_indices(ref):
    """'BL2' -> (63, 1) : indices 0-based (colonne, ligne)."""
    m = RE_CELLULE.match(ref.strip())
    if not m:
        raise ValueError("référence de cellule invalide : %r" % ref)
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    row = int(m.group(2))
    if not (1 <= col <= MAX_COL and 1 <= row <= MAX_ROW):
        raise ValueError("cellule hors de la feuille : %r" % ref)
    return col - 1, row - 1


def plage_vers_indices(plage):
    """'BL2:BM4' -> (from_col, from_row, to_col, to_row) = (63, 1, 65, 4).

    Les bornes de fin sont EXCLUSIVES (sémantique de <xdr:to> avec décalage nul) : le
    bouton couvre exactement les cellules de la plage. Une cellule seule ('BL2') est admise."""
    morceaux = plage.split(":")
    if len(morceaux) not in (1, 2):
        raise ValueError("plage invalide : %r" % plage)
    c1, r1 = _cellule_vers_indices(morceaux[0])
    c2, r2 = _cellule_vers_indices(morceaux[-1])
    if c2 < c1 or r2 < r1:
        raise ValueError("plage %r : coin haut-gauche attendu en premier" % plage)
    return c1, r1, c2 + 1, r2 + 1


def bouton_plage(sheet, plage, text, macro, inset_px=0, **options):
    """Décrit un bouton couvrant la plage « plage » (ex. "BL2:BM4") de la feuille « sheet ».

    inset_px : marge intérieure (pixels) entre le bouton et les bords de la plage.
    options  : fill, font_color, bold, font_size, font_name, name, edit_as.
    Retourne le dict attendu par make_xlsm(..., buttons=[...])."""
    fc, fr, tc, tr = plage_vers_indices(plage)
    b = {"sheet": sheet, "text": text, "macro": macro,
         "from_col": fc, "from_row": fr, "to_col": tc, "to_row": tr}
    if inset_px:
        b["inset_px"] = inset_px
    b.update(options)
    return b


# --------------------------------------------------------------------------------------
# Repérage des éléments XML par position d'octets (sans ré-sérialisation)
# --------------------------------------------------------------------------------------
class _Elem:
    """Élément repéré dans un document XML : nom qualifié, attributs, positions en octets.

    start     : position du « < » de la balise ouvrante
    stag_end  : position juste après le « > » de la balise ouvrante
    etag      : position du « < » de la balise fermante (None si élément vide <x/>)
    end       : position juste après la fin de l'élément
    """
    __slots__ = ("ns", "local", "qname", "attrs", "start", "stag_end", "etag", "end",
                 "depth", "children")

    def __init__(self):
        self.children = []

    @property
    def prefix(self):
        return self.qname.split(":", 1)[0] if ":" in self.qname else ""

    def attr(self, local, ns=""):
        """Valeur d'un attribut (ns = URI de l'espace de noms, '' pour sans préfixe)."""
        return self.attrs.get((ns + " " + local) if ns else local)


def _verifier_encodage(data, nom):
    """Les positions d'octets et les fragments insérés supposent de l'UTF-8 (encodage de
    toutes les parties écrites par Excel et openpyxl) : tout autre encodage est refusé."""
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        raise ValueError("%s : partie XML en UTF-16, non gérée" % nom)
    debut = data[3:] if data[:3] == b"\xef\xbb\xbf" else data
    m = re.match(rb"<\?xml[^>]*?encoding\s*=\s*[\"']([A-Za-z0-9._-]+)[\"']", debut)
    if m and m.group(1).lower() not in (b"utf-8", b"utf8"):
        raise ValueError("%s : encodage %s non géré (UTF-8 attendu)" % (nom, m.group(1).decode()))


def _fin_balise_ouvrante(data, pos):
    """Position juste après le « > » de la balise commençant en pos (guillemets gérés,
    car « > » est autorisé tel quel dans une valeur d'attribut)."""
    quote = None
    n = len(data)
    i = pos + 1
    while i < n:
        c = data[i]
        if quote is not None:
            if c == quote:
                quote = None
        elif c in (0x22, 0x27):          # " ou '
            quote = c
        elif c == 0x3E:                  # >
            return i + 1
        i += 1
    raise ValueError("balise non terminée à la position %d" % pos)


def _indexer(data, max_depth=None, nom="?"):
    """Analyse « data » (octets) avec expat et retourne l'élément racine, ses descendants
    jusqu'à la profondeur max_depth (racine = 0 ; None = tout) étant rangés dans .children.
    Lève une erreur si le document n'est pas bien formé."""
    _verifier_encodage(data, nom)
    p = xml.parsers.expat.ParserCreate(namespace_separator=" ")
    pile = []
    racine = []

    def debut(nom_el, attrs):
        depth = len(pile)
        if max_depth is not None and depth > max_depth:
            pile.append(None)
            return
        e = _Elem()
        e.start = p.CurrentByteIndex
        if data[e.start:e.start + 1] != b"<":
            raise ValueError("position de balise ouvrante incohérente (%d)" % e.start)
        e.stag_end = _fin_balise_ouvrante(data, e.start)
        e.qname = re.match(rb"<([^\s/>]+)", data[e.start:e.start + 300]).group(1).decode("utf-8")
        if " " in nom_el:
            e.ns, e.local = nom_el.split(" ", 1)
        else:
            e.ns, e.local = "", nom_el
        e.attrs = attrs
        e.depth = depth
        if pile:
            parent = pile[-1]
            if parent is not None:
                parent.children.append(e)
        else:
            racine.append(e)
        pile.append(e)

    def fin(nom_el):
        e = pile.pop()
        if e is None:
            return
        if data[e.stag_end - 2:e.stag_end] == b"/>":
            e.etag = None
            e.end = e.stag_end
        else:
            idx = p.CurrentByteIndex
            if data[idx:idx + 2] != b"</":
                raise ValueError("position de balise fermante incohérente (%d)" % idx)
            e.etag = idx
            e.end = data.index(b">", idx) + 1

    p.StartElementHandler = debut
    p.EndElementHandler = fin
    try:
        p.Parse(data, True)
    except xml.parsers.expat.ExpatError as exc:
        raise ValueError("%s : XML mal formé (%s)" % (nom, exc))
    return racine[0]


def _iter(e):
    """Parcours en profondeur de l'arbre repéré."""
    yield e
    for c in e.children:
        yield from _iter(c)


def _remplacer(data, edits):
    """Applique des éditions [(début, fin, nouveaux_octets)] sans recouvrement."""
    edits = sorted(edits, key=lambda t: (t[0], t[1]))
    out, pos = [], 0
    for a, b, s in edits:
        if a < pos:
            raise ValueError("éditions qui se recouvrent")
        out.append(data[pos:a])
        out.append(s)
        pos = b
    out.append(data[pos:])
    return b"".join(out)


# Un attribut dans une balise ouvrante : blancs, nom, « = », valeur entre guillemets.
_RE_ATTRIBUT = re.compile(rb"""\s+([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')""")


def _attributs_bruts(data, e):
    """[(nom qualifié, début de la valeur, fin de la valeur)] — positions guillemets inclus.
    Les attributs sont lus l'un après l'autre depuis le nom de l'élément : un texte
    ressemblant à un attribut à l'intérieur d'une valeur ne peut donc pas être confondu."""
    pos = e.start + 1 + len(e.qname.encode("utf-8"))
    out = []
    while True:
        m = _RE_ATTRIBUT.match(data, pos, e.stag_end)
        if not m:
            break
        out.append((m.group(1).decode("utf-8"), m.start(2), m.end(2)))
        pos = m.end()
    return out


def _edit_attribut(data, e, nom, valeur):
    """Édition qui pose nom="valeur" sur la balise ouvrante de e (attribut sans préfixe) :
    remplace la valeur si l'attribut existe, sinon l'insère juste après le nom de l'élément."""
    v = _xml_escape(valeur, {'"': "&quot;"}).encode("utf-8")
    for qn, a, b in _attributs_bruts(data, e):
        if qn == nom:
            return (a, b, b'"' + v + b'"')
    if e.attr(nom) is not None:          # garde-fou : expat le voit mais pas le découpage
        raise ValueError("attribut %s introuvable dans la balise <%s>" % (nom, e.qname))
    pos = e.start + 1 + len(e.qname.encode("utf-8"))
    return (pos, pos, b" " + nom.encode() + b'="' + v + b'"')


def _edit_ajout_enfant_fin(data, e, fragment):
    """Édition qui ajoute « fragment » comme dernier enfant de e (gère <x/> -> <x>...</x>)."""
    if e.etag is not None:
        return (e.etag, e.etag, fragment)
    # élément vide : on remplace « /> » par « > fragment </qname> »
    return (e.stag_end - 2, e.stag_end, b">" + fragment + b"</" + e.qname.encode() + b">")


# --------------------------------------------------------------------------------------
# Relations OPC
# --------------------------------------------------------------------------------------
def _chemin_rels(part):
    """xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels"""
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _resoudre_cible(part_source, cible):
    """Résout la cible (relative ou absolue) d'une relation en nom de partie sans « / » initial."""
    if cible.startswith("/"):
        return posixpath.normpath(cible.lstrip("/"))
    base = posixpath.dirname(part_source)
    return posixpath.normpath(posixpath.join(base, cible))


def _lire_rels(data, nom="?"):
    """Retourne (racine repérée, liste de dicts {Id, Type, Target, TargetMode})."""
    racine = _indexer(data, 1, nom)
    rels = []
    for c in racine.children:
        if c.ns == NS_PKG_REL and c.local == "Relationship":
            rels.append({k: c.attrs.get(k) for k in ("Id", "Type", "Target", "TargetMode")})
    return racine, rels


def _nouvel_id(rels):
    """Premier identifiant « rIdN » libre (N > tous les rIdN existants)."""
    n = 0
    for r in rels:
        m = re.match(r"^rId(\d+)$", r["Id"] or "")
        if m:
            n = max(n, int(m.group(1)))
    existants = {r["Id"] for r in rels}
    while True:
        n += 1
        if "rId%d" % n not in existants:
            return "rId%d" % n


def _xml_relation(rid, rtype, cible):
    return ('<Relationship Id="%s" Type="%s" Target="%s"/>'
            % (rid, rtype, _xml_escape(cible, {'"': "&quot;"}))).encode("utf-8")


def _ajouter_relation(data, rtype, cible, nom="?"):
    """Ajoute une relation dans un fichier .rels existant (octets) ; retourne (octets, Id)."""
    racine, rels = _lire_rels(data, nom)
    rid = _nouvel_id(rels)
    frag = _xml_relation(rid, rtype, cible)
    return _remplacer(data, [_edit_ajout_enfant_fin(data, racine, frag)]), rid


def _nouveau_rels(rtype, cible, rid="rId1"):
    return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<Relationships xmlns="' + NS_PKG_REL.encode() + b'">'
            + _xml_relation(rid, rtype, cible) + b"</Relationships>")


def _cible_relative(part_source, part_cible):
    """Cible relative d'une relation (style Excel : « ../drawings/drawing3.xml »)."""
    return posixpath.relpath(part_cible, posixpath.dirname(part_source))


# --------------------------------------------------------------------------------------
# Lecture minimale du fichier binaire composé (CFB, [MS-CFB]) pour le contrôle de cohérence
# --------------------------------------------------------------------------------------
_CFB_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
_FIN_CHAINE = 0xFFFFFFFE
_LIBRE = 0xFFFFFFFF


def _cfb_lire_flux(data, nom_flux):
    """Retourne le contenu du flux « nom_flux » situé à la racine d'un fichier CFB, ou None.
    Lecture seule et volontairement minimale ([MS-CFB] §2.2 en-tête, §2.3 FAT, §2.6 répertoire,
    §2.5 mini-flux)."""
    if len(data) < 512 or data[:8] != _CFB_SIGNATURE:
        raise ValueError("vbaProject.bin n'est pas un fichier binaire composé (signature CFB absente)")
    (major,) = struct.unpack_from("<H", data, 0x1A)
    (sect_shift, mini_shift) = struct.unpack_from("<HH", data, 0x1E)
    if (major, sect_shift) not in ((3, 9), (4, 12)) or mini_shift != 6:
        raise ValueError("en-tête CFB inattendu (version %d, secteurs 2^%d)" % (major, sect_shift))
    taille_sect = 1 << sect_shift
    (nb_fat, dir_start, _tx, mini_cutoff, minifat_start, nb_minifat, difat_start, nb_difat) = \
        struct.unpack_from("<IIIIIIII", data, 0x2C)

    def secteur(n):
        off = (n + 1) * taille_sect
        if off >= len(data):                 # un dernier secteur tronqué est toléré
            raise ValueError("secteur CFB hors du fichier : %d" % n)
        return data[off:off + taille_sect]

    # DIFAT : 109 entrées dans l'en-tête puis chaîne de secteurs DIFAT
    difat = list(struct.unpack_from("<109I", data, 0x4C))
    s, garde = difat_start, 0
    while s not in (_FIN_CHAINE, _LIBRE) and garde < nb_difat + 1:
        vals = struct.unpack("<%dI" % (taille_sect // 4), secteur(s))
        difat.extend(vals[:-1])
        s = vals[-1]
        garde += 1
    fat = []
    for s in difat[:nb_fat]:
        fat.extend(struct.unpack("<%dI" % (taille_sect // 4), secteur(s).ljust(taille_sect, b"\0")))

    def chaine(debut):
        out, s, vus = [], debut, set()
        while s not in (_FIN_CHAINE, _LIBRE):
            if s in vus or s >= len(fat):
                raise ValueError("chaîne FAT invalide")
            vus.add(s)
            out.append(s)
            s = fat[s]
        return out

    repertoire = b"".join(secteur(s) for s in chaine(dir_start))
    entrees = []
    for i in range(len(repertoire) // 128):
        e = repertoire[i * 128:(i + 1) * 128]
        (lg,) = struct.unpack_from("<H", e, 64)
        nom = e[:max(min(lg, 64) - 2, 0)].decode("utf-16-le", "replace")
        typ = e[66]
        gauche, droite, enfant = struct.unpack_from("<III", e, 68)
        (debut, taille) = struct.unpack_from("<IQ", e, 116)
        if major == 3:
            taille &= 0xFFFFFFFF
        entrees.append((nom, typ, gauche, droite, enfant, debut, taille))
    racine = entrees[0]
    if racine[1] != 5:
        raise ValueError("entrée racine CFB absente")

    # enfants directs de la racine : parcours de l'arbre rouge-noir (frères gauche/droite)
    acc, a_voir, vus = [], [racine[4]], set()
    while a_voir:
        idx = a_voir.pop()
        if idx == _LIBRE or idx >= len(entrees) or idx in vus:
            continue
        vus.add(idx)
        acc.append(idx)
        a_voir.extend((entrees[idx][2], entrees[idx][3]))
    cible = None
    for idx in acc:
        if entrees[idx][1] == 2 and entrees[idx][0].upper() == nom_flux.upper():
            cible = entrees[idx]
    if cible is None:
        return None
    _, _, _, _, _, debut, taille = cible
    if taille < mini_cutoff:
        mini_flux = b"".join(secteur(s) for s in chaine(racine[5]))
        minifat = []
        for s in chaine(minifat_start):
            minifat.extend(struct.unpack("<%dI" % (taille_sect // 4), secteur(s)))
        out, s, vus = [], debut, set()
        while s not in (_FIN_CHAINE, _LIBRE):
            if s in vus or s >= len(minifat):
                raise ValueError("chaîne mini-FAT invalide")
            vus.add(s)
            out.append(mini_flux[s * 64:(s + 1) * 64])
            s = minifat[s]
        return b"".join(out)[:taille]
    return b"".join(secteur(s) for s in chaine(debut))[:taille]


def lire_projet_vba(vba_bin_bytes):
    """Lit le flux PROJECT ([MS-OVBA] §2.3.1, texte non compressé, page de codes du projet)
    et retourne {"documents": [...], "modules": [...], "classes": [...], "nom": ...}."""
    flux = _cfb_lire_flux(vba_bin_bytes, "PROJECT")
    if flux is None:
        raise ValueError("vbaProject.bin ne contient pas de flux PROJECT : ce n'est pas un projet VBA")
    try:
        texte = flux.decode("cp1252")    # page de codes des postes Windows en français
    except UnicodeDecodeError:
        texte = flux.decode("latin-1")
    info = {"documents": [], "modules": [], "classes": [], "nom": None}
    for ligne in texte.split("\r\n"):
        if ligne.startswith("["):        # début des sections [Host Extender Info] / [Workspace]
            break
        cle, _, val = ligne.partition("=")
        if cle == "Document":
            info["documents"].append(val.split("/", 1)[0])
        elif cle == "Module":
            info["modules"].append(val)
        elif cle in ("Class", "BaseClass"):
            info["classes"].append(val)
        elif cle == "Name":
            info["nom"] = val.strip('"')
    return info


# --------------------------------------------------------------------------------------
# Géométrie : largeur des colonnes (px) et hauteur des lignes (pt) d'une feuille
# --------------------------------------------------------------------------------------
def _px_largeur(w):
    """Largeur de colonne (en caractères, attribut width) -> pixels (ECMA-376 §18.3.1.13)."""
    return int(((256 * w + int(128 / LARGEUR_CHIFFRE_MAX)) / 256) * LARGEUR_CHIFFRE_MAX)


def _metrique_feuille(sheet_racine):
    """Retourne (fonction px_colonne(index0), fonction pt_ligne(index0))."""
    largeurs, hauteurs = {}, {}
    defaut_px, defaut_pt = None, 15.0
    for c in sheet_racine.children:
        if c.ns != NS_MAIN:
            continue
        if c.local == "sheetFormatPr":
            if c.attr("defaultRowHeight"):
                defaut_pt = float(c.attr("defaultRowHeight"))
            if c.attr("defaultColWidth"):
                defaut_px = _px_largeur(float(c.attr("defaultColWidth")))
            else:
                base = float(c.attr("baseColWidth") or 8)
                # ECMA-376 §18.3.1.81 : base*MDW + 5 px de marge, arrondi au multiple de 8
                defaut_px = -(-int(base * LARGEUR_CHIFFRE_MAX + 5) // 8) * 8
        elif c.local == "cols":
            for col in c.children:
                if col.local != "col":
                    continue
                cache = col.attr("hidden") in ("1", "true")
                w = float(col.attr("width") or 0)
                px = 0 if cache else _px_largeur(w)
                for i in range(int(col.attr("min")), int(col.attr("max")) + 1):
                    largeurs[i - 1] = px
        elif c.local == "sheetData":
            for row in c.children:
                if row.local != "row" or not row.attr("r"):
                    continue
                r = int(row.attr("r")) - 1
                if row.attr("hidden") in ("1", "true"):
                    hauteurs[r] = 0.0
                elif row.attr("ht"):
                    hauteurs[r] = float(row.attr("ht"))
    if defaut_px is None:
        defaut_px = 64
    return (lambda i: largeurs.get(i, defaut_px)), (lambda i: hauteurs.get(i, defaut_pt))


def _police_normale(parts, wb_part, wb_rels):
    """(nom, taille) de la première police de styles.xml (police du style Normal, qui fixe la
    largeur des colonnes en pixels), ou None si elle est introuvable."""
    rel = [r for r in wb_rels if r["Type"] == REL_STYLES]
    if not rel:
        return None
    sp = _resoudre_cible(wb_part, rel[0]["Target"])
    if sp not in parts:
        return None
    for c in _indexer(parts[sp], 3, sp).children:
        if c.ns == NS_MAIN and c.local == "fonts" and c.children:
            vals = {e.local: e.attr("val") for e in c.children[0].children}
            try:
                return vals.get("name"), float(vals.get("sz") or 0)
            except ValueError:
                return vals.get("name"), None
    return None


def _normaliser_point(index, off, taille_emu, maximum):
    """Ramène un point (cellule, décalage EMU) à un décalage compris dans la cellule :
    reporte le surplus sur les cellules suivantes, ou un décalage négatif sur les
    précédentes. taille_emu(i) = taille de la cellule i en EMU."""
    while off < 0:
        index -= 1
        if index < 0:
            raise ValueError("ancre avant la première cellule")
        off += taille_emu(index)
    while index < maximum and off > 0 and off >= taille_emu(index):
        off -= taille_emu(index)
        index += 1
    return index, off


def _resoudre_ancres(b, px_col, pt_row):
    """Applique inset_px puis normalise les quatre points d'ancrage du bouton b (dict
    normalisé) ; retourne un nouveau dict."""
    n = dict(b)
    w = lambda i: px_col(i) * EMU_PAR_PIXEL
    h = lambda i: int(round(pt_row(i) * EMU_PAR_POINT))
    m = n["inset_px"] * EMU_PAR_PIXEL
    erreur = ValueError("bouton %r : zone trop petite pour la marge inset_px=%d"
                        % (b["text"], b["inset_px"]))
    try:
        n["from_col"], n["from_col_off"] = _normaliser_point(n["from_col"], n["from_col_off"] + m, w, MAX_COL - 1)
        n["from_row"], n["from_row_off"] = _normaliser_point(n["from_row"], n["from_row_off"] + m, h, MAX_ROW - 1)
        n["to_col"], n["to_col_off"] = _normaliser_point(n["to_col"], n["to_col_off"] - m, w, MAX_COL - 1)
        n["to_row"], n["to_row_off"] = _normaliser_point(n["to_row"], n["to_row_off"] - m, h, MAX_ROW - 1)
    except ValueError:
        raise erreur
    if (n["to_col"], n["to_col_off"]) <= (n["from_col"], n["from_col_off"]) or \
       (n["to_row"], n["to_row_off"]) <= (n["from_row"], n["from_row_off"]):
        raise erreur
    # position et taille absolues (EMU) pour <a:xfrm>
    x1 = sum(w(i) for i in range(n["from_col"])) + n["from_col_off"]
    x2 = sum(w(i) for i in range(n["to_col"])) + n["to_col_off"]
    y1 = sum(h(i) for i in range(n["from_row"])) + n["from_row_off"]
    y2 = sum(h(i) for i in range(n["to_row"])) + n["to_row_off"]
    n["xfrm"] = (x1, y1, max(x2 - x1, 0), max(y2 - y1, 0))
    return n


# --------------------------------------------------------------------------------------
# XML d'un bouton (forme DrawingML avec macro)
# --------------------------------------------------------------------------------------
def _xml_bouton(p, shape_id, b):
    """Retourne le fragment <twoCellAnchor> du bouton (b : dict normalisé et résolu).

    p : préfixe lié à l'espace de noms spreadsheetDrawing là où le fragment est inséré
        ('' = espace par défaut, comme dans les dessins écrits par openpyxl). L'espace de
        noms DrawingML principal (préfixe a) est déclaré sur <sp> : le fragment ne dépend
        donc d'aucune déclaration de la racine autre que celle du préfixe p."""
    q = (p + ":") if p else ""
    esc = lambda s: _xml_escape(s, {'"': "&quot;"})
    x, y, cx, cy = b["xfrm"]
    police = ('<a:latin typeface="%s"/>' % esc(b["font_name"])) if b["font_name"] else ""
    paragraphes = "".join(
        '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="fr-FR" sz="%d" b="%d"><a:solidFill>'
        '<a:srgbClr val="%s"/></a:solidFill>%s</a:rPr><a:t>%s</a:t></a:r></a:p>'
        % (int(round(b["font_size"] * 100)), 1 if b["bold"] else 0, b["font_color"], police,
           esc(ligne))
        for ligne in b["text"].split("\n"))
    return (
        '<{q}twoCellAnchor editAs="{edit_as}">'
        '<{q}from><{q}col>{fc}</{q}col><{q}colOff>{fco}</{q}colOff><{q}row>{fr}</{q}row>'
        '<{q}rowOff>{fro}</{q}rowOff></{q}from>'
        '<{q}to><{q}col>{tc}</{q}col><{q}colOff>{tco}</{q}colOff><{q}row>{tr}</{q}row>'
        '<{q}rowOff>{tro}</{q}rowOff></{q}to>'
        '<{q}sp xmlns:a="{ns_a}" macro="{macro}" textlink="">'
        '<{q}nvSpPr><{q}cNvPr id="{sid}" name="{nom}" descr="{descr}"/><{q}cNvSpPr/></{q}nvSpPr>'
        '<{q}spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        '<a:ln w="9525"><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:ln>'
        '</{q}spPr>'
        '<{q}txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="45720" '
        'tIns="0" rIns="45720" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/>{paras}</{q}txBody>'
        '</{q}sp>'
        '<{q}clientData fPrintsWithSheet="0"/>'
        '</{q}twoCellAnchor>'
    ).format(q=q, edit_as=b["edit_as"], fc=b["from_col"], fco=b["from_col_off"],
             fr=b["from_row"], fro=b["from_row_off"], tc=b["to_col"], tco=b["to_col_off"],
             tr=b["to_row"], tro=b["to_row_off"], ns_a=NS_A, macro=esc(b["macro_attr"]),
             sid=shape_id, nom=esc(b["name"]), descr=esc(b["text"].replace("\n", " ")),
             x=x, y=y, cx=cx, cy=cy, fill=b["fill"], paras=paragraphes).encode("utf-8")


def _nouveau_dessin(fragments):
    return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<xdr:wsDr xmlns:xdr="' + NS_XDR.encode() + b'" xmlns:a="' + NS_A.encode() + b'">'
            + b"".join(fragments) + b"</xdr:wsDr>")


# --------------------------------------------------------------------------------------
# Validation des paramètres
# --------------------------------------------------------------------------------------
_CLES_BOUTON = ("sheet", "text", "macro")
_CLES_INDICES = ("from_col", "from_row", "to_col", "to_row")
_CLES_FACULTATIVES = {"range", "from_col_off", "from_row_off", "to_col_off", "to_row_off",
                      "inset_px", "fill", "font_color", "bold", "font_size", "font_name",
                      "name", "edit_as"}


def _entier(v):
    return isinstance(v, int) and not isinstance(v, bool)


def _normaliser_bouton(b, i):
    if not isinstance(b, dict):
        raise ValueError("bouton n°%d : dict attendu" % i)
    n = dict(b)
    if "range" in n:
        if any(k in n for k in _CLES_INDICES):
            raise ValueError("bouton n°%d : donner « range » OU les quatre indices, pas les deux" % i)
        n["from_col"], n["from_row"], n["to_col"], n["to_row"] = plage_vers_indices(n.pop("range"))
    manquants = [k for k in _CLES_BOUTON + _CLES_INDICES if k not in n]
    if manquants:
        raise ValueError("bouton n°%d : clés manquantes %s" % (i, manquants))
    inconnues = set(b) - set(_CLES_BOUTON) - set(_CLES_INDICES) - _CLES_FACULTATIVES
    if inconnues:
        raise ValueError("bouton n°%d : clés inconnues %s" % (i, sorted(inconnues)))
    macro = n["macro"]
    if not isinstance(macro, str):
        raise ValueError("bouton n°%d : nom de macro attendu" % i)
    if macro.startswith("[0]!"):
        macro = macro[4:]
    if not RE_MACRO.match(macro):
        raise ValueError("bouton n°%d : nom de macro invalide %r (NomMacro ou Module.NomMacro)"
                         % (i, n["macro"]))
    n["macro_name"] = macro
    n["macro_attr"] = "[0]!" + macro      # format écrit par Excel pour une macro du classeur
    if len(n["macro_attr"]) > 256:        # [MS-OI29500] : 256 caractères au plus
        raise ValueError("bouton n°%d : nom de macro trop long" % i)
    for k in _CLES_INDICES:
        if not _entier(n[k]) or n[k] < 0:
            raise ValueError("bouton n°%d : %s doit être un entier >= 0" % (i, k))
    for k in ("from_col_off", "from_row_off", "to_col_off", "to_row_off", "inset_px"):
        n[k] = n.get(k, 0)
        if not _entier(n[k]) or n[k] < 0:
            raise ValueError("bouton n°%d : %s doit être un entier >= 0" % (i, k))
    if n["from_col"] >= MAX_COL or n["to_col"] > MAX_COL or n["from_row"] >= MAX_ROW \
            or n["to_row"] > MAX_ROW:
        raise ValueError("bouton n°%d : hors des limites de la feuille" % i)
    if (n["to_col"], n["to_col_off"]) <= (n["from_col"], n["from_col_off"]) or \
       (n["to_row"], n["to_row_off"]) <= (n["from_row"], n["from_row_off"]):
        raise ValueError("bouton n°%d : le coin bas-droit doit être après le coin haut-gauche" % i)
    n["fill"] = str(n.get("fill", "1F2A44")).lstrip("#").upper()
    n["font_color"] = str(n.get("font_color", "FFFFFF")).lstrip("#").upper()
    for k in ("fill", "font_color"):
        if not RE_COULEUR.match(n[k]):
            raise ValueError("bouton n°%d : couleur %s invalide %r (RRGGBB)" % (i, k, n[k]))
    n["bold"] = bool(n.get("bold", True))
    n["font_size"] = float(n.get("font_size", 10))
    if not 1 <= n["font_size"] <= 400:
        raise ValueError("bouton n°%d : font_size hors de [1, 400]" % i)
    n["font_name"] = n.get("font_name") or None
    if n["font_name"] is not None and (not isinstance(n["font_name"], str)
                                       or len(n["font_name"]) > 31):
        raise ValueError("bouton n°%d : font_name invalide" % i)
    n["edit_as"] = n.get("edit_as", "oneCell")
    if n["edit_as"] not in ("twoCell", "oneCell", "absolute"):
        raise ValueError("bouton n°%d : edit_as invalide" % i)
    if not isinstance(n["text"], str) or not n["text"].strip():
        raise ValueError("bouton n°%d : texte vide" % i)
    if re.search(r"[\x00-\x08\x0b-\x1f￾￿]", n["text"]):
        raise ValueError("bouton n°%d : caractère de contrôle interdit en XML" % i)
    n["name"] = n.get("name") or ("Bouton " + macro)
    if not isinstance(n["name"], str) or re.search(r"[\x00-\x1f]", n["name"]):
        raise ValueError("bouton n°%d : nom de forme invalide" % i)
    return n


def _verifier_codenames(codenames, feuilles):
    if not isinstance(codenames, dict) or "__workbook__" not in codenames:
        raise ValueError('codenames doit contenir "__workbook__" (en général "ThisWorkbook")')
    inconnues = set(codenames) - set(feuilles) - {"__workbook__"}
    if inconnues:
        raise ValueError("codenames : feuilles inexistantes %s" % sorted(inconnues))
    manquantes = [f for f in feuilles if f not in codenames]
    if manquantes:
        raise ValueError("codenames : aucun codeName pour les feuilles %s" % manquantes)
    vus = {}
    for cle, nom in codenames.items():
        if not isinstance(nom, str) or not RE_IDENT.match(nom) or len(nom) > 31:
            raise ValueError("codeName invalide pour %r : %r (identifiant VBA de 31 caractères "
                             "au plus : lettre puis lettres, chiffres ou _)" % (cle, nom))
        if nom.lower() in vus:
            raise ValueError("codeName %r utilisé deux fois (%r et %r) — la casse ne compte pas en VBA"
                             % (nom, vus[nom.lower()], cle))
        vus[nom.lower()] = cle


# --------------------------------------------------------------------------------------
# Fonction principale
# --------------------------------------------------------------------------------------
def make_xlsm(src_xlsx, dst_xlsm, vba_bin_bytes, codenames, buttons):
    """Crée dst_xlsm à partir de src_xlsx. Retourne un rapport (dict) ; lève ValueError si
    les paramètres sont incohérents (rien n'est alors écrit). Les avertissements sont émis
    via warnings.warn et recopiés dans rapport["avertissements"]."""
    rapport = {"avertissements": [], "parties_modifiees": [], "parties_ajoutees": [],
               "boutons": []}

    def avertir(msg):
        rapport["avertissements"].append(msg)
        warnings.warn(msg, stacklevel=3)

    if not str(dst_xlsm).lower().endswith(".xlsm"):
        avertir("le fichier de sortie devrait porter l'extension .xlsm (Excel refuse d'ouvrir "
                "un classeur prenant en charge les macros sous l'extension .xlsx)")
    vba_bin_bytes = bytes(vba_bin_bytes)
    projet = lire_projet_vba(vba_bin_bytes)       # lève ValueError si ce n'est pas un projet VBA
    rapport["projet_vba"] = projet

    with zipfile.ZipFile(src_xlsx) as zin:
        infos = zin.infolist()
        noms = [i.filename for i in infos]
        if len({n.lower() for n in noms}) != len(noms):
            raise ValueError("archive source avec des entrées en double")
        parts = {i.filename: zin.read(i) for i in infos}
    noms_min = {n.lower() for n in noms}
    if "[content_types].xml" not in noms_min:
        raise ValueError("[Content_Types].xml absent : ce n'est pas un paquetage OPC")
    nouveaux = {}          # nom -> octets (parties modifiées ou ajoutées)
    ordre_ajouts = []

    def lire(nom):
        return nouveaux.get(nom, parts.get(nom))

    def ecrire(nom, octets):
        if nom not in parts and nom not in nouveaux:
            if nom.lower() in noms_min or any(nom.lower() == a.lower() for a in ordre_ajouts):
                raise ValueError("conflit de nom de partie : %s" % nom)
            ordre_ajouts.append(nom)
        nouveaux[nom] = octets

    # ---- partie principale (classeur) via _rels/.rels
    _, rels_pkg = _lire_rels(parts["_rels/.rels"], "_rels/.rels")
    principaux = [r for r in rels_pkg if r["Type"] == REL_OFFICE_DOC]
    if len(principaux) != 1:
        raise ValueError("relation officeDocument introuvable ou multiple")
    wb_part = _resoudre_cible("", principaux[0]["Target"])
    wb_rels_part = _chemin_rels(wb_part)
    if any(n.lower().endswith("vbaproject.bin") for n in noms):
        raise ValueError("le classeur source contient déjà un projet VBA")

    # ---- feuilles : nom -> partie
    wb_data = parts[wb_part]
    wb_racine = _indexer(wb_data, 2, wb_part)
    _, wb_rels = _lire_rels(parts[wb_rels_part], wb_rels_part)
    rel_par_id = {r["Id"]: r for r in wb_rels}
    feuilles = {}               # titre -> partie
    ordre_feuilles = []
    for c in wb_racine.children:
        if c.ns == NS_MAIN and c.local == "sheets":
            for s in c.children:
                if s.local != "sheet":
                    continue
                rid = s.attr("id", NS_R)
                rel = rel_par_id.get(rid)
                if rel is None:
                    raise ValueError("relation %s de la feuille %r introuvable" % (rid, s.attr("name")))
                if rel["Type"] != REL_WORKSHEET:
                    raise ValueError("feuille %r : type non géré (%s)" % (s.attr("name"), rel["Type"]))
                feuilles[s.attr("name")] = _resoudre_cible(wb_part, rel["Target"])
                ordre_feuilles.append(s.attr("name"))
    _verifier_codenames(codenames, ordre_feuilles)

    # ---- cohérence avec les modules document du projet VBA (avertissements seulement)
    docs = {d.lower(): d for d in projet["documents"]}
    wb_code = codenames["__workbook__"]
    if wb_code.lower() not in docs:
        avertir("le projet VBA ne contient pas de module document %r pour le classeur : Excel en "
                "créera un vide (le code de ThisWorkbook ne serait pas relié au classeur)" % wb_code)
    elif docs[wb_code.lower()] != wb_code:
        avertir("codeName %r du classeur : casse différente du module %r du projet VBA"
                % (wb_code, docs[wb_code.lower()]))
    for titre in ordre_feuilles:
        if codenames[titre].lower() not in docs:
            avertir("feuille %r : pas de module document %r dans le projet VBA ; Excel le créera "
                    "(vide) à l'ouverture — sans conséquence si la feuille n'a pas de code"
                    % (titre, codenames[titre]))
        elif docs[codenames[titre].lower()] != codenames[titre]:
            avertir("feuille %r : codeName %r de casse différente du module %r du projet VBA "
                    "(à aligner)" % (titre, codenames[titre], docs[codenames[titre].lower()]))
    utilises = {v.lower() for v in codenames.values()}
    orphelins = [d for d in projet["documents"] if d.lower() not in utilises]
    if orphelins:
        avertir("modules document du projet VBA sans feuille correspondante : %s (Excel les "
                "conserve tels quels ; à éviter dans le projet VBA définitif)" % orphelins)
    autres = {m.lower() for m in projet["modules"] + projet["classes"]}
    conflits = [v for v in codenames.values() if v.lower() in autres]
    if conflits:
        raise ValueError("codeName identique à un module standard/classe du projet : %s" % conflits)

    # ---- 1. workbook.xml : codeName du classeur
    wbpr = [c for c in wb_racine.children if c.ns == NS_MAIN and c.local == "workbookPr"]
    if wbpr:
        edit = _edit_attribut(wb_data, wbpr[0], "codeName", wb_code)
    else:
        p = wb_racine.prefix
        frag = ("<%sworkbookPr codeName=\"%s\"/>" % ((p + ":") if p else "", wb_code)).encode()
        suivants = [c for c in wb_racine.children if not (c.ns == NS_MAIN and c.local in AVANT_WORKBOOKPR)]
        pos = suivants[0].start if suivants else wb_racine.etag
        edit = (pos, pos, frag)
    ecrire(wb_part, _remplacer(wb_data, [edit]))
    rapport["parties_modifiees"].append(wb_part)

    # ---- 2. vbaProject.bin + relation
    vba_part = posixpath.join(posixpath.dirname(wb_part), "vbaProject.bin")
    ecrire(vba_part, vba_bin_bytes)
    rapport["parties_ajoutees"].append(vba_part)
    data, _ = _ajouter_relation(parts[wb_rels_part], REL_VBA, _cible_relative(wb_part, vba_part),
                                wb_rels_part)
    ecrire(wb_rels_part, data)
    rapport["parties_modifiees"].append(wb_rels_part)

    # ---- 3. feuilles : codeName + boutons
    boutons = [_normaliser_bouton(b, i) for i, b in enumerate(buttons or [])]
    for b in boutons:
        if b["sheet"] not in feuilles:
            raise ValueError("bouton %r : feuille %r inexistante" % (b["text"], b["sheet"]))
    police = _police_normale(parts, wb_part, wb_rels)
    if boutons and police is not None and police != ("Calibri", 11.0):
        avertir("police Normal du classeur %s %s pt : les marges inset_px sont calculées avec "
                "7 px par chiffre (Calibri 11) ; écart possible de quelques pixels" % police)
    nouveaux_dessins = []       # parties de dessin créées (pour [Content_Types].xml)
    for titre in ordre_feuilles:
        sp = feuilles[titre]
        sdata = parts[sp]
        mes_boutons = [b for b in boutons if b["sheet"] == titre]
        sracine = _indexer(sdata, 2 if mes_boutons else 1, sp)
        if not (sracine.ns == NS_MAIN and sracine.local == "worksheet"):
            raise ValueError("partie %s : racine <worksheet> attendue" % sp)
        edits = []
        q = (sracine.prefix + ":") if sracine.prefix else ""
        spr = [c for c in sracine.children if c.ns == NS_MAIN and c.local == "sheetPr"]
        if spr:
            if spr[0] is not sracine.children[0]:
                raise ValueError("partie %s : <sheetPr> n'est pas le premier enfant" % sp)
            edits.append(_edit_attribut(sdata, spr[0], "codeName", codenames[titre]))
        else:
            premier = sracine.children[0] if sracine.children else None
            if premier is not None and premier.ns != NS_MAIN:
                raise ValueError("partie %s : premier enfant inattendu <%s>" % (sp, premier.qname))
            frag = ('<%ssheetPr codeName="%s"/>' % (q, codenames[titre])).encode()
            edits.append((sracine.stag_end, sracine.stag_end, frag))

        if mes_boutons:
            px_col, pt_row = _metrique_feuille(sracine)
            resolus = [_resoudre_ancres(b, px_col, pt_row) for b in mes_boutons]
            if len({b["name"] for b in resolus}) != len(resolus):
                raise ValueError("%s : noms de boutons en double" % titre)
            drawing = [c for c in sracine.children if c.ns == NS_MAIN and c.local == "drawing"]
            srels_part = _chemin_rels(sp)
            if drawing:
                rid = drawing[0].attr("id", NS_R)
                if srels_part not in parts:
                    raise ValueError("%s : <drawing> sans fichier de relations" % sp)
                _, srels = _lire_rels(parts[srels_part], srels_part)
                rel = [r for r in srels if r["Id"] == rid]
                if not rel or rel[0]["Type"] != REL_DRAWING or rel[0]["TargetMode"] == "External":
                    raise ValueError("%s : relation de dessin %s invalide" % (sp, rid))
                dpart = _resoudre_cible(sp, rel[0]["Target"])
                ddata = lire(dpart)
                if ddata is None:
                    raise ValueError("partie de dessin %s absente" % dpart)
                dracine = _indexer(ddata, None, dpart)
                if not (dracine.ns == NS_XDR and dracine.local == "wsDr"):
                    raise ValueError("%s : racine <wsDr> attendue" % dpart)
                # identifiants de forme : uniques dans le dessin (tous les cNvPr, y compris
                # ceux des groupes et des mc:Fallback)
                ids = [int(e.attr("id")) for e in _iter(dracine)
                       if e.local == "cNvPr" and (e.attr("id") or "").isdigit()]
                noms_formes = {e.attr("name") for e in _iter(dracine) if e.local == "cNvPr"}
                prochain = max(ids + [0]) + 1
                # la racine est <wsDr> dans l'espace spreadsheetDrawing : on réutilise son
                # préfixe (vide pour openpyxl, « xdr » pour Excel)
                p = dracine.prefix
                frags = []
                for b in resolus:
                    if b["name"] in noms_formes:
                        raise ValueError("%s : nom de forme %r déjà utilisé" % (titre, b["name"]))
                    noms_formes.add(b["name"])
                    frags.append(_xml_bouton(p, prochain, b))
                    rapport["boutons"].append(_resume_bouton(titre, dpart, prochain, b))
                    prochain += 1
                ecrire(dpart, _remplacer(ddata, [_edit_ajout_enfant_fin(ddata, dracine, b"".join(frags))]))
                rapport["parties_modifiees"].append(dpart)
            else:
                # nouveau dessin : xl/drawings/drawingN.xml (N libre)
                n = 1
                existants = {x.lower() for x in list(parts) + ordre_ajouts}
                while ("xl/drawings/drawing%d.xml" % n) in existants:
                    n += 1
                dpart = "xl/drawings/drawing%d.xml" % n
                frags = []
                for k, b in enumerate(resolus):
                    frags.append(_xml_bouton("xdr", k + 2, b))      # ids à partir de 2, comme Excel
                    rapport["boutons"].append(_resume_bouton(titre, dpart, k + 2, b))
                ecrire(dpart, _nouveau_dessin(frags))
                nouveaux_dessins.append(dpart)
                rapport["parties_ajoutees"].append(dpart)
                cible = _cible_relative(sp, dpart)
                if srels_part in parts:
                    srdata, rid = _ajouter_relation(parts[srels_part], REL_DRAWING, cible, srels_part)
                    rapport["parties_modifiees"].append(srels_part)
                else:
                    rid = "rId1"
                    srdata = _nouveau_rels(REL_DRAWING, cible, rid)
                    rapport["parties_ajoutees"].append(srels_part)
                ecrire(srels_part, srdata)
                frag = ('<%sdrawing xmlns:r="%s" r:id="%s"/>' % (q, NS_R, rid)).encode()
                avant = [c for c in sracine.children if c.ns == NS_MAIN and c.local in AVANT_DRAWING]
                edits.append((avant[-1].end, avant[-1].end, frag))
        ecrire(sp, _remplacer(sdata, edits))
        rapport["parties_modifiees"].append(sp)

    # ---- 4. [Content_Types].xml
    ct_nom = [n for n in noms if n.lower() == "[content_types].xml"][0]
    ct = parts[ct_nom]
    ctr = _indexer(ct, 1, ct_nom)
    edits = []
    over_wb = [c for c in ctr.children if c.local == "Override"
               and (c.attr("PartName") or "").lower() == "/" + wb_part.lower()]
    frags = []
    if over_wb:
        if over_wb[0].attr("ContentType") != CT_MAIN_XLSX:
            raise ValueError("type de contenu du classeur inattendu : %s" % over_wb[0].attr("ContentType"))
        edits.append(_edit_attribut(ct, over_wb[0], "ContentType", CT_MAIN_XLSM))
    else:
        frags.append('<Override PartName="/%s" ContentType="%s"/>' % (wb_part, CT_MAIN_XLSM))
    defaut_bin = [c for c in ctr.children if c.local == "Default"
                  and (c.attr("Extension") or "").lower() == "bin"]
    autres_bin = [n for n in noms if n.lower().endswith(".bin")]
    if defaut_bin and defaut_bin[0].attr("ContentType") != CT_VBA:
        frags.append('<Override PartName="/%s" ContentType="%s"/>' % (vba_part, CT_VBA))
    elif not defaut_bin:
        if autres_bin:   # d'autres .bin existent sans Default « bin » : Override ciblé
            frags.append('<Override PartName="/%s" ContentType="%s"/>' % (vba_part, CT_VBA))
        else:            # comme Excel et XlsxWriter
            frags.insert(0, '<Default Extension="bin" ContentType="%s"/>' % CT_VBA)
    for d in nouveaux_dessins:
        frags.append('<Override PartName="/%s" ContentType="%s"/>' % (d, CT_DRAWING))
    if frags:
        # Les Default sont placés après le dernier Default existant (lisibilité ; CT_Types
        # accepte Default et Override dans n'importe quel ordre).
        defs = [f for f in frags if f.startswith("<Default")]
        overs = [f for f in frags if not f.startswith("<Default")]
        if defs:
            derniers = [c for c in ctr.children if c.local == "Default"]
            pos = derniers[-1].end if derniers else ctr.stag_end
            edits.append((pos, pos, "".join(defs).encode()))
        if overs:
            edits.append(_edit_ajout_enfant_fin(ct, ctr, "".join(overs).encode()))
    ecrire(ct_nom, _remplacer(ct, edits))
    rapport["parties_modifiees"].append(ct_nom)

    # ---- vérification : toutes les parties XML modifiées/ajoutées restent bien formées
    for nom, octets in nouveaux.items():
        if nom.endswith((".xml", ".rels")):
            _indexer(octets, 0, nom)

    # ---- écriture de l'archive (ordre et métadonnées des entrées d'origine conservés)
    date_ref = next((i.date_time for i in infos if i.filename == wb_part), (1980, 1, 1, 0, 0, 0))
    with zipfile.ZipFile(dst_xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, nouveaux.get(info.filename, parts[info.filename]))
        for nom in ordre_ajouts:
            zi = zipfile.ZipInfo(nom, date_time=date_ref)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, nouveaux[nom])
    return rapport


def _resume_bouton(titre, dpart, sid, b):
    """Entrée du rapport pour un bouton inséré."""
    return {"feuille": titre, "dessin": dpart, "id": sid, "nom": b["name"], "texte": b["text"],
            "macro": b["macro_attr"],
            "from": [b["from_col"], b["from_col_off"], b["from_row"], b["from_row_off"]],
            "to": [b["to_col"], b["to_col_off"], b["to_row"], b["to_row_off"]]}


# --------------------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 5:
        print(__doc__)
        sys.exit(2)
    src, dst, vba, conf = sys.argv[1:5]
    with open(vba, "rb") as f:
        vba_bytes = f.read()
    with open(conf, encoding="utf-8") as f:
        config = json.load(f)
    r = make_xlsm(src, dst, vba_bytes, config["codenames"], config.get("buttons", []))
    print(json.dumps(r, ensure_ascii=False, indent=1))

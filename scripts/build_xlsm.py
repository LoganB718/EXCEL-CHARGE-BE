# -*- coding: utf-8 -*-
"""
build_xlsm.py — assemble le classeur « Plan_de_charge_BE_Dashboard.xlsm ».

Entrées (par défaut) :
    - le classeur .xlsx à jour produit par openpyxl (5 onglets : Charge, Data, Dashboard,
      Dessinateurs, Calculs ; Dashboard et Dessinateurs ont déjà un dessin avec graphiques) ;
    - le module VBA « PlanDeCharge.bas » du dépôt (UTF-8, ligne « Attribute VB_Name »).

Sortie : un classeur .xlsm contenant
    - un projet VBA (xl/vbaProject.bin) : ThisWorkbook + un module de document par feuille
      + le module standard PlanDeCharge (stocké en page de codes 1252, fins de ligne CRLF) ;
    - deux boutons sur Dashboard et sur Dessinateurs, à droite de la zone d'impression :
        « Éditer PDF »             couvrant BL2:BM4 -> macro <Onglet>_PDF
        « Afficher / masquer l'aide »  couvrant BL6:BM7 -> macro <Onglet>_Aide
      bleu foncé 1F2A44, texte blanc gras 10 pt, marge intérieure de 3 px.

Le travail est délégué aux deux modules voisins :
    vba_project.py : build_vba_project_bin()  -> octets du vbaProject.bin ([MS-OVBA], [MS-CFB])
    make_xlsm.py   : make_xlsm()              -> chirurgie minimale du paquetage OPC
Ce script ajoute les contrôles de cohérence propres à ce classeur (macros des boutons présentes
dans le module, noms définis utilisés par le module, boutons hors zone d'impression,
identifiants réservés de VBA) et écrit la sortie de façon atomique.

Choix des codeName (noms de code VBA des feuilles) : shCharge, shData, shDashboard,
shDessinateurs, shCalculs (et ThisWorkbook pour le classeur), plutôt que Feuil1…Feuil5 :
    - ils restent justes si un onglet est renommé ou déplacé (Feuil3 ne dit rien de l'onglet) ;
    - l'explorateur de projets VBA affiche « shDashboard (Dashboard) » : lisible d'emblée ;
    - ASCII pur, sans accent : aucune dépendance à la page de codes ni à la langue d'Excel ;
    - aucun risque de collision avec le nom « FeuilN » qu'Excel français donnera à une feuille
      ajoutée plus tard (il prend le premier FeuilN libre) ;
    - « ThisWorkbook » est le nom de code qu'Excel donne au classeur (y compris en français).
Le module PlanDeCharge ne dépend d'aucun codeName (il passe par des noms définis) : ce choix
est sans effet sur son fonctionnement.

Utilisation (depuis la racine du dépôt) :
    python scripts/build_dashboard.py classeur.xlsx
    python scripts/build_xlsm.py --xlsx classeur.xlsx --sortie Plan_de_charge_BE_Dashboard.xlsm
    options : [--bas MODULE.bas] [--bin COPIE_vbaProject.bin] [--forcer]
Entre les deux, le classeur livré passe aussi par un recalcul LibreOffice et
inject_cached_values.py (valeurs affichées avant le premier calcul d'Excel).
build_xlsm.py doit rester la DERNIÈRE étape : réenregistrer le .xlsm avec openpyxl ou
LibreOffice supprime les boutons ou leur macro.

Bibliothèque standard uniquement (plus les deux composants ci-dessus).
"""

import argparse
import json
import os
import re
import sys
import tempfile
import warnings
import xml.etree.ElementTree as ET
import zipfile

# --------------------------------------------------------------------------------------
# Emplacements par défaut
# --------------------------------------------------------------------------------------
ICI = os.path.dirname(os.path.abspath(__file__))     # dossier scripts/ du dépôt

XLSX_DEFAUT = "Plan_de_charge_BE_Dashboard.xlsx"
BAS_DEFAUT = os.path.join(ICI, "macros", "PlanDeCharge.bas")
SORTIE_DEFAUT = "Plan_de_charge_BE_Dashboard.xlsm"

if ICI not in sys.path:
    sys.path.insert(0, ICI)
from vba_project import build_vba_project_bin           # noqa: E402
from make_xlsm import bouton_plage, make_xlsm, plage_vers_indices   # noqa: E402

# --------------------------------------------------------------------------------------
# Paramètres du classeur
# --------------------------------------------------------------------------------------
NOM_MODULE = "PlanDeCharge"

# Titre de l'onglet -> codeName (ordre des onglets du classeur).
CODENAMES = {
    "__workbook__": "ThisWorkbook",
    "Charge": "shCharge",
    "Data": "shData",
    "Dashboard": "shDashboard",
    "Dessinateurs": "shDessinateurs",
    "Calculs": "shCalculs",
}

ONGLETS_BOUTONS = ("Dashboard", "Dessinateurs")

# (plage couverte, texte, suffixe de la macro <Onglet>_<suffixe>)
BOUTONS = (
    ("BL2:BM4", "Éditer PDF", "PDF"),
    ("BL6:BM7", "Afficher / masquer l'aide", "Aide"),
)
STYLE_BOUTON = {"fill": "1F2A44", "font_color": "FFFFFF", "bold": True, "font_size": 10,
                "font_name": "Arial"}          # police du classeur
MARGE_PX = 3          # marge intérieure entre le bouton et les bords des cellules

# Code des modules de document (événements : ouverture, activation d'un onglet, modification),
# fichiers du dossier macros/ ; les autres modules de document restent vides.
DOCUMENTS_CODE = {
    "ThisWorkbook": "ThisWorkbook.vba",
    "shDashboard": "shDashboard.vba",
    "shDessinateurs": "shDessinateurs.vba",
}

# Noms définis lus par le module (cellule unique de l'onglet indiqué).
NOMS_REQUIS = {
    "aide_Dashboard": "Calculs", "edite_Dashboard": "Dashboard",
    "aide_Dessinateurs": "Calculs", "edite_Dessinateurs": "Dessinateurs",
}
# Noms définis lus par AjusterLignes (plages, éventuellement en plusieurs morceaux, de l'onglet).
PLAGES_REQUISES = {"lignes_masquables_Dashboard": "Dashboard",
                   "lignes_masquables_Dessinateurs": "Dessinateurs"}

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Identifiants réservés de VBA ([MS-VBAL] §3.3.5.2 : reserved-name, special-form,
# reserved-type-identifier, littéraux) : ils ne peuvent pas nommer une variable, une
# constante ou un paramètre (« Identificateur attendu » à la compilation dans Excel).
# Remarque : LibreOffice Basic les accepte ; ses essais ne détectent donc pas cette erreur.
RESERVES_VBA = {m.lower() for m in (
    "Abs CBool CByte CCur CDate CDbl CDec CInt CLng CLngLng CLngPtr CSng CStr CVar CVErr "
    "Date Debug DoEvents Fix Int Len LenB Me PSet Scale Sgn String "
    "Array Circle Input InputB LBound UBound "
    "Boolean Byte Currency Double Integer Long LongLong LongPtr Single Variant "
    "True False Nothing Empty Null").split()}


class ErreurControle(ValueError):
    """Incohérence bloquante détectée avant l'assemblage."""


# --------------------------------------------------------------------------------------
# Lecture et contrôle du module VBA
# --------------------------------------------------------------------------------------
def lire_bas(chemin):
    """Lit un fichier .bas en UTF-8 (BOM accepté). Retourne (nom du module, texte)."""
    with open(chemin, "rb") as f:
        brut = f.read()
    try:
        texte = brut.decode("utf-8")
    except UnicodeDecodeError:
        raise ErreurControle("%s : fichier attendu en UTF-8" % chemin) from None
    texte = texte.lstrip("﻿")
    premiere = texte.split("\n", 1)[0].strip()
    m = re.match(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', premiere, re.IGNORECASE)
    nom = m.group(1) if m else os.path.splitext(os.path.basename(chemin))[0]
    return nom, texte


def _sans_chaines_ni_commentaire(ligne):
    """Remplace les chaînes "..." par "" et coupe le commentaire (' ou Rem)."""
    ligne = re.sub(r'"(?:[^"]|"")*"', '""', ligne)
    ligne = ligne.split("'", 1)[0]
    return re.sub(r"(^|:)\s*Rem\b.*$", r"\1", ligne, flags=re.IGNORECASE)


def _lignes_logiques(texte):
    """Lignes logiques (continuations « _ » recollées), sans chaînes ni commentaires,
    avec le numéro de la première ligne physique."""
    lignes, tampon, debut = [], "", None
    for numero, brute in enumerate(texte.replace("\r\n", "\n").split("\n"), 1):
        net = _sans_chaines_ni_commentaire(brute)
        if debut is None:
            debut = numero
        if re.search(r"\s_\s*$", net):
            tampon += re.sub(r"\s_\s*$", " ", net)
            continue
        lignes.append((debut, tampon + net))
        tampon, debut = "", None
    if tampon:
        lignes.append((debut, tampon))
    return lignes


def _decouper(liste):
    """Découpe « a As X, b(1 To 2) As Y » selon les virgules hors parenthèses."""
    morceaux, prof, cour = [], 0, ""
    for ch in liste:
        if ch == "(":
            prof += 1
        elif ch == ")":
            prof -= 1
        if ch == "," and prof == 0:
            morceaux.append(cour)
            cour = ""
        else:
            cour += ch
    morceaux.append(cour)
    return [m.strip() for m in morceaux if m.strip()]


def identifiants_declares(texte):
    """Liste (ligne, identifiant, nature) des variables, constantes et paramètres déclarés.
    Contrôle léger (pas un analyseur VBA complet) : suffisant pour un module de macros."""
    trouves = []
    re_proc = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*"
                         r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\((.*)\)",
                         re.IGNORECASE)
    re_decl = re.compile(r"^\s*(?:Dim|Static|Global|Const|ReDim(?:\s+Preserve)?|"
                         r"(?:Public|Private)(?:\s+Const)?)\s+(?!Sub\b|Function\b|Property\b|"
                         r"Declare\b|Type\b|Enum\b|Event\b|Const\b)(.*)$", re.IGNORECASE)
    for numero, ligne in _lignes_logiques(texte):
        for instr in re.split(r":(?!=)", ligne):    # « : » sépare les instructions (pas « := »)
            m = re_proc.match(instr)
            if m:
                for p in _decouper(m.group(2)):
                    p = re.sub(r"^(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)+", "", p,
                               flags=re.IGNORECASE)
                    ident = re.match(r"\w+", p)
                    if ident:
                        trouves.append((numero, ident.group(0), "paramètre"))
                continue
            m = re_decl.match(instr)
            if m:
                for item in _decouper(m.group(1)):
                    item = re.sub(r"^WithEvents\s+", "", item, flags=re.IGNORECASE)
                    ident = re.match(r"\w+", item)
                    if ident:
                        trouves.append((numero, ident.group(0), "déclaration"))
    return trouves


def macros_publiques(texte):
    """Noms des Sub publiques sans paramètre (seules affectables à un bouton)."""
    noms = set()
    for _, ligne in _lignes_logiques(texte):
        m = re.match(r"^\s*(?:Public\s+)?Sub\s+(\w+)\s*\(\s*\)", ligne, re.IGNORECASE)
        if m and not re.match(r"^\s*Private\b", ligne, re.IGNORECASE):
            noms.add(m.group(1))
    return noms


def controler_module(nom, texte, macros_attendues):
    """Retourne la liste des problèmes (chaînes) du module VBA."""
    problemes = []
    if nom != NOM_MODULE:
        problemes.append("le module s'appelle %r (Attribute VB_Name) au lieu de %r"
                         % (nom, NOM_MODULE))
    if not re.search(r"^\s*Option\s+Explicit\b", texte, re.IGNORECASE | re.MULTILINE):
        problemes.append("« Option Explicit » absent")
    for numero, ident, nature in identifiants_declares(texte):
        if ident.lower() in RESERVES_VBA:
            problemes.append("ligne %d : %s « %s » = identifiant réservé de VBA ([MS-VBAL] "
                             "§3.3.5.2) -> erreur de compilation « Identificateur attendu » "
                             "dans Excel (LibreOffice l'accepte)" % (numero, nature, ident))
    publiques = {m.lower(): m for m in macros_publiques(texte)}
    for macro in macros_attendues:
        if macro.lower() not in publiques:
            problemes.append("macro %s introuvable (Sub publique sans paramètre attendue)" % macro)
        elif publiques[macro.lower()] != macro:
            problemes.append("macro %s : casse différente dans le module (%s)"
                             % (macro, publiques[macro.lower()]))
    return problemes


# --------------------------------------------------------------------------------------
# Contrôles sur le classeur source
# --------------------------------------------------------------------------------------
def _col_index(lettres):
    n = 0
    for ch in lettres.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def controler_classeur(xlsx):
    """Vérifie la présence des onglets, des noms définis requis, et que les boutons sont
    à droite de la zone d'impression. Retourne la liste des problèmes."""
    problemes = []
    with zipfile.ZipFile(xlsx) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
    feuilles = [s.get("name") for s in wb.iter(NS_MAIN + "sheet")]
    for titre in CODENAMES:
        if titre != "__workbook__" and titre not in feuilles:
            problemes.append("onglet %r absent du classeur" % titre)
    globaux, zones = {}, {}
    for dn in wb.iter(NS_MAIN + "definedName"):
        texte = (dn.text or "").strip()
        if dn.get("localSheetId") is None:
            globaux[dn.get("name")] = texte
        elif dn.get("name") == "_xlnm.Print_Area":
            zones[feuilles[int(dn.get("localSheetId"))]] = texte
    for nom, onglet in NOMS_REQUIS.items():
        cible = globaux.get(nom)
        motif = r"^'?%s'?!\$?[A-Z]{1,3}\$?\d+$" % re.escape(onglet)
        if cible is None:
            problemes.append("nom défini %s absent (utilisé par le module)" % nom)
        elif not re.match(motif, cible):
            problemes.append("nom défini %s = %s : une cellule de l'onglet %s est attendue"
                             % (nom, cible, onglet))
    for nom, onglet in PLAGES_REQUISES.items():
        cible = globaux.get(nom)
        morceau = r"'?%s'?!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?" % re.escape(onglet)
        if cible is None:
            problemes.append("nom défini %s absent (utilisé par AjusterLignes)" % nom)
        elif not re.match(r"^%s(?:,%s)*$" % (morceau, morceau), cible):
            problemes.append("nom défini %s = %s : des plages de l'onglet %s sont attendues"
                             % (nom, cible, onglet))
    for onglet in ONGLETS_BOUTONS:
        zone = zones.get(onglet)
        if zone is None:
            problemes.append("onglet %s : pas de zone d'impression" % onglet)
            continue
        derniere = max(_col_index(c) for c in re.findall(r"\$?([A-Z]{1,3})\$?\d+", zone))
        for plage, texte, _ in BOUTONS:
            premiere_col = plage_vers_indices(plage)[0]
            if premiere_col <= derniere:
                problemes.append("bouton « %s » (%s) dans la zone d'impression %s de %s"
                                 % (texte, plage, zone, onglet))
    return problemes


# --------------------------------------------------------------------------------------
# Assemblage
# --------------------------------------------------------------------------------------
def boutons():
    """Description des boutons pour make_xlsm()."""
    liste = []
    for onglet in ONGLETS_BOUTONS:
        for plage, texte, suffixe in BOUTONS:
            liste.append(bouton_plage(onglet, plage, texte, "%s_%s" % (onglet, suffixe),
                                      inset_px=MARGE_PX,
                                      name="Bouton %s %s" % (onglet, suffixe), **STYLE_BOUTON))
    return liste


def construire(xlsx=XLSX_DEFAUT, bas=BAS_DEFAUT, sortie=SORTIE_DEFAUT, copie_bin=None,
               forcer=False):
    """Construit le classeur .xlsm. Retourne le rapport de make_xlsm() complété.
    Lève ErreurControle si un contrôle échoue (sauf forcer=True : avertissements)."""
    nom, texte = lire_bas(bas)
    macros = ["%s_%s" % (o, s) for o in ONGLETS_BOUTONS for _, _, s in BOUTONS]
    problemes = controler_module(nom, texte, macros) + controler_classeur(xlsx)
    codes = {}
    for module, fichier in DOCUMENTS_CODE.items():
        with open(os.path.join(os.path.dirname(os.path.abspath(bas)), fichier), encoding="utf-8-sig") as f:
            codes[module] = f.read()
        for numero, ident, nature in identifiants_declares(codes[module]):
            if ident.lower() in RESERVES_VBA:
                problemes.append("%s, ligne %d : %s « %s » = identifiant réservé de VBA"
                                 % (fichier, numero, nature, ident))
    if problemes and not forcer:
        raise ErreurControle("contrôles en échec :\n  - " + "\n  - ".join(problemes)
                             + "\n(--forcer pour construire malgré tout)")

    # 1. Projet VBA : modules de document dans l'ordre des onglets, puis le module standard.
    documents = [(CODENAMES["__workbook__"], "workbook", codes.get(CODENAMES["__workbook__"]))]
    documents += [(code, "worksheet", codes.get(code)) for titre, code in CODENAMES.items()
                  if titre != "__workbook__"]
    vba_bin = build_vba_project_bin({nom: texte}, documents)
    if copie_bin:
        with open(copie_bin, "wb") as f:
            f.write(vba_bin)

    # 2. Classeur : écriture dans un fichier temporaire du même dossier, puis
    #    remplacement atomique (jamais de .xlsm à moitié écrit).
    dossier = os.path.dirname(os.path.abspath(sortie))
    fd, temporaire = tempfile.mkstemp(prefix=".build_", suffix=".xlsm", dir=dossier)
    os.close(fd)
    try:
        with warnings.catch_warnings(record=True) as alertes:
            warnings.simplefilter("always")
            rapport = make_xlsm(xlsx, temporaire, vba_bin, CODENAMES, boutons())
        with zipfile.ZipFile(temporaire) as z:          # relecture de contrôle
            if z.testzip() is not None:
                raise ErreurControle("archive produite corrompue")
        # mkstemp crée le fichier en 0600 : droits usuels (0666 moins le umask) rétablis.
        masque = os.umask(0)
        os.umask(masque)
        os.chmod(temporaire, 0o666 & ~masque)
        os.replace(temporaire, sortie)
    finally:
        if os.path.exists(temporaire):
            os.remove(temporaire)
    rapport["controles"] = problemes
    rapport["alertes_make_xlsm"] = [str(a.message) for a in alertes]
    rapport["sortie"] = os.path.abspath(sortie)
    rapport["taille_vbaProject_bin"] = len(vba_bin)
    return rapport


def main(argv=None):
    ap = argparse.ArgumentParser(description="Assemble Plan_de_charge_BE_Dashboard.xlsm")
    ap.add_argument("--xlsx", default=XLSX_DEFAUT, help="classeur source (.xlsx)")
    ap.add_argument("--bas", default=BAS_DEFAUT, help="module VBA PlanDeCharge (.bas, UTF-8)")
    ap.add_argument("--sortie", default=SORTIE_DEFAUT, help="classeur produit (.xlsm)")
    ap.add_argument("--bin", default=None, help="copie facultative du vbaProject.bin produit")
    ap.add_argument("--forcer", action="store_true",
                    help="construire malgré des contrôles en échec (signalés dans le rapport)")
    args = ap.parse_args(argv)
    try:
        rapport = construire(args.xlsx, args.bas, args.sortie, args.bin, args.forcer)
    except ErreurControle as exc:
        print("ERREUR : %s" % exc, file=sys.stderr)
        return 1
    for p in rapport["controles"]:
        print("ATTENTION (--forcer) : %s" % p, file=sys.stderr)
    print(json.dumps(rapport, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())

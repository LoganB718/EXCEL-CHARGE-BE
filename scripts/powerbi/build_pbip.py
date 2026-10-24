"""Génération du projet Power BI « PlanDeCharge_BI » (PBIP : modèle sémantique model.bim + rapport PBIR).

Usage (depuis la racine du dépôt) : python3 scripts/powerbi/build_pbip.py powerbi

Fichiers de scripts/powerbi/ :
    build_pbip.py      rapport (pages, visuels) et écriture du projet ; pbip_lib.py : écriture PBIR / PBIP
    model_def.py       modèle sémantique : requêtes Power Query (M), tables, relations, mesures DAX
    mirror.py          miroir Python des requêtes M et des mesures DAX (Power BI ne s'exécute pas hors Windows)
    test_mirror.py     miroir comparé aux valeurs d'Excel sur des classeurs recalculés (--racine ou liste .xlsx)
    validate_pbip.py   fichiers PBIR contre les schémas officiels (copies dans schemas/), références au modèle
    validate_bim.py    contrôles structurels de model.bim (--donnees : échantillon écrit par exporter_donnees.py)
    check_model.py     analyse M (check_m.js, outils/analyse_m.js) et DAX (pbi_parsers) du modèle
Dépendances des contrôles (pas de la génération) : jsonschema et pbi_parsers (pip), node et
@microsoft/powerquery-parser (npm install dans scripts/powerbi, ou variable PBI_NODE_MODULES).
Contrôles :
    python3 scripts/powerbi/validate_pbip.py powerbi --bim
    python3 scripts/powerbi/check_model.py
    python3 scripts/powerbi/test_mirror.py <classeur recalculé.xlsx> [...]

Le modèle vient de model_def.build_model() ; le rapport reproduit les onglets Dashboard et Dessinateurs du classeur
Plan_de_charge_BE_Dashboard.xlsm (scripts/build_dashboard.py), avec une fiche d'extraction par affaire et une page
de contrôle Power BI / Excel. Pages :
    1. Tableau de bord   (onglet Dashboard)
    2. Ressources        (onglet Dessinateurs)
    3. Fiche affaire     (extraction, masquée)
    4. Contrôle          (comparaison avec l'onglet Calculs)
    + deux pages d'info-bulle masquées (projet, occupation).
"""
import os
import sys

ICI = os.path.dirname(os.path.abspath(__file__))
if ICI not in sys.path:
    sys.path.insert(0, ICI)

from pbip_lib import (Rapport, Visuel, ColonneTableau, Couleur, Entier, theme_personnalise, col, mes, agr,
                      carte, segment, segment_dates, zone_texte, paragraphe, forme_rectangle, graphique_combo,
                      graphique_barres, anneau, tableau, matrice, bouton, filtre_categoriel, filtre_avance, filtre_comparaison, couleur, literal,
                      lit_texte, lit_valeur, hex_couleur, sel_metadata, sel_par_point, objet, ecrire_projet,
                      _habiller)
from model_def import build_model

NOM = "PlanDeCharge_BI"
T = "Mesures"

# Couleurs du classeur (build_dashboard.py)
NAVY, INK, GREY, MUTED, LINE, BG = "1F2A44", "1F2937", "6B7280", "9CA3AF", "D6DCE4", "EEF1F5"
ACCENT, AMBER, RED, GREEN = "2A78D6", "EDA100", "B42318", "1E7B34"
BANDEAU_TXT = "C9D3E3"
SLOT_COLORS = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948", "9CA3AF"]
FIX_C, PROJ_C, OVER_C, CAP_C = "B8C0CC", "2A78D6", "D03B3B", "1F2937"   # graphique d'une ressource
BARRE_LISTE = "5B6B8C"          # barres neutres (chefs, types) : ne pas confondre avec les couleurs des projets

# Grille de page : 1280 px de large, marges de 16 px, gouttières de 12 px
LARGEUR = 1280
M, G = 16, 12
UTILE = LARGEUR - 2 * M
BANDEAU_H = 92

# Synchronisation des segments entre les pages
SYNC_CHEF, SYNC_PROJET, SYNC_TYPE, SYNC_RES = "sync_chef", "sync_projet", "sync_type", "sync_ressource"
EXCEL = "Comme dans Excel"
PERIODE_EXCEL = "12 mois du Dashboard Excel"

P_TDB, P_RES, P_FICHE, P_CTRL = "tableau_de_bord", "ressources", "fiche_affaire", "controle"
P_IB_PROJET, P_IB_OCC = "infobulle_projet", "infobulle_occupation"
# Pages à imprimer : A4 paysage (1280 × 905 = rapport 1,414), une page du rapport = une page du PDF
P_PDF_TDB1, P_PDF_TDB2, P_PDF_RES1, P_PDF_RES2 = "pdf_tableau_de_bord_1", "pdf_tableau_de_bord_2", "pdf_ressources_1", \
    "pdf_ressources_2"
HAUT_PDF, BANDEAU_PDF = 905, 74
N_AFFAIRES_PDF = 30


# ---------------------------------------------------------------------------------------------- outils
def grille(n, x0=M, largeur=UTILE, g=G):
    """n colonnes égales : liste de (x, largeur)."""
    w = (largeur - (n - 1) * g) / n
    return [(round(x0 + i * (w + g), 1), round(w, 1)) for i in range(n)]


def m(nom):
    return mes(T, nom)


def condition(champ, operateur, valeur):
    """Condition de règle de mise en forme : champ <op> littéral."""
    k = {"=": 0, ">": 1, ">=": 2, "<": 3, "<=": 4}[operateur]
    c = {"Comparison": {"ComparisonKind": k, "Left": champ.expr(), "Right": literal(lit_valeur(valeur))}}
    if isinstance(valeur, str):
        # comparaison de textes : annotations que Power BI Desktop écrit toujours dans ce cas (corpus FUAM, FCA)
        c["Annotations"] = {"PowerBI.SQExprEvaluationKind": 1, "PowerBI.SQExprTextOperatorOption": 2}
    return c


def regle(*cas):
    """Couleur par règles (mise en forme conditionnelle « Règles ») : cas = (condition, couleur hex)."""
    return Couleur({"Conditional": {"Cases": [
        {"Condition": c, "Value": literal(lit_texte(hex_couleur(h)))} for c, h in cas]}})


def couleur_champ(champ):
    """Couleur « valeur de champ » : mesure renvoyant #RRGGBB, ou colonne texte (première valeur)."""
    if champ.genre == "Measure":
        return Couleur(champ.expr())
    return Couleur({"Aggregation": {"Expression": champ.expr_base(), "Function": 3}})


def props_sans_selecteur(v, nom_objet, props):
    """Complète l'entrée sans sélecteur d'un objet du visuel (la crée au besoin)."""
    for e in v.objets.setdefault(nom_objet, []):
        if "selector" not in e:
            e["properties"].update(objet(props)["properties"])
            return v
    v.objets[nom_objet].append(objet(props))
    return v


def titre_dynamique(v, champ):
    """Titre du visuel donné par une mesure texte."""
    for e in v.conteneur.get("title", []):
        e["properties"]["text"] = {"expr": champ.expr()}
    return v


def sous_titre(v, texte_ou_champ, taille=9, couleur_txt=GREY):
    texte = {"expr": texte_ou_champ.expr()} if hasattr(texte_ou_champ, "expr") else texte_ou_champ
    v.objet_conteneur("subTitle", {"show": True, "text": texte, "fontSize": float(taille),
                                   "fontColor": couleur(couleur_txt)})
    return v


def couleur_valeur_carte(v, coul):
    """Couleur de la valeur d'une carte (Couleur : règle ou mesure)."""
    v.objets["labels"][0]["properties"]["color"] = coul.pbir()
    return v


def barres_donnees(v, champ, teinte, maximum=None):
    """Barres de données d'une colonne de tableau (forme relevée dans les rapports de Desktop).
    maximum : valeur d'une barre pleine (1 = 100 % pour un avancement, comme l'Excel) ; par défaut la plus grande
    valeur affichée."""
    bornes = {} if maximum is None else {"maxValue": {"expr": literal(lit_valeur(float(maximum)))}}
    v.objet("columnFormatting", {"dataBars": {
        **bornes,
        "positiveColor": couleur(teinte).pbir(),
        "negativeColor": {"solid": {"color": {"expr": literal("'minColor'")}}},
        "axisColor": {"solid": {"color": {"expr": literal("'foreground'")}}},
        "reverseDirection": {"expr": literal("false")},
        "hideText": {"expr": literal("false")},
        "totalMatchingOption": {"expr": literal("1L")},
    }}, sel_metadata(champ))
    return v


def couleur_colonne(v, champ, police=None, fond=None):
    pr = {}
    if police is not None:
        pr["fontColor"] = police
    if fond is not None:
        pr["backColor"] = fond
    v.objet("values", pr, sel_par_point(champ))
    return v


def alignement(v, champ, sens):
    v.objet("columnFormatting", {"alignment": sens}, sel_metadata(champ))
    return v


def texte(paragraphes, x, y, w, h, cle, **opts):
    return zone_texte(paragraphes, x, y, w, h, cle=cle, marges=(0, 4, 0, 4), **opts)


def ligne_texte(contenu, x, y, w, h, cle, taille=9, couleur_txt=GREY, gras=False, italique=False,
                alignement_txt=None):
    return texte([paragraphe(dict(texte=contenu, taille=taille, couleur=couleur_txt, gras=gras,
                                  italique=italique), alignement=alignement_txt)], x, y, w, h, cle)


def texte_mesure(champ, x, y, w, h, cle, taille=9, couleur_txt=INK, gras=False, nom_affiche=None):
    """Texte dynamique aligné à gauche : carte à plusieurs lignes sans étiquette ni cadre."""
    v = Visuel("multiRowCard", x, y, w, h)
    v.ajouter_role("Values", champ, nom_affiche=nom_affiche)
    v.objet("dataLabels", {"fontSize": float(taille), "color": couleur(couleur_txt), "bold": bool(gras)})
    v.objet("categoryLabels", {"show": False})
    v.objet("card", {"barShow": False, "outlineStyle": 0.0, "cardPadding": 0.0})
    v.libelle = champ.nom
    v.cle = cle
    return _habiller(v, fond=False, entete=False, marges=(0, 0, 0, 0), infobulle=False)


def valeur(champ, x, y, w, h, cle, taille=9, couleur_txt=GREY, gras=False, etiquette=None, decimales=None,
           taille_etiquette=8.5, couleur_etiquette=GREY, fond=False, **opts):
    """Valeur centrée (carte classique sans titre)."""
    return carte(champ, x, y, w, h, etiquette=etiquette, afficher_etiquette=etiquette is not None,
                 taille_valeur=taille, couleur_valeur=couleur_txt, couleur_etiquette=couleur_etiquette,
                 taille_etiquette=taille_etiquette, gras=gras, decimales=decimales, fond=fond, cle=cle,
                 marges=(0, 0, 0, 0), **opts)


def rectangle(x, y, w, h, cle, remplissage="FFFFFF", contour=LINE, arrondi=0):
    return forme_rectangle(x, y, w, h, remplissage, contour=contour, arrondi=arrondi, fond=False, cle=cle)


def section(p, texte_ou_champ, y, cle, largeur=UTILE):
    """Titre de section (comme les en-têtes de l'Excel) et filet sous le titre."""
    if hasattr(texte_ou_champ, "expr"):
        p.ajouter(texte_mesure(texte_ou_champ, M, y, 760, 26, cle + "_titre", taille=11, couleur_txt=NAVY,
                               gras=True))
    else:
        p.ajouter(ligne_texte(texte_ou_champ, M, y, largeur - 230 if largeur > 900 else largeur, 26,
                              cle + "_titre", taille=11, couleur_txt=NAVY, gras=True))
    p.ajouter(rectangle(M, y + 27, largeur, 2, cle + "_filet", remplissage="B8C2CF", contour=None))
    return y + 36


def bandeau(p, titre, sous_titre_txt, cle="bandeau", ligne3=None, droite=None, boutons=()):
    """Bandeau bleu nuit : titre, sous-titre, 3e ligne dynamique ; à droite période / date de lecture / boutons."""
    p.ajouter(forme_rectangle(0, 0, LARGEUR, BANDEAU_H, NAVY, fond=False, cle=cle))
    p.ajouter(texte([paragraphe(dict(texte=titre, gras=True, couleur="FFFFFF", taille=20))],
                    M, 8, 820, 36, cle + "_titre"))
    p.ajouter(ligne_texte(sous_titre_txt, M, 44, 820, 20, cle + "_soustitre", couleur_txt=BANDEAU_TXT))
    if ligne3 is not None:
        p.ajouter(texte_mesure(ligne3, M, 66, 820, 20, cle + "_ligne3", couleur_txt=BANDEAU_TXT, taille=9.5,
                               gras=True))
    for i, (champ, taille, coul, gras) in enumerate(droite or []):
        y, h = (6, 28) if i == 0 else (34, 22)
        p.ajouter(valeur(champ, 852, y, 412, h, "%s_droite%d" % (cle, i), taille=taille, couleur_txt=coul,
                         gras=gras))
    n = len(boutons)
    for i, (txt, cible) in enumerate(boutons):
        w = 132
        x = LARGEUR - M - (n - i) * w - (n - 1 - i) * 8
        p.ajouter(bouton(txt, x, 60, w, 26, page_cible=cible, remplissage="33476E", taille_texte=9,
                         arrondi=4, info="Aller à la page « %s » (dans Power BI Desktop : Ctrl + clic)"
                                          % txt.rstrip(" ›"), fond=False,
                         cle="%s_nav%d" % (cle, i)))


def filtres_reglages():
    """Périmètre et répartition « Comme dans Excel » : filtres du rapport (volet Filtres, replié), pas de segment
    visible : le lecteur du rapport n'a pas à connaître ces réglages internes."""
    return [filtre_categoriel(col("Périmètre", "Périmètre"), [EXCEL], selection_unique=True, nom="filtre_perimetre"),
            filtre_categoriel(col("Répartition", "Répartition"), [EXCEL], selection_unique=True,
                              nom="filtre_repartition")]


def filtre_periode():
    """Les 12 mois du Dashboard Excel (filtre de page, modifiable dans le volet Filtres)."""
    return filtre_categoriel(col("Mois", "Période"), [PERIODE_EXCEL], nom="filtre_periode")


def segments_filtres(p, y, masque=False, h=52):
    """Segments Chef de projet, Projet, Type d'affaire (listes déroulantes) synchronisés entre les pages ;
    masque : copie invisible (pages PDF) qui applique la sélection faite à l'écran."""
    w3 = (UTILE - 2 * G) / 3
    for k, (colonne, titre, groupe, cle) in enumerate((("Chef de projet", "Chef de projet", SYNC_CHEF, "seg_chef"),
                                                        ("Projet", "Projet", SYNC_PROJET, "seg_projet"),
                                                        ("Type", "Type d'affaire", SYNC_TYPE, "seg_type"))):
        v = p.ajouter(segment(col("Lignes", colonne), M + k * (w3 + G), y, w3, h, titre=titre, mode="Dropdown",
                              groupe_sync=groupe, cle=cle))
        v.masque = masque
    return y + h


def tuile(p, xw, y, h, cle, titre, champ, sous1, sous2, couleur_val=None, accent=ACCENT, taille=22, aide=None,
          couleurs_sous=(None, None)):
    """Tuile d'indicateur (comme les tuiles de l'Excel) : filet d'accent, titre, valeur, deux lignes.
    xw = (x, largeur) ; couleurs_sous : couleur (règle) des deux lignes dynamiques, gris par défaut."""
    x, w = xw
    p.ajouter(rectangle(x, y, w, h, cle + "_fond"))
    p.ajouter(rectangle(x, y, w, 4, cle + "_accent", remplissage=accent, contour=None))
    p.ajouter(ligne_texte(titre, x + 8, y + 10, w - 16, 20, cle + "_titre", taille=9, gras=True,
                          alignement_txt="center"))
    compact = h < 116           # tuiles des pages PDF
    y_val, h_val, y_sous, pas = (30, 34, 64, 19) if compact else (32, 40, 74, 22)
    v = p.ajouter(valeur(champ, x + 8, y + y_val, w - 16, h_val, cle + "_valeur", taille=taille, couleur_txt=NAVY,
                         gras=True, aide=aide))
    if couleur_val is not None:
        couleur_valeur_carte(v, couleur_val)
    for i, s in enumerate((sous1, sous2)):
        yy = y + y_sous + pas * i
        if s is None:
            continue
        if isinstance(s, str):
            p.ajouter(ligne_texte(s, x + 8, yy, w - 16, pas - 2, "%s_sous%d" % (cle, i), taille=8.5,
                                  alignement_txt="center"))
        else:
            vs = p.ajouter(valeur(s, x + 8, yy, w - 16, pas - 2, "%s_sous%d" % (cle, i), taille=8.5))
            if couleurs_sous[i] is not None:
                couleur_valeur_carte(vs, couleurs_sous[i])
    return v


def rouge_si_positif(champ):
    return regle((condition(champ, ">", 0.0), RED), (condition(champ, "<=", 0.0), NAVY))


def tableau_projets(x, y, w, h, cle, avec_part=True, titre=None):
    """Tableau légende des projets : n° (pastille de la couleur du projet), projet, heures, part."""
    rang, projet = col("Lignes", "Ordre groupe projet"), col("Lignes", "Groupe projet")
    cols = [ColonneTableau(rang, 28, "N°", alignement="Center"),
            ColonneTableau(projet, w - 28 - 60 - (52 if avec_part else 0) - 16, "Projet"),
            ColonneTableau(m("Charge identifiée (h)"), 60, "Heures")]
    if avec_part:
        cols.append(ColonneTableau(m("Part de la charge (%)"), 52, "Part"))
    v = tableau(cols, x, y, w, h, titre=titre, tri=[(rang, "Ascending")], taille_texte=9, cle=cle,
                fond=False, entete_fond="FFFFFF", entete_texte=GREY, entete=False,
                page_infobulle=P_IB_PROJET)
    teinte = couleur_champ(col("Lignes", "Couleur projet"))
    couleur_colonne(v, rang, police=couleur("FFFFFF"), fond=teinte)
    charge = m("Charge identifiée (h)")
    v.filtres.append(filtre_avance(charge, filtre_comparaison(charge, ">", 0.0), nom=cle + "_sans_zero",
                                   masque=True))
    return v


# ---------------------------------------------------------------------------------------------- page 1
def bloc_cartes(p, y, H=284, alerte=True):
    """Section 1 : titre du périmètre et les quatre cartes (chefs, programmes, charge, plus gros reste)."""
    section(p, m("Titre du périmètre"), y, "s1")
    if alerte:
        p.ajouter(ligne_texte("Lignes de Charge à corriger (dessinateur ET ressource externe) :", 780, y, 420, 26,
                              "s1_alerte_txt", taille=9, gras=True, alignement_txt="right"))
        v = p.ajouter(valeur(m("Lignes avec dessinateur et externe"), 1204, y, 60, 26, "s1_alerte_val", taille=12,
                             gras=True, couleur_txt=GREEN,
                             aide="Une ligne de l'onglet Charge = une ressource : seule la colonne K (dessinateur) "
                                  "est comptée quand K et L sont remplies."))
        couleur_valeur_carte(v, regle((condition(m("Lignes avec dessinateur et externe"), ">", 0.0), RED),
                                      (condition(m("Lignes avec dessinateur et externe"), "<=", 0.0), GREEN)))
    y += 36
    cartes = grille(4)
    # carte 1 : chefs de projet
    x, w = cartes[0]
    p.ajouter(rectangle(x, y, w, H, "c1_fond"))
    p.ajouter(ligne_texte("CHEFS DE PROJET", x + 8, y + 10, w - 16, 20, "c1_titre", gras=True,
                          alignement_txt="center"))
    p.ajouter(valeur(m("Nombre de chefs de projet"), x + 8, y + 32, w - 16, 42, "c1_valeur", taille=24,
                     couleur_txt=NAVY, gras=True))
    p.ajouter(valeur(m("Sous-titre chefs"), x + 8, y + 76, w - 16, 20, "c1_sous", taille=8.5))
    p.ajouter(graphique_barres(col("Lignes", "Chef de projet"), [m("Charge identifiée (h)")], x + 4, y + 102,
                               w - 8, H - 108, couleur_barres=BARRE_LISTE, fond=False, cle="c1_barres",
                               infobulles=[m("Nombre d'affaires"), m("Reste à faire (h)")],
                               texte_alt="Charge identifiée par chef de projet, par charge décroissante"))
    props_sans_selecteur(_par_cle(p, "c1_barres"), "categoryAxis", {"maxMarginFactor": Entier(45)})
    # carte 2 : programmes suivis
    x, w = cartes[1]
    p.ajouter(rectangle(x, y, w, H, "c2_fond"))
    p.ajouter(ligne_texte("PROGRAMMES SUIVIS", x + 8, y + 10, w - 16, 20, "c2_titre", gras=True,
                          alignement_txt="center"))
    p.ajouter(valeur(m("Nombre de projets"), x + 8, y + 32, w - 16, 42, "c2_valeur", taille=24, couleur_txt=NAVY,
                     gras=True))
    p.ajouter(ligne_texte("projets distincts · classés par charge", x + 8, y + 76, w - 16, 20, "c2_sous",
                          taille=8.5, alignement_txt="center"))
    v = p.ajouter(graphique_barres(col("Lignes", "Groupe projet"), [m("Charge identifiée (h)")], x + 4, y + 102,
                                   w - 8, H - 108, fond=False, cle="c2_barres",
                                   infobulles=[m("Part de la charge (%)"), m("Nombre d'affaires")],
                                   texte_alt="Charge identifiée par projet, dans la couleur du projet"))
    v.objet("dataPoint", {"fill": couleur_champ(col("Lignes", "Couleur projet"))}, sel_par_point())
    # carte 3 : charge identifiée
    x, w = cartes[2]
    p.ajouter(rectangle(x, y, w, H, "c3_fond"))
    p.ajouter(ligne_texte("CHARGE IDENTIFIÉE (H)", x + 8, y + 10, w - 16, 20, "c3_titre", gras=True,
                          alignement_txt="center"))
    p.ajouter(valeur(m("Charge identifiée (h)"), x + 8, y + 32, w - 16, 42, "c3_valeur", taille=24,
                     couleur_txt=NAVY, gras=True,
                     aide="Somme des heures (colonne M) des lignes du périmètre, toutes dates confondues."))
    p.ajouter(valeur(m("Sous-titre charge identifiée"), x + 8, y + 76, w - 16, 20, "c3_sous", taille=8.5))
    wd = (w - 24) / 2
    p.ajouter(valeur(m("Réalisé (h)"), x + 8, y + 100, wd, 46, "c3_realise", taille=13, couleur_txt=INK, gras=True,
                     etiquette="Réalisé (selon avancement)", decimales=0))
    p.ajouter(valeur(m("Reste à faire (h)"), x + 16 + wd, y + 100, wd, 46, "c3_reste", taille=13, couleur_txt=INK,
                     gras=True, etiquette="Reste à faire", decimales=0))
    p.ajouter(graphique_barres(col("Lignes", "Type"), [m("Charge identifiée (h)")], x + 4, y + 150, w - 8,
                               H - 156, titre="Par type d'affaire (h)", style_titre={"fontSize": 9, "fontColor": GREY},
                               couleur_barres=BARRE_LISTE, fond=False, cle="c3_types",
                               infobulles=[m("Part de la charge (%)"), m("Nombre d'affaires")]))
    props_sans_selecteur(_par_cle(p, "c3_types"), "categoryAxis", {"maxMarginFactor": Entier(45)})
    # carte 4 : plus gros reste à faire (mise en avant, fond bleu nuit)
    x, w = cartes[3]
    p.ajouter(rectangle(x, y, w, H, "c4_fond", remplissage=NAVY, contour=NAVY))
    p.ajouter(rectangle(x, y, w, 4, "c4_accent", remplissage=AMBER, contour=None))
    p.ajouter(ligne_texte("★  PLUS GROS RESTE À FAIRE (H)", x + 8, y + 10, w - 16, 20, "c4_titre", gras=True,
                          couleur_txt=AMBER, alignement_txt="center"))
    p.ajouter(valeur(m("Plus gros reste à faire (h)"), x + 8, y + 30, w - 16, 40, "c4_valeur", taille=24,
                     couleur_txt="FFFFFF", gras=True))
    p.ajouter(valeur(m("Sous-titre plus gros reste"), x + 8, y + 70, w - 16, 20, "c4_sous", taille=8.5,
                     couleur_txt=BANDEAU_TXT))
    details = [("Projet", "Plus gros reste à faire – projet"), ("Sous-projet", "Plus gros reste à faire – sous-projet"),
               ("Produit", "Plus gros reste à faire – produit"),
               ("Chef de projet", "Plus gros reste à faire – chef de projet"),
               ("Ressources", "Plus gros reste à faire – ressources"),
               ("Début planifié", "Plus gros reste à faire – début planifié"),
               ("Date de besoin", "Plus gros reste à faire – date de besoin"),
               ("Avancement", "Plus gros reste à faire – avancement")]
    retard = m("Plus gros reste à faire – retard")
    p.ajouter(fiche_valeurs([(lib, m(n)) for lib, n in details], x + 10, y + 94, w - 20, H - 100, "c4_details",
                            fond_cellules=NAVY, texte_libelles="AEB9CC", texte_valeurs="FFFFFF", taille=8.5,
                            couleurs={len(details) - 1: regle((condition(retard, ">", 0.0), "F97066"),
                                                              (condition(retard, "<=", 0.0), "FFFFFF"))}))

    return y + H

def bloc_indicateurs(p, y, TH=120):
    """Section 2 : les cinq tuiles d'indicateurs."""
    y = section(p, "INDICATEURS CLÉS", y, "s2")
    tuiles = grille(5)
    tuile(p, tuiles[0], y, TH, "t1", "MOIS EN SURCHARGE", m("Mois en surcharge"), m("Sous-titre capacité"),
          m("Sous-titre surcharge"), couleur_val=rouge_si_positif(m("Mois en surcharge")),
          aide="Mois où la charge répartie (arrondie) dépasse la capacité projets de l'équipe (arrondie).")
    tuile(p, tuiles[1], y, TH, "t2", "PIC DE CHARGE (H)", m("Pic de charge (h)"), m("Sous-titre pic (mois)"),
          m("Sous-titre pic (affaires)"))
    tuile(p, tuiles[2], y, TH, "t3", "AFFAIRES EN RETARD", m("Affaires en retard"), m("Sous-titre hors délai"),
          m("Sous-titre écart"), couleur_val=rouge_si_positif(m("Affaires en retard")),
          aide="Hors délai : date de besoin dépassée sans être à 100 %. Écart : avancement inférieur de plus de "
               "15 points à l'avancement attendu à la date de l'actualisation.")
    tuile(p, tuiles[3], y, TH, "t4", "NOMBRE D'AFFAIRES", m("Nombre d'affaires"), m("Sous-titre projets"),
          m("Sous-titre lignes terminées"))
    tuile(p, tuiles[4], y, TH, "t5", "TAUX D'AVANCEMENT MOYEN", m("Avancement moyen pondéré"),
          "pondéré par les heures des affaires", m("Sous-titre avancement"))

    return y + TH

def bloc_graphiques(p, y, GH=392):
    """Section 3 : histogramme mensuel par projet + capacité, anneau et tableau des projets."""
    y = section(p, "RÉPARTITION DE LA CHARGE", y, "s3")
    wl = 724
    p.ajouter(rectangle(M, y, wl, GH, "g1_fond"))
    v = p.ajouter(graphique_combo(
        col("Mois", "Libellé mois"), [m("Charge répartie (h)")], M + 8, y + 6, wl - 16, GH - 44,
        serie=col("Lignes", "Ordre groupe projet"), lignes=[m("Capacité projets de l'équipe (h/mois)")],
        titre="Charge en heures par mois et par projet",
        couleurs_series={Entier(k + 1): c for k, c in enumerate(SLOT_COLORS)},
        couleurs_lignes={"Capacité projets de l'équipe (h/mois)": CAP_C}, legende=None, fond=False,
        page_infobulle=P_IB_PROJET, cle="g1_combo",
        texte_alt="Heures réparties par mois, empilées par projet (couleurs du tableau des projets), et capacité "
                  "projets de l'équipe en pointillés"))
    v.roles["Series"][0]["displayName"] = "N° de projet"
    sous_titre(v, m("Sous-titre graphique mensuel"))
    v.objet("totals", {"show": True, "fontSize": 8.0, "color": couleur("374151"), "bold": True,
                       "labelDisplayUnits": 1.0, "labelPrecision": Entier(0)})
    p.ajouter(texte([paragraphe(dict(texte="– – –", gras=True, taille=10, couleur=CAP_C),
                                dict(texte="  pointillés :", taille=9, couleur=GREY))],
                    M + 8, y + GH - 34, 104, 24, "g1_pointilles"))
    p.ajouter(texte_mesure(m("Sous-titre capacité"), M + 112, y + GH - 30, 300, 22, "g1_capacite", taille=9,
                           couleur_txt=INK, gras=True))
    p.ajouter(valeur(m("Texte hors période"), M + 420, y + GH - 36, wl - 428, 28, "g1_hors", taille=9))
    xr, wr = M + wl + G, UTILE - wl - G
    p.ajouter(rectangle(xr, y, wr, GH, "g2_fond"))
    p.ajouter(ligne_texte("Répartition de la charge par projet", xr + 8, y + 8, wr - 16, 24, "g2_titre", taille=11,
                          couleur_txt=NAVY, gras=True))
    p.ajouter(anneau(col("Lignes", "Ordre groupe projet"), m("Charge identifiée (h)"), xr + 8, y + 40, 220, GH - 96,
                     couleurs={Entier(k + 1): c for k, c in enumerate(SLOT_COLORS)},
                     style_etiquettes="Percent of total", position_etiquettes="preferOutside", legende=False,
                     total=m("Charge identifiée (h)"), libelle_total="h au total", fond=False,
                     page_infobulle=P_IB_PROJET, cle="g2_anneau",
                     texte_alt="Part de chaque projet dans la charge identifiée ; total au centre"))
    props_sans_selecteur(_par_cle(p, "g2_anneau"), "labels", {"percentageLabelPrecision": Entier(0)})
    p.ajouter(tableau_projets(xr + 228, y + 44, wr - 236, GH - 100, "g2_legende"))
    p.ajouter(texte([paragraphe(dict(texte="Total au centre · une couleur par projet, commune aux deux graphiques "
                                     "(n° 1 à 9) · au-delà de 9 projets, « Autres projets » en gris.",
                                     taille=8.5, couleur=GREY, italique=True))],
                    xr + 8, y + GH - 44, wr - 16, 36, "g2_note"))

    return y + GH

def bloc_liste_affaires(p, y, h=468, bouton_fiche=True, n_max=None, titre_section=None, taille_texte=None):
    """Section 4 : tableau des affaires (n_max : n° d'affaire maximal affiché, pour le PDF)."""
    if bouton_fiche:
        p.ajouter(bouton("Ouvrir la fiche de l'affaire", LARGEUR - M - 220, y - 2, 220, 26, action="Drillthrough",
                         page_cible=P_FICHE, remplissage=ACCENT, taille_texte=9, arrondi=4,
                         info="Sélectionner une affaire dans le tableau, puis cliquer (Ctrl + clic dans Power BI "
                              "Desktop), ou clic droit sur l'affaire > Extraire",
                         fond=False, cle="s4_extraction"))
    y = section(p, titre_section or "LISTE DES AFFAIRES (un sous-projet par ligne, toutes ressources regroupées, "
                   "par charge décroissante)", y, "s4")
    p.ajouter(tableau_affaires(M, y, UTILE, h, n_max=n_max, taille_texte=taille_texte))
    return y + h



def page_tableau_de_bord(r):
    p = r.page("Tableau de bord", nom=P_TDB, largeur=LARGEUR, hauteur=1806, fond=BG, papier_peint=BG,
               filtres=[filtre_periode()])
    bandeau(p, "Plan de charge du bureau d'études",
            "Synthèse de l'onglet Charge du classeur Excel : saisir dans Charge et Data, enregistrer le classeur, "
            "puis actualiser le rapport.",
            ligne3=m("Rappel des réglages"),
            droite=[(m("Période affichée"), 12, "FFFFFF", True), (m("Données lues le"), 9, BANDEAU_TXT, False)],
            boutons=[("Ressources ›", P_RES), ("Contrôle ›", P_CTRL)])

    y = segments_filtres(p, BANDEAU_H + 12) + 16
    y = bloc_cartes(p, y) + 16
    y = bloc_indicateurs(p, y) + 16
    y = bloc_graphiques(p, y) + 16
    y = bloc_liste_affaires(p, y) + 16
    notes = [
        "Affaire = un sous-projet (projet + sous-projet + produit) : ses lignes de l'onglet Charge sont additionnées ; "
        "chef = celui de sa 1re ligne ; avancement = toutes ses lignes pondérées par leurs heures ; reste à faire = "
        "heures × (1 − avancement).",
        "Périmètre « Projets en cours » = lignes dont l'avancement (colonne O) est inférieur à 100 % : une ressource "
        "qui a terminé sa part sort du périmètre, même si d'autres continuent l'affaire. Périmètre, répartition et "
        "période sont ceux enregistrés dans le classeur (volet Filtres pour les changer).",
        "Charge par mois : heures (M) réparties au prorata des jours ouvrés (lundi au vendredi, hors jours non "
        "travaillés de l'onglet Data) entre le début planifié (N) et la date de besoin (F) ; sans date de besoin "
        "valable, tout va sur le mois de début.",
        "Affaire en retard (avancement en rouge) : date de besoin dépassée sans être à 100 %, ou avancement inférieur "
        "de plus de 15 points à l'avancement attendu. La date du jour est celle de la dernière actualisation.",
        "Pointillés : capacité projets de l'équipe = somme des (capacité − charge fixe) des ressources des listes "
        "Data. ETP = heures ÷ (capacité − charge fixe) d'une personne. Au-delà de 9 projets : « Autres projets » (gris).",
        "Cliquer sur une barre, une part ou une ligne filtre les autres visuels (cliquer à nouveau pour annuler). "
        "Clic droit sur une affaire > Extraire > Fiche affaire. Les chiffres sont ceux du classeur enregistré, lu "
        "à la dernière actualisation.",
    ]
    p.ajouter(notes_lecture(notes, M, y, UTILE, 150, "notes"))
    return p


def notes_lecture(notes, x, y, w, h, cle):
    paras = [paragraphe(dict(texte="NOTES DE LECTURE", gras=True, taille=8.5, couleur=GREY))]
    paras += [paragraphe(dict(texte="•  " + n, taille=8, couleur=GREY, italique=True)) for n in notes]
    return texte(paras, x, y, w, h, cle)


def fiche_valeurs(champs, x, y, w, h, cle, titre=None, fond_cellules="FFFFFF", texte_libelles=GREY,
                  texte_valeurs=INK, taille=9, couleurs=None, part_libelles=0.45):
    """Liste « libellé : valeur » verticale : une zone de texte par libellé et une carte à plusieurs lignes par
    valeur (une matrice à valeurs en lignes s'affichait mal dans Desktop : fond et couleurs ignorés, défilement).
    couleurs : {indice du champ: Couleur} pour colorer une valeur (règle ou mesure). Renvoie la liste des visuels."""
    visuels = []
    if titre:
        visuels.append(rectangle(x, y, w, h, cle + "_fond", remplissage=fond_cellules))
        visuels.append(ligne_texte(titre, x + 6, y + 6, w - 12, 24, cle + "_titre", taille=11, couleur_txt=NAVY,
                                   gras=True))
        y, h = y + 34, h - 38
    pas = h / max(1, len(champs))
    wl = (w - 12) * part_libelles
    for k, (lib, champ) in enumerate(champs):
        yk = y + k * pas
        visuels.append(ligne_texte(lib, x + 6, yk, wl, pas, "%s_lib%d" % (cle, k), taille=taille,
                                   couleur_txt=texte_libelles))
        v = texte_mesure(champ, x + 6 + wl, yk + 2, w - 12 - wl, pas - 2, "%s_val%d" % (cle, k), taille=taille,
                         couleur_txt=texte_valeurs, gras=True, nom_affiche=lib)
        if couleurs and k in couleurs:
            v.objets["dataLabels"][0]["properties"]["color"] = couleurs[k].pbir()
        visuels.append(v)
    return visuels


def tableau_affaires(x, y, w, h, n_max=None, taille_texte=None):
    A = "Affaires"
    coul = m("Couleur avancement")
    cols = [
        ColonneTableau(col(A, "N°"), 40, "N°", alignement="Center"),
        ColonneTableau(col(A, "Affaire"), 300, "Affaire (projet · sous-projet · produit)"),
        ColonneTableau(col(A, "Chef de projet"), 120, "Chef de projet"),
        ColonneTableau(col(A, "Ressources"), 96, "Ressources", alignement="Center"),
        ColonneTableau(col(A, "Charge (h)"), 76, "Charge (h)"),
        ColonneTableau(col(A, "% de la charge totale"), 100, "% de la charge totale"),
        ColonneTableau(col(A, "Début (affiché)"), 86, "Début planifié", alignement="Center"),
        ColonneTableau(col(A, "Date de besoin"), 86, "Date de besoin", alignement="Center"),
        ColonneTableau(col(A, "Avancement"), 96, "Avancement", couleur_police=coul),
        ColonneTableau(col(A, "Reste à faire de l'affaire (h)"), 84, "Reste à faire (h)"),
        ColonneTableau(col(A, "Statut"), 150, "Statut (retard)", couleur_police=coul),
    ]
    v = tableau(cols, x, y, w, h, titre="Affaires du périmètre (clic droit sur une ligne > Extraire > Fiche affaire)",
                tri=[(col(A, "N°"), "Ascending")], cle="tableau_affaires",
                **({} if taille_texte is None else {"taille_texte": taille_texte}),
                aide="Une ligne par affaire. Charge et % : lignes de l'affaire dans le périmètre. Avancement et reste "
                     "à faire : toutes les lignes de l'affaire. Rouge : affaire en retard.")
    barres_donnees(v, col(A, "% de la charge totale"), "86B6EF")
    barres_donnees(v, col(A, "Avancement"), "7FD1AE", maximum=1)
    if n_max is not None:       # page PDF : les n_max premières affaires (n° = rang par charge décroissante)
        v.filtres.append(filtre_avance(col(A, "N°"), filtre_comparaison(col(A, "N°"), "<=", int(n_max)),
                                       nom="tableau_affaires_n_max", masque=True))
    return v


# ---------------------------------------------------------------------------------------------- page 2
def bloc_occupation(p, y, HH=330):
    """Occupation de l'équipe : carte de chaleur ressources × mois et heures non affectées."""
    y = section(p, "OCCUPATION DE L'ÉQUIPE (charge fixe + projets, en % de la capacité)", y, "s1")
    p.ajouter(texte([paragraphe(
        dict(texte="■ ", couleur="1E7B34", taille=10), dict(texte="Vert : moins de 90 %     ", taille=8.5, couleur=GREY),
        dict(texte="■ ", couleur="EDA100", taille=10), dict(texte="Orange : de 90 à 100 %     ", taille=8.5, couleur=GREY),
        dict(texte="■ ", couleur="B42318", taille=10),
        dict(texte="Rouge : au-delà de 100 % (surcharge)     ·     ressources des listes Data (dessinateurs puis "
                   "externes), puis ressources saisies dans Charge hors listes (réglages par défaut)",
             taille=8.5, couleur=GREY))], M, y, UTILE, 22, "legende_occ"))
    y += 26
    occ = m("Occupation (%)")
    heat = matrice([col("Ressources", "Ressource (avec mention)")], [col("Mois", "Libellé mois")], [occ], M, y, UTILE,
                   HH,
                   titre="Occupation par ressource et par mois", couleur_fond=m("Couleur de fond (occupation)"),
                   couleur_police=m("Couleur du texte (occupation)"), totaux_colonnes=True,
                   largeurs={col("Ressources", "Ressource (avec mention)"): 180}, cle="carte_chaleur",
                   page_infobulle=P_IB_OCC,
                   aide="Survoler une case pour le détail du mois. Colonne Moyenne : moyenne des mois affichés.",
                   filtres=[filtre_categoriel(col("Ressources", "Type de ressource"),
                                              ["Interne", "Externe", "Hors liste"], masque=True,
                                              nom="filtre_occ_types")])
    heat.roles["Rows"][0]["displayName"] = "Ressource"
    props_sans_selecteur(heat, "subTotals", {"columnSubtotalsLabel": "Moyenne"})
    props_sans_selecteur(heat, "values", {"bold": True})
    p.ajouter(heat)
    y += HH + 12
    na = matrice([col("Ressources", "Ressource")], [col("Mois", "Libellé mois")], [m("Charge répartie (h)")],
                 M, y, UTILE, 84, titre="Heures non affectées à une ressource (lignes sans dessinateur ni externe)",
                 largeurs={col("Ressources", "Ressource"): 180}, cle="non_affecte",
                 filtres=[filtre_categoriel(col("Ressources", "Type de ressource"), ["Non affecté"], masque=True,
                                            nom="filtre_non_affecte")])
    p.ajouter(na)
    return y + 84, heat, na



def bloc_choix_ressource(p, y, masque_segment=False):
    """Section « ressource choisie » : segment Ressource (synchronisé) et caractéristiques de la ressource ;
    masque_segment : copie invisible du segment (page PDF), le nom de la ressource est alors affiché."""
    y = section(p, "RESSOURCE CHOISIE", y, "s2")
    seg_res = p.ajouter(segment(col("Ressources", "Ressource"), M, y, 300, 56, titre="Ressource", mode="Dropdown",
                                selection_unique=True, cle="seg_ressource", groupe_sync=SYNC_RES,
                                aide="Dessinateurs (Data E) puis ressources externes (Data H).",
                                filtres=[filtre_categoriel(col("Ressources", "Type de ressource"),
                                                           ["Interne", "Externe", "Hors liste"], masque=True,
                                                           nom="filtre_seg_ressource")]))
    seg_res.masque = masque_segment
    if masque_segment:
        p.ajouter(rectangle(M, y, 300, 56, "nom_ressource_fond"))
        p.ajouter(ligne_texte("Ressource", M + 4, y + 4, 292, 18, "nom_ressource_lib", taille=9, gras=True,
                              couleur_txt=INK))
        p.ajouter(texte_mesure(col("Ressources", "Ressource"), M + 8, y + 24, 288, 28, "nom_ressource", taille=13,
                               couleur_txt=NAVY, gras=True))
    xi, wi = M + 300 + G, UTILE - 300 - G
    p.ajouter(rectangle(xi, y, wi, 56, "info_fond"))
    p.ajouter(rectangle(xi, y, 4, 56, "info_accent", remplissage=ACCENT, contour=None))
    infos = (("Origine", "Origine de la ressource", 2.2), ("Capacité (h/mois)", "Capacité de la ressource (h/mois)", 1),
             ("Charge fixe hors projet (h/mois)", "Charge fixe de la ressource (h/mois)", 1),
             ("Disponible pour les projets (h/mois)", "Capacité projets de la ressource (h/mois)", 1))
    total = sum(k for _l, _n, k in infos)
    xk = xi + 12
    for i_info, (lib, n, k) in enumerate(infos):
        wk = (wi - 20) * k / total
        p.ajouter(valeur(m(n), xk, y + 4, wk - 8, 48, "info_ressource%d" % i_info, taille=11, couleur_txt=NAVY,
                         gras=True, etiquette=lib))
        xk += wk
    return y + 56, seg_res


def bloc_detail_ressource(p, y, TH=120, CH=330, TABH=196):
    """Tuiles, graphique mensuel et tableau mensuel de la ressource choisie."""
    tuiles = grille(5)
    occ = m("Occupation (%)")
    # « Attention : charge fixe ≥ capacité » en rouge (capacité projets = max(0, capacité - charge fixe) = 0)
    cap_proj = m("Capacité projets de la ressource (h/mois)")
    tuile(p, tuiles[0], y, TH, "t1", "CAPACITÉ MENSUELLE (H)", m("Capacité de la ressource (h/mois)"),
          m("Sous-titre charge fixe de la ressource"), m("Sous-titre capacité de la ressource"),
          couleurs_sous=(None, regle((condition(cap_proj, "<=", 0.0), RED), (condition(cap_proj, ">", 0.0), GREY))))
    tuile(p, tuiles[1], y, TH, "t2", "CHARGE PROJETS AFFECTÉE (H)", m("Charge identifiée (h)"),
          m("Sous-titre affaires de la ressource"), m("Sous-titre charge sur la période"))
    occ_pic = m("Occupation au pic de la ressource (%)")
    tuile(p, tuiles[2], y, TH, "t3", "MOIS LE PLUS CHARGÉ", m("Mois le plus chargé de la ressource (texte)"),
          m("Sous-titre pic de la ressource"), m("Sous-titre occupation au pic"), taille=18,
          couleur_val=regle((condition(occ_pic, ">", 1.0), RED), (condition(occ_pic, "<=", 1.0), NAVY)),
          couleurs_sous=(None, regle((condition(occ_pic, ">", 1.0), RED), (condition(occ_pic, "<=", 1.0), GREY))))
    tuile(p, tuiles[3], y, TH, "t4", "MOIS EN SURCHARGE", m("Mois en surcharge de la ressource"),
          m("Sous-titre dépassement cumulé"), m("Sous-titre surcharge de la ressource"),
          couleur_val=regle((condition(m("Mois en surcharge de la ressource"), ">", 0.0), RED),
                            (condition(m("Mois en surcharge de la ressource"), "<=", 0.0), GREEN)))
    tuile(p, tuiles[4], y, TH, "t5", "OCCUPATION MOYENNE", m("Occupation (%)"), m("Sous-titre heures disponibles"),
          "sur les mois affichés",
          couleur_val=regle((condition(occ, ">", 1.0), RED), (condition(occ, ">=", 0.9), "8A5A00"),
                            (condition(occ, "<", 0.9), GREEN)))
    y += TH + 12
    combo = graphique_combo(
        col("Mois", "Libellé mois"),
        [m("Charge fixe (h)"), m("Projets dans la capacité (h)"), m("Dépassement de capacité (h)")],
        M, y, UTILE, CH, lignes=[m("Capacité de la ressource (h/mois)")], titre="Charge mensuelle (heures)",
        couleurs_lignes={"Capacité de la ressource (h/mois)": CAP_C}, legende="Top", cle="combo_ressource",
        infobulles=[m("Charge totale de la ressource (h)"), occ],
        texte_alt="Charge mensuelle de la ressource choisie : charge fixe, projets dans la capacité, dépassement, "
                  "et capacité mensuelle en pointillés")
    for c, teinte in ((m("Charge fixe (h)"), FIX_C), (m("Projets dans la capacité (h)"), PROJ_C),
                      (m("Dépassement de capacité (h)"), OVER_C)):
        combo.objet("dataPoint", {"fill": couleur(teinte)}, sel_metadata(c))
    noms = {"Charge fixe (h)": "Charge fixe hors projet", "Projets dans la capacité (h)": "Charge projets",
            "Dépassement de capacité (h)": "Dépassement de capacité",
            "Capacité de la ressource (h/mois)": "Capacité mensuelle"}
    for role in ("Y", "Y2", "Tooltips"):
        for proj in combo.roles.get(role, []):
            if proj["nativeQueryRef"] in noms:
                proj["displayName"] = noms[proj["nativeQueryRef"]]
    titre_dynamique(combo, m("Titre de la ressource"))
    sous_titre(combo, m("Sous-titre graphique mensuel"))
    p.ajouter(combo)
    y += CH + 12
    # tableau mensuel (valeurs en lignes, mois en colonnes), comme sous le graphique de l'Excel
    lignes_tab = [("Charge fixe hors projet", m("Charge fixe (h)")), ("Charge projets", m("Charge répartie (h)")),
                  ("Charge totale", m("Charge totale de la ressource (h)")),
                  ("Capacité", m("Capacité de la ressource (h/mois)")),
                  ("Marge (+) / dépassement (−)", m("Marge de la ressource (h)")), ("Occupation", occ)]
    tab = matrice([], [col("Mois", "Libellé mois")], [c for _l, c in lignes_tab], M, y, UTILE, TABH,
                  titre="Heures par mois de la ressource choisie", cle="tableau_mensuel")
    for (lib, c), proj in zip(lignes_tab, tab.roles["Values"]):
        proj["displayName"] = lib
    props_sans_selecteur(tab, "values", {"valuesOnRow": True})
    marge = m("Marge de la ressource (h)")
    couleur_colonne(tab, marge, police=regle((condition(marge, "<", 0.0), RED), (condition(marge, ">", 0.0), GREEN),
                                             (condition(marge, "=", 0.0), INK)))
    couleur_colonne(tab, occ, police=couleur_champ(m("Couleur du texte (occupation)")),
                    fond=couleur_champ(m("Couleur de fond (occupation)")))
    props_sans_selecteur(tab, "rowHeaders", {"bold": True, "fontColor": couleur(INK)})
    p.ajouter(tab)
    return y + TABH



def bloc_lignes_ressource(p, y, h=300, bouton_fiche=True):
    """Lignes de l'onglet Charge de la ressource choisie."""
    if bouton_fiche:
        p.ajouter(bouton("Ouvrir la fiche de l'affaire", LARGEUR - M - 220, y - 2, 220, 26, action="Drillthrough",
                         page_cible=P_FICHE, remplissage=ACCENT, taille_texte=9, arrondi=4,
                         info="Sélectionner une ligne dans le tableau, puis cliquer (Ctrl + clic dans Power BI "
                              "Desktop), ou clic droit sur la ligne > Extraire",
                         fond=False, cle="s3_extraction"))
    y = section(p, "AFFAIRES DE LA RESSOURCE (ses lignes de l'onglet Charge, par charge décroissante)", y, "s3")
    L = "Lignes"
    cols = [
        ColonneTableau(col(L, "N° ligne"), 64, "Ligne Charge", alignement="Center"),
        ColonneTableau(col("Affaires", "Affaire"), 330, "Affaire (projet · sous-projet · produit)"),
        ColonneTableau(col(L, "Chef de projet"), 130, "Chef de projet"),
        ColonneTableau(col(L, "Type"), 140, "Type"),
        ColonneTableau(col(L, "Heures (h)"), 80, "Charge (h)"),
        # mesure vide seulement pour une combinaison ligne × affaire inexistante, 0 pour une ligne sans heure sur la
        # période (sans date de début, 0 h, hors des mois choisis) : toutes les lignes de la ressource restent listées
        ColonneTableau(m("Heures de la ligne sur la période (h)"), 100, "Heures sur la période"),
        ColonneTableau(col(L, "Début (affiché)"), 92, "Début planifié", alignement="Center"),
        ColonneTableau(col(L, "Date de besoin"), 92, "Date de besoin", alignement="Center"),
        ColonneTableau(col(L, "Avancement"), 90, "Avancement"),
    ]
    aff = tableau(cols, M, y, UTILE, h, titre="Lignes de la ressource choisie (clic droit > Extraire > Fiche affaire)",
                  tri=[(col(L, "Heures (h)"), "Descending")], cle="affaires_ressource")
    aff.objet("columnFormatting", {"labelPrecision": Entier(0)}, sel_metadata(m("Heures de la ligne sur la période (h)")))
    sous_titre(aff, m("Sous-titre lignes de la ressource"))
    barres_donnees(aff, col(L, "Avancement"), "7FD1AE", maximum=1)
    p.ajouter(aff)
    return y + h


def page_ressources(r):
    p = r.page("Ressources", nom=P_RES, largeur=LARGEUR, hauteur=1880, fond=BG, papier_peint=BG,
               filtres=[filtre_periode()])
    bandeau(p, "Charge par ressource (dessinateurs et externes)",
            "Charge fixe hors projet + heures de ses lignes dans l'onglet Charge, comparées à la capacité mensuelle "
            "de la personne.",
            ligne3=m("Rappel des réglages"),
            droite=[(m("Période affichée"), 12, "FFFFFF", True), (m("Données lues le"), 9, BANDEAU_TXT, False)],
            boutons=[("Tableau de bord ›", P_TDB), ("Contrôle ›", P_CTRL)])
    # la ressource se choisit en tête de page ; l'occupation de toute l'équipe vient ensuite
    y, seg_res = bloc_choix_ressource(p, BANDEAU_H + 12)
    y = bloc_detail_ressource(p, y + 12) + 16
    y = bloc_lignes_ressource(p, y) + 16
    y, heat, na = bloc_occupation(p, y)
    y += 16
    notes = [
        "Charge d'une ressource = charge fixe hors projet + heures des lignes de l'onglet Charge où elle figure "
        "(colonne K pour un dessinateur, L pour une ressource externe), réparties par mois comme sur le tableau de bord.",
        "Une ligne = une ressource : plusieurs personnes sur un même sous-projet = plusieurs lignes, chacune avec ses "
        "heures. Capacité et charge fixe : Data (F-G ou I-J), sinon valeurs par défaut du Dashboard Excel.",
        "Toutes les ressources des listes Data sont affichées (l'Excel n'en montre que 16) ; une ressource saisie dans "
        "Charge mais absente de Data apparaît avec les réglages par défaut. Lignes avec dessinateur ET externe : seule "
        "la colonne K compte.",
    ]
    p.ajouter(notes_lecture(notes, M, y, UTILE, 110, "notes"))
    # le segment Ressource ne filtre pas l'occupation de l'équipe ; l'occupation ne filtre pas le détail
    for cible in ("carte_chaleur", "non_affecte"):
        p.interaction(seg_res, _par_cle(p, cible), "NoFilter")
    for source in (heat, na):
        for v in p.visuels:
            if v is not source and v.type not in ("shape", "textbox", "actionButton"):
                p.interaction(source, v, "NoFilter")
    return p


def _par_cle(p, cle):
    for v in p.visuels:
        if v.cle == cle:
            return v
    raise KeyError(cle)


# ---------------------------------------------------------------------------------------------- page 3
def page_fiche(r):
    A, L = "Affaires", "Lignes"
    p = r.page("Fiche affaire", nom=P_FICHE, largeur=LARGEUR, hauteur=1146, fond=BG, papier_peint=BG,
               masquee=True, extraction=[col(A, "Affaire")], garder_tous_filtres=False)
    p.ajouter(forme_rectangle(0, 0, LARGEUR, BANDEAU_H, NAVY, fond=False, cle="bandeau"))
    p.ajouter(ligne_texte("FICHE AFFAIRE", M, 8, 600, 24, "bandeau_titre", taille=11, couleur_txt=AMBER, gras=True))
    p.ajouter(texte_mesure(col(A, "Affaire"), M, 32, 820, 32, "bandeau_affaire", taille=18, couleur_txt="FFFFFF",
                           gras=True))
    p.ajouter(texte_mesure(m("Rappel des réglages"), M, 66, 820, 20, "bandeau_reglages", taille=9.5,
                           couleur_txt=BANDEAU_TXT, gras=True))
    p.ajouter(valeur(m("Données lues le"), 852, 56, 412, 24, "bandeau_lecture", taille=9, couleur_txt=BANDEAU_TXT))
    p.ajouter(bouton("‹  Retour", LARGEUR - M - 120, 12, 120, 32, action="Back", icone="back", remplissage="33476E",
                     taille_texte=10, arrondi=4,
                     info="Revenir à la page précédente (dans Power BI Desktop : Ctrl + clic)", fond=False,
                     cle="retour"))
    y = BANDEAU_H + 8
    p.ajouter(ligne_texte("Fiche ouverte par clic droit sur une affaire > Extraire > Fiche affaire (tableau de bord ou "
                          "page Ressources) ; tous les mois de l'affaire sont affichés.", M, y, UTILE, 20, "aide",
                          taille=8.5, italique=True))
    y += 20 + 12
    # tuiles
    tuiles = grille(5)
    TH = 110
    coul = couleur_champ(m("Couleur avancement"))
    tuile(p, tuiles[0], y, TH, "t1", "AVANCEMENT", agr(A, "Avancement", "min"),
          "toutes les lignes, pondéré par les heures", None, couleur_val=coul)
    tuile(p, tuiles[1], y, TH, "t2", "AVANCEMENT ATTENDU", agr(A, "Avancement attendu", "min"),
          "prorata des jours ouvrés écoulés", None)
    tuile(p, tuiles[2], y, TH, "t3", "CHARGE DU PÉRIMÈTRE (H)", agr(A, "Charge (h)", "min"),
          "lignes de l'affaire dans le périmètre", None)
    tuile(p, tuiles[3], y, TH, "t4", "RESTE À FAIRE (H)", agr(A, "Reste à faire de l'affaire (h)", "min"),
          "toutes les lignes de l'affaire", None)
    tuile(p, tuiles[4], y, TH, "t5", "STATUT", agr(A, "Statut", "min"), "retard à la date de l'actualisation",
          None, couleur_val=coul, taille=12)
    y += TH + 16
    # carte d'identité + répartition mensuelle
    ident = [("Projet", agr(A, "Projet", "min")), ("Sous-projet", agr(A, "Sous-projet", "min")),
             ("Produit", agr(A, "Produit", "min")), ("Chef de projet", agr(A, "Chef de projet", "min")),
             ("Ressources", agr(A, "Ressources", "min")), ("Début planifié", agr(A, "Début planifié", "min")),
             ("Date de besoin", agr(A, "Date de besoin", "min")),
             ("Heures toutes lignes (h)", agr(A, "Heures toutes lignes (h)", "min")),
             ("Réalisé toutes lignes (h)", agr(A, "Réalisé toutes lignes (h)", "min")),
             ("Lignes dans le périmètre", agr(A, "Lignes de l'affaire", "min")),
             ("Personnes internes", agr(A, "Personnes internes", "min")),
             ("Personnes externes", agr(A, "Personnes externes", "min"))]
    p.ajouter(fiche_valeurs(ident, M, y, 400, 300, "identite", titre="Carte d'identité", texte_valeurs=INK))
    rep = graphique_barres(col("Mois", "Libellé mois"), [m("Charge répartie sans les mois vides (h)")], M + 400 + G, y,
                           UTILE - 400 - G,
                           300, type_visuel="columnChart", serie=col(L, "Ressource de la ligne"),
                           titre="Répartition mensuelle de l'affaire (h), par ressource", legende="Top",
                           tri=(col("Mois", "Libellé mois"), "Ascending"), etiquettes=False, cle="repartition",
                           montrer_axe_valeurs=True,
                           texte_alt="Heures de l'affaire réparties par mois, empilées par ressource")
    rep.roles["Y"][0]["displayName"] = "Heures"
    sous_titre(rep, "Répartition choisie dans le segment ; tous les mois où l'affaire a des heures")
    p.ajouter(rep)
    y += 300 + 12
    w2 = (UTILE - G) / 2
    p.ajouter(tableau([ColonneTableau(col(L, "Type"), w2 - 110, "Type d'affaire"),
                       ColonneTableau(m("Charge identifiée (h)"), 80, "Heures")],
                      M, y, w2, 140, titre="Type(s) d'affaire des lignes", tri=[(m("Charge identifiée (h)"), "Descending")],
                      cle="types"))
    p.ajouter(texte([
        paragraphe(dict(texte="Lecture", gras=True, taille=9, couleur=NAVY)),
        paragraphe(dict(texte="Avancement attendu = jours ouvrés écoulés depuis le début planifié ÷ jours ouvrés "
                              "jusqu'à la date de besoin. En retard si la date de besoin est dépassée sans être à "
                              "100 %, ou si l'avancement est inférieur de plus de 15 points à l'attendu.",
                        taille=8.5, couleur=GREY)),
        paragraphe(dict(texte="Les lignes terminées comptent dans l'avancement de l'affaire mais n'apparaissent "
                              "ci-dessous qu'avec le périmètre « Tous les projets ».", taille=8.5, couleur=GREY)),
    ], M + w2 + G, y, w2, 140, "lecture", fond="FFFFFF"))
    y += 140 + 16
    y = section(p, "LIGNES DE L'AFFAIRE (onglet Charge)", y, "s1")
    cols = [
        ColonneTableau(col(L, "N° ligne"), 52, "Ligne", alignement="Center"),
        ColonneTableau(col(L, "Ressource de la ligne"), 112, "Ressource"),
        ColonneTableau(col(L, "Type de ressource"), 64, "Interne / externe"),
        ColonneTableau(col(L, "Heures (h)"), 56, "Heures (h)"),
        ColonneTableau(col(L, "Avancement"), 76, "Avancement"),
        ColonneTableau(col(L, "Réalisé de la ligne (h)"), 60, "Réalisé (h)"),
        ColonneTableau(col(L, "Reste de la ligne (h)"), 56, "Reste (h)"),
        ColonneTableau(col(L, "Début (affiché)"), 76, "Début planifié", alignement="Center"),
        ColonneTableau(col(L, "Date de besoin"), 76, "Date de besoin", alignement="Center"),
        ColonneTableau(col(L, "Type"), 104, "Type"),
        ColonneTableau(col(L, "Coordinateur"), 90, "Coordinateur"),
        ColonneTableau(col(L, "Description"), 150, "Description"),
        ColonneTableau(col(L, "Contraintes"), 118, "Contraintes"),
        ColonneTableau(col(L, "Autres informations"), 118, "Autres informations"),
        ColonneTableau(col(L, "Dessinateur et externe"), 58, "À corriger (K et L)", alignement="Center"),
    ]
    lig = tableau(cols, M, y, UTILE, 300, titre="Lignes de l'affaire dans le périmètre",
                  tri=[(col(L, "N° ligne"), "Ascending")], cle="lignes")
    props_sans_selecteur(lig, "values", {"wordWrap": True})
    barres_donnees(lig, col(L, "Avancement"), "7FD1AE", maximum=1)
    p.ajouter(lig)
    y += 300 + 12
    p.ajouter(ligne_texte("Description, contraintes et autres informations : colonnes G, H et I de l'onglet Charge. "
                          "« À corriger » = ligne avec un dessinateur (K) ET une ressource externe (L) : seule K compte.",
                          M, y, UTILE, 28, "note", taille=8.5, italique=True))
    # ouverte sans extraction (onglet visible dans Desktop), la fiche mélangerait toutes les affaires : consigne et
    # visuels vides tant qu'une seule affaire n'est pas choisie
    choisie = m("Affaire choisie")
    for v in p.visuels:
        if v.type not in ("shape", "textbox", "actionButton", "slicer") and v.cle != "bandeau_lecture":
            v.filtres.append(filtre_avance(choisie, filtre_comparaison(choisie, "=", 1.0),
                                           nom="%s_affaire_choisie" % v.cle, masque=True))
    p.ajouter(texte_mesure(m("Message fiche affaire"), M, 34, 820, 28, "message", taille=12, couleur_txt="F97066",
                           gras=True))
    return p


# ---------------------------------------------------------------------------------------------- page 4
def page_controle(r):
    C = "Contrôle Excel"
    p = r.page("Contrôle", nom=P_CTRL, largeur=LARGEUR, hauteur=1130, fond=BG, papier_peint=BG)
    bandeau(p, "Contrôle : Power BI comparé au classeur Excel",
            "Chiffres recalculés par Power BI avec les réglages du classeur, comparés aux valeurs enregistrées dans "
            "l'onglet Calculs.",
            droite=[(m("Données lues le"), 10, "FFFFFF", True)],
            boutons=[("Tableau de bord ›", P_TDB), ("Ressources ›", P_RES)])
    y = BANDEAU_H + 12
    statut_n = m("Contrôles en écart")
    tuile(p, (M, 300), y, 120, "t_ecarts", "CONTRÔLES EN ÉCART", statut_n,
          "indicateurs et 12 mois comparés ci-dessous", "✔ = identique · ✘ = écart",
          couleur_val=regle((condition(statut_n, ">", 0.0), RED), (condition(statut_n, "<=", 0.0), GREEN)),
          accent=GREEN)
    reg = Visuel("multiRowCard", M + 300 + G, y, UTILE - 300 - G, 120)
    for lib, c in (("Réglages enregistrés dans le classeur", "Rappel des réglages (Excel)"),
                   ("Période du Dashboard Excel", "Période affichée (Excel)"),
                   ("Capacité projets de l'équipe (h/mois)", "Capacité projets de l'équipe (Excel)"),
                   ("Lignes valides de l'onglet Charge", "Lignes valides (Excel)")):
        reg.ajouter_role("Values", col("Réglages", c), nom_affiche=lib)
    reg.objet("dataLabels", {"fontSize": 11.0, "color": couleur(NAVY), "bold": True})
    reg.objet("categoryLabels", {"show": True, "fontSize": 8.5, "color": couleur(GREY)})
    reg.objet("card", {"barShow": True, "barColor": couleur(ACCENT), "outlineStyle": 0.0})
    reg.cle = "reglages_excel"
    _habiller(reg, titre="Réglages lus dans le classeur Excel (cellules du Dashboard)", fond="FFFFFF", entete=False)
    p.ajouter(reg)
    y += 120 + 12
    p.ajouter(texte([
        paragraphe(dict(texte="Comment lire cette page", gras=True, taille=10, couleur=NAVY)),
        paragraphe(dict(texte="Power BI recalcule les chiffres du Dashboard Excel avec les réglages enregistrés dans le "
                              "classeur (variante « Comme dans Excel », 12 mois du Dashboard Excel), sans tenir compte "
                              "des segments, et les compare aux valeurs de l'onglet Calculs lues dans le même fichier. "
                              "✔ = identique à l'arrondi près ; ✘ = écart.", taille=9, couleur=INK)),
        paragraphe(dict(texte="Écarts possibles et normaux : les affaires en retard (et le 1er mois automatique quand aucune "
                              "ligne n'a de date de début) dépendent de la date du jour. Excel la fige au dernier "
                              "enregistrement du classeur, Power BI à l'actualisation : si le classeur a été enregistré "
                              "un autre jour, ces lignes peuvent différer.", taille=9, couleur=INK)),
        paragraphe(dict(texte="Un classeur modifié mais pas encore enregistré n'est pas pris en compte : enregistrer dans "
                              "Excel (les valeurs de l'onglet Calculs sont alors à jour), puis actualiser le rapport.",
                        taille=9, couleur=INK)),
    ], M, y, UTILE, 112, "explications", fond="FFFFFF"))
    y += 112 + 12
    statut = m("Statut du contrôle")
    vert_rouge = regle((condition(statut, "=", "✘"), RED), (condition(statut, "=", "✔"), GREEN))
    wm = 560
    mois = tableau([ColonneTableau(col(C, "Contrôle"), 170, "Mois du Dashboard"),
                    ColonneTableau(col(C, "Cellule Excel"), 92, "Cellule Excel"),
                    ColonneTableau(m("Valeur Excel"), 80, "Excel (h)"),
                    ColonneTableau(m("Valeur Power BI"), 88, "Power BI (h)"),
                    ColonneTableau(m("Écart"), 60, "Écart"),
                    ColonneTableau(statut, 50, "Statut", alignement="Center")],
                   M, y, wm, 340, titre="Charge des 12 mois du Dashboard Excel (h)", tri=[(col(C, "Contrôle"), "Ascending")],
                   cle="controle_mois",
                   filtres=[filtre_categoriel(col(C, "Catégorie"), ["Charge par mois"], masque=True,
                                              nom="filtre_ctrl_mois")])
    couleur_colonne(mois, statut, police=vert_rouge)
    props_sans_selecteur(mois, "values", {"bold": False})
    p.ajouter(mois)
    ind = tableau([ColonneTableau(col(C, "Contrôle"), 250, "Indicateur (libellé de l'onglet Calculs)"),
                   ColonneTableau(col(C, "Cellule Excel"), 80, "Cellule Excel"),
                   ColonneTableau(m("Excel (affiché)"), 118, "Excel"),
                   ColonneTableau(m("Power BI (affiché)"), 118, "Power BI"),
                   ColonneTableau(m("Écart"), 56, "Écart"),
                   ColonneTableau(statut, 46, "Statut", alignement="Center")],
                  M + wm + G, y, UTILE - wm - G, 760, titre="Indicateurs principaux (onglet Calculs, colonne AL)",
                  tri=[(col(C, "Contrôle"), "Ascending")], cle="controle_indicateurs",
                  filtres=[filtre_categoriel(col(C, "Catégorie"), ["Indicateurs"], masque=True,
                                             nom="filtre_ctrl_indicateurs")])
    couleur_colonne(ind, statut, police=vert_rouge)
    props_sans_selecteur(ind, "values", {"wordWrap": True})
    p.ajouter(ind)
    p.ajouter(texte([
        paragraphe(dict(texte="Valeurs lues", gras=True, taille=9, couleur=NAVY)),
        paragraphe(dict(texte="Excel : onglet Calculs, colonnes AO-AP (mois) et AK-AL (indicateurs), plages nommées "
                              "bi_Mois et bi_Indicateurs. Power BI : mesures du modèle avec les réglages Excel, sur les "
                              "12 mois du Dashboard Excel.", taille=8.5, couleur=GREY)),
        paragraphe(dict(texte="Tolérance : 0,01 h pour les heures, 0,001 pour les nombres, 0,01 point pour les "
                              "pourcentages.", taille=8.5, couleur=GREY)),
    ], M, y + 340 + 12, wm, 96, "valeurs_lues", fond="FFFFFF"))
    return p


# ---------------------------------------------------------------------------------------------- info-bulles
# ---------------------------------------------------------------------------------------------- pages PDF
def page_pdf(r, nom_affiche, nom, titre, page_n):
    """Page A4 paysage à exporter en PDF : bandeau compact (réglages, filtres, période, date de lecture)."""
    p = r.page(nom_affiche, nom=nom, largeur=LARGEUR, hauteur=HAUT_PDF, affichage="FitToPage", fond="FFFFFF",
               papier_peint="FFFFFF", filtres=[filtre_periode()])
    p.ajouter(forme_rectangle(0, 0, LARGEUR, BANDEAU_PDF, NAVY, fond=False, cle="bandeau"))
    p.ajouter(texte([paragraphe(dict(texte=titre, gras=True, couleur="FFFFFF", taille=17))], M, 4, 820, 32,
                    "bandeau_titre"))
    p.ajouter(texte_mesure(m("Rappel des réglages"), M, 36, 820, 18, "bandeau_reglages", couleur_txt=BANDEAU_TXT,
                           taille=9, gras=True))
    p.ajouter(texte_mesure(m("Filtres appliqués"), M, 54, 820, 18, "bandeau_filtres", couleur_txt=BANDEAU_TXT,
                           taille=9))
    p.ajouter(valeur(m("Période affichée"), 852, 6, 412, 26, "bandeau_periode", taille=12, couleur_txt="FFFFFF",
                     gras=True))
    p.ajouter(valeur(m("Données lues le"), 852, 34, 412, 18, "bandeau_lecture", taille=9, couleur_txt=BANDEAU_TXT))
    p.ajouter(ligne_texte(page_n, 852, 54, 412, 18, "bandeau_page", taille=9, couleur_txt=BANDEAU_TXT,
                          alignement_txt="right"))
    return p


def segment_ressource_masque(p):
    """Copie invisible du segment Ressource : la page suit la ressource choisie sur la page Ressources."""
    v = p.ajouter(segment(col("Ressources", "Ressource"), M, 0, 300, 56, titre="Ressource", mode="Dropdown",
                          selection_unique=True, cle="seg_ressource", groupe_sync=SYNC_RES))
    v.masque = True
    return v


def pages_pdf(r):
    # tableau de bord : synthèse puis liste des affaires
    p = page_pdf(r, "PDF – Tableau de bord (1/2)", P_PDF_TDB1, "Plan de charge du bureau d'études", "Page 1 / 2")
    segments_filtres(p, 0, masque=True)
    y = bloc_cartes(p, BANDEAU_PDF + 8, H=250, alerte=False) + 10
    y = bloc_indicateurs(p, y, TH=100) + 10
    bloc_graphiques(p, y, GH=322)

    p = page_pdf(r, "PDF – Tableau de bord (2/2)", P_PDF_TDB2, "Plan de charge du bureau d'études : liste des affaires",
                 "Page 2 / 2")
    segments_filtres(p, 0, masque=True)
    y = bloc_liste_affaires(p, BANDEAU_PDF + 8, h=752, bouton_fiche=False, n_max=N_AFFAIRES_PDF, taille_texte=8.5,
                            titre_section="LISTE DES AFFAIRES (les %d premières, par charge décroissante)"
                                          % N_AFFAIRES_PDF)
    p.ajouter(ligne_texte("Nombre d'affaires du périmètre :", M, y + 6, 230, 20, "total_lib", taille=9, gras=True,
                          couleur_txt=INK))
    p.ajouter(texte_mesure(m("Nombre d'affaires"), M + 232, y + 6, 80, 20, "total_val", taille=9, gras=True))
    p.ajouter(ligne_texte("au-delà de la %de, voir le rapport en ligne." % N_AFFAIRES_PDF, M + 320, y + 6, 600, 20,
                          "total_note", taille=9, italique=True))

    # ressources : ressource choisie, puis ses lignes et l'occupation de l'équipe
    p = page_pdf(r, "PDF – Ressources (1/2)", P_PDF_RES1, "Charge par ressource : ressource choisie", "Page 1 / 2")
    y, _seg = bloc_choix_ressource(p, BANDEAU_PDF + 8, masque_segment=True)
    bloc_detail_ressource(p, y + 10, TH=104, CH=300, TABH=196)

    p = page_pdf(r, "PDF – Ressources (2/2)", P_PDF_RES2, "Charge par ressource : affaires et occupation de l'équipe",
                 "Page 2 / 2")
    seg = segment_ressource_masque(p)
    y = bloc_lignes_ressource(p, BANDEAU_PDF + 8, h=250, bouton_fiche=False) + 10
    _y, heat, na = bloc_occupation(p, y, HH=290)
    for cible in (heat, na):            # l'occupation montre toute l'équipe, comme à l'écran
        p.interaction(seg, cible, "NoFilter")


def page_infobulle_projet(r):
    p = r.page("Info-bulle projet", nom=P_IB_PROJET, largeur=320, hauteur=200, affichage="ActualSize",
               masquee=True, type_page="Tooltip", fond="FFFFFF")
    p.ajouter(texte_mesure(m("Titre info-bulle projet"), 8, 6, 304, 22, "titre", taille=10, couleur_txt=NAVY,
                           gras=True))
    v = tableau([ColonneTableau(col("Lignes", "Groupe projet"), 104, "Projet"),
                 ColonneTableau(m("Charge répartie (h)"), 64, "Heures réparties (h)"),
                 ColonneTableau(m("Affaires actives"), 56, "Affaires actives"),
                 ColonneTableau(m("Charge identifiée (h)"), 64, "Charge totale (h)")],
                8, 30, 304, 162, tri=[(m("Charge répartie (h)"), "Descending")], cle="detail", taille_texte=8.5,
                entete=False)
    p.ajouter(v)
    return p


def page_infobulle_occupation(r):
    p = r.page("Info-bulle occupation", nom=P_IB_OCC, largeur=320, hauteur=220, affichage="ActualSize",
               masquee=True, type_page="Tooltip", fond="FFFFFF")
    p.ajouter(texte_mesure(col("Ressources", "Ressource"), 8, 6, 304, 24, "ressource", taille=11, couleur_txt=NAVY,
                           gras=True))
    occ = m("Occupation (%)")
    p.ajouter(fiche_valeurs([("Charge fixe hors projet (h)", m("Charge fixe (h)")),
                       ("Charge projets (h)", m("Charge répartie (h)")),
                       ("Charge totale (h)", m("Charge totale de la ressource (h)")),
                       ("Capacité (h/mois)", m("Capacité de la ressource (h/mois)")),
                       ("Marge (h)", m("Marge de la ressource (h)")),
                       ("Occupation", occ)], 8, 34, 304, 178, "detail",
                      couleurs={5: couleur_champ(m("Couleur du texte (occupation)"))}))
    return p


# ---------------------------------------------------------------------------------------------- rapport
class RapportPlanDeCharge(Rapport):
    def _fichier_theme(self):
        return "Plan_de_charge_BE.json"     # convention de Desktop : espaces remplacées par « _ »


def theme():
    t = theme_personnalise("Plan de charge BE", SLOT_COLORS, fond="#FFFFFF", texte="#" + INK, accent="#" + ACCENT,
                           bon="#" + GREEN, neutre="#" + AMBER, mauvais="#" + RED, police="Segoe UI",
                           couleur_titres="#" + NAVY, fond_page="#" + BG)
    commun = t["visualStyles"]["*"]["*"]
    commun["background"] = [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}]
    commun["title"] = [{"fontColor": {"solid": {"color": "#" + NAVY}}, "fontSize": 11, "bold": True,
                        "fontFamily": "Segoe UI"}]
    return t


def rapport():
    r = RapportPlanDeCharge(NOM, theme=theme(), volet_filtres_deplie=False)
    r.filtres.extend(filtres_reglages())
    page_tableau_de_bord(r)
    page_ressources(r)
    page_fiche(r)
    page_controle(r)
    pages_pdf(r)
    page_infobulle_projet(r)
    page_infobulle_occupation(r)
    r.page_active = P_TDB
    texte_alternatif(r)
    return r


def texte_alternatif(r):
    """Texte de remplacement (lecteurs d'écran, volet Sélection) des visuels de données sans titre affiché :
    nom du premier champ affiché."""
    for p in r.pages:
        for v in p.visuels:
            if v.type in ("shape", "textbox", "actionButton") or "general" in v.conteneur:
                continue
            titre = any(e["properties"].get("show", {}).get("expr", {}).get("Literal", {}).get("Value") == "true"
                        for e in v.conteneur.get("title", []))
            if titre or not v.roles:
                continue
            proj = next(iter(v.roles.values()))[0]
            v.objet_conteneur("general", {"altText": proj.get("displayName") or proj["nativeQueryRef"]})


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    dossier = sys.argv[1]
    ecrire_projet(dossier, NOM, rapport(), model_bim=build_model())
    print("projet écrit :", dossier)
    return 0


if __name__ == "__main__":
    sys.exit(main())

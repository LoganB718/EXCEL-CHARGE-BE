# -*- coding: utf-8 -*-
"""Modèle sémantique Power BI « Plan de charge BE » : requêtes Power Query (M), tables, relations et mesures DAX.

build_model() renvoie le contenu de model.bim (TMSL) prêt à json.dump. Toutes les données viennent du classeur
Excel stocké dans Teams (paramètre UrlFichier), lu par les noms définis bi_* et p_* ; la logique de l'onglet Calculs
est reproduite dans les requêtes M, les mesures DAX restent des agrégations simples.
"""
import re
import uuid

COMPATIBILITY_LEVEL = 1567
CULTURE = "fr-FR"
EXCEL = "Comme dans Excel"                      # variante par défaut des segments Périmètre et Répartition
PERIODE_EXCEL = "12 mois du Dashboard Excel"    # mois de la fenêtre du Dashboard Excel
RETARD_COULEUR, ENCRE = "#B42318", "#1F2937"     # avancement en rouge si retard (RED_T), sinon encre (INK)
# occupation (OCC_RULES de build_dashboard.py) : (fond, texte) rouge > 100 %, orange >= 90 %, vert sinon
OCC_ROUGE, OCC_ORANGE, OCC_VERT = ("#F8D7D7", "#B42318"), ("#FFF1CC", "#8A5A00"), ("#E3F4E8", "#1E7B34")
SLOT_COLORS = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948", "9CA3AF"]
N_TABLEAU_EXCEL = 16                            # ressources affichées dans le tableau d'occupation Excel

_NS = uuid.UUID("5b0c7a52-3f7e-4a52-9a0e-6d3f1c2b8e41")


def tag(*parts):
    """lineageTag stable (identique d'une génération à l'autre)."""
    return str(uuid.uuid5(_NS, "/".join(parts)))


# --------------------------------------------------------------------------------------------- types de colonnes
M_TYPE = {"text": "type text", "int": "Int64.Type", "num": "type number", "date": "type date",
          "bool": "type logical", "datetime": "type datetime"}
BIM_TYPE = {"text": "string", "int": "int64", "num": "double", "date": "dateTime", "bool": "boolean",
            "datetime": "dateTime"}
F_H, F_H1, F_PCT, F_PCT1, F_DATE, F_MOIS, F_INT, F_DEC = ("#,0", "#,0.0", "0 %", "0.0 %", "dd/mm/yyyy",
                                                          "mmmm yyyy", "0", "0.0")


class Col:
    def __init__(self, name, kind, fmt=None, hidden=False, summarize="none", desc="", sort_by=None, folder=None):
        self.name, self.kind, self.fmt, self.hidden = name, kind, fmt, hidden
        self.summarize, self.desc, self.sort_by, self.folder = summarize, desc, sort_by, folder
        if fmt is None and kind == "date":
            self.fmt = F_DATE


def C(name, kind, fmt=None, **kw):
    return Col(name, kind, fmt, **kw)


def H(name, kind, desc="", **kw):
    """Colonne technique masquée."""
    return Col(name, kind, None if kind != "date" else F_DATE, hidden=True, desc=desc, **kw)


# --------------------------------------------------------------------------------------------- utilitaires M
def m_lines(text):
    return text.strip("\n").split("\n")


def final_steps(cols, source="Résultat"):
    """Étapes finales d'une requête chargée : colonnes du modèle dans l'ordre, puis typage explicite."""
    names = ", ".join('"%s"' % c.name for c in cols)
    types = ", ".join('{"%s", %s}' % (c.name, M_TYPE[c.kind]) for c in cols)
    return (f",\n    Sélection = Table.SelectColumns({source}, {{{names}}}, MissingField.UseNull),\n"
            f"    Typée = Table.TransformColumnTypes(Sélection, {{{types}}})\nin\n    Typée")


# =============================================================================================== expressions M
# Paramètre, fonctions et requêtes intermédiaires (non chargées). Les noms des requêtes sont ceux de l'éditeur
# Power Query ; les fonctions reproduisent les fonctions Excel utilisées par l'onglet Calculs.
EXPRESSIONS = []   # (nom, groupe, type de résultat, description, code M)


def expr(name, group, result, desc, code):
    EXPRESSIONS.append((name, group, result, desc, code.strip("\n")))


expr("UrlFichier", "Paramètres", "Text",
     "Chemin SharePoint / Teams du classeur Plan_de_charge_BE_Dashboard.xlsm (Excel bureau > Fichier > Informations > "
     "Copier le chemin d'accès, sans « ?web=1 »).",
     '"https://VOTRE-ENTREPRISE.sharepoint.com/sites/VOTRE-EQUIPE/Shared%20Documents/General/'
     'Plan_de_charge_BE_Dashboard.xlsm" meta [IsParameterQuery = true, Type = "Text", '
     'IsParameterQueryRequired = true]')

expr("Classeur", "Requêtes intermédiaires", "Table", "Contenu du classeur (valeurs enregistrées par Excel).", r'''
let
    // une seule source : le fichier stocké dans Teams / SharePoint (actualisation planifiée sans passerelle)
    Source = Excel.Workbook(Web.Contents(UrlFichier), null, true)
in
    Source''')

expr("fnLireNom", "Fonctions", "Function", "Plage d'un nom défini du classeur (colonnes Column1, Column2...).", r'''
(nom as text) as table =>
let
    Trouvés = Table.SelectRows(Classeur, each [Kind] = "DefinedName" and [Item] = nom),
    Données =
        if Table.IsEmpty(Trouvés) then
            error Error.Record(
                "Nom défini absent",
                "Le classeur ne contient pas le nom défini " & nom
                    & " : l'enregistrer avec la version à jour du générateur (scripts/build_dashboard.py).",
                nom
            )
        else
            Trouvés{0}[Data]
in
    Données''')

expr("fnValeurNom", "Fonctions", "Function", "Valeur d'un nom défini d'une seule cellule (null si vide).", r'''
(nom as text) as any =>
let
    Plage = fnLireNom(nom),
    Valeur =
        if Table.IsEmpty(Plage) or Table.ColumnCount(Plage) = 0 then null
        else try Table.FirstValue(Plage) otherwise null
in
    Valeur''')

expr("fnSérieExcel", "Fonctions", "Function", "Numéro de série Excel d'une date (null pour une autre valeur).", r'''
(valeur as any) as nullable number =>
    if valeur is date then Duration.Days(valeur - #date(1899, 12, 30))
    else if valeur is datetime then Duration.TotalDays(valeur - #datetime(1899, 12, 30, 0, 0, 0))
    else null''')

expr("fnNombreExcel", "Fonctions", "Function",
     "Valeur numérique d'une cellule au sens de ESTNUM (nombre ou date), sinon null.", r'''
(valeur as any) as nullable number =>
    if valeur is number then valeur
    else if valeur is date or valeur is datetime then fnSérieExcel(valeur)
    else null''')

expr("fnNExcel", "Fonctions", "Function", "Fonction N() d'Excel : nombre, date, VRAI = 1, sinon 0.", r'''
(valeur as any) as number =>
let
    Nombre = fnNombreExcel(valeur)
in
    if valeur is logical then (if valeur then 1 else 0)
    else if Nombre = null then 0
    else Nombre''')

expr("fnDateExcel", "Fonctions", "Function",
     "Date d'une cellule si ESTNUM (date ou numéro de série), sinon null (texte, vide).", r'''
(valeur as any) as nullable date =>
    if valeur is date then valeur
    else if valeur is datetime then Date.From(valeur)
    else if valeur is number then (try Date.From(valeur) otherwise null)
    else null''')

expr("fnTexteExcel", "Fonctions", "Function",
     "Texte d'une cellule tel qu'Excel le convertit (nombres au format standard français).", r'''
(valeur as any) as text =>
    if valeur = null then ""
    else if valeur is text then valeur
    else if valeur is number then Number.ToText(valeur, "G", "fr-FR")
    else if valeur is logical then (if valeur then "VRAI" else "FAUX")
    else if valeur is date or valeur is datetime then Number.ToText(fnSérieExcel(valeur), "G", "fr-FR")
    else Text.From(valeur)''')

expr("fnSupprEspaces", "Fonctions", "Function",
     "SUPPRESPACE d'Excel : espaces retirés en tête et en fin, espaces multiples réduits à un seul.", r'''
(texte as text) as text =>
    Text.Combine(List.Select(Text.Split(texte, " "), each _ <> ""), " ")''')

expr("fnCléAffaire", "Fonctions", "Function",
     "Clé d'affaire « projet | sous-projet | produit » comme la colonne EN de Calculs (casse ignorée).", r'''
(projet as any, sousProjet as any, produit as any) as text =>
let
    // espaces insécables remplacés par des espaces, puis SUPPRESPACE
    Partie = (valeur as any) as text => fnSupprEspaces(Text.Replace(fnTexteExcel(valeur), "#(00A0)", " ")),
    Brute = Partie(projet) & " | " & Partie(sousProjet) & " | " & Partie(produit),
    // caractères génériques de NB.SI neutralisés comme dans Excel
    SansJokers = Text.Replace(Text.Replace(Text.Replace(Brute, "~", "-"), "*", "+"), "?", "!"),
    Tronquée = Text.Start(SansJokers, 240)
in
    Text.Lower(Tronquée, "fr-FR")''')

expr("fnArrondiExcel", "Fonctions", "Function",
     "ARRONDI d'Excel : demi-unité arrondie en s'éloignant de zéro.", r'''
(nombre as number, décimales as number) as number =>
let
    Puissance = Number.Power(10, décimales),
    // arrondi intermédiaire à 6 décimales : neutralise les erreurs de représentation binaire (2,675 -> 2,68 comme Excel)
    Décalé = Number.Round(nombre * Puissance, 6, RoundingMode.AwayFromZero)
in
    Number.Round(Décalé, 0, RoundingMode.AwayFromZero) / Puissance''')

expr("fnJoursOuvrés", "Fonctions", "Function",
     "NB.JOURS.OUVRES d'Excel : jours du lundi au vendredi entre deux dates incluses, hors jours non travaillés.",
     r'''
(début as date, fin as date, joursNonTravaillés as list) as number =>
let
    // joursNonTravaillés : dates distinctes tombant un jour de semaine
    Sens = if fin < début then -1 else 1,
    Premier = if Sens = 1 then début else fin,
    Dernier = if Sens = 1 then fin else début,
    NbJours = Duration.Days(Dernier - Premier) + 1,
    Semaines = Number.IntegerDivide(NbJours, 7),
    Reste = Number.Mod(NbJours, 7),
    JourPremier = Date.DayOfWeek(Premier, Day.Monday),
    OuvrésReste = List.Count(List.Select(List.Numbers(JourPremier, Reste), each Number.Mod(_, 7) < 5)),
    Retirés = List.Count(List.Select(joursNonTravaillés, each _ >= Premier and _ <= Dernier))
in
    Sens * (Semaines * 5 + OuvrésReste - Retirés)''')

expr("NomsDesMois", "Requêtes intermédiaires", "Record", "Noms français des mois (abrégés et complets).", r'''
[
    Courts = {"janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."},
    Longs = {"janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
        "novembre", "décembre"}
]''')

expr("HeureParis", "Requêtes intermédiaires", "DateTime",
     "Date et heure de l'actualisation à Paris (le service Power BI tourne en UTC).", r'''
let
    Utc = DateTimeZone.RemoveZone(DateTimeZone.UtcNow()),
    Année = Date.Year(Utc),
    DernierDimanche = (mois as number) as date =>
        let
            Fin = Date.EndOfMonth(#date(Année, mois, 1))
        in
            Date.AddDays(Fin, - Date.DayOfWeek(Fin, Day.Sunday)),
    // heure d'été : du dernier dimanche de mars 02:00 (01:00 UTC) au dernier dimanche d'octobre 03:00 (01:00 UTC)
    DébutHeureÉté = DernierDimanche(3) & #time(1, 0, 0),
    FinHeureÉté = DernierDimanche(10) & #time(1, 0, 0),
    Décalage = if Utc >= DébutHeureÉté and Utc < FinHeureÉté then 2 else 1,
    Paris = Utc + #duration(0, Décalage, 0, 0)
in
    Paris''')

expr("DateDuJour", "Requêtes intermédiaires", "Date",
     "Aujourd'hui = date de l'actualisation à Paris (AUJOURDHUI() du classeur).", r'''
let
    Source = Date.From(HeureParis)
in
    Source''')

expr("ParamètresExcel", "Requêtes intermédiaires", "Record",
     "Réglages du Dashboard Excel (noms p_Perimetre, p_Repartition, p_PremierMois, p_HeuresETP, p_ChargeFixe).", r'''
let
    ValeurPérimètre = fnValeurNom("p_Perimetre"),
    ValeurRépartition = fnValeurNom("p_Repartition"),
    PremierMois = fnValeurNom("p_PremierMois"),
    HeuresETP = fnValeurNom("p_HeuresETP"),
    ChargeFixe = fnValeurNom("p_ChargeFixe"),
    PremierMoisDate = fnDateExcel(PremierMois),
    Résultat = [
        #"Périmètre Excel" = fnTexteExcel(ValeurPérimètre),
        #"Répartition Excel" = fnTexteExcel(ValeurRépartition),
        // comparaisons de textes insensibles à la casse, comme le signe = d'Excel
        #"Tous les projets" = (Text.Lower(fnTexteExcel(ValeurPérimètre), "fr-FR") = "tous les projets"),
        #"Lissée" = (Text.Lower(fnTexteExcel(ValeurRépartition), "fr-FR") = "lissée jusqu'à la date de besoin"),
        #"Premier mois saisi" = if PremierMoisDate = null then null else Date.StartOfMonth(PremierMoisDate),
        #"Capacité par défaut" = fnNExcel(HeuresETP),
        #"Charge fixe par défaut" = fnNExcel(ChargeFixe),
        #"Capacité (texte)" = fnTexteExcel(HeuresETP),
        #"Charge fixe (texte)" = fnTexteExcel(ChargeFixe)
    ]
in
    Résultat''')

expr("DonnéesData", "Requêtes intermédiaires", "Table",
     "Onglet Data, lignes 4 à 203 (nom bi_Data), colonnes A à K par position.", r'''
let
    Source = fnLireNom("bi_Data"),
    Colonnes = List.Transform({1 .. 11}, each "Column" & Text.From(_)),
    Complète = Table.SelectColumns(Source, Colonnes, MissingField.UseNull),
    SansErreurs = Table.ReplaceErrorValues(Complète, List.Transform(Colonnes, each {_, null})),
    Numérotée = Table.AddIndexColumn(SansErreurs, "N° ligne", 4, 1, Int64.Type),
    Résultat = Table.Buffer(Numérotée)
in
    Résultat''')

expr("JoursNonTravaillés", "Requêtes intermédiaires", "List",
     "Jours non travaillés de Data (colonne K) tombant un jour de semaine, sans doublon.", r'''
let
    Dates = List.Transform(DonnéesData[Column11], fnDateExcel),
    JoursDeSemaine = List.Select(Dates, each _ <> null and Date.DayOfWeek(_, Day.Monday) < 5),
    Résultat = List.Buffer(List.Sort(List.Distinct(JoursDeSemaine)))
in
    Résultat''')

expr("DonnéesCharge", "Requêtes intermédiaires", "Table",
     "Lignes valides de l'onglet Charge (nom bi_Charge, colonnes A à O par position) et leurs valeurs nettoyées.",
     r'''
let
    Source = fnLireNom("bi_Charge"),
    Colonnes = List.Transform({1 .. 15}, each "Column" & Text.From(_)),
    Lettres = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"},
    Complète = Table.SelectColumns(Source, Colonnes, MissingField.UseNull),
    SansErreurs = Table.ReplaceErrorValues(Complète, List.Transform(Colonnes, each {_, null})),
    Renommée = Table.RenameColumns(SansErreurs, List.Zip({Colonnes, Lettres})),
    Numérotée = Table.AddIndexColumn(Renommée, "N° ligne", 4, 1, Int64.Type),
    // règle 1 : ligne valide = nom du projet (B) non vide après SUPPRESPACE
    Valides = Table.SelectRows(Numérotée, each fnSupprEspaces(fnTexteExcel([B])) <> ""),
    AjoutChef = Table.AddColumn(Valides, "Chef de projet (ligne)",
        each let t = fnSupprEspaces(fnTexteExcel([A])) in if t = "" then "(non renseigné)" else t, type text),
    AjoutProjet = Table.AddColumn(AjoutChef, "Projet (ligne)", each fnSupprEspaces(fnTexteExcel([B])), type text),
    AjoutSousProjet = Table.AddColumn(AjoutProjet, "Sous-projet", each fnSupprEspaces(fnTexteExcel([C])), type text),
    AjoutProduit = Table.AddColumn(AjoutSousProjet, "Produit", each fnSupprEspaces(fnTexteExcel([D])), type text),
    AjoutType = Table.AddColumn(AjoutProduit, "Type (ligne)",
        each let t = fnSupprEspaces(fnTexteExcel([E])) in if t = "" then "(non renseigné)" else t, type text),
    AjoutBesoin = Table.AddColumn(AjoutType, "Date de besoin", each fnDateExcel([F]), type nullable date),
    AjoutDescription = Table.AddColumn(AjoutBesoin, "Description", each fnSupprEspaces(fnTexteExcel([G])), type text),
    AjoutContraintes = Table.AddColumn(AjoutDescription, "Contraintes", each fnSupprEspaces(fnTexteExcel([H])), type text),
    AjoutInfos = Table.AddColumn(AjoutContraintes, "Autres informations", each fnSupprEspaces(fnTexteExcel([I])), type text),
    AjoutCoordinateur = Table.AddColumn(AjoutInfos, "Coordinateur", each fnSupprEspaces(fnTexteExcel([J])), type text),
    AjoutDessinateur = Table.AddColumn(AjoutCoordinateur, "Dessinateur", each fnSupprEspaces(fnTexteExcel([K])), type text),
    AjoutExterne = Table.AddColumn(AjoutDessinateur, "Ressource externe", each fnSupprEspaces(fnTexteExcel([L])), type text),
    // règle 3 : heures = max(0, M numérique sinon 0) ; règle 1 : avancement borné entre 0 et 1
    AjoutHeures = Table.AddColumn(AjoutExterne, "Heures (h)", each List.Max({0, fnNExcel([M])}), type number),
    AjoutDébut = Table.AddColumn(AjoutHeures, "Début planifié", each fnDateExcel([N]), type nullable date),
    AjoutAvancement = Table.AddColumn(AjoutDébut, "Avancement", each List.Min({1, List.Max({0, fnNExcel([O])})}), type number),
    // règle 4 : ressource = dessinateur (K), sinon ressource externe (L), sinon « (non affecté) »
    AjoutRessource = Table.AddColumn(AjoutAvancement, "Ressource de la ligne",
        each if [Dessinateur] <> "" then [Dessinateur]
            else if [Ressource externe] <> "" then [Ressource externe]
            else "(non affecté)", type text),
    AjoutTypeRessource = Table.AddColumn(AjoutRessource, "Type de ressource",
        each if [Dessinateur] <> "" then "Interne" else if [Ressource externe] <> "" then "Externe" else "Non affecté",
        type text),
    AjoutDouble = Table.AddColumn(AjoutTypeRessource, "Dessinateur et externe",
        each [Dessinateur] <> "" and [Ressource externe] <> "", type logical),
    AjoutCléRessource = Table.AddColumn(AjoutDouble, "Clé ressource (texte)",
        each Text.Lower([#"Ressource de la ligne"], "fr-FR"), type text),
    AjoutCléAffaire = Table.AddColumn(AjoutCléRessource, "Clé affaire (texte)", each fnCléAffaire([B], [C], [D]), type text),
    AjoutRéalisé = Table.AddColumn(AjoutCléAffaire, "Réalisé de la ligne (h)", each [#"Heures (h)"] * [Avancement], type number),
    AjoutTerminée = Table.AddColumn(AjoutRéalisé, "Terminée", each [Avancement] >= 1, type logical),
    SansBrutes = Table.RemoveColumns(AjoutTerminée, Lettres),
    Résultat = Table.Buffer(SansBrutes)
in
    Résultat''')

expr("RépartitionLignes", "Requêtes intermédiaires", "Table",
     "Heures de chaque ligne réparties par mois en jours ouvrés, pour les deux modes de répartition (règle 6).", r'''
let
    // lignes sans date de début ou sans heures : non réparties (comptées « hors période ou sans date de début »)
    Source = Table.SelectRows(DonnéesCharge, each [#"Début planifié"] <> null and [#"Heures (h)"] <> 0),
    Colonnes = Table.SelectColumns(Source, {"N° ligne", "Heures (h)", "Début planifié", "Date de besoin"}),
    AvecModes = Table.AddColumn(Colonnes, "Code répartition", each {"Lissée", "Début"}),
    Modes = Table.ExpandListColumn(AvecModes, "Code répartition"),
    // fin = date de besoin si lissage et besoin >= début, sinon début
    AjoutFin = Table.AddColumn(Modes, "Fin",
        each if [#"Code répartition"] = "Lissée" and [#"Date de besoin"] <> null
                and [#"Date de besoin"] >= [#"Début planifié"]
            then [#"Date de besoin"] else [#"Début planifié"], type date),
    AjoutJours = Table.AddColumn(AjoutFin, "Jours ouvrés",
        each fnJoursOuvrés([#"Début planifié"], [Fin], JoursNonTravaillés), type number),
    AjoutMois = Table.AddColumn(AjoutJours, "Mois",
        each
            let
                Premier = Date.StartOfMonth([#"Début planifié"]),
                NbMois = (Date.Year([Fin]) - Date.Year(Premier)) * 12 + Date.Month([Fin]) - Date.Month(Premier) + 1
            in
                List.Transform({0 .. NbMois - 1}, (i) => Date.AddMonths(Premier, i))),
    ParMois = Table.ExpandListColumn(AjoutMois, "Mois"),
    // prorata des jours ouvrés du mois ; aucun jour ouvré sur la période : tout sur le mois de début
    AjoutHeures = Table.AddColumn(ParMois, "Heures du mois",
        each
            let
                DébutMois = [Mois],
                FinMois = Date.EndOfMonth([Mois])
            in
                if [#"Jours ouvrés"] = 0 then
                    (if [#"Début planifié"] >= DébutMois and [#"Début planifié"] <= FinMois then [#"Heures (h)"] else 0)
                else
                    [#"Heures (h)"]
                        * fnJoursOuvrés(List.Max({[#"Début planifié"], DébutMois}), List.Min({[Fin], FinMois}),
                            JoursNonTravaillés)
                        / [#"Jours ouvrés"],
        type number),
    Positives = Table.SelectRows(AjoutHeures, each [#"Heures du mois"] > 0),
    Résultat = Table.Buffer(Table.SelectColumns(Positives, {"N° ligne", "Code répartition", "Mois", "Heures du mois"}))
in
    Résultat''')

expr("LignesPérimètres", "Requêtes intermédiaires", "Table",
     "Lignes de chaque variante du segment Périmètre, avec chef, projet et type regroupés (casse ignorée), "
     "affaire, rangs et groupe de couleur du projet.", r'''
let
    Options = Table.Buffer(Périmètre),
    Couleurs = {"#2A78D6", "#EB6834", "#1BAF7A", "#EDA100", "#E87BA4", "#008300", "#6250D6", "#E34948", "#9CA3AF"},
    // regroupement insensible à la casse (NB.SI d'Excel) : libellé de la 1re ligne, total arrondi, rang décroissant
    Grouper = (lignes as table, colonne as text, nom as text) as table =>
        let
            AvecClé = Table.AddColumn(lignes, "Clé groupe", each Text.Lower(Record.Field(_, colonne), "fr-FR"), type text),
            Groupes = Table.Group(AvecClé, {"Clé groupe"}, {
                {nom, each Record.Field(Table.Min(_, "N° ligne"), colonne), type text},
                {"Première ligne", each List.Min([#"N° ligne"]), Int64.Type},
                {"Total", each fnArrondiExcel(List.Sum([#"Heures (h)"]), 2), type number}
            }),
            Triés = Table.Sort(Groupes, {{"Total", Order.Descending}, {"Première ligne", Order.Ascending}}),
            Rangés = Table.AddIndexColumn(Triés, "Rang", 1, 1, Int64.Type)
        in
            Table.SelectColumns(Rangés, {"Clé groupe", nom, "Rang"}),
    Associer = (lignes as table, colonne as text, nom as text, rang as text) as table =>
        let
            Groupes = Grouper(lignes, colonne, nom),
            AvecClé = Table.AddColumn(lignes, "Clé groupe", each Text.Lower(Record.Field(_, colonne), "fr-FR"), type text),
            Jointes = Table.NestedJoin(AvecClé, {"Clé groupe"}, Groupes, {"Clé groupe"}, "Groupe", JoinKind.LeftOuter),
            Développées = Table.ExpandTableColumn(Jointes, "Groupe", {nom, "Rang"}, {nom, rang})
        in
            Table.RemoveColumns(Développées, {"Clé groupe"}),
    ParOption = (option as record) as table =>
        let
            Ordre = option[Ordre],
            // règle 2 : « Projets en cours » = lignes d'avancement < 100 %
            LignesOption = if option[#"Code périmètre"] = "Tous" then DonnéesCharge
                else Table.SelectRows(DonnéesCharge, each [Avancement] < 1),
            AjoutOrdre = Table.AddColumn(LignesOption, "Ordre périmètre", each Ordre, Int64.Type),
            AjoutPérimètre = Table.AddColumn(AjoutOrdre, "Périmètre", each option[Périmètre], type text),
            AjoutCléLigne = Table.AddColumn(AjoutPérimètre, "Clé ligne", each Ordre * 10000 + [#"N° ligne"], Int64.Type),
            // règle 5 : affaire = 1re ligne du périmètre ayant la même clé
            PremièresLignes = Table.Group(AjoutCléLigne, {"Clé affaire (texte)"},
                {{"Première ligne de l'affaire", each List.Min([#"N° ligne"]), Int64.Type}}),
            JointesAffaires = Table.NestedJoin(AjoutCléLigne, {"Clé affaire (texte)"}, PremièresLignes,
                {"Clé affaire (texte)"}, "Affaire", JoinKind.LeftOuter),
            AvecAffaire = Table.ExpandTableColumn(JointesAffaires, "Affaire", {"Première ligne de l'affaire"}),
            AjoutCléAffaire = Table.AddColumn(AvecAffaire, "Clé affaire",
                each Ordre * 10000 + [#"Première ligne de l'affaire"], Int64.Type),
            AvecChef = Associer(AjoutCléAffaire, "Chef de projet (ligne)", "Chef de projet", "Rang chef"),
            AvecType = Associer(AvecChef, "Type (ligne)", "Type", "Rang type"),
            NbProjets = Table.RowCount(Grouper(AvecType, "Projet (ligne)", "Projet")),
            AvecProjet = Associer(AvecType, "Projet (ligne)", "Projet", "Rang projet"),
            // au-delà de 9 projets, les projets de rang >= 9 sont regroupés dans « Autres projets (n) »
            AjoutOrdreGroupe = Table.AddColumn(AvecProjet, "Ordre groupe projet",
                each if NbProjets > 9 then List.Min({[Rang projet], 9}) else [Rang projet], Int64.Type),
            AjoutGroupe = Table.AddColumn(AjoutOrdreGroupe, "Groupe projet",
                each if NbProjets > 9 and [Ordre groupe projet] = 9
                    then "Autres projets (" & Text.From(NbProjets - 8) & ")" else [Projet], type text),
            AjoutCouleur = Table.AddColumn(AjoutGroupe, "Couleur projet", each Couleurs{[Ordre groupe projet] - 1}, type text),
            // rang de la ligne : heures décroissantes, puis ordre de l'onglet Charge
            TriLignes = Table.Sort(AjoutCouleur, {{"Heures (h)", Order.Descending}, {"N° ligne", Order.Ascending}}),
            RangLignes = Table.AddIndexColumn(TriLignes, "Rang ligne", 1, 1, Int64.Type)
        in
            Table.Sort(RangLignes, {{"N° ligne", Order.Ascending}}),
    Tables = List.Transform(Table.ToRecords(Options), ParOption),
    Résultat = Table.Buffer(Table.Combine(Tables))
in
    Résultat''')

expr("FenêtreExcel", "Requêtes intermédiaires", "Record",
     "Les 12 mois du Dashboard Excel (règle 7) calculés avec les réglages du classeur.", r'''
let
    LignesExcel = Table.SelectRows(LignesPérimètres, each [#"Ordre périmètre"] = 1),
    Débuts = List.RemoveNulls(LignesExcel[#"Début planifié"]),
    DébutAuto = if List.IsEmpty(Débuts) then Date.StartOfMonth(DateDuJour) else Date.StartOfMonth(List.Min(Débuts)),
    Saisi = ParamètresExcel[#"Premier mois saisi"],
    Début = if Saisi <> null then Saisi else DébutAuto,
    Résultat = [
        #"1er mois automatique" = DébutAuto,
        #"1er mois affiché" = Début,
        #"Dernier mois affiché" = Date.AddMonths(Début, 11)
    ]
in
    Résultat''')

# =============================================================================================== tables chargées
TABLES = []   # dict(name, desc, cols, m, measures)


def table(name, desc, cols, m_body, hidden=False):
    names = [c.name for c in cols]
    assert len(names) == len(set(names)), name
    m = m_body.strip("\n") + final_steps(cols)
    TABLES.append(dict(name=name, desc=desc, cols=cols, m=m, hidden=hidden))


table("Périmètre", "Segment Périmètre (sélection unique) : « Comme dans Excel » suit le réglage p_Perimetre du classeur.", [
    C("Périmètre", "text", sort_by="Ordre", desc="Variante de périmètre proposée par le segment."),
    H("Ordre", "int"),
    C("Périmètre appliqué", "text", desc="Périmètre réellement appliqué (Projets en cours ou Tous les projets)."),
    H("Code périmètre", "text"),
    C("Description", "text"),
], r'''
let
    Tous = ParamètresExcel[#"Tous les projets"],
    Appliqué = if Tous then "Tous les projets" else "Projets en cours",
    Résultat = #table(
        type table [Périmètre = text, Ordre = Int64.Type, #"Périmètre appliqué" = text, #"Code périmètre" = text,
            Description = text],
        {
            {"Comme dans Excel", 1, Appliqué, if Tous then "Tous" else "En cours",
                "Réglage du classeur Excel : " & Appliqué},
            {"Projets en cours", 2, "Projets en cours", "En cours", "Lignes dont l'avancement est inférieur à 100 %"},
            {"Tous les projets", 3, "Tous les projets", "Tous", "Toutes les lignes, y compris les lignes terminées"}
        })''')

table("Répartition", "Segment Répartition (sélection unique) : « Comme dans Excel » suit le réglage p_Repartition.", [
    C("Répartition", "text", sort_by="Ordre", desc="Variante de répartition mensuelle proposée par le segment."),
    H("Ordre", "int"),
    C("Répartition appliquée", "text", desc="Mode réellement appliqué."),
    H("Code répartition", "text"),
    C("Description", "text"),
], r'''
let
    Lissée = ParamètresExcel[#"Lissée"],
    Appliquée = if Lissée then "Lissée jusqu'à la date de besoin" else "Tout sur le mois de début",
    Résultat = #table(
        type table [Répartition = text, Ordre = Int64.Type, #"Répartition appliquée" = text,
            #"Code répartition" = text, Description = text],
        {
            {"Comme dans Excel", 1, Appliquée, if Lissée then "Lissée" else "Début",
                "Réglage du classeur Excel : " & Appliquée},
            {"Lissée jusqu'au besoin", 2, "Lissée jusqu'à la date de besoin", "Lissée",
                "Prorata des jours ouvrés entre le début planifié et la date de besoin"},
            {"Tout sur le mois de début", 3, "Tout sur le mois de début", "Début",
                "Toute la charge sur le mois du début planifié"}
        })''')

table("Ressources", "Dessinateurs (Data E) puis ressources externes (Data H), ressources saisies dans Charge hors "
      "listes, et « (non affecté) ».", [
    H("Clé ressource", "int"),
    C("Ressource", "text", sort_by="Ordre"),
    H("Ordre", "int"),
    C("Type de ressource", "text", desc="Interne (Data E), Externe (Data H), Hors liste (absente des listes Data) "
      "ou Non affecté."),
    C("Mention", "text", desc="« ext. » pour une ressource externe, « hors liste » pour une ressource absente de Data."),
    C("Ressource (avec mention)", "text", sort_by="Ordre", desc="Nom suivi de la mention (« · ext. », « · hors liste ») : "
      "lignes du tableau d'occupation, comme la colonne « ext. » de l'onglet Dessinateurs."),
    C("Position Excel", "int", F_INT, desc="Rang dans la liste commune de Calculs (FA) ; les 16 premiers forment le "
      "tableau d'occupation de l'onglet Dessinateurs."),
    C("Dans le tableau Excel", "bool", desc="Ressource affichée dans le tableau d'occupation de l'onglet Dessinateurs."),
    C("Dans l'équipe", "bool", desc="Ressource des listes Data, comptée dans la capacité projets de l'équipe."),
    C("Capacité (h/mois)", "num", F_H, summarize="sum"),
    C("Charge fixe (h/mois)", "num", F_H, summarize="sum"),
    C("Capacité projets (h/mois)", "num", F_H, summarize="sum", desc="max(0, capacité - charge fixe)."),
    H("Clé ressource (texte)", "text"),
], r'''
let
    Data = DonnéesData,
    Défauts = ParamètresExcel,
    // NBVAL de chaque liste : Excel prend les n premières cellules de la colonne (liste saisie sans ligne vide)
    Compter = (colonne as text) as number => List.Count(List.Select(Table.Column(Data, colonne), each _ <> null)),
    NbInternes = Compter("Column5"),
    NbExternes = Compter("Column8"),
    Lire = (lignes as table, nom as text, capacité as text, fixe as text, typeRessource as text) as list =>
        List.Transform(Table.ToRecords(lignes), (r) =>
            let
                Capacité = fnNombreExcel(Record.Field(r, capacité)),
                Fixe = fnNombreExcel(Record.Field(r, fixe))
            in
                {
                    fnSupprEspaces(fnTexteExcel(Record.Field(r, nom))),
                    typeRessource,
                    if Capacité = null then Défauts[#"Capacité par défaut"] else Capacité,
                    if Fixe = null then Défauts[#"Charge fixe par défaut"] else Fixe
                }),
    Liste = #table(
        type table [Ressource = text, #"Type de ressource" = text, #"Capacité (h/mois)" = number,
            #"Charge fixe (h/mois)" = number],
        Lire(Table.FirstN(Data, NbInternes), "Column5", "Column6", "Column7", "Interne")
            & Lire(Table.FirstN(Data, NbExternes), "Column8", "Column9", "Column10", "Externe")),
    Positions = Table.AddIndexColumn(Liste, "Position Excel", 1, 1, Int64.Type),
    NonVides = Table.SelectRows(Positions, each [Ressource] <> ""),
    AvecClé = Table.AddColumn(NonVides, "Clé ressource (texte)", each Text.Lower([Ressource], "fr-FR"), type text),
    ClésListe = List.Buffer(AvecClé[#"Clé ressource (texte)"]),
    // même nom dans les deux listes ou en double (casse ignorée) : la 1re occurrence gagne
    Indexées = Table.AddIndexColumn(AvecClé, "Index", 0, 1, Int64.Type),
    Uniques = Table.SelectRows(Indexées, each List.PositionOf(ClésListe, [#"Clé ressource (texte)"]) = [Index]),
    Équipe = Table.AddColumn(Table.RemoveColumns(Uniques, {"Index"}), "Dans l'équipe", each true, type logical),
    ClésÉquipe = List.Buffer(Équipe[#"Clé ressource (texte)"]),
    // ressources saisies dans Charge mais absentes des listes Data : réglages par défaut
    LignesHors = Table.SelectRows(DonnéesCharge, each [#"Clé ressource (texte)"] <> "(non affecté)"
        and not List.Contains(ClésÉquipe, [#"Clé ressource (texte)"])),
    PremièresHors = Table.Group(LignesHors, {"Clé ressource (texte)"}, {
        {"Ressource", each Table.Min(_, "N° ligne")[#"Ressource de la ligne"], type text},
        {"Première ligne", each List.Min([#"N° ligne"]), Int64.Type}}),
    HorsTriées = Table.Sort(PremièresHors, {{"Première ligne", Order.Ascending}}),
    Hors = Table.FromColumns(
        {HorsTriées[Ressource], List.Repeat({"Hors liste"}, Table.RowCount(HorsTriées)),
            List.Repeat({Défauts[#"Capacité par défaut"]}, Table.RowCount(HorsTriées)),
            List.Repeat({Défauts[#"Charge fixe par défaut"]}, Table.RowCount(HorsTriées)),
            List.Repeat({null}, Table.RowCount(HorsTriées)), HorsTriées[#"Clé ressource (texte)"],
            List.Repeat({false}, Table.RowCount(HorsTriées))},
        Table.ColumnNames(Équipe)),
    NonAffectée = #table(Table.ColumnNames(Équipe),
        if List.Contains(ClésÉquipe, "(non affecté)") then {}
        else {{"(non affecté)", "Non affecté", null, null, null, "(non affecté)", false}}),
    Toutes = Table.Combine({Équipe, Hors, NonAffectée}),
    AjoutClé = Table.AddIndexColumn(Toutes, "Clé ressource", 1, 1, Int64.Type),
    AjoutOrdre = Table.AddColumn(AjoutClé, "Ordre", each [Clé ressource], Int64.Type),
    AjoutCapacitéProjets = Table.AddColumn(AjoutOrdre, "Capacité projets (h/mois)",
        each if [#"Capacité (h/mois)"] = null then null
            else List.Max({0, [#"Capacité (h/mois)"] - [#"Charge fixe (h/mois)"]}), type nullable number),
    AjoutTableau = Table.AddColumn(AjoutCapacitéProjets, "Dans le tableau Excel",
        each [Position Excel] <> null and [Position Excel] <= 16, type logical),
    AjoutMention = Table.AddColumn(AjoutTableau, "Mention",
        each if [#"Type de ressource"] = "Externe" then "ext."
            else if [#"Type de ressource"] = "Hors liste" then "hors liste" else "", type text),
    AjoutLibellé = Table.AddColumn(AjoutMention, "Ressource (avec mention)",
        each if [Mention] = "" then [Ressource] else [Ressource] & " · " & [Mention], type text),
    Résultat = AjoutLibellé''')

table("Lignes", "Une ligne par variante de périmètre et par ligne valide de l'onglet Charge.", [
    H("Clé ligne", "int", "Ordre de la variante × 10000 + n° de ligne."),
    H("Clé affaire", "int"),
    H("Clé ressource", "int"),
    H("Périmètre", "text"),
    C("N° ligne", "int", F_INT, desc="Numéro de ligne dans l'onglet Charge."),
    C("Chef de projet", "text", desc="Chef de la ligne (« (non renseigné) » si vide), libellé de sa 1re ligne du "
      "périmètre (casse ignorée)."),
    C("Projet", "text"),
    C("Sous-projet", "text"),
    C("Produit", "text"),
    C("Type", "text"),
    C("Description", "text", folder="Détail"),
    C("Contraintes", "text", folder="Détail"),
    C("Autres informations", "text", folder="Détail"),
    C("Coordinateur", "text", folder="Détail"),
    C("Dessinateur", "text", folder="Détail", desc="Colonne K de Charge."),
    C("Ressource externe", "text", folder="Détail", desc="Colonne L de Charge."),
    C("Ressource de la ligne", "text", desc="Dessinateur, sinon ressource externe, sinon « (non affecté) »."),
    C("Type de ressource", "text", desc="Interne (K), Externe (L) ou Non affecté."),
    C("Dessinateur et externe", "bool", desc="Ligne à corriger : dessinateur ET ressource externe (seul K compte)."),
    C("Heures (h)", "num", F_H, summarize="sum"),
    C("Avancement", "num", F_PCT),
    C("Réalisé de la ligne (h)", "num", F_H, summarize="sum", desc="Heures × avancement."),
    C("Reste de la ligne (h)", "num", F_H, summarize="sum", desc="Heures - réalisé."),
    C("Début planifié", "date"),
    C("Début (affiché)", "text", sort_by="Début planifié", desc="Date de début jj/mm/aaaa, ou « — » sans date de début "
      "(comme l'onglet Dessinateurs) ; triée par date."),
    C("Date de besoin", "date"),
    C("Terminée", "bool", desc="Avancement de 100 %."),
    C("Sans date de début", "bool"),
    C("Rang chef", "int", F_INT, folder="Rangs"),
    C("Rang projet", "int", F_INT, folder="Rangs"),
    C("Rang type", "int", F_INT, folder="Rangs"),
    C("Rang ligne", "int", F_INT, folder="Rangs", desc="Rang de la ligne par heures décroissantes (ordre de "
      "l'onglet Charge en cas d'égalité)."),
    C("Groupe projet", "text", desc="Projet, ou « Autres projets (n) » au-delà du 8e projet quand il y en a plus "
      "de 9 (couleurs du Dashboard Excel)."),
    C("Ordre groupe projet", "int", F_INT, folder="Rangs"),
    C("Couleur projet", "text", desc="Couleur du groupe projet (palette du Dashboard Excel)."),
    H("Clé affaire (texte)", "text"),
], r'''
let
    Source = LignesPérimètres,
    Clés = Table.SelectColumns(Ressources, {"Clé ressource (texte)", "Clé ressource"}),
    Jointes = Table.NestedJoin(Source, {"Clé ressource (texte)"}, Clés, {"Clé ressource (texte)"}, "Jointure ressource",
        JoinKind.LeftOuter),
    AvecRessource = Table.ExpandTableColumn(Jointes, "Jointure ressource", {"Clé ressource"}),
    AjoutReste = Table.AddColumn(AvecRessource, "Reste de la ligne (h)",
        each [#"Heures (h)"] - [#"Réalisé de la ligne (h)"], type number),
    AjoutSansDébut = Table.AddColumn(AjoutReste, "Sans date de début", each [#"Début planifié"] = null, type logical),
    AjoutDébutAffiché = Table.AddColumn(AjoutSansDébut, "Début (affiché)",
        each if [#"Début planifié"] = null then "—"
            else Text.PadStart(Text.From(Date.Day([#"Début planifié"])), 2, "0") & "/"
                & Text.PadStart(Text.From(Date.Month([#"Début planifié"])), 2, "0") & "/"
                & Text.From(Date.Year([#"Début planifié"])), type text),
    Résultat = AjoutDébutAffiché''')

table("Affaires", "Une ligne par variante de périmètre et par affaire (projet | sous-projet | produit) ayant au "
      "moins une ligne dans le périmètre.", [
    H("Clé affaire", "int"),
    H("Périmètre", "text"),
    C("N°", "int", F_INT, desc="Rang dans le tableau des affaires du Dashboard (charge décroissante, puis ordre de "
      "l'onglet Charge)."),
    C("Affaire", "text", desc="Projet · sous-projet · produit (libellés de la 1re ligne du périmètre, « — » si "
      "vide) : champ d'extraction vers la fiche affaire."),
    C("Projet", "text"),
    C("Sous-projet", "text"),
    C("Produit", "text"),
    C("Chef de projet", "text", desc="Chef de la 1re ligne de l'affaire."),
    C("Ressources", "text", desc="« N dont X ext. », « X ext. », « à affecter », « + à affecter »."),
    C("Charge (h)", "num", F_H, summarize="sum", desc="Heures des lignes du périmètre (arrondies à 0,01 h)."),
    C("% de la charge totale", "num", F_PCT1, desc="Charge de l'affaire / charge identifiée du périmètre."),
    C("Début planifié", "date", desc="Plus petit début des lignes du périmètre."),
    C("Début (affiché)", "text", sort_by="Début planifié", desc="Date de début, ou « à planifier » ; triée par date."),
    C("Date de besoin", "date", desc="Plus grande date de besoin des lignes du périmètre."),
    C("Avancement", "num", F_PCT, desc="Moyenne pondérée par les heures de TOUTES les lignes de l'affaire (moyenne "
      "simple sans heures)."),
    C("Heures toutes lignes (h)", "num", F_H, summarize="sum"),
    C("Réalisé toutes lignes (h)", "num", F_H, summarize="sum"),
    C("Reste à faire de l'affaire (h)", "num", F_H, summarize="sum", desc="max(0, heures - réalisé) de toutes les "
      "lignes de l'affaire (arrondi à 0,01 h)."),
    C("Avancement attendu", "num", F_PCT, desc="Prorata des jours ouvrés écoulés entre début et besoin (vide si non "
      "calculable)."),
    C("Retard", "int", F_INT, desc="2 = date de besoin dépassée sans être à 100 %, 1 = avancement inférieur de plus "
      "de 15 points à l'attendu, 0 = dans les temps."),
    C("Statut", "text"),
    C("Couleur d'alerte", "text", desc="Rouge B42318 si retard, sinon encre du Dashboard (1F2937)."),
    C("Lignes de l'affaire", "int", F_INT, summarize="sum", desc="Lignes de l'affaire dans le périmètre."),
    C("Personnes internes", "int", F_INT, summarize="sum"),
    C("Personnes externes", "int", F_INT, summarize="sum"),
    C("À affecter", "bool", desc="Au moins une ligne sans ressource."),
    C("N° 1re ligne", "int", F_INT, desc="Ligne de Charge de la 1re ligne de l'affaire dans le périmètre."),
    H("Clé affaire (texte)", "text"),
], r'''
let
    LignesP = LignesPérimètres,
    Aujourdhui = DateDuJour,
    // avancement, réalisé et reste à faire : TOUTES les lignes valides de l'affaire (périmètre ignoré)
    Toutes = Table.Group(DonnéesCharge, {"Clé affaire (texte)"}, {
        {"Heures toutes lignes (h)", each fnArrondiExcel(List.Sum([#"Heures (h)"]), 2), type number},
        {"Réalisé toutes lignes (h)", each fnArrondiExcel(List.Sum([#"Réalisé de la ligne (h)"]), 2), type number},
        {"Moyenne des avancements", each List.Average([Avancement]), type number}}),
    // personnes distinctes (type de leur 1re ligne dans l'affaire) et lignes sans ressource
    Personnes = (lignes as table) as record =>
        let
            AvecPersonne = Table.SelectRows(lignes, each [#"Clé ressource (texte)"] <> "(non affecté)"),
            ParPersonne = Table.Group(AvecPersonne, {"Clé ressource (texte)"},
                {{"Type", each Table.Min(_, "N° ligne")[#"Type de ressource"], type text}}),
            Internes = Table.RowCount(Table.SelectRows(ParPersonne, each [Type] = "Interne")),
            Externes = Table.RowCount(Table.SelectRows(ParPersonne, each [Type] = "Externe")),
            NonAffectées = Table.RowCount(Table.SelectRows(lignes, each [#"Type de ressource"] = "Non affecté")),
            Texte =
                if Internes + Externes = 0 then "à affecter"
                else (if Internes = 0 then Text.From(Externes) & " ext."
                    else Text.From(Internes + Externes) & (if Externes > 0 then " dont " & Text.From(Externes) & " ext." else ""))
                    & (if NonAffectées > 0 then " + à affecter" else "")
        in
            [Ressources = Texte, #"Personnes internes" = Internes, #"Personnes externes" = Externes,
                #"À affecter" = NonAffectées > 0],
    Groupes = Table.Group(LignesP, {"Ordre périmètre", "Périmètre", "Clé affaire", "Clé affaire (texte)"}, {
        {"Première", each Table.Min(_, "N° ligne"), type record},
        {"Charge (h)", each fnArrondiExcel(List.Sum([#"Heures (h)"]), 2), type number},
        {"Début planifié", each List.Min(List.RemoveNulls([#"Début planifié"])), type nullable date},
        {"Date de besoin", each List.Max(List.RemoveNulls([#"Date de besoin"])), type nullable date},
        {"Lignes de l'affaire", each Table.RowCount(_), Int64.Type},
        {"Personnes", each Personnes(_), type record}}),
    AvecPersonnes = Table.ExpandRecordColumn(Groupes, "Personnes",
        {"Ressources", "Personnes internes", "Personnes externes", "À affecter"}),
    JointesToutes = Table.NestedJoin(AvecPersonnes, {"Clé affaire (texte)"}, Toutes, {"Clé affaire (texte)"}, "Toutes",
        JoinKind.LeftOuter),
    AvecToutes = Table.ExpandTableColumn(JointesToutes, "Toutes",
        {"Heures toutes lignes (h)", "Réalisé toutes lignes (h)", "Moyenne des avancements"}),
    Totaux = Table.Group(LignesP, {"Ordre périmètre"}, {{"Total du périmètre", each List.Sum([#"Heures (h)"]), type number}}),
    JointesTotaux = Table.NestedJoin(AvecToutes, {"Ordre périmètre"}, Totaux, {"Ordre périmètre"}, "Totaux",
        JoinKind.LeftOuter),
    AvecTotaux = Table.ExpandTableColumn(JointesTotaux, "Totaux", {"Total du périmètre"}),
    // libellés et chef de la 1re ligne de l'affaire dans le périmètre
    AjoutProjet = Table.AddColumn(AvecTotaux, "Projet", each [Première][Projet], type text),
    AjoutSousProjet = Table.AddColumn(AjoutProjet, "Sous-projet", each [Première][#"Sous-projet"], type text),
    AjoutProduit = Table.AddColumn(AjoutSousProjet, "Produit", each [Première][Produit], type text),
    AjoutChef = Table.AddColumn(AjoutProduit, "Chef de projet", each [Première][#"Chef de projet"], type text),
    AjoutPremièreLigne = Table.AddColumn(AjoutChef, "N° 1re ligne", each [Première][#"N° ligne"], Int64.Type),
    AjoutAffaire = Table.AddColumn(AjoutPremièreLigne, "Affaire",
        each Text.Combine(List.Transform({[Projet], [#"Sous-projet"], [Produit]}, each if _ = "" then "—" else _),
            " · "), type text),
    AjoutPart = Table.AddColumn(AjoutAffaire, "% de la charge totale",
        each if [Total du périmètre] = 0 then 0 else [#"Charge (h)"] / [Total du périmètre], type number),
    AjoutAvancement = Table.AddColumn(AjoutPart, "Avancement",
        each if [#"Heures toutes lignes (h)"] > 0 then [#"Réalisé toutes lignes (h)"] / [#"Heures toutes lignes (h)"]
            else [Moyenne des avancements], type number),
    AjoutReste = Table.AddColumn(AjoutAvancement, "Reste à faire de l'affaire (h)",
        each List.Max({0, fnArrondiExcel([#"Heures toutes lignes (h)"] - [#"Réalisé toutes lignes (h)"], 2)}), type number),
    // avancement attendu : prorata des jours ouvrés écoulés entre le début et la date de besoin
    AjoutAttendu = Table.AddColumn(AjoutReste, "Avancement attendu",
        each
            let
                Début = [#"Début planifié"],
                Besoin = [#"Date de besoin"],
                Total = if Début = null or Besoin = null or Besoin <= Début then 0
                    else fnJoursOuvrés(Début, Besoin, JoursNonTravaillés)
            in
                if Début = null or Besoin = null or Besoin <= Début then null
                else if Aujourdhui <= Début then 0
                else if Aujourdhui > Besoin then 1
                else if Total = 0 then null
                else fnJoursOuvrés(Début, Date.AddDays(Aujourdhui, -1), JoursNonTravaillés) / Total,
        type nullable number),
    AjoutRetard = Table.AddColumn(AjoutAttendu, "Retard",
        each if [#"Date de besoin"] <> null and [#"Date de besoin"] < Aujourdhui and fnArrondiExcel([Avancement], 4) < 1 then 2
            else if [Avancement attendu] = null then 0
            else if fnArrondiExcel([Avancement] - [Avancement attendu], 4) < -0.15 then 1
            else 0, Int64.Type),
    AjoutStatut = Table.AddColumn(AjoutRetard, "Statut",
        each if [Retard] = 2 then "Hors délai (besoin dépassé)"
            else if [Retard] = 1 then "En retard (écart > 15 pts)" else "Dans les temps", type text),
    AjoutCouleur = Table.AddColumn(AjoutStatut, "Couleur d'alerte",
        each if [Retard] > 0 then "#B42318" else "#1F2937", type text),
    AjoutDébutAffiché = Table.AddColumn(AjoutCouleur, "Début (affiché)",
        each if [#"Début planifié"] = null then "à planifier"
            else Text.PadStart(Text.From(Date.Day([#"Début planifié"])), 2, "0") & "/"
                & Text.PadStart(Text.From(Date.Month([#"Début planifié"])), 2, "0") & "/"
                & Text.From(Date.Year([#"Début planifié"])), type text),
    // rang du tableau des affaires : charge décroissante, puis ordre de l'onglet Charge
    ParPérimètre = Table.Group(AjoutDébutAffiché, {"Ordre périmètre"}, {{"Affaires",
        each Table.AddIndexColumn(Table.Sort(_, {{"Charge (h)", Order.Descending}, {"N° 1re ligne", Order.Ascending}}),
            "N°", 1, 1, Int64.Type), type table}}),
    Résultat = Table.Combine(ParPérimètre[Affaires])''')

table("Charge mensuelle", "Heures réparties par variante de périmètre et de répartition, par ligne et par mois.", [
    H("Clé ligne", "int"),
    H("Clé affaire", "int"),
    H("Répartition", "text"),
    H("Mois", "date"),
    C("Heures (h)", "num", F_H1, summarize="sum"),
], r'''
let
    LignesP = Table.SelectColumns(LignesPérimètres, {"Clé ligne", "Clé affaire", "N° ligne"}),
    Options = Table.SelectColumns(Répartition, {"Répartition", "Code répartition"}),
    Croisées = Table.AddColumn(LignesP, "Options", each Options, type table),
    Développées = Table.ExpandTableColumn(Croisées, "Options", {"Répartition", "Code répartition"}),
    Jointes = Table.NestedJoin(Développées, {"N° ligne", "Code répartition"}, RépartitionLignes,
        {"N° ligne", "Code répartition"}, "Mois répartis", JoinKind.Inner),
    ParMois = Table.ExpandTableColumn(Jointes, "Mois répartis", {"Mois", "Heures du mois"}, {"Mois", "Heures (h)"}),
    Résultat = ParMois''')

table("Mois", "Calendrier mensuel couvrant la fenêtre du Dashboard Excel et tous les mois de charge répartie.", [
    C("Mois", "date", F_MOIS, desc="1er jour du mois."),
    C("Libellé mois", "text", sort_by="Clé de tri", desc="Ex. « sept. 2026 »."),
    C("Mois (long)", "text", sort_by="Clé de tri", desc="Ex. « septembre 2026 »."),
    H("Clé de tri", "int"),
    C("Année", "int", F_INT),
    H("Fin de mois", "date"),
    C("N° mois Excel", "int", F_INT, desc="Rang du mois dans les 12 mois du Dashboard Excel (vide hors fenêtre)."),
    C("Période", "text", sort_by="Ordre période", desc="« 12 mois du Dashboard Excel », « Avant » ou « Après »."),
    H("Ordre période", "int"),
], r'''
let
    Fenêtre = FenêtreExcel,
    Noms = NomsDesMois,
    MoisRépartis = List.Distinct(RépartitionLignes[Mois]),
    Premier = List.Min({Fenêtre[#"1er mois affiché"]} & MoisRépartis),
    Dernier = List.Max({Fenêtre[#"Dernier mois affiché"]} & MoisRépartis),
    NbMois = (Date.Year(Dernier) - Date.Year(Premier)) * 12 + Date.Month(Dernier) - Date.Month(Premier) + 1,
    Liste = List.Transform({0 .. NbMois - 1}, each Date.AddMonths(Premier, _)),
    Source = Table.FromList(Liste, Splitter.SplitByNothing(), {"Mois"}),
    AjoutLibellé = Table.AddColumn(Source, "Libellé mois",
        each Noms[Courts]{Date.Month([Mois]) - 1} & " " & Text.From(Date.Year([Mois])), type text),
    AjoutLong = Table.AddColumn(AjoutLibellé, "Mois (long)",
        each Noms[Longs]{Date.Month([Mois]) - 1} & " " & Text.From(Date.Year([Mois])), type text),
    AjoutTri = Table.AddColumn(AjoutLong, "Clé de tri", each Date.Year([Mois]) * 100 + Date.Month([Mois]), Int64.Type),
    AjoutAnnée = Table.AddColumn(AjoutTri, "Année", each Date.Year([Mois]), Int64.Type),
    AjoutFin = Table.AddColumn(AjoutAnnée, "Fin de mois", each Date.EndOfMonth([Mois]), type date),
    AjoutNuméro = Table.AddColumn(AjoutFin, "N° mois Excel",
        each
            let
                n = (Date.Year([Mois]) - Date.Year(Fenêtre[#"1er mois affiché"])) * 12
                    + Date.Month([Mois]) - Date.Month(Fenêtre[#"1er mois affiché"]) + 1
            in
                if n >= 1 and n <= 12 then n else null, type nullable number),
    AjoutPériode = Table.AddColumn(AjoutNuméro, "Période",
        each if [#"N° mois Excel"] <> null then "12 mois du Dashboard Excel"
            else if [Mois] < Fenêtre[#"1er mois affiché"] then "Avant" else "Après", type text),
    AjoutOrdrePériode = Table.AddColumn(AjoutPériode, "Ordre période",
        each if [Période] = "Avant" then 1 else if [Période] = "Après" then 3 else 2, Int64.Type),
    Résultat = AjoutOrdrePériode''')

table("Réglages", "Réglages lus dans le classeur (une ligne) et valeurs dérivées. Suffixe (Excel) : valeur du "
      "classeur, distincte des mesures du même nom qui suivent les segments.", [
    C("Périmètre Excel", "text", desc="Valeur de p_Perimetre."),
    C("Répartition Excel", "text", desc="Valeur de p_Repartition."),
    C("Périmètre appliqué (Excel)", "text"),
    C("Répartition appliquée (Excel)", "text"),
    C("1er mois saisi", "date", F_MOIS, desc="p_PremierMois ramené au 1er du mois (vide = automatique)."),
    C("1er mois automatique", "date", F_MOIS, desc="Mois du plus petit début planifié du périmètre Excel."),
    C("1er mois affiché (Excel)", "date", F_MOIS, desc="1er des 12 mois du Dashboard Excel."),
    C("Dernier mois affiché (Excel)", "date", F_MOIS, desc="12e mois du Dashboard Excel."),
    C("Période affichée (Excel)", "text", desc="Bandeau des onglets Dashboard et Dessinateurs."),
    C("Capacité par défaut (h/mois)", "num", F_H, desc="p_HeuresETP."),
    C("Charge fixe par défaut (h/mois)", "num", F_H, desc="p_ChargeFixe."),
    C("Capacité projets par personne (Excel)", "num", F_H, desc="Capacité par défaut - charge fixe par défaut "
      "(diviseur des ETP, h/mois)."),
    C("Capacité projets de l'équipe (Excel)", "num", F_H, desc="Somme des capacités projets des ressources des "
      "listes Data (h/mois)."),
    C("Rappel capacité", "text", desc="« 120 h/pers. dont 20 h hors projet »."),
    C("Rappel des réglages (Excel)", "text", desc="Rappel des réglages du Dashboard Excel (même texte que l'Excel)."),
    C("Lignes valides (Excel)", "int", F_INT, desc="Lignes de Charge dont le nom de projet est renseigné."),
    C("Lignes terminées (Excel)", "int", F_INT, desc="Lignes valides à 100 % (tous périmètres)."),
    C("Lignes à corriger (Excel)", "int", F_INT, desc="Lignes avec un dessinateur ET une ressource externe."),
], r'''
let
    P = ParamètresExcel,
    F = FenêtreExcel,
    Noms = NomsDesMois,
    MoisLong = (d as date) as text => Noms[Longs]{Date.Month(d) - 1} & " " & Text.From(Date.Year(d)),
    Équipe = Table.SelectRows(Ressources, each [#"Dans l'équipe"]),
    SommeÉquipe = List.Sum(Équipe[#"Capacité projets (h/mois)"]),
    RappelCapacité = P[#"Capacité (texte)"] & " h/pers. dont " & P[#"Charge fixe (texte)"] & " h hors projet",
    Rappel = P[#"Périmètre Excel"] & "  ·  "
        & (if P[#"Lissée"] then "lissage en jours ouvrés jusqu'au besoin" else "tout sur le mois de début")
        & "  ·  " & RappelCapacité,
    Résultat = #table(
        {"Périmètre Excel", "Répartition Excel", "Périmètre appliqué (Excel)", "Répartition appliquée (Excel)",
            "1er mois saisi", "1er mois automatique", "1er mois affiché (Excel)", "Dernier mois affiché (Excel)",
            "Période affichée (Excel)", "Capacité par défaut (h/mois)", "Charge fixe par défaut (h/mois)",
            "Capacité projets par personne (Excel)", "Capacité projets de l'équipe (Excel)", "Rappel capacité",
            "Rappel des réglages (Excel)", "Lignes valides (Excel)", "Lignes terminées (Excel)",
            "Lignes à corriger (Excel)"},
        {{
            P[#"Périmètre Excel"],
            P[#"Répartition Excel"],
            if P[#"Tous les projets"] then "Tous les projets" else "Projets en cours",
            if P[#"Lissée"] then "Lissée jusqu'à la date de besoin" else "Tout sur le mois de début",
            P[#"Premier mois saisi"],
            F[#"1er mois automatique"],
            F[#"1er mois affiché"],
            F[#"Dernier mois affiché"],
            "Période " & MoisLong(F[#"1er mois affiché"]) & " – " & MoisLong(F[#"Dernier mois affiché"]),
            P[#"Capacité par défaut"],
            P[#"Charge fixe par défaut"],
            P[#"Capacité par défaut"] - P[#"Charge fixe par défaut"],
            if SommeÉquipe = null then 0 else SommeÉquipe,
            RappelCapacité,
            Rappel,
            Table.RowCount(DonnéesCharge),
            Table.RowCount(Table.SelectRows(DonnéesCharge, each [Terminée])),
            Table.RowCount(Table.SelectRows(DonnéesCharge, each [Dessinateur et externe]))
        }})''')

table("Actualisation", "Date et heure de la dernière actualisation (heure de Paris).", [
    C("Actualisé le", "datetime", "dd/mm/yyyy hh:nn"),
    C("Aujourd'hui", "date", desc="Date utilisée pour l'avancement attendu et les retards."),
    C("Libellé", "text"),
], r'''
let
    Heure = HeureParis,
    Deux = (n as number) as text => Text.PadStart(Text.From(n), 2, "0"),
    Libellé = "Données lues le " & Deux(Date.Day(Heure)) & "/" & Deux(Date.Month(Heure)) & "/"
        & Text.From(Date.Year(Heure)) & " à " & Deux(Time.Hour(Heure)) & ":" & Deux(Time.Minute(Heure))
        & " (heure de Paris)",
    Résultat = #table({"Actualisé le", "Aujourd'hui", "Libellé"}, {{Heure, Date.From(Heure), Libellé}})''')

CONTROLES = [  # (ligne de Calculs!AK:AL, code, nature, tolérance)
    (4, "total", "Heures", 0.01), (5, "nbAff", "Nombre", 0.001), (6, "nbChefs", "Nombre", 0.001),
    (7, "nbProj", "Nombre", 0.001), (8, "nbTypes", "Nombre", 0.001), (9, "realise", "Heures", 0.01),
    (10, "reste", "Heures", 0.01), (11, "avgSimple", "Pourcentage", 0.0001), (12, "avgPond", "Pourcentage", 0.0001),
    (13, "autoStart", "Mois", 0), (14, "start", "Mois", 0), (15, "periode", "Heures", 0.01),
    (16, "hors", "Heures", 0.01), (17, "pic", "Heures", 0.01), (19, "picMois", "Mois", 0),
    (20, "picActives", "Nombre", 0.001), (21, "picETP", "Décimal", 0.001), (22, "totalETP", "Décimal", 0.001),
    (24, "nbTermines", "Nombre", 0.001), (25, "finPeriode", "Mois", 0), (26, "picPart", "Pourcentage", 0.0001),
    (27, "periodeTxt", "Texte", 0), (33, "capEquipe", "Heures", 0.01), (34, "nbSurcharge", "Nombre", 0.001),
    (35, "premSurcharge", "Mois", 0), (36, "picCap", "Pourcentage", 0.0001), (37, "nbRetard", "Nombre", 0.001),
    (38, "nbDepasse", "Nombre", 0.001), (39, "nbEcart", "Nombre", 0.001), (40, "reglagesTxt", "Texte", 0),
]
_controles_m = ",\n        ".join('{%d, "%s", "%s", %s}' % (r, c, n, repr(t)) for r, c, n, t in CONTROLES)

table("Contrôle Excel", "Indicateurs et charge mensuelle calculés par Excel (onglet Calculs), comparés aux mesures "
      "Power BI avec les réglages Excel.", [
    C("Ordre", "int", F_INT, hidden=True),
    C("Catégorie", "text"),
    C("Contrôle", "text", sort_by="Ordre"),
    H("Code", "text"),
    C("Nature", "text", desc="Heures, Nombre, Décimal, Pourcentage, Mois ou Texte."),
    C("Cellule Excel", "text"),
    C("Mois", "date", F_MOIS, desc="Mois de la ligne (charge par mois) ou valeur Excel d'un indicateur de type Mois."),
    C("Valeur lue", "num", "#,0.00", desc="Valeur numérique lue dans l'onglet Calculs."),
    C("Valeur lue (texte)", "text", desc="Valeur lue (indicateurs de type Texte ou Mois)."),
    C("Tolérance", "num", "0.0000"),
], r'''
let
    Noms = NomsDesMois,
    MoisCourt = (d as nullable date) as text =>
        if d = null then "" else Noms[Courts]{Date.Month(d) - 1} & " " & Text.From(Date.Year(d)),
    // indicateurs : Calculs!AK4:AL40 (libellé | valeur), lus par position
    Indicateurs = Table.SelectColumns(fnLireNom("bi_Indicateurs"), {"Column1", "Column2"}, MissingField.UseNull),
    IndSansErreurs = Table.ReplaceErrorValues(Indicateurs, {{"Column1", null}, {"Column2", null}}),
    IndNumérotés = Table.AddIndexColumn(IndSansErreurs, "Ligne", 4, 1, Int64.Type),
    Comparés = #table(type table [Ligne = Int64.Type, Code = text, Nature = text, Tolérance = number], {
        ''' + _controles_m + r'''
    }),
    Joints = Table.NestedJoin(Comparés, {"Ligne"}, IndNumérotés, {"Ligne"}, "Excel", JoinKind.LeftOuter),
    Développés = Table.ExpandTableColumn(Joints, "Excel", {"Column1", "Column2"}, {"Libellé Excel", "Valeur brute"}),
    LignesIndicateurs = Table.FromRecords(List.Transform(Table.ToRecords(Développés), (r) => [
        Ordre = r[Ligne],
        Catégorie = "Indicateurs",
        // libellé unique (colonne triée par Ordre) même si la cellule de l'onglet Calculs est vide
        Contrôle = let l = fnTexteExcel(r[#"Libellé Excel"]) in if l = "" then "Indicateur " & r[Code] else l,
        Code = r[Code],
        Nature = r[Nature],
        #"Cellule Excel" = "Calculs!AL" & Text.From(r[Ligne]),
        #"Mois" = if r[Nature] = "Mois" then fnDateExcel(r[#"Valeur brute"]) else null,
        #"Valeur lue" = if r[Nature] = "Mois" or r[Nature] = "Texte" then null else fnNombreExcel(r[#"Valeur brute"]),
        #"Valeur lue (texte)" =
            if r[Nature] = "Texte" then fnTexteExcel(r[#"Valeur brute"])
            else if r[Nature] = "Mois" then MoisCourt(fnDateExcel(r[#"Valeur brute"]))
            else null,
        Tolérance = r[Tolérance]
    ]), type table [Ordre = Int64.Type, Catégorie = text, Contrôle = text, Code = text, Nature = text,
        #"Cellule Excel" = text, Mois = nullable date, #"Valeur lue" = nullable number,
        #"Valeur lue (texte)" = nullable text, Tolérance = number]),
    // charge des 12 mois affichés : Calculs!AO4:AP15 (mois | charge)
    MoisExcel = Table.SelectColumns(fnLireNom("bi_Mois"), {"Column1", "Column2"}, MissingField.UseNull),
    MoisSansErreurs = Table.ReplaceErrorValues(MoisExcel, {{"Column1", null}, {"Column2", null}}),
    MoisNumérotés = Table.AddIndexColumn(MoisSansErreurs, "N°", 1, 1, Int64.Type),
    LignesMois = Table.FromRecords(List.Transform(Table.ToRecords(MoisNumérotés), (r) => [
        Ordre = 100 + r[#"N°"],
        Catégorie = "Charge par mois",
        Contrôle = "Mois " & Text.PadStart(Text.From(r[#"N°"]), 2, "0") & " : " & MoisCourt(fnDateExcel(r[Column1])),
        Code = "mois",
        Nature = "Heures",
        #"Cellule Excel" = "Calculs!AP" & Text.From(r[#"N°"] + 3),
        #"Mois" = fnDateExcel(r[Column1]),
        #"Valeur lue" = fnNombreExcel(r[Column2]),
        #"Valeur lue (texte)" = null,
        Tolérance = 0.01
    ]), Value.Type(LignesIndicateurs)),
    Résultat = Table.Combine({LignesIndicateurs, LignesMois})''')

table("Mesures", "Table des mesures du rapport.", [
    H("Colonne technique", "text"),
], r'''
let
    Résultat = #table(type table [#"Colonne technique" = text], {})''')

# =============================================================================================== mesures DAX
MEASURES = []   # (nom, dossier, format, description, expression)
_P = "VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], \"%s\" )" % EXCEL
_R = "VAR _r = SELECTEDVALUE ( 'Répartition'[Répartition], \"%s\" )" % EXCEL


def mesure(name, folder, fmt, desc, dax):
    MEASURES.append((name, folder, fmt, desc, dax.strip("\n")))


def fp(expr_, plus0=True, extra=""):
    """Agrégation dans la variante de périmètre sélectionnée (« Comme dans Excel » si aucune ou plusieurs)."""
    return (f"{_P}\nRETURN\n    CALCULATE ( {expr_}, 'Périmètre'[Périmètre] = _p{extra} )" + (" + 0" if plus0 else ""))


def fpr(expr_, plus0=True):
    """Agrégation dans les variantes de périmètre et de répartition sélectionnées."""
    return (f"{_P}\n{_R}\nRETURN\n    CALCULATE ( {expr_}, 'Périmètre'[Périmètre] = _p, 'Répartition'[Répartition] = _r )"
            + (" + 0" if plus0 else ""))


def top_affaire(expr_):
    """Valeur de l'affaire au plus gros reste à faire (égalité : 1re dans l'ordre de l'onglet Charge)."""
    return (f"{_P}\nRETURN\n    CALCULATE (\n        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, "
            f"Affaires[N° 1re ligne], ASC )\n        RETURN {expr_},\n        'Périmètre'[Périmètre] = _p\n    )")


# ---- Charge
D_CH = "Charge"
mesure("Charge identifiée (h)", D_CH, F_H,
       "Somme des heures des lignes du périmètre (onglet Charge, colonne M).", fp("SUM ( Lignes[Heures (h)] )"))
mesure("Réalisé (h)", D_CH, F_H, "Heures × avancement des lignes du périmètre.", fp("SUM ( Lignes[Réalisé de la ligne (h)] )"))
mesure("Reste à faire (h)", D_CH, F_H, "Charge identifiée - réalisé.", "[Charge identifiée (h)] - [Réalisé (h)]")
mesure("Capacité projets par personne (h/mois)", D_CH, F_H,
       "Capacité par défaut - charge fixe par défaut (réglages du Dashboard Excel).",
       "MAX ( Réglages[Capacité projets par personne (Excel)] )")
mesure("ETP-mois", D_CH, F_DEC, "Charge identifiée / (capacité - charge fixe) d'une personne (0 si diviseur <= 0).", """
VAR _c = [Capacité projets par personne (h/mois)]
RETURN
    IF ( _c <= 0, 0, DIVIDE ( [Charge identifiée (h)], _c ) )""")
mesure("Nombre de chefs de projet", D_CH, F_INT, "Chefs distincts des lignes du périmètre.",
       fp("DISTINCTCOUNT ( Lignes[Chef de projet] )"))
mesure("Nombre de projets", D_CH, F_INT, "Projets distincts (casse ignorée).", fp("DISTINCTCOUNT ( Lignes[Projet] )"))
mesure("Nombre de types", D_CH, F_INT, "Types d'affaire distincts.", fp("DISTINCTCOUNT ( Lignes[Type] )"))
mesure("Nombre de lignes", D_CH, F_INT, "Lignes de l'onglet Charge dans le périmètre.", fp("COUNTROWS ( Lignes )"))
mesure("Part de la charge (%)", D_CH, F_PCT1, "Part de la charge identifiée dans le total de la sélection "
       "(segments et filtres de la page).", """
DIVIDE ( [Charge identifiée (h)], CALCULATE ( [Charge identifiée (h)], ALLSELECTED () ) )""")

# ---- Affaires
D_AF = "Affaires"
mesure("Nombre d'affaires", D_AF, F_INT, "Affaires (sous-projets regroupés) ayant au moins une ligne dans le "
       "périmètre.", fp("COUNTROWS ( Affaires )"))
mesure("Avancement moyen simple", D_AF, F_PCT, "Moyenne simple des avancements des affaires.", f"""
{_P}
RETURN
    CALCULATE (
        VAR _n = COUNTROWS ( Affaires )
        RETURN IF ( _n = 0, 0, DIVIDE ( SUM ( Affaires[Avancement] ), _n ) ),
        'Périmètre'[Périmètre] = _p
    )""")
mesure("Avancement moyen pondéré", D_AF, F_PCT, "Réalisé / heures de toutes les lignes des affaires du périmètre "
       "(repli : moyenne simple).", f"""
{_P}
RETURN
    CALCULATE (
        VAR _h = SUM ( Affaires[Heures toutes lignes (h)] )
        RETURN IF ( _h = 0, [Avancement moyen simple], DIVIDE ( SUM ( Affaires[Réalisé toutes lignes (h)] ), _h ) ),
        'Périmètre'[Périmètre] = _p
    )""")
mesure("Affaires en retard", D_AF, F_INT, "Affaires hors délai ou en écart de plus de 15 points.",
       fp("COUNTROWS ( Affaires )", extra=", KEEPFILTERS ( Affaires[Retard] > 0 )"))
mesure("Affaires hors délai", D_AF, F_INT, "Date de besoin dépassée sans être à 100 %.",
       fp("COUNTROWS ( Affaires )", extra=", KEEPFILTERS ( Affaires[Retard] = 2 )"))
mesure("Affaires en écart", D_AF, F_INT, "Avancement inférieur de plus de 15 points à l'avancement attendu.",
       fp("COUNTROWS ( Affaires )", extra=", KEEPFILTERS ( Affaires[Retard] = 1 )"))
mesure("Plus gros reste à faire (h)", D_AF, F_H, "Reste à faire de l'affaire qui en a le plus.",
       f"{_P}\nRETURN\n    COALESCE ( CALCULATE ( MAX ( Affaires[Reste à faire de l'affaire (h)] ), 'Périmètre'[Périmètre] = _p ), 0 )")
for _nom, _expr, _fmt in (
        ("affaire", "CONCATENATEX ( _t, Affaires[Affaire] )", None),
        ("projet", "CONCATENATEX ( _t, Affaires[Projet] )", None),
        ("sous-projet", "CONCATENATEX ( _t, Affaires[Sous-projet] )", None),
        ("produit", "CONCATENATEX ( _t, Affaires[Produit] )", None),
        ("chef de projet", "CONCATENATEX ( _t, Affaires[Chef de projet] )", None),
        ("ressources", "CONCATENATEX ( _t, Affaires[Ressources] )", None),
        ("total de l'affaire (h)", "SUMX ( _t, Affaires[Heures toutes lignes (h)] )", F_H),
        ("début planifié", "MAXX ( _t, Affaires[Début planifié] )", F_DATE),
        ("date de besoin", "MAXX ( _t, Affaires[Date de besoin] )", F_DATE),
        ("avancement", "MAXX ( _t, Affaires[Avancement] )", F_PCT),
        ("retard", "MAXX ( _t, Affaires[Retard] )", F_INT)):
    mesure(f"Plus gros reste à faire – {_nom}", D_AF, _fmt,
           f"Affaire au plus gros reste à faire : {_nom} (égalité : 1re dans l'ordre de l'onglet Charge).",
           top_affaire(_expr))
mesure("Couleur avancement", D_AF, None, "Couleur de l'avancement d'une affaire : rouge si retard (mise en forme "
       "conditionnelle, valeur de champ).", f"""
IF ( SELECTEDVALUE ( Affaires[Retard], 0 ) > 0, "{RETARD_COULEUR}", "{ENCRE}" )""")
mesure("Lignes terminées", D_AF, F_INT, "Lignes valides à 100 % de toutes les lignes, quel que soit le périmètre "
       "(comme l'Excel) ; suit les segments et clics sur les colonnes de Lignes (chef, projet, type), pas un clic sur "
       "une affaire.", """
CALCULATE (
    COUNTROWS ( Lignes ),
    REMOVEFILTERS ( Affaires ),
    'Périmètre'[Périmètre] = "Tous les projets",
    KEEPFILTERS ( Lignes[Terminée] = TRUE () )
) + 0""")
mesure("Lignes avec dessinateur et externe", D_AF, F_INT, "Lignes ayant un dessinateur ET une ressource externe "
       "(à corriger dans Charge), parmi toutes les lignes valides ; suit les segments chef, projet, type.", """
CALCULATE (
    COUNTROWS ( Lignes ),
    REMOVEFILTERS ( Affaires ),
    'Périmètre'[Périmètre] = "Tous les projets",
    KEEPFILTERS ( Lignes[Dessinateur et externe] = TRUE () )
) + 0""")

# ---- Mois et capacité
D_MO = "Mois et capacité"
mesure("Charge répartie (h)", D_MO, F_H, "Heures réparties sur les mois du contexte (période, mois du graphique).",
       fpr("SUM ( 'Charge mensuelle'[Heures (h)] )"))
mesure("Charge répartie sans les mois vides (h)", D_MO, F_H, "Comme [Charge répartie (h)], mais vide (et non 0) "
       "sans heure répartie : un graphique n'affiche alors que les mois et ressources qui ont des heures.",
       fpr("SUM ( 'Charge mensuelle'[Heures (h)] )", plus0=False))
mesure("Heures de la ligne sur la période (h)", D_MO, F_H, "Charge répartie des lignes du contexte, 0 pour une ligne "
       "sans heure sur la période (sans date de début, 0 h, hors des mois choisis) ; vide pour une combinaison "
       "ligne × affaire qui n'existe pas.", """
IF ( NOT ( ISEMPTY ( Lignes ) ), [Charge répartie (h)] )""")
mesure("Affaires actives", D_MO, F_INT, "Affaires ayant au moins une heure répartie sur les mois du contexte.",
       fpr("DISTINCTCOUNT ( 'Charge mensuelle'[Clé affaire] )"))
mesure("Charge hors période ou sans date de début (h)", D_MO, F_H, "Charge identifiée non répartie sur les mois "
       "sélectionnés (arrondie à 0,1 h).", """
MAX ( 0, ROUND ( [Charge identifiée (h)] - [Charge répartie (h)], 1 ) )""")
mesure("Capacité projets de l'équipe (h/mois)", D_MO, F_H, "Somme des (capacité - charge fixe) des ressources des "
       "listes Data : constante par mois (ligne pointillée).", "MAX ( Réglages[Capacité projets de l'équipe (Excel)] )")
mesure("Mois en surcharge", D_MO, F_INT, "Mois où la charge arrondie dépasse la capacité projets arrondie.", """
VAR _cap = ROUND ( [Capacité projets de l'équipe (h/mois)], 0 )
RETURN
    COUNTROWS ( FILTER ( VALUES ( Mois[Mois] ), ROUND ( [Charge répartie (h)], 0 ) > _cap ) ) + 0""")
mesure("Premier mois en surcharge", D_MO, F_MOIS, "Premier mois en surcharge (vide si aucun).", """
VAR _cap = ROUND ( [Capacité projets de l'équipe (h/mois)], 0 )
RETURN
    MINX ( FILTER ( VALUES ( Mois[Mois] ), ROUND ( [Charge répartie (h)], 0 ) > _cap ), Mois[Mois] )""")
mesure("Pic de charge (h)", D_MO, F_H, "Charge du mois le plus chargé.", """
COALESCE ( MAXX ( VALUES ( Mois[Mois] ), [Charge répartie (h)] ), 0 )""")
mesure("Mois du pic", D_MO, F_MOIS, "Mois le plus chargé (le premier en cas d'égalité).", """
VAR _pic = [Pic de charge (h)]
RETURN
    MINX ( FILTER ( VALUES ( Mois[Mois] ), [Charge répartie (h)] = _pic ), Mois[Mois] )""")
mesure("Affaires actives au pic", D_MO, F_INT, "Affaires actives sur le mois le plus chargé.", """
VAR _m = [Mois du pic]
RETURN
    IF ( ISBLANK ( _m ), 0, CALCULATE ( [Affaires actives], Mois[Mois] = _m ) )""")
mesure("ETP au pic", D_MO, F_DEC, "Pic / (capacité - charge fixe) d'une personne.", """
VAR _c = [Capacité projets par personne (h/mois)]
RETURN
    IF ( _c <= 0, 0, DIVIDE ( [Pic de charge (h)], _c ) )""")
mesure("Pic en % de la capacité", D_MO, F_PCT, "Pic / capacité projets de l'équipe (vide si capacité nulle).", """
VAR _cap = [Capacité projets de l'équipe (h/mois)]
RETURN
    IF ( _cap <= 0, BLANK (), DIVIDE ( [Pic de charge (h)], _cap ) )""")
mesure("Part du pic dans la période", D_MO, F_PCT, "Charge du mois de pic / charge répartie sur la période.", """
VAR _t = [Charge répartie (h)]
RETURN
    IF ( _t = 0, 0, DIVIDE ( [Pic de charge (h)], _t ) )""")
mesure("Premier mois affiché", D_MO, F_MOIS, "Premier mois de la période sélectionnée.", "MIN ( Mois[Mois] )")
mesure("Dernier mois affiché", D_MO, F_MOIS, "Dernier mois de la période sélectionnée.", "MAX ( Mois[Mois] )")

# ---- Groupes de projets : les 9 séries de couleur fixe des graphiques du Dashboard Excel (projets par charge
# décroissante, « Autres projets (n) » en 9e position au-delà de 9 projets)
D_GR = "Groupes de projets"
for _k, _coul in enumerate(SLOT_COLORS, 1):
    _rang = f"projet de rang {_k} par charge" if _k < 9 else "9e projet, ou « Autres projets (n) » au-delà de 9"
    mesure(f"Groupe {_k} – libellé", D_GR, None, f"Nom du groupe {_k} ({_rang}) ; couleur #{_coul}.",
           fp("SELECTEDVALUE ( Lignes[Groupe projet] )", plus0=False, extra=f", Lignes[Ordre groupe projet] = {_k}"))
    mesure(f"Groupe {_k} – charge (h)", D_GR, F_H, f"Charge identifiée du groupe {_k} ({_rang}) ; couleur #{_coul} "
           "(anneau).", f"CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = {_k} )")
    mesure(f"Groupe {_k} – charge répartie (h)", D_GR, F_H, f"Charge répartie du groupe {_k} ({_rang}) sur les mois "
           f"du contexte ; couleur #{_coul} (histogramme empilé).",
           f"CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = {_k} )")

# ---- Ressource (une ressource dans le contexte : segment, ligne de matrice)
D_RE = "Ressource"
_RF = "VAR _f = [Charge fixe de la ressource (h/mois)]\nVAR _c = [Capacité de la ressource (h/mois)]"
mesure("Capacité de la ressource (h/mois)", D_RE, F_H, "Capacité mensuelle d'une ressource (vide pour « (non "
       "affecté) » ou plusieurs ressources).", """
IF ( HASONEVALUE ( Ressources[Clé ressource] ), MAX ( Ressources[Capacité (h/mois)] ) )""")
mesure("Charge fixe de la ressource (h/mois)", D_RE, F_H, "Charge fixe hors projet d'une ressource.", """
IF ( HASONEVALUE ( Ressources[Clé ressource] ), MAX ( Ressources[Charge fixe (h/mois)] ) )""")
mesure("Capacité projets de la ressource (h/mois)", D_RE, F_H, "max(0, capacité - charge fixe).", """
IF ( HASONEVALUE ( Ressources[Clé ressource] ), MAX ( Ressources[Capacité projets (h/mois)] ) )""")
mesure("Charge fixe (h)", D_RE, F_H, "Charge fixe de la ressource sur les mois du contexte.", """
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), _f * COUNTROWS ( VALUES ( Mois[Mois] ) ) )""")
mesure("Charge totale de la ressource (h)", D_RE, F_H, "Charge fixe + heures projets de la ressource.", """
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), SUMX ( VALUES ( Mois[Mois] ), _f + [Charge répartie (h)] ) )""")
mesure("Projets dans la capacité (h)", D_RE, F_H, "Heures projets tenant dans la capacité (bleu du graphique "
       "Excel).", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _h = [Charge répartie (h)]
            VAR _marge = _c - ROUND ( _f + _h, 0 )
            RETURN IF ( _marge < 0, MAX ( 0, MIN ( _h, _c - _f ) ), _h )
        )
    )""")
mesure("Dépassement de capacité (h)", D_RE, F_H, "Heures projets au-delà de la capacité (rouge du graphique "
       "Excel).", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _h = [Charge répartie (h)]
            VAR _marge = _c - ROUND ( _f + _h, 0 )
            RETURN _h - IF ( _marge < 0, MAX ( 0, MIN ( _h, _c - _f ) ), _h )
        )
    )""")
mesure("Marge de la ressource (h)", D_RE, "+#,0;-#,0;0", "Capacité - charge totale arrondie (négatif = "
       "dépassement).", f"""
{_RF}
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), SUMX ( VALUES ( Mois[Mois] ), _c - ROUND ( _f + [Charge répartie (h)], 0 ) ) )""")
mesure("Dépassement cumulé (h)", D_RE, F_H, "Somme des dépassements mensuels.", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _m = _c - ROUND ( _f + [Charge répartie (h)], 0 )
            RETURN IF ( _m < 0, - _m, 0 )
        )
    )""")
mesure("Heures encore disponibles (h)", D_RE, F_H, "Somme des marges mensuelles positives.", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _m = _c - ROUND ( _f + [Charge répartie (h)], 0 )
            RETURN IF ( _m > 0, _m, 0 )
        )
    )""")
_OCC = """
            VAR _t = ROUND ( _f + [Charge répartie (h)], 0 )
            RETURN IF ( _c <= 0, IF ( _t > 0, 9.99, 0 ), DIVIDE ( _t, _c ) )"""
mesure("Occupation (%)", D_RE, F_PCT, "(Charge fixe + heures projets, arrondi à l'heure) / capacité ; moyenne des "
       "mois si plusieurs mois ; 999 % si capacité nulle avec de la charge.", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        AVERAGEX (
            VALUES ( Mois[Mois] ),{_OCC}
        )
    )""")
mesure("Occupation maximale (%)", D_RE, F_PCT, "Occupation du mois le plus occupé.", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        MAXX (
            VALUES ( Mois[Mois] ),{_OCC}
        )
    )""")
mesure("Couleur de fond (occupation)", D_RE, None, "Vert < 90 %, orange 90-100 %, rouge > 100 % (fonds de "
       "l'Excel).", f"""
VAR _o = [Occupation (%)]
RETURN
    SWITCH ( TRUE (), ISBLANK ( _o ), BLANK (), _o > 1, "{OCC_ROUGE[0]}", _o >= 0.9, "{OCC_ORANGE[0]}", "{OCC_VERT[0]}" )""")
mesure("Couleur du texte (occupation)", D_RE, None, "Texte rouge, orange foncé ou vert (couleurs de l'Excel).", f"""
VAR _o = [Occupation (%)]
RETURN
    SWITCH ( TRUE (), ISBLANK ( _o ), BLANK (), _o > 1, "{OCC_ROUGE[1]}", _o >= 0.9, "{OCC_ORANGE[1]}", "{OCC_VERT[1]}" )""")
mesure("Mois en surcharge de la ressource", D_RE, F_INT, "Mois où la charge totale arrondie dépasse la "
       "capacité.", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        COUNTROWS ( FILTER ( VALUES ( Mois[Mois] ), _c - ROUND ( _f + [Charge répartie (h)], 0 ) < 0 ) ) + 0
    )""")
mesure("Premier mois en surcharge de la ressource", D_RE, F_MOIS, "Premier mois en surcharge (vide si aucun).", f"""
{_RF}
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        MINX ( FILTER ( VALUES ( Mois[Mois] ), _c - ROUND ( _f + [Charge répartie (h)], 0 ) < 0 ), Mois[Mois] )
    )""")
mesure("Charge au pic de la ressource (h)", D_RE, F_H, "Charge totale (fixe + projets) du mois le plus chargé.", """
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), MAXX ( VALUES ( Mois[Mois] ), _f + [Charge répartie (h)] ) )""")
mesure("Mois le plus chargé de la ressource", D_RE, F_MOIS, "Mois de charge totale maximale (le premier en cas "
       "d'égalité).", """
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _pic = [Charge au pic de la ressource (h)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        MINX ( FILTER ( VALUES ( Mois[Mois] ), _f + [Charge répartie (h)] = _pic ), Mois[Mois] )
    )""")
mesure("Occupation au pic de la ressource (%)", D_RE, F_PCT, "Occupation du mois le plus chargé.", """
VAR _m = [Mois le plus chargé de la ressource]
RETURN
    IF ( NOT ( ISBLANK ( _m ) ), CALCULATE ( [Occupation (%)], Mois[Mois] = _m ) )""")

# ---- Textes (cartes, sous-titres)
D_TX = "Textes"


def mois_long(var):
    """Libellé long (« septembre 2026 ») du mois contenu dans la variable DAX `var`, lu dans la table Mois sans tenir
    compte des filtres de la page."""
    return f"CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], {var} ), REMOVEFILTERS ( Mois ) )"


mesure("Périmètre en vigueur", D_TX, None, "Projets en cours ou Tous les projets.", f"""
{_P}
RETURN
    LOOKUPVALUE ( 'Périmètre'[Périmètre appliqué], 'Périmètre'[Périmètre], _p )""")
mesure("Répartition en vigueur", D_TX, None, "Lissée jusqu'à la date de besoin ou Tout sur le mois de début.", f"""
{_R}
RETURN
    LOOKUPVALUE ( 'Répartition'[Répartition appliquée], 'Répartition'[Répartition], _r )""")
mesure("Rappel des réglages", D_TX, None, "Réglages appliqués par les segments (format du bandeau Excel).", """
[Périmètre en vigueur] & "  ·  "
    & IF ( [Répartition en vigueur] = "Tout sur le mois de début", "tout sur le mois de début",
        "lissage en jours ouvrés jusqu'au besoin" )
    & "  ·  " & SELECTEDVALUE ( Réglages[Rappel capacité] )""")
mesure("Données lues le", D_TX, None, "Date et heure de l'actualisation (heure de Paris).",
       "SELECTEDVALUE ( Actualisation[Libellé] )")
mesure("Titre du périmètre", D_TX, None, "PROJETS EN COURS ou TOUS LES PROJETS.", """
IF ( [Périmètre en vigueur] = "Tous les projets", "TOUS LES PROJETS", "PROJETS EN COURS" )""")
mesure("Sous-titre chefs", D_TX, None, "Sous-titre de la carte des chefs de projet.", """
IF ( [Périmètre en vigueur] = "Tous les projets", "sur l'ensemble des projets", "sur les projets en cours" )""")
mesure("Sous-titre charge identifiée", D_TX, None, "« soit ≈ x ETP-mois ».", """
"soit ≈ " & FORMAT ( [ETP-mois], "0.0", "fr-FR" ) & " ETP-mois\"""")
mesure("Sous-titre capacité", D_TX, None, "« Capacité projets : x h/mois ».", """
"Capacité projets : " & FORMAT ( [Capacité projets de l'équipe (h/mois)], "#,0", "fr-FR" ) & " h/mois\"""")
mesure("Sous-titre surcharge", D_TX, None, "« Dès <mois> » ou « Pic à x % de la capacité ».", f"""
VAR _m = [Premier mois en surcharge]
VAR _pc = [Pic en % de la capacité]
RETURN
    IF (
        NOT ( ISBLANK ( _m ) ),
        "Dès " & {mois_long("_m")},
        IF ( ISBLANK ( _pc ), "", "Pic à " & FORMAT ( ROUND ( _pc * 100, 0 ), "0" ) & " % de la capacité" )
    )""")
mesure("Sous-titre pic (mois)", D_TX, None, "« en <mois du pic> ».", f"""
VAR _m = [Mois du pic]
RETURN
    IF ( NOT ( ISBLANK ( _m ) ), "en " & {mois_long("_m")} )""")
mesure("Sous-titre pic (affaires)", D_TX, None, "« n affaire(s) active(s) · ≈ x ETP ».", """
VAR _e = ROUND ( [ETP au pic], 1 )
VAR _n = [Affaires actives au pic]
RETURN
    _n & " affaire(s) active(s)  ·  ≈ "
        & IF ( _e = INT ( _e ), FORMAT ( _e, "0" ), FORMAT ( _e, "0.0", "fr-FR" ) ) & " ETP\"""")
mesure("Sous-titre hors délai", D_TX, None, "« dont n hors délai (besoin dépassé) ».", """
"dont " & [Affaires hors délai] & " hors délai (besoin dépassé)\"""")
mesure("Sous-titre écart", D_TX, None, "« n en retard de plus de 15 pts ».", """
[Affaires en écart] & " en retard de plus de 15 pts\"""")
mesure("Sous-titre projets", D_TX, None, "« réparties sur n projet(s) ».", """
"réparties sur " & [Nombre de projets] & " projet(s)\"""")
mesure("Sous-titre lignes terminées", D_TX, None, "Lignes terminées comptées ou non selon le périmètre.", """
IF (
    [Périmètre en vigueur] = "Tous les projets",
    "dont " & [Lignes terminées] & " ligne(s) terminée(s)",
    [Lignes terminées] & " ligne(s) terminée(s) non comptée(s)"
)""")
mesure("Sous-titre avancement", D_TX, None, "« Moyenne simple par affaire : x % ».", """
"Moyenne simple par affaire : " & FORMAT ( ROUND ( [Avancement moyen simple] * 100, 0 ), "0" ) & " %\"""")
mesure("Période affichée", D_TX, None, "« Période <premier mois> – <dernier mois> » de la sélection.", f"""
VAR _d = MIN ( Mois[Mois] )
VAR _f = MAX ( Mois[Mois] )
RETURN
    IF ( NOT ( ISBLANK ( _d ) ), "Période " & {mois_long("_d")} & " – " & {mois_long("_f")} )""")
mesure("Texte hors période", D_TX, None, "Charge hors période ou sans date de début (sous-titre du graphique).", """
VAR _h = [Charge hors période ou sans date de début (h)]
RETURN
    IF ( _h > 0, "Hors période ou sans date de début : " & FORMAT ( _h, "#,0", "fr-FR" ) & " h",
        "Toute la charge est visible sur la période" )""")
mesure("Titre de la ressource", D_TX, None, "Titre du graphique mensuel de la ressource sélectionnée.", """
IF (
    HASONEVALUE ( Ressources[Clé ressource] ),
    "CHARGE MENSUELLE DE " & UPPER ( SELECTEDVALUE ( Ressources[Ressource] ) ) & " (HEURES)",
    "CHARGE MENSUELLE (HEURES) — choisir une ressource"
)""")
mesure("Sous-titre graphique mensuel", D_TX, None, "« n mois à partir de <premier mois> » (période sélectionnée).", f"""
VAR _n = COUNTROWS ( VALUES ( Mois[Mois] ) )
VAR _m = MIN ( Mois[Mois] )
RETURN
    IF ( _n > 0, _n & " mois à partir de " & {mois_long("_m")} )""")
mesure("Sous-titre plus gros reste", D_TX, None, "« sur x h au total pour cette affaire ».", """
"sur " & FORMAT ( [Plus gros reste à faire – total de l'affaire (h)] + 0, "#,0", "fr-FR" )
    & " h au total pour cette affaire\"""")
mesure("Origine de la ressource", D_TX, None, "Liste Data d'origine de la ressource sélectionnée.", """
SWITCH (
    SELECTEDVALUE ( Ressources[Type de ressource] ),
    "Interne", "Dessinateur interne (Data, colonne E)",
    "Externe", "Ressource externe (Data, colonne H)",
    "Hors liste", "Absent des listes Data : réglages par défaut",
    "Non affecté", "Lignes sans dessinateur ni ressource externe",
    BLANK ()
)""")
mesure("Sous-titre charge fixe de la ressource", D_TX, None, "« dont x h de charge fixe ».", """
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), "dont " & FORMAT ( _f, "#,0", "fr-FR" ) & " h de charge fixe" )""")
mesure("Sous-titre capacité de la ressource", D_TX, None, "« soit x h/mois pour les projets » ou alerte si la charge "
       "fixe atteint la capacité.", """
VAR _c = [Capacité de la ressource (h/mois)]
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _c ) ),
        IF (
            _f >= _c,
            "Attention : charge fixe ≥ capacité",
            "soit " & FORMAT ( [Capacité projets de la ressource (h/mois)], "#,0", "fr-FR" ) & " h/mois pour les projets"
        )
    )""")
mesure("Sous-titre affaires de la ressource", D_TX, None, "« sur n affaire(s) ».", """
"sur " & [Nombre d'affaires] & " affaire(s)\"""")
mesure("Sous-titre charge sur la période", D_TX, None, "« dont x h sur la période » (mois sélectionnés).", """
"dont " & FORMAT ( [Charge répartie (h)], "#,0", "fr-FR" ) & " h sur la période\"""")
mesure("Sous-titre lignes de la ressource", D_TX, None, "Nombre de lignes de la ressource dans le périmètre, ou "
       "« Aucune affaire affectée à cette ressource dans le périmètre. » (comme l'onglet Dessinateurs).", """
VAR _n = [Nombre de lignes]
RETURN
    IF (
        _n = 0,
        "Aucune affaire affectée à cette ressource dans le périmètre.",
        _n & " ligne(s) dans le périmètre · heures sur la période = mois sélectionnés (0 sans date de début)"
    )""")
mesure("Titre info-bulle projet", D_TX, None, "Titre de l'info-bulle projet : mois survolé, ou toute la période "
       "affichée (anneau, tableau des projets).", f"""
VAR _m = IF ( HASONEVALUE ( Mois[Mois] ), VALUES ( Mois[Mois] ) )
RETURN
    IF (
        ISBLANK ( _m ),
        "Projet survolé · période affichée",
        "Projet survolé · " & {mois_long("_m")}
    )""")
mesure("Mois le plus chargé de la ressource (texte)", D_TX, None, "Mois le plus chargé, « — » sans charge projets sur "
       "la période.", f"""
VAR _m = [Mois le plus chargé de la ressource]
RETURN
    IF (
        ISBLANK ( _m ) || [Charge répartie (h)] = 0,
        "—",
        {mois_long("_m")}
    )""")
mesure("Sous-titre pic de la ressource", D_TX, None, "« x h au total (fixe + projets) ».", """
VAR _p = [Charge au pic de la ressource (h)]
RETURN
    IF ( NOT ( ISBLANK ( _p ) ), FORMAT ( _p, "#,0", "fr-FR" ) & " h au total (fixe + projets)" )""")
mesure("Sous-titre occupation au pic", D_TX, None, "« x % de la capacité ».", """
VAR _o = [Occupation au pic de la ressource (%)]
RETURN
    IF ( NOT ( ISBLANK ( _o ) ), FORMAT ( ROUND ( _o * 100, 0 ), "0" ) & " % de la capacité" )""")
mesure("Sous-titre dépassement cumulé", D_TX, None, "« Dépassement cumulé : x h ».", """
VAR _d = [Dépassement cumulé (h)]
RETURN
    IF ( NOT ( ISBLANK ( _d ) ), "Dépassement cumulé : " & FORMAT ( _d, "#,0", "fr-FR" ) & " h" )""")
mesure("Sous-titre surcharge de la ressource", D_TX, None, "« Aucun dépassement prévu » ou « Dès <mois> ».", f"""
VAR _n = [Mois en surcharge de la ressource]
VAR _m = [Premier mois en surcharge de la ressource]
RETURN
    IF (
        NOT ( ISBLANK ( _n ) ),
        IF ( _n = 0, "Aucun dépassement prévu", "Dès " & {mois_long("_m")} )
    )""")
mesure("Sous-titre heures disponibles", D_TX, None, "« x h encore disponibles ».", """
VAR _d = [Heures encore disponibles (h)]
RETURN
    IF ( NOT ( ISBLANK ( _d ) ), FORMAT ( _d, "#,0", "fr-FR" ) & " h encore disponibles" )""")

# ---- Contrôle (réglages Excel, période du Dashboard Excel, aucun autre filtre)
D_CO = "Contrôle"
_CODES = {
    "total": "[Charge identifiée (h)]", "nbAff": "[Nombre d'affaires]", "nbChefs": "[Nombre de chefs de projet]",
    "nbProj": "[Nombre de projets]", "nbTypes": "[Nombre de types]", "realise": "[Réalisé (h)]",
    "reste": "[Reste à faire (h)]", "avgSimple": "[Avancement moyen simple]", "avgPond": "[Avancement moyen pondéré]",
    "periode": "[Charge répartie (h)]", "hors": "[Charge hors période ou sans date de début (h)]",
    "pic": "[Pic de charge (h)]", "picActives": "[Affaires actives au pic]", "picETP": "[ETP au pic]",
    "totalETP": "[ETP-mois]", "nbTermines": "[Lignes terminées]", "picPart": "[Part du pic dans la période]",
    "capEquipe": "[Capacité projets de l'équipe (h/mois)]", "nbSurcharge": "[Mois en surcharge]",
    "picCap": "[Pic en % de la capacité]", "nbRetard": "[Affaires en retard]", "nbDepasse": "[Affaires hors délai]",
    "nbEcart": "[Affaires en écart]", "mois": "CALCULATE ( [Charge répartie (h)], Mois[Mois] = _mois )",
}
_CODES_MOIS = {"autoStart": "MAX ( Réglages[1er mois automatique] )", "start": "[Premier mois affiché]",
               "picMois": "[Mois du pic]", "finPeriode": "[Dernier mois affiché]",
               "premSurcharge": "[Premier mois en surcharge]"}
_CODES_TEXTE = {"periodeTxt": "[Période affichée]", "reglagesTxt": "SELECTEDVALUE ( Réglages[Rappel des réglages (Excel)] )"}
_CTX_EXCEL = """
        REMOVEFILTERS ( Lignes ),
        REMOVEFILTERS ( Affaires ),
        REMOVEFILTERS ( Ressources ),
        REMOVEFILTERS ( 'Charge mensuelle' ),
        REMOVEFILTERS ( Mois ),
        'Périmètre'[Périmètre] = "Comme dans Excel",
        'Répartition'[Répartition] = "Comme dans Excel",
        Mois[Période] = "12 mois du Dashboard Excel\""""


def _switch(codes):
    lines = ",\n".join(f'            "{k}", {v}' for k, v in codes.items())
    return f"SWITCH (\n            _code,\n{lines},\n            BLANK ()\n        )"


mesure("Valeur Power BI", D_CO, "#,0.00", "Valeur numérique recalculée par Power BI avec les réglages Excel sur les "
       "12 mois du Dashboard Excel.", f"""
VAR _code = SELECTEDVALUE ( 'Contrôle Excel'[Code] )
VAR _mois = SELECTEDVALUE ( 'Contrôle Excel'[Mois] )
RETURN
    CALCULATE (
        {_switch(_CODES)},{_CTX_EXCEL}
    )""")
mesure("Mois Power BI", D_CO, F_MOIS, "Mois recalculé par Power BI (indicateurs de type Mois).", f"""
VAR _code = SELECTEDVALUE ( 'Contrôle Excel'[Code] )
RETURN
    CALCULATE (
        {_switch(_CODES_MOIS)},{_CTX_EXCEL}
    )""")
mesure("Texte Power BI", D_CO, None, "Texte recalculé par Power BI (indicateurs de type Texte).", f"""
VAR _code = SELECTEDVALUE ( 'Contrôle Excel'[Code] )
RETURN
    CALCULATE (
        {_switch(_CODES_TEXTE)},{_CTX_EXCEL}
    )""")
mesure("Valeur Excel", D_CO, "#,0.00", "Valeur lue dans l'onglet Calculs.", "SUM ( 'Contrôle Excel'[Valeur lue] )")
mesure("Écart", D_CO, "#,0.00", "Power BI - Excel (indicateurs numériques).", """
VAR _n = SELECTEDVALUE ( 'Contrôle Excel'[Nature] )
RETURN
    IF ( _n IN { "Heures", "Nombre", "Décimal", "Pourcentage" }, [Valeur Power BI] - [Valeur Excel] )""")
mesure("Excel (affiché)", D_CO, None, "Valeur Excel mise en forme.", """
VAR _n = SELECTEDVALUE ( 'Contrôle Excel'[Nature] )
VAR _v = [Valeur Excel]
RETURN
    SWITCH (
        _n,
        "Texte", SELECTEDVALUE ( 'Contrôle Excel'[Valeur lue (texte)] ),
        "Mois", SELECTEDVALUE ( 'Contrôle Excel'[Valeur lue (texte)] ),
        "Pourcentage", IF ( ISBLANK ( _v ), "", FORMAT ( _v, "0.0 %", "fr-FR" ) ),
        "Heures", FORMAT ( _v, "#,0.0", "fr-FR" ),
        "Décimal", FORMAT ( _v, "0.00", "fr-FR" ),
        FORMAT ( _v, "0", "fr-FR" )
    )""")
mesure("Power BI (affiché)", D_CO, None, "Valeur Power BI mise en forme.", """
VAR _n = SELECTEDVALUE ( 'Contrôle Excel'[Nature] )
VAR _v = [Valeur Power BI]
VAR _m = [Mois Power BI]
RETURN
    SWITCH (
        _n,
        "Texte", [Texte Power BI],
        "Mois", IF ( ISBLANK ( _m ), "", FORMAT ( _m, "mmm yyyy", "fr-FR" ) ),
        "Pourcentage", IF ( ISBLANK ( _v ), "", FORMAT ( _v, "0.0 %", "fr-FR" ) ),
        "Heures", FORMAT ( _v, "#,0.0", "fr-FR" ),
        "Décimal", FORMAT ( _v, "0.00", "fr-FR" ),
        FORMAT ( _v, "0", "fr-FR" )
    )""")
mesure("Statut du contrôle", D_CO, None, "✔ si Power BI et Excel concordent (tolérance de la ligne), sinon ✘.", """
VAR _n = SELECTEDVALUE ( 'Contrôle Excel'[Nature] )
VAR _tol = SELECTEDVALUE ( 'Contrôle Excel'[Tolérance] )
VAR _pbi = [Valeur Power BI]
VAR _xl = [Valeur Excel]
VAR _ok =
    SWITCH (
        _n,
        "Texte", [Texte Power BI] = SELECTEDVALUE ( 'Contrôle Excel'[Valeur lue (texte)] ),
        "Mois", [Mois Power BI] = SELECTEDVALUE ( 'Contrôle Excel'[Mois] ),
        IF (
            ISBLANK ( _pbi ) && ISBLANK ( _xl ),
            TRUE (),
            IF ( ISBLANK ( _pbi ) || ISBLANK ( _xl ), FALSE (), ABS ( _pbi - _xl ) <= _tol )
        )
    )
RETURN
    IF ( ISBLANK ( _n ), BLANK (), IF ( _ok, "✔", "✘" ) )""")
mesure("Contrôles en écart", D_CO, F_INT, "Nombre de contrôles ✘.", """
COUNTROWS ( FILTER ( VALUES ( 'Contrôle Excel'[Ordre] ), [Statut du contrôle] = "✘" ) ) + 0""")

# ---- relations (plusieurs -> un)
RELATIONSHIPS = [
    # (de (côté plusieurs), colonne, vers (côté un), colonne, filtre dans les deux sens)
    ("Lignes", "Clé affaire", "Affaires", "Clé affaire", True),
    ("Affaires", "Périmètre", "Périmètre", "Périmètre", False),
    ("Lignes", "Clé ressource", "Ressources", "Clé ressource", False),
    ("Charge mensuelle", "Clé ligne", "Lignes", "Clé ligne", False),
    ("Charge mensuelle", "Répartition", "Répartition", "Répartition", False),
    ("Charge mensuelle", "Mois", "Mois", "Mois", False),
]
QUERY_GROUPS = ["Paramètres", "Fonctions", "Requêtes intermédiaires"]


# =============================================================================================== assemblage
def _col_json(tname, c):
    d = {"name": c.name, "dataType": BIM_TYPE[c.kind], "sourceColumn": c.name, "lineageTag": tag(tname, "col", c.name)}
    if c.fmt:
        d["formatString"] = c.fmt
    if c.hidden:
        d["isHidden"] = True
    d["summarizeBy"] = c.summarize
    if c.desc:
        d["description"] = c.desc
    if c.sort_by:
        d["sortByColumn"] = c.sort_by
    if c.folder:
        d["displayFolder"] = c.folder
    ann = [{"name": "SummarizationSetBy", "value": "User"}]
    if c.kind == "date":
        ann.append({"name": "UnderlyingDateTimeDataType", "value": "Date"})
    d["annotations"] = ann
    return d


def _measure_json(name, folder, fmt, desc, dax):
    d = {"name": name, "expression": m_lines(dax) if "\n" in dax else dax, "lineageTag": tag("Mesures", "mes", name),
         "displayFolder": folder, "description": desc}
    if fmt:
        d["formatString"] = fmt
    return d


def build_model():
    """Contenu complet de model.bim (TMSL)."""
    tables = []
    for t in TABLES:
        tj = {"name": t["name"], "description": t["desc"], "lineageTag": tag(t["name"]),
              "columns": [_col_json(t["name"], c) for c in t["cols"]],
              "partitions": [{"name": t["name"], "mode": "import",
                              "source": {"type": "m", "expression": m_lines(t["m"])}}],
              "annotations": [{"name": "PBI_ResultType", "value": "Table"},
                              {"name": "PBI_NavigationStepName", "value": "Navigation"}]}
        if t["name"] == "Mesures":
            tj["measures"] = [_measure_json(*m) for m in MEASURES]
        tables.append(tj)
    rels = []
    for frm, fcol, to, tcol, both in RELATIONSHIPS:
        r = {"name": tag("rel", frm, fcol, to, tcol), "fromTable": frm, "fromColumn": fcol, "toTable": to,
             "toColumn": tcol}
        if both:
            r["crossFilteringBehavior"] = "bothDirections"
        rels.append(r)
    expressions = []
    for name, group, result, desc, code in EXPRESSIONS:
        expressions.append({"name": name, "kind": "m", "expression": m_lines(code) if "\n" in code else code,
                            "lineageTag": tag("expr", name), "queryGroup": group, "description": desc,
                            "annotations": [{"name": "PBI_NavigationStepName", "value": "Navigation"},
                                            {"name": "PBI_ResultType", "value": result}]})
    query_order = [e[0] for e in EXPRESSIONS] + [t["name"] for t in TABLES]
    model = {
        "culture": CULTURE,
        "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "sourceQueryCulture": CULTURE,
        "tables": tables,
        "relationships": rels,
        "expressions": expressions,
        "queryGroups": [{"folder": g, "annotations": [{"name": "PBI_QueryGroupOrder", "value": str(i)}]}
                        for i, g in enumerate(QUERY_GROUPS)],
        "annotations": [
            {"name": "PBI_QueryOrder", "value": "[" + ",".join('"%s"' % n for n in query_order) + "]"},
            {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
            {"name": "PBI_ProTooling", "value": "[\"DevMode\"]"},
        ],
    }
    return {"compatibilityLevel": COMPATIBILITY_LEVEL, "model": model}


def m_expressions(model=None):
    """Toutes les expressions M du modèle : {nom: code} (partitions des tables et expressions partagées)."""
    model = model or build_model()["model"]
    out = {}
    for t in model["tables"]:
        for p in t["partitions"]:
            e = p["source"]["expression"]
            out[t["name"]] = "\n".join(e) if isinstance(e, list) else e
    for e in model["expressions"]:
        x = e["expression"]
        out[e["name"]] = "\n".join(x) if isinstance(x, list) else x
    return out


def dax_expressions(model=None):
    """Toutes les expressions DAX du modèle : {nom de mesure: code}."""
    model = model or build_model()["model"]
    out = {}
    for t in model["tables"]:
        for m in t.get("measures", []):
            x = m["expression"]
            out[m["name"]] = "\n".join(x) if isinstance(x, list) else x
    return out


if __name__ == "__main__":
    import json
    import sys
    json.dump(build_model(), sys.stdout, ensure_ascii=False, indent=2)

# Modèle sémantique Power BI « Plan de charge BE » — notes

Modèle de données du rapport Power BI qui reproduit les onglets **Dashboard** et **Dessinateurs** du classeur
`Plan_de_charge_BE_Dashboard.xlsm`. Toutes les données viennent de ce classeur, stocké dans Teams (SharePoint) ; Power
BI ne modifie rien. La logique de l'onglet **Calculs** est refaite dans Power Query (M) ; les mesures DAX restent des
agrégations simples de colonnes précalculées.

## 1. Fichiers

| Fichier (dossier `pbi/build`) | Rôle |
|---|---|
| `model_def.py` | `build_model()` renvoie le contenu complet de `model.bim` (TMSL, niveau de compatibilité 1567, culture fr-FR, `sourceQueryCulture` fr-FR, `defaultPowerBIDataSourceVersion` powerBI_V3, annotations de Power BI Desktop) ; `m_expressions()` et `dax_expressions()` extraient les codes. |
| `check_m.js` | Analyse les expressions M avec l'analyseur officiel `@microsoft/powerquery-parser` 4.0.0 (lexeur + analyseur, position des erreurs) et résout chaque identifiant dans sa portée (let, champs d'enregistrement, paramètres, `each`) ou parmi les requêtes. |
| `check_model.py` | Contrôles statiques : M (syntaxe, identifiants, fonctions de bibliothèque connues, une seule source), DAX (`pbi_parsers` 0.9.5, fonctions et nombre d'arguments, références de colonnes, mesures et tables), modèle (noms, relations, chemins de filtre, tris, formats, annotations). Témoins positifs et négatifs des deux analyseurs. |
| `fautes_check_model.py` | Injecte 16 fautes typiques dans une copie de `model_def.py` : chacune doit être signalée par `check_model.py`. |
| `mirror.py` | Miroir Python fidèle, étape par étape, des requêtes M (mêmes noms d'étapes et de colonnes) et des 136 mesures DAX dans un contexte de filtre simplifié. |
| `test_mirror.py` | Compare le miroir aux valeurs calculées par Excel sur 26 classeurs recalculés et 14 variantes de segments. |
| `mutations_mirror.py` | Introduit 14 erreurs plausibles dans le miroir : chacune doit faire échouer `test_mirror.py`. |
| `scen_variantes.py` | Crée et recalcule (LibreOffice) les scénarios supplémentaires : variantes des réglages, 1er mois saisi, arrondis, seuils du retard. |
| `notes_modele.py` | Écrit ce document. |

Commandes (venv `pbi/tools/venv`, Node.js) :

```
<venv>/bin/python -I check_model.py <dossier>          # contrôles statiques, écrit model.bim
<venv>/bin/python -I fautes_check_model.py <dossier_neuf>
python3 test_mirror.py                                  # miroir = Excel
python3 mutations_mirror.py                             # sensibilité du test
```

## 2. Source, paramètre et lecture du classeur

- Paramètre `UrlFichier` (texte, requis) : chemin SharePoint du classeur (Excel bureau > Fichier > Informations >
  Copier le chemin d'accès, sans `?web=1`). Valeur par défaut factice
  `https://VOTRE-ENTREPRISE.sharepoint.com/sites/VOTRE-EQUIPE/Shared%20Documents/General/Plan_de_charge_BE_Dashboard.xlsm`.
- Requête `Classeur` = `Excel.Workbook(Web.Contents(UrlFichier), null, true)` : c'est la **seule** requête qui accède à une
  source (contrôlé) ; toutes les autres en dérivent, donc pas d'erreur Formula.Firewall et une seule authentification
  (« Compte d'organisation ») ; l'actualisation planifiée du service fonctionne sans passerelle.
- Les plages sont lues par les **noms définis** du classeur (`fnLireNom`) : `bi_Charge` (Charge!A4:O503),
  `bi_Data` (Data!A4:K203), `bi_Indicateurs` (Calculs!AK4:AL40), `bi_Mois` (Calculs!AO4:AP15) ; réglages `p_Perimetre`,
  `p_Repartition`, `p_PremierMois`, `p_HeuresETP`, `p_ChargeFixe` (Dashboard!F6, R6, AK6, AU6, BD6). Colonnes par
  position (Column1…), erreurs de cellule remplacées par null, message explicite si un nom manque.
- `Excel.Workbook` lit les **valeurs enregistrées** : le classeur doit avoir été enregistré par Excel après la saisie
  (cas normal). Dates : date, date/heure ou numéro de série acceptés, texte ignoré (null).

## 3. Requêtes Power Query (non chargées sauf les tables du § 4)

| Requête | Groupe | Rôle |
|---|---|---|
| `UrlFichier` | Paramètres | Chemin SharePoint / Teams du classeur Plan_de_charge_BE_Dashboard.xlsm (Excel bureau > Fichier > Informations > Copier le chemin d'accès, sans « ?web=1 »). |
| `Classeur` | Requêtes intermédiaires | Contenu du classeur (valeurs enregistrées par Excel). |
| `fnLireNom` | Fonctions | Plage d'un nom défini du classeur (colonnes Column1, Column2...). |
| `fnValeurNom` | Fonctions | Valeur d'un nom défini d'une seule cellule (null si vide). |
| `fnSérieExcel` | Fonctions | Numéro de série Excel d'une date (null pour une autre valeur). |
| `fnNombreExcel` | Fonctions | Valeur numérique d'une cellule au sens de ESTNUM (nombre ou date), sinon null. |
| `fnNExcel` | Fonctions | Fonction N() d'Excel : nombre, date, VRAI = 1, sinon 0. |
| `fnDateExcel` | Fonctions | Date d'une cellule si ESTNUM (date ou numéro de série), sinon null (texte, vide). |
| `fnTexteExcel` | Fonctions | Texte d'une cellule tel qu'Excel le convertit (nombres au format standard français). |
| `fnSupprEspaces` | Fonctions | SUPPRESPACE d'Excel : espaces retirés en tête et en fin, espaces multiples réduits à un seul. |
| `fnCléAffaire` | Fonctions | Clé d'affaire « projet | sous-projet | produit » comme la colonne EN de Calculs (casse ignorée). |
| `fnArrondiExcel` | Fonctions | ARRONDI d'Excel : demi-unité arrondie en s'éloignant de zéro. |
| `fnJoursOuvrés` | Fonctions | NB.JOURS.OUVRES d'Excel : jours du lundi au vendredi entre deux dates incluses, hors jours non travaillés. |
| `NomsDesMois` | Requêtes intermédiaires | Noms français des mois (abrégés et complets). |
| `HeureParis` | Requêtes intermédiaires | Date et heure de l'actualisation à Paris (le service Power BI tourne en UTC). |
| `DateDuJour` | Requêtes intermédiaires | Aujourd'hui = date de l'actualisation à Paris (AUJOURDHUI() du classeur). |
| `ParamètresExcel` | Requêtes intermédiaires | Réglages du Dashboard Excel (noms p_Perimetre, p_Repartition, p_PremierMois, p_HeuresETP, p_ChargeFixe). |
| `DonnéesData` | Requêtes intermédiaires | Onglet Data, lignes 4 à 203 (nom bi_Data), colonnes A à K par position. |
| `JoursNonTravaillés` | Requêtes intermédiaires | Jours non travaillés de Data (colonne K) tombant un jour de semaine, sans doublon. |
| `DonnéesCharge` | Requêtes intermédiaires | Lignes valides de l'onglet Charge (nom bi_Charge, colonnes A à O par position) et leurs valeurs nettoyées. |
| `RépartitionLignes` | Requêtes intermédiaires | Heures de chaque ligne réparties par mois en jours ouvrés, pour les deux modes de répartition (règle 6). |
| `LignesPérimètres` | Requêtes intermédiaires | Lignes de chaque variante du segment Périmètre, avec chef, projet et type regroupés (casse ignorée), affaire, rangs et groupe de couleur du projet. |
| `FenêtreExcel` | Requêtes intermédiaires | Les 12 mois du Dashboard Excel (règle 7) calculés avec les réglages du classeur. |

Correspondance avec les règles de la spécification :

| Règle | Où |
|---|---|
| 1. ligne valide, avancement borné | `DonnéesCharge` (étapes `Valides`, `AjoutAvancement`) |
| 2. périmètre | `LignesPérimètres` (`LignesOption`) — 3 variantes : Comme dans Excel, Projets en cours, Tous les projets |
| 3. heures, chef, projet, type, casse ignorée | `DonnéesCharge` (`AjoutHeures`, `AjoutChef`…) ; regroupements insensibles à la casse dans `LignesPérimètres` (`Grouper`) |
| 4. ressource de la ligne, lignes K+L | `DonnéesCharge` (`AjoutRessource`, `AjoutTypeRessource`, `AjoutDouble`) |
| 5. affaire et libellés | `fnCléAffaire` (SUPPRESPACE, CAR(160) remplacé, jokers ~ * ? neutralisés, 240 caractères, casse ignorée) ; `LignesPérimètres` (1re ligne) ; `Affaires` (`Première`) |
| 6. répartition mensuelle | `RépartitionLignes` (deux modes, tous les mois) avec `fnJoursOuvrés` (= NB.JOURS.OUVRES) |
| 7. fenêtre de 12 mois | `FenêtreExcel` puis `Mois[Période]` |
| 8. capacité | `Ressources` (dessinateurs puis externes, doublons comptés une fois, valeurs propres ou par défaut) ; `Réglages` (somme) |
| 9. surcharge, pic | mesures `[Mois en surcharge]`, `[Pic de charge (h)]`, `[ETP au pic]`, `[Pic en % de la capacité]` |
| 10. occupation | mesures `[Occupation (%)]` et couleurs |
| 11. affaires | `Affaires` (charge, %, ressources, début, besoin, avancement toutes lignes, réalisé, reste, attendu, retard) |
| 12. indicateurs | mesures des dossiers Charge, Affaires, Mois et capacité |

Pièges de M traités : `Number.Round` arrondit au pair par défaut → `fnArrondiExcel` (RoundingMode.AwayFromZero après
pré-arrondi à 6 décimales, 2,675 → 2,68 comme Excel) ; lundi-vendredi = `Date.DayOfWeek(d, Day.Monday) < 5` ;
`Text.Trim` ne réduit pas les espaces internes → `fnSupprEspaces` (SUPPRESPACE garde l'espace insécable, comme Excel ;
il n'est remplacé que dans la clé d'affaire, comme SUBSTITUE(…;CAR(160);" ")) ; comparaisons insensibles à la casse
par `Text.Lower(…, "fr-FR")` ; null géré à chaque comparaison ; `Table.Buffer` / `List.Buffer` sur les données
réutilisées ; colonnes typées à la dernière étape (`Typée`) de chaque table chargée.

Date du jour : `HeureParis` convertit `DateTimeZone.UtcNow()` (le service tourne en UTC) en heure de Paris (heure
d'été du dernier dimanche de mars 01:00 UTC au dernier dimanche d'octobre 01:00 UTC ; contrôlé contre le fuseau
Europe/Paris de 2024 à 2031) ; `DateDuJour` = sa date = AUJOURDHUI() du classeur.

## 4. Tables chargées

Les tables `Lignes` et `Affaires` contiennent une copie des données **par variante du segment Périmètre** (colonne
masquée `Périmètre`) et `Charge mensuelle` une copie par variante de périmètre **et** de répartition. Les clés
numériques intègrent la variante (`Clé ligne` = ordre de la variante × 10000 + n° de ligne ; `Clé affaire` = ordre ×
10000 + n° de la 1re ligne de l'affaire). Volume négligeable (≤ 500 lignes × 3 variantes ; ≤ 500 × 3 × 3 × mois).

### Périmètre

Segment Périmètre (sélection unique) : « Comme dans Excel » suit le réglage p_Perimetre du classeur.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Périmètre | texte |  | oui | Variante de périmètre proposée par le segment. Triée par Ordre. |
| Ordre | entier |  | non |  |
| Périmètre appliqué | texte |  | oui | Périmètre réellement appliqué (Projets en cours ou Tous les projets). |
| Code périmètre | texte |  | non |  |
| Description | texte |  | oui |  |

### Répartition

Segment Répartition (sélection unique) : « Comme dans Excel » suit le réglage p_Repartition.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Répartition | texte |  | oui | Variante de répartition mensuelle proposée par le segment. Triée par Ordre. |
| Ordre | entier |  | non |  |
| Répartition appliquée | texte |  | oui | Mode réellement appliqué. |
| Code répartition | texte |  | non |  |
| Description | texte |  | oui |  |

### Ressources

Dessinateurs (Data E) puis ressources externes (Data H), ressources saisies dans Charge hors listes, et « (non affecté) ».

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Clé ressource | entier |  | non |  |
| Ressource | texte |  | oui | Triée par Ordre. |
| Ordre | entier |  | non |  |
| Type de ressource | texte |  | oui | Interne (Data E), Externe (Data H), Hors liste (absente des listes Data) ou Non affecté. |
| Mention | texte |  | oui | « ext. » pour une ressource externe, « hors liste » pour une ressource absente de Data. |
| Ressource (avec mention) | texte |  | oui | Nom suivi de la mention (« · ext. », « · hors liste ») : lignes du tableau d'occupation, comme la colonne « ext. » de l'onglet Dessinateurs. Triée par Ordre. |
| Position Excel | entier | 0 | oui | Rang dans la liste commune de Calculs (FA) ; les 16 premiers forment le tableau d'occupation de l'onglet Dessinateurs. |
| Dans le tableau Excel | vrai/faux |  | oui | Ressource affichée dans le tableau d'occupation de l'onglet Dessinateurs. |
| Dans l'équipe | vrai/faux |  | oui | Ressource des listes Data, comptée dans la capacité projets de l'équipe. |
| Capacité (h/mois) | décimal | #,0 | oui |  |
| Charge fixe (h/mois) | décimal | #,0 | oui |  |
| Capacité projets (h/mois) | décimal | #,0 | oui | max(0, capacité - charge fixe). |
| Clé ressource (texte) | texte |  | non |  |

### Lignes

Une ligne par variante de périmètre et par ligne valide de l'onglet Charge.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Clé ligne | entier |  | non | Ordre de la variante × 10000 + n° de ligne. |
| Clé affaire | entier |  | non |  |
| Clé ressource | entier |  | non |  |
| Périmètre | texte |  | non |  |
| N° ligne | entier | 0 | oui | Numéro de ligne dans l'onglet Charge. |
| Chef de projet | texte |  | oui | Chef de la ligne (« (non renseigné) » si vide), libellé de sa 1re ligne du périmètre (casse ignorée). |
| Projet | texte |  | oui |  |
| Sous-projet | texte |  | oui |  |
| Produit | texte |  | oui |  |
| Type | texte |  | oui |  |
| Description | texte |  | oui | Dossier Détail. |
| Contraintes | texte |  | oui | Dossier Détail. |
| Autres informations | texte |  | oui | Dossier Détail. |
| Coordinateur | texte |  | oui | Dossier Détail. |
| Dessinateur | texte |  | oui | Colonne K de Charge. Dossier Détail. |
| Ressource externe | texte |  | oui | Colonne L de Charge. Dossier Détail. |
| Ressource de la ligne | texte |  | oui | Dessinateur, sinon ressource externe, sinon « (non affecté) ». |
| Type de ressource | texte |  | oui | Interne (K), Externe (L) ou Non affecté. |
| Dessinateur et externe | vrai/faux |  | oui | Ligne à corriger : dessinateur ET ressource externe (seul K compte). |
| Heures (h) | décimal | #,0 | oui |  |
| Avancement | décimal | 0 % | oui |  |
| Réalisé de la ligne (h) | décimal | #,0 | oui | Heures × avancement. |
| Reste de la ligne (h) | décimal | #,0 | oui | Heures - réalisé. |
| Début planifié | date | dd/mm/yyyy | oui |  |
| Début (affiché) | texte |  | oui | Date de début jj/mm/aaaa, ou « — » sans date de début (comme l'onglet Dessinateurs) ; triée par date. Triée par Début planifié. |
| Date de besoin | date | dd/mm/yyyy | oui |  |
| Terminée | vrai/faux |  | oui | Avancement de 100 %. |
| Sans date de début | vrai/faux |  | oui |  |
| Rang chef | entier | 0 | oui | Dossier Rangs. |
| Rang projet | entier | 0 | oui | Dossier Rangs. |
| Rang type | entier | 0 | oui | Dossier Rangs. |
| Rang ligne | entier | 0 | oui | Rang de la ligne par heures décroissantes (ordre de l'onglet Charge en cas d'égalité). Dossier Rangs. |
| Groupe projet | texte |  | oui | Projet, ou « Autres projets (n) » au-delà du 8e projet quand il y en a plus de 9 (couleurs du Dashboard Excel). |
| Ordre groupe projet | entier | 0 | oui | Dossier Rangs. |
| Couleur projet | texte |  | oui | Couleur du groupe projet (palette du Dashboard Excel). |
| Clé affaire (texte) | texte |  | non |  |

### Affaires

Une ligne par variante de périmètre et par affaire (projet | sous-projet | produit) ayant au moins une ligne dans le périmètre.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Clé affaire | entier |  | non |  |
| Périmètre | texte |  | non |  |
| N° | entier | 0 | oui | Rang dans le tableau des affaires du Dashboard (charge décroissante, puis ordre de l'onglet Charge). |
| Affaire | texte |  | oui | Projet · sous-projet · produit (libellés de la 1re ligne du périmètre, « — » si vide) : champ d'extraction vers la fiche affaire. |
| Projet | texte |  | oui |  |
| Sous-projet | texte |  | oui |  |
| Produit | texte |  | oui |  |
| Chef de projet | texte |  | oui | Chef de la 1re ligne de l'affaire. |
| Ressources | texte |  | oui | « N dont X ext. », « X ext. », « à affecter », « + à affecter ». |
| Charge (h) | décimal | #,0 | oui | Heures des lignes du périmètre (arrondies à 0,01 h). |
| % de la charge totale | décimal | 0.0 % | oui | Charge de l'affaire / charge identifiée du périmètre. |
| Début planifié | date | dd/mm/yyyy | oui | Plus petit début des lignes du périmètre. |
| Début (affiché) | texte |  | oui | Date de début, ou « à planifier » ; triée par date. Triée par Début planifié. |
| Date de besoin | date | dd/mm/yyyy | oui | Plus grande date de besoin des lignes du périmètre. |
| Avancement | décimal | 0 % | oui | Moyenne pondérée par les heures de TOUTES les lignes de l'affaire (moyenne simple sans heures). |
| Heures toutes lignes (h) | décimal | #,0 | oui |  |
| Réalisé toutes lignes (h) | décimal | #,0 | oui |  |
| Reste à faire de l'affaire (h) | décimal | #,0 | oui | max(0, heures - réalisé) de toutes les lignes de l'affaire (arrondi à 0,01 h). |
| Avancement attendu | décimal | 0 % | oui | Prorata des jours ouvrés écoulés entre début et besoin (vide si non calculable). |
| Retard | entier | 0 | oui | 2 = date de besoin dépassée sans être à 100 %, 1 = avancement inférieur de plus de 15 points à l'attendu, 0 = dans les temps. |
| Statut | texte |  | oui |  |
| Couleur d'alerte | texte |  | oui | Rouge B42318 si retard, sinon encre du Dashboard (1F2937). |
| Lignes de l'affaire | entier | 0 | oui | Lignes de l'affaire dans le périmètre. |
| Personnes internes | entier | 0 | oui |  |
| Personnes externes | entier | 0 | oui |  |
| À affecter | vrai/faux |  | oui | Au moins une ligne sans ressource. |
| N° 1re ligne | entier | 0 | oui | Ligne de Charge de la 1re ligne de l'affaire dans le périmètre. |
| Clé affaire (texte) | texte |  | non |  |

### Charge mensuelle

Heures réparties par variante de périmètre et de répartition, par ligne et par mois.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Clé ligne | entier |  | non |  |
| Clé affaire | entier |  | non |  |
| Répartition | texte |  | non |  |
| Mois | date | dd/mm/yyyy | non |  |
| Heures (h) | décimal | #,0.0 | oui |  |

### Mois

Calendrier mensuel couvrant la fenêtre du Dashboard Excel et tous les mois de charge répartie.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Mois | date | mmmm yyyy | oui | 1er jour du mois. |
| Libellé mois | texte |  | oui | Ex. « sept. 2026 ». Triée par Clé de tri. |
| Mois (long) | texte |  | oui | Ex. « septembre 2026 ». Triée par Clé de tri. |
| Clé de tri | entier |  | non |  |
| Année | entier | 0 | oui |  |
| Fin de mois | date | dd/mm/yyyy | non |  |
| N° mois Excel | entier | 0 | oui | Rang du mois dans les 12 mois du Dashboard Excel (vide hors fenêtre). |
| Période | texte |  | oui | « 12 mois du Dashboard Excel », « Avant » ou « Après ». Triée par Ordre période. |
| Ordre période | entier |  | non |  |

### Réglages

Réglages lus dans le classeur (une ligne) et valeurs dérivées. Suffixe (Excel) : valeur du classeur, distincte des mesures du même nom qui suivent les segments.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Périmètre Excel | texte |  | oui | Valeur de p_Perimetre. |
| Répartition Excel | texte |  | oui | Valeur de p_Repartition. |
| Périmètre appliqué (Excel) | texte |  | oui |  |
| Répartition appliquée (Excel) | texte |  | oui |  |
| 1er mois saisi | date | mmmm yyyy | oui | p_PremierMois ramené au 1er du mois (vide = automatique). |
| 1er mois automatique | date | mmmm yyyy | oui | Mois du plus petit début planifié du périmètre Excel. |
| 1er mois affiché (Excel) | date | mmmm yyyy | oui | 1er des 12 mois du Dashboard Excel. |
| Dernier mois affiché (Excel) | date | mmmm yyyy | oui | 12e mois du Dashboard Excel. |
| Période affichée (Excel) | texte |  | oui | Bandeau des onglets Dashboard et Dessinateurs. |
| Capacité par défaut (h/mois) | décimal | #,0 | oui | p_HeuresETP. |
| Charge fixe par défaut (h/mois) | décimal | #,0 | oui | p_ChargeFixe. |
| Capacité projets par personne (Excel) | décimal | #,0 | oui | Capacité par défaut - charge fixe par défaut (diviseur des ETP, h/mois). |
| Capacité projets de l'équipe (Excel) | décimal | #,0 | oui | Somme des capacités projets des ressources des listes Data (h/mois). |
| Rappel capacité | texte |  | oui | « 120 h/pers. dont 20 h hors projet ». |
| Rappel des réglages (Excel) | texte |  | oui | Rappel des réglages du Dashboard Excel (même texte que l'Excel). |
| Lignes valides (Excel) | entier | 0 | oui | Lignes de Charge dont le nom de projet est renseigné. |
| Lignes terminées (Excel) | entier | 0 | oui | Lignes valides à 100 % (tous périmètres). |
| Lignes à corriger (Excel) | entier | 0 | oui | Lignes avec un dessinateur ET une ressource externe. |

### Actualisation

Date et heure de la dernière actualisation (heure de Paris).

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Actualisé le | date/heure | dd/mm/yyyy hh:nn | oui |  |
| Aujourd'hui | date | dd/mm/yyyy | oui | Date utilisée pour l'avancement attendu et les retards. |
| Libellé | texte |  | oui |  |

### Contrôle Excel

Indicateurs et charge mensuelle calculés par Excel (onglet Calculs), comparés aux mesures Power BI avec les réglages Excel.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Ordre | entier | 0 | non |  |
| Catégorie | texte |  | oui |  |
| Contrôle | texte |  | oui | Triée par Ordre. |
| Code | texte |  | non |  |
| Nature | texte |  | oui | Heures, Nombre, Décimal, Pourcentage, Mois ou Texte. |
| Cellule Excel | texte |  | oui |  |
| Mois | date | mmmm yyyy | oui | Mois de la ligne (charge par mois) ou valeur Excel d'un indicateur de type Mois. |
| Valeur lue | décimal | #,0.00 | oui | Valeur numérique lue dans l'onglet Calculs. |
| Valeur lue (texte) | texte |  | oui | Valeur lue (indicateurs de type Texte ou Mois). |
| Tolérance | décimal | 0.0000 | oui |  |

### Mesures

Table des mesures du rapport.

| Colonne | Type | Format | Visible | Description |
|---|---|---|---|---|
| Colonne technique | texte |  | non |  |

## 5. Relations

| De (plusieurs) | Vers (un) | Filtre |
|---|---|---|
| Lignes[Clé affaire] | Affaires[Clé affaire] | deux sens |
| Affaires[Périmètre] | Périmètre[Périmètre] | un sens |
| Lignes[Clé ressource] | Ressources[Clé ressource] | un sens |
| Charge mensuelle[Clé ligne] | Lignes[Clé ligne] | un sens |
| Charge mensuelle[Répartition] | Répartition[Répartition] | un sens |
| Charge mensuelle[Mois] | Mois[Mois] | un sens |

```
Périmètre 1 ──► * Affaires 1 ◄──► * Lignes * ◄── 1 Ressources
                                     1
                                     │
                                     ▼ *
     Répartition 1 ──► * Charge mensuelle * ◄── 1 Mois
```

Tables sans relation : `Réglages` (une ligne), `Actualisation` (une ligne), `Contrôle Excel`, `Mesures` (table des
mesures).

**Filtre dans les deux sens entre `Lignes` et `Affaires`** (seule exception, nécessaire) : les segments Chef, Projet,
Type (colonnes de `Lignes`) et Ressource (`Ressources`) doivent restreindre les affaires (nombre d'affaires, retards,
avancement moyen, tableau des affaires, « ses affaires » sur la page Ressources). Le contrôle des chemins confirme
qu'il n'existe qu'un seul chemin de filtre entre deux tables quelconques (pas d'ambiguïté) ; `Ressources` ne filtre que
`Lignes` (sens unique), et `Mois` / `Répartition` ne filtrent que `Charge mensuelle`.

## 6. Variantes, segments et contexte

- **Périmètre** et **Répartition** : tables de 3 lignes triées par `Ordre`. Toutes les mesures lisent
  `SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )` (idem Répartition) : aucune sélection ou plusieurs
  sélections → « Comme dans Excel ». Les segments doivent être en **sélection unique**, valeur initiale « Comme dans
  Excel », synchronisés entre les pages : un visuel qui affiche directement des lignes de `Affaires` ou `Lignes` (tableau
  des affaires, lignes d'une ressource) montrerait sinon les 3 variantes.
- **Période** : `Mois[Période]` (« 12 mois du Dashboard Excel », « Avant », « Après », triée par `Ordre période`),
  multi-sélection, valeur initiale « 12 mois du Dashboard Excel » ; ou plage de `Mois[Mois]`. La table `Mois` couvre la
  fenêtre Excel et tous les mois où des heures sont réparties (quelle que soit la variante).
- **Chef / Projet / Type** : utiliser les colonnes de `Lignes` (`Lignes[Chef de projet]`, `Lignes[Projet]`,
  `Lignes[Type]`), comme les listes et les comptes de l'Excel (chef, projet et type de chaque ligne). Les libellés
  sont ceux de la 1re ligne de chaque groupe insensible à la casse.
- **Ressource** : `Ressources[Ressource]` (triée par `Ressources[Ordre]` = ordre des listes Data), `Ressources[Type de
  ressource]`. La table contient aussi les ressources saisies dans Charge mais absentes de Data (« Hors liste »,
  réglages par défaut) et « (non affecté) ».
- **Aujourd'hui** = `Actualisation[Aujourd'hui]` (date de l'actualisation à Paris), utilisé pour l'avancement attendu et
  les retards ; affiché par `[Données lues le]`.

## 7. Mesures

136 mesures dans la table `Mesures`, rangées en dossiers d'affichage. Formats : heures `#,0` (« h » dans les
titres), % `0 %` / `0.0 %` (affichés « 45 % » / « 45,3 % » en culture fr-FR), dates `dd/mm/yyyy`, mois `mmmm yyyy`.
Le code DAX complet est en annexe.

### Dossier « Charge »

| Mesure | Format | Signification |
|---|---|---|
| Charge identifiée (h) | #,0 | Somme des heures des lignes du périmètre (onglet Charge, colonne M). |
| Réalisé (h) | #,0 | Heures × avancement des lignes du périmètre. |
| Reste à faire (h) | #,0 | Charge identifiée - réalisé. |
| Capacité projets par personne (h/mois) | #,0 | Capacité par défaut - charge fixe par défaut (réglages du Dashboard Excel). |
| ETP-mois | 0.0 | Charge identifiée / (capacité - charge fixe) d'une personne (0 si diviseur <= 0). |
| Nombre de chefs de projet | 0 | Chefs distincts des lignes du périmètre. |
| Nombre de projets | 0 | Projets distincts (casse ignorée). |
| Nombre de types | 0 | Types d'affaire distincts. |
| Nombre de lignes | 0 | Lignes de l'onglet Charge dans le périmètre. |
| Part de la charge (%) | 0.0 % | Part de la charge identifiée dans le total de la sélection (segments et filtres de la page). |

### Dossier « Affaires »

| Mesure | Format | Signification |
|---|---|---|
| Nombre d'affaires | 0 | Affaires (sous-projets regroupés) ayant au moins une ligne dans le périmètre. |
| Avancement moyen simple | 0 % | Moyenne simple des avancements des affaires. |
| Avancement moyen pondéré | 0 % | Réalisé / heures de toutes les lignes des affaires du périmètre (repli : moyenne simple). |
| Affaires en retard | 0 | Affaires hors délai ou en écart de plus de 15 points. |
| Affaires hors délai | 0 | Date de besoin dépassée sans être à 100 %. |
| Affaires en écart | 0 | Avancement inférieur de plus de 15 points à l'avancement attendu. |
| Plus gros reste à faire (h) | #,0 | Reste à faire de l'affaire qui en a le plus. |
| Plus gros reste à faire – affaire | texte | Affaire au plus gros reste à faire : affaire (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – projet | texte | Affaire au plus gros reste à faire : projet (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – sous-projet | texte | Affaire au plus gros reste à faire : sous-projet (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – produit | texte | Affaire au plus gros reste à faire : produit (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – chef de projet | texte | Affaire au plus gros reste à faire : chef de projet (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – ressources | texte | Affaire au plus gros reste à faire : ressources (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – total de l'affaire (h) | #,0 | Affaire au plus gros reste à faire : total de l'affaire (h) (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – début planifié | dd/mm/yyyy | Affaire au plus gros reste à faire : début planifié (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – date de besoin | dd/mm/yyyy | Affaire au plus gros reste à faire : date de besoin (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – avancement | 0 % | Affaire au plus gros reste à faire : avancement (égalité : 1re dans l'ordre de l'onglet Charge). |
| Plus gros reste à faire – retard | 0 | Affaire au plus gros reste à faire : retard (égalité : 1re dans l'ordre de l'onglet Charge). |
| Couleur avancement | texte | Couleur de l'avancement d'une affaire : rouge si retard (mise en forme conditionnelle, valeur de champ). |
| Lignes terminées | 0 | Lignes valides à 100 % de toutes les lignes, quel que soit le périmètre (comme l'Excel) ; suit les segments et clics sur les colonnes de Lignes (chef, projet, type), pas un clic sur une affaire. |
| Lignes avec dessinateur et externe | 0 | Lignes ayant un dessinateur ET une ressource externe (à corriger dans Charge), parmi toutes les lignes valides ; suit les segments chef, projet, type. |

### Dossier « Mois et capacité »

| Mesure | Format | Signification |
|---|---|---|
| Charge répartie (h) | #,0 | Heures réparties sur les mois du contexte (période, mois du graphique). |
| Charge répartie sans les mois vides (h) | #,0 | Comme [Charge répartie (h)], mais vide (et non 0) sans heure répartie : un graphique n'affiche alors que les mois et ressources qui ont des heures. |
| Heures de la ligne sur la période (h) | #,0 | Charge répartie des lignes du contexte, 0 pour une ligne sans heure sur la période (sans date de début, 0 h, hors des mois choisis) ; vide pour une combinaison ligne × affaire qui n'existe pas. |
| Affaires actives | 0 | Affaires ayant au moins une heure répartie sur les mois du contexte. |
| Charge hors période ou sans date de début (h) | #,0 | Charge identifiée non répartie sur les mois sélectionnés (arrondie à 0,1 h). |
| Capacité projets de l'équipe (h/mois) | #,0 | Somme des (capacité - charge fixe) des ressources des listes Data : constante par mois (ligne pointillée). |
| Mois en surcharge | 0 | Mois où la charge arrondie dépasse la capacité projets arrondie. |
| Premier mois en surcharge | mmmm yyyy | Premier mois en surcharge (vide si aucun). |
| Pic de charge (h) | #,0 | Charge du mois le plus chargé. |
| Mois du pic | mmmm yyyy | Mois le plus chargé (le premier en cas d'égalité). |
| Affaires actives au pic | 0 | Affaires actives sur le mois le plus chargé. |
| ETP au pic | 0.0 | Pic / (capacité - charge fixe) d'une personne. |
| Pic en % de la capacité | 0 % | Pic / capacité projets de l'équipe (vide si capacité nulle). |
| Part du pic dans la période | 0 % | Charge du mois de pic / charge répartie sur la période. |
| Premier mois affiché | mmmm yyyy | Premier mois de la période sélectionnée. |
| Dernier mois affiché | mmmm yyyy | Dernier mois de la période sélectionnée. |

### Dossier « Groupes de projets »

| Mesure | Format | Signification |
|---|---|---|
| Groupe 1 – libellé | texte | Nom du groupe 1 (projet de rang 1 par charge) ; couleur #2A78D6. |
| Groupe 1 – charge (h) | #,0 | Charge identifiée du groupe 1 (projet de rang 1 par charge) ; couleur #2A78D6 (anneau). |
| Groupe 1 – charge répartie (h) | #,0 | Charge répartie du groupe 1 (projet de rang 1 par charge) sur les mois du contexte ; couleur #2A78D6 (histogramme empilé). |
| Groupe 2 – libellé | texte | Nom du groupe 2 (projet de rang 2 par charge) ; couleur #EB6834. |
| Groupe 2 – charge (h) | #,0 | Charge identifiée du groupe 2 (projet de rang 2 par charge) ; couleur #EB6834 (anneau). |
| Groupe 2 – charge répartie (h) | #,0 | Charge répartie du groupe 2 (projet de rang 2 par charge) sur les mois du contexte ; couleur #EB6834 (histogramme empilé). |
| Groupe 3 – libellé | texte | Nom du groupe 3 (projet de rang 3 par charge) ; couleur #1BAF7A. |
| Groupe 3 – charge (h) | #,0 | Charge identifiée du groupe 3 (projet de rang 3 par charge) ; couleur #1BAF7A (anneau). |
| Groupe 3 – charge répartie (h) | #,0 | Charge répartie du groupe 3 (projet de rang 3 par charge) sur les mois du contexte ; couleur #1BAF7A (histogramme empilé). |
| Groupe 4 – libellé | texte | Nom du groupe 4 (projet de rang 4 par charge) ; couleur #EDA100. |
| Groupe 4 – charge (h) | #,0 | Charge identifiée du groupe 4 (projet de rang 4 par charge) ; couleur #EDA100 (anneau). |
| Groupe 4 – charge répartie (h) | #,0 | Charge répartie du groupe 4 (projet de rang 4 par charge) sur les mois du contexte ; couleur #EDA100 (histogramme empilé). |
| Groupe 5 – libellé | texte | Nom du groupe 5 (projet de rang 5 par charge) ; couleur #E87BA4. |
| Groupe 5 – charge (h) | #,0 | Charge identifiée du groupe 5 (projet de rang 5 par charge) ; couleur #E87BA4 (anneau). |
| Groupe 5 – charge répartie (h) | #,0 | Charge répartie du groupe 5 (projet de rang 5 par charge) sur les mois du contexte ; couleur #E87BA4 (histogramme empilé). |
| Groupe 6 – libellé | texte | Nom du groupe 6 (projet de rang 6 par charge) ; couleur #008300. |
| Groupe 6 – charge (h) | #,0 | Charge identifiée du groupe 6 (projet de rang 6 par charge) ; couleur #008300 (anneau). |
| Groupe 6 – charge répartie (h) | #,0 | Charge répartie du groupe 6 (projet de rang 6 par charge) sur les mois du contexte ; couleur #008300 (histogramme empilé). |
| Groupe 7 – libellé | texte | Nom du groupe 7 (projet de rang 7 par charge) ; couleur #6250D6. |
| Groupe 7 – charge (h) | #,0 | Charge identifiée du groupe 7 (projet de rang 7 par charge) ; couleur #6250D6 (anneau). |
| Groupe 7 – charge répartie (h) | #,0 | Charge répartie du groupe 7 (projet de rang 7 par charge) sur les mois du contexte ; couleur #6250D6 (histogramme empilé). |
| Groupe 8 – libellé | texte | Nom du groupe 8 (projet de rang 8 par charge) ; couleur #E34948. |
| Groupe 8 – charge (h) | #,0 | Charge identifiée du groupe 8 (projet de rang 8 par charge) ; couleur #E34948 (anneau). |
| Groupe 8 – charge répartie (h) | #,0 | Charge répartie du groupe 8 (projet de rang 8 par charge) sur les mois du contexte ; couleur #E34948 (histogramme empilé). |
| Groupe 9 – libellé | texte | Nom du groupe 9 (9e projet, ou « Autres projets (n) » au-delà de 9) ; couleur #9CA3AF. |
| Groupe 9 – charge (h) | #,0 | Charge identifiée du groupe 9 (9e projet, ou « Autres projets (n) » au-delà de 9) ; couleur #9CA3AF (anneau). |
| Groupe 9 – charge répartie (h) | #,0 | Charge répartie du groupe 9 (9e projet, ou « Autres projets (n) » au-delà de 9) sur les mois du contexte ; couleur #9CA3AF (histogramme empilé). |

### Dossier « Ressource »

| Mesure | Format | Signification |
|---|---|---|
| Capacité de la ressource (h/mois) | #,0 | Capacité mensuelle d'une ressource (vide pour « (non affecté) » ou plusieurs ressources). |
| Charge fixe de la ressource (h/mois) | #,0 | Charge fixe hors projet d'une ressource. |
| Capacité projets de la ressource (h/mois) | #,0 | max(0, capacité - charge fixe). |
| Charge fixe (h) | #,0 | Charge fixe de la ressource sur les mois du contexte. |
| Charge totale de la ressource (h) | #,0 | Charge fixe + heures projets de la ressource. |
| Projets dans la capacité (h) | #,0 | Heures projets tenant dans la capacité (bleu du graphique Excel). |
| Dépassement de capacité (h) | #,0 | Heures projets au-delà de la capacité (rouge du graphique Excel). |
| Marge de la ressource (h) | +#,0;-#,0;0 | Capacité - charge totale arrondie (négatif = dépassement). |
| Dépassement cumulé (h) | #,0 | Somme des dépassements mensuels. |
| Heures encore disponibles (h) | #,0 | Somme des marges mensuelles positives. |
| Occupation (%) | 0 % | (Charge fixe + heures projets, arrondi à l'heure) / capacité ; moyenne des mois si plusieurs mois ; 999 % si capacité nulle avec de la charge. |
| Occupation maximale (%) | 0 % | Occupation du mois le plus occupé. |
| Couleur de fond (occupation) | texte | Vert < 90 %, orange 90-100 %, rouge > 100 % (fonds de l'Excel). |
| Couleur du texte (occupation) | texte | Texte rouge, orange foncé ou vert (couleurs de l'Excel). |
| Mois en surcharge de la ressource | 0 | Mois où la charge totale arrondie dépasse la capacité. |
| Premier mois en surcharge de la ressource | mmmm yyyy | Premier mois en surcharge (vide si aucun). |
| Charge au pic de la ressource (h) | #,0 | Charge totale (fixe + projets) du mois le plus chargé. |
| Mois le plus chargé de la ressource | mmmm yyyy | Mois de charge totale maximale (le premier en cas d'égalité). |
| Occupation au pic de la ressource (%) | 0 % | Occupation du mois le plus chargé. |

### Dossier « Textes »

| Mesure | Format | Signification |
|---|---|---|
| Périmètre en vigueur | texte | Projets en cours ou Tous les projets. |
| Répartition en vigueur | texte | Lissée jusqu'à la date de besoin ou Tout sur le mois de début. |
| Rappel des réglages | texte | Réglages appliqués par les segments (format du bandeau Excel). |
| Données lues le | texte | Date et heure de l'actualisation (heure de Paris). |
| Titre du périmètre | texte | PROJETS EN COURS ou TOUS LES PROJETS. |
| Sous-titre chefs | texte | Sous-titre de la carte des chefs de projet. |
| Sous-titre charge identifiée | texte | « soit ≈ x ETP-mois ». |
| Sous-titre capacité | texte | « Capacité projets : x h/mois ». |
| Sous-titre surcharge | texte | « Dès <mois> » ou « Pic à x % de la capacité ». |
| Sous-titre pic (mois) | texte | « en <mois du pic> ». |
| Sous-titre pic (affaires) | texte | « n affaire(s) active(s) · ≈ x ETP ». |
| Sous-titre hors délai | texte | « dont n hors délai (besoin dépassé) ». |
| Sous-titre écart | texte | « n en retard de plus de 15 pts ». |
| Sous-titre projets | texte | « réparties sur n projet(s) ». |
| Sous-titre lignes terminées | texte | Lignes terminées comptées ou non selon le périmètre. |
| Sous-titre avancement | texte | « Moyenne simple par affaire : x % ». |
| Période affichée | texte | « Période <premier mois> – <dernier mois> » de la sélection. |
| Texte hors période | texte | Charge hors période ou sans date de début (sous-titre du graphique). |
| Titre de la ressource | texte | Titre du graphique mensuel de la ressource sélectionnée. |
| Sous-titre graphique mensuel | texte | « n mois à partir de <premier mois> » (période sélectionnée). |
| Sous-titre plus gros reste | texte | « sur x h au total pour cette affaire ». |
| Origine de la ressource | texte | Liste Data d'origine de la ressource sélectionnée. |
| Sous-titre charge fixe de la ressource | texte | « dont x h de charge fixe ». |
| Sous-titre capacité de la ressource | texte | « soit x h/mois pour les projets » ou alerte si la charge fixe atteint la capacité. |
| Sous-titre affaires de la ressource | texte | « sur n affaire(s) ». |
| Sous-titre charge sur la période | texte | « dont x h sur la période » (mois sélectionnés). |
| Sous-titre lignes de la ressource | texte | Nombre de lignes de la ressource dans le périmètre, ou « Aucune affaire affectée à cette ressource dans le périmètre. » (comme l'onglet Dessinateurs). |
| Titre info-bulle projet | texte | Titre de l'info-bulle projet : mois survolé, ou toute la période affichée (anneau, tableau des projets). |
| Mois le plus chargé de la ressource (texte) | texte | Mois le plus chargé, « — » sans charge projets sur la période. |
| Sous-titre pic de la ressource | texte | « x h au total (fixe + projets) ». |
| Sous-titre occupation au pic | texte | « x % de la capacité ». |
| Sous-titre dépassement cumulé | texte | « Dépassement cumulé : x h ». |
| Sous-titre surcharge de la ressource | texte | « Aucun dépassement prévu » ou « Dès <mois> ». |
| Sous-titre heures disponibles | texte | « x h encore disponibles ». |

### Dossier « Contrôle »

| Mesure | Format | Signification |
|---|---|---|
| Valeur Power BI | #,0.00 | Valeur numérique recalculée par Power BI avec les réglages Excel sur les 12 mois du Dashboard Excel. |
| Mois Power BI | mmmm yyyy | Mois recalculé par Power BI (indicateurs de type Mois). |
| Texte Power BI | texte | Texte recalculé par Power BI (indicateurs de type Texte). |
| Valeur Excel | #,0.00 | Valeur lue dans l'onglet Calculs. |
| Écart | #,0.00 | Power BI - Excel (indicateurs numériques). |
| Excel (affiché) | texte | Valeur Excel mise en forme. |
| Power BI (affiché) | texte | Valeur Power BI mise en forme. |
| Statut du contrôle | texte | ✔ si Power BI et Excel concordent (tolérance de la ligne), sinon ✘. |
| Contrôles en écart | 0 | Nombre de contrôles ✘. |


## 8. Utilisation dans le rapport (champs par visuel)

Couleurs du classeur (`build_dashboard.py`) : bleu nuit `#1F2A44` (bandeaux, en-têtes), encre `#1F2937`, gris
`#6B7280`, accent `#2A78D6`, ambre `#EDA100`, rouge d'alerte `#B42318`, vert `#1E7B34`, fond `#EEF1F5`. Projets
(groupes 1 à 9) : `#2A78D6`, `#EB6834`, `#1BAF7A`, `#EDA100`, `#E87BA4`, `#008300`, `#6250D6`, `#E34948`, `#9CA3AF`
(« Autres projets »). Occupation : fond / texte `#E3F4E8` / `#1E7B34` (< 90 %), `#FFF1CC` / `#8A5A00` (90-100 %),
`#F8D7D7` / `#B42318` (> 100 %). Graphique d'une ressource : charge fixe `#B8C0CC`, projets `#2A78D6`, dépassement
`#D03B3B`, capacité `#1F2937` en pointillés. Barres de données : % de la charge `#86B6EF`, avancement `#7FD1AE`.

### Segments (synchronisés sur toutes les pages)
- `'Périmètre'[Périmètre]` et `'Répartition'[Répartition]` : sélection unique, « Comme dans Excel ».
- `Mois[Période]` (multi-sélection, « 12 mois du Dashboard Excel ») ; en option plage de `Mois[Mois]`.
- `Lignes[Chef de projet]`, `Lignes[Projet]`, `Lignes[Type]` ; page Ressources : `Ressources[Ressource]` (sélection unique).

### Page « Tableau de bord » (onglet Dashboard)
- Bandeau : `[Rappel des réglages]`, `[Période affichée]`, `[Données lues le]` ; titre de section `[Titre du périmètre]` ;
  alerte si `[Lignes avec dessinateur et externe]` > 0 (lignes à corriger dans Charge).
- Carte CHEFS DE PROJET : `[Nombre de chefs de projet]`, `[Sous-titre chefs]` ; liste : `Lignes[Chef de projet]` +
  `[Charge identifiée (h)]`, tri décroissant, 8 premiers.
- Carte PROGRAMMES SUIVIS : `[Nombre de projets]` ; liste : `Lignes[Projet]` + `[Charge identifiée (h)]` (ou
  `Lignes[Groupe projet]` avec pastille `Lignes[Couleur projet]` en mise en forme par valeur de champ).
- Carte CHARGE IDENTIFIÉE : `[Charge identifiée (h)]`, `[Sous-titre charge identifiée]`, `[Réalisé (h)]`,
  `[Reste à faire (h)]` ; par type : `Lignes[Type]` + `[Charge identifiée (h)]`, 4 premiers.
- Carte ★ PLUS GROS RESTE À FAIRE : `[Plus gros reste à faire (h)]`, `[Sous-titre plus gros reste]`,
  `[Plus gros reste à faire – projet]`, `[Plus gros reste à faire – sous-projet]`, `[Plus gros reste à faire – produit]`,
  `[Plus gros reste à faire – chef de projet]`, `[Plus gros reste à faire – ressources]`,
  `[Plus gros reste à faire – début planifié]`, `[Plus gros reste à faire – date de besoin]`,
  `[Plus gros reste à faire – avancement]` (en rouge si `[Plus gros reste à faire – retard]` > 0).
- Tuile MOIS EN SURCHARGE : `[Mois en surcharge]` (rouge `#B42318` si > 0), `[Sous-titre capacité]`,
  `[Sous-titre surcharge]`.
- Tuile PIC DE CHARGE : `[Pic de charge (h)]`, `[Sous-titre pic (mois)]`, `[Sous-titre pic (affaires)]`.
- Tuile AFFAIRES EN RETARD : `[Affaires en retard]` (rouge si > 0), `[Sous-titre hors délai]`, `[Sous-titre écart]`.
- Tuile NOMBRE D'AFFAIRES : `[Nombre d'affaires]`, `[Sous-titre projets]`, `[Sous-titre lignes terminées]`.
- Tuile TAUX D'AVANCEMENT MOYEN : `[Avancement moyen pondéré]`, texte « pondéré par les heures des affaires »,
  `[Sous-titre avancement]`.
- Histogramme empilé « charge en heures par mois et par projet » : axe `Mois[Libellé mois]` (trié par `Mois[Clé de tri]`),
  valeurs `[Groupe 1 – charge répartie (h)]` … `[Groupe 9 – charge répartie (h)]` avec les 9 couleurs des projets dans
  l'ordre (comme les 9 séries de l'Excel ; légende remplacée par le tableau légende ci-dessous), ligne
  `[Capacité projets de l'équipe (h/mois)]` en pointillés `#1F2937` ; étiquette de total : `[Charge répartie (h)]`.
  Variante avec clic croisé par projet : légende `Lignes[Groupe projet]` et valeur `[Charge répartie (h)]` (les couleurs
  suivent alors l'ordre du thème, pas forcément celui de l'Excel). Sous-titres : `[Sous-titre graphique mensuel]`,
  `[Texte hors période]`.
- Anneau « répartition par projet » : valeurs `[Groupe 1 – charge (h)]` … `[Groupe 9 – charge (h)]` (mêmes couleurs),
  total au centre `[Charge identifiée (h)]`. Tableau légende : `Lignes[Groupe projet]`, `Lignes[Ordre groupe projet]`
  (tri), pastille `Lignes[Couleur projet]`, `[Charge identifiée (h)]`, `[Part de la charge (%)]` ; ou les mesures
  `[Groupe 1 – libellé]` …
- Tableau des affaires (une ligne par affaire, tri par `Affaires[N°]`) : `Affaires[N°]`, `Affaires[Projet]`,
  `Affaires[Sous-projet]`, `Affaires[Produit]`, `Affaires[Chef de projet]`, `Affaires[Ressources]`, `Affaires[Charge (h)]`,
  `Affaires[% de la charge totale]` (barre `#86B6EF`), `Affaires[Avancement]` (barre `#7FD1AE`, couleur de police par
  valeur de champ `[Couleur avancement]`), `Affaires[Début (affiché)]` (ou `Affaires[Début planifié]`),
  `Affaires[Date de besoin]`, `Affaires[Statut]`. Les colonnes `Affaires[Charge (h)]` et `Affaires[% de la charge totale]`
  reprennent exactement l'Excel (périmètre entier) ; avec un segment Chef / Projet / Type actif, préférer
  `[Charge identifiée (h)]` et `[Part de la charge (%)]` qui suivent la sélection. Extraction vers la fiche affaire :
  champ `Affaires[Affaire]`.

### Page « Ressources » (onglet Dessinateurs)
- Carte de chaleur (matrice) : lignes `Ressources[Ressource]` (+ `Ressources[Mention]`), colonnes `Mois[Libellé mois]`,
  valeur `[Occupation (%)]`, fond `[Couleur de fond (occupation)]`, police `[Couleur du texte (occupation)]` (mise en
  forme par valeur de champ) ; total de ligne = moyenne des mois ; `[Occupation maximale (%)]` en info-bulle. Filtre
  facultatif `Ressources[Dans le tableau Excel]` = vrai pour retrouver les 16 lignes de l'Excel. Heures non
  affectées : ligne « (non affecté) » avec `[Charge répartie (h)]` (son occupation est vide).
- Ressource sélectionnée (segment `Ressources[Ressource]`, sélection unique) : titre `[Titre de la ressource]`,
  `[Origine de la ressource]`.
  - Tuile CAPACITÉ MENSUELLE : `[Capacité de la ressource (h/mois)]`, `[Sous-titre charge fixe de la ressource]`,
    `[Sous-titre capacité de la ressource]` (rouge s'il commence par « Attention »).
  - Tuile CHARGE PROJETS AFFECTÉE : `[Charge identifiée (h)]`, `[Sous-titre affaires de la ressource]`,
    `[Sous-titre charge sur la période]`.
  - Tuile MOIS LE PLUS CHARGÉ : `[Mois le plus chargé de la ressource (texte)]` (rouge si
    `[Occupation au pic de la ressource (%)]` > 100 %), `[Sous-titre pic de la ressource]`,
    `[Sous-titre occupation au pic]`.
  - Tuile MOIS EN SURCHARGE : `[Mois en surcharge de la ressource]` (rouge si > 0, vert si 0),
    `[Sous-titre dépassement cumulé]`, `[Sous-titre surcharge de la ressource]`.
  - Tuile OCCUPATION MOYENNE : `[Occupation (%)]`, `[Sous-titre heures disponibles]`, texte « sur les mois affichés ».
  - Histogramme empilé par `Mois[Libellé mois]` : `[Charge fixe (h)]` (`#B8C0CC`), `[Projets dans la capacité (h)]`
    (`#2A78D6`), `[Dépassement de capacité (h)]` (`#D03B3B`) ; ligne `[Capacité de la ressource (h/mois)]` (`#1F2937`
    pointillés).
  - Tableau mensuel (valeurs en lignes, mois en colonnes) : `[Charge fixe (h)]`, `[Charge répartie (h)]` (charge
    projets), `[Charge totale de la ressource (h)]`, `[Capacité de la ressource (h/mois)]`,
    `[Marge de la ressource (h)]` (format `+#,0;-#,0;0`, rouge si négatif, vert si positif), `[Occupation (%)]`
    (couleurs d'occupation).
  - Ses lignes (tri par `Lignes[Rang ligne]`) : `Lignes[Projet]`, `Lignes[Sous-projet]`, `Lignes[Produit]`,
    `Lignes[Chef de projet]`, `Lignes[Heures (h)]`, `[Charge répartie (h)]` (heures sur la période),
    `Lignes[Début planifié]`, `Lignes[Date de besoin]`, `Lignes[Avancement]` ; ses affaires : `Affaires[Affaire]` +
    `[Charge identifiée (h)]` (filtre dans les deux sens).

### Page « Fiche affaire » (extraction)
- Champ d'extraction `Affaires[Affaire]` (garder tous les filtres : la variante du segment Périmètre suit).
- En-tête : `Affaires[Affaire]`, `Affaires[Chef de projet]`, `Affaires[Ressources]`, `Affaires[Statut]` (couleur
  `Affaires[Couleur d'alerte]`), `Affaires[Début planifié]`, `Affaires[Date de besoin]`, `Affaires[Avancement]`,
  `Affaires[Avancement attendu]`, `Affaires[Heures toutes lignes (h)]`, `Affaires[Réalisé toutes lignes (h)]`,
  `Affaires[Reste à faire de l'affaire (h)]`, `Affaires[Charge (h)]`, `Affaires[Lignes de l'affaire]`.
- Ses lignes (tri `Lignes[N° ligne]`) : `Lignes[N° ligne]`, `Lignes[Ressource de la ligne]`, `Lignes[Type de ressource]`,
  `Lignes[Heures (h)]`, `Lignes[Avancement]`, `Lignes[Réalisé de la ligne (h)]`, `Lignes[Reste de la ligne (h)]`,
  `Lignes[Début planifié]`, `Lignes[Date de besoin]`, `Lignes[Type]`, `Lignes[Coordinateur]`, `Lignes[Description]`,
  `Lignes[Contraintes]`, `Lignes[Autres informations]`, `Lignes[Dessinateur et externe]` (alerte).
- Charge par mois de l'affaire : axe `Mois[Libellé mois]`, valeur `[Charge répartie (h)]`, légende facultative
  `Lignes[Ressource de la ligne]`.
- Les lignes terminées hors du périmètre « Projets en cours » comptent dans l'avancement de l'affaire mais
  n'apparaissent qu'avec le périmètre « Tous les projets ».

### Page « Contrôle »
- Tableau trié par `'Contrôle Excel'[Ordre]` : `'Contrôle Excel'[Catégorie]`, `'Contrôle Excel'[Contrôle]`,
  `'Contrôle Excel'[Cellule Excel]`, `[Excel (affiché)]`, `[Power BI (affiché)]`, `[Écart]`, `[Statut du contrôle]`
  (✔ / ✘, police verte / rouge). Carte `[Contrôles en écart]`. Rappel : `Réglages[Rappel des réglages (Excel)]`,
  `Réglages[Période affichée (Excel)]`, `[Données lues le]`.
- Les mesures de contrôle ignorent tous les segments : réglages « Comme dans Excel », 12 mois du Dashboard Excel.

## 9. Choix et limites

- **Date du jour** : Power BI utilise la date de l'actualisation (heure de Paris), l'Excel la date de son dernier
  recalcul. L'avancement attendu et les retards sont donc figés entre deux actualisations ; sur la page Contrôle, les
  lignes « affaires en retard » peuvent différer si le classeur a été enregistré un autre jour que l'actualisation
  (idem pour le 1er mois automatique quand aucune ligne n'a de date de début).
- **Espace insécable** : la clé d'affaire remplace CAR(160) par une espace comme Excel Windows (règle 3 de la
  spécification). Le recalcul de référence par LibreOffice ne le fait pas (son CAR(160) n'est pas U+00A0) : les tests
  rejouent donc la règle d'Excel Windows sur les clés lues et la vérifient par des tests unitaires.
- **Jokers de NB.SI** (* ? ~) : neutralisés dans la clé d'affaire comme dans l'Excel. Dans les noms de chef, de projet
  ou de type, l'Excel les interprète comme des jokers (anomalie de ses listes de classement) ; Power BI compare les
  textes exactement (casse ignorée). Défaut du classeur, signalé dans LISEZMOI.md (scénario de relecture
  r_jokers_noms : « VEG* » regroupé avec VEGA par Excel, 3 contrôles ✘).
- **Date de début avec une heure** : Excel garde l'heure (Calculs I) et exclut la journée (ou la ligne) quand le début
  tombe le dernier jour d'un mois avec une heure ; `fnDateExcel` ne garde que la date (règle de la spécification).
  Défaut du classeur, signalé dans LISEZMOI.md (scénario r_heure, 5 contrôles ✘).
- **Libellés** : chef, projet et type affichés = texte de la 1re ligne du groupe (casse ignorée), SUPPRESPACE
  appliqué ; le chef vide devient « (non renseigné) » (l'Excel affiche une cellule vide dans le tableau des affaires).
  Le moteur de Power BI compare les textes sans tenir compte de la casse : `Lignes[Projet]` « ATLAS » et « atlas »
  sont une seule valeur.
- **Arrondis** : charge, heures, réalisé et reste des affaires arrondis à 0,01 h comme l'Excel (ARRONDI). Les mesures
  (`[Charge identifiée (h)]`, `[Part de la charge (%)]`) ne sont pas arrondies : écart < 0,01 h / 0,01 point.
- **Lignes terminées** : `[Lignes terminées]` et `[Lignes avec dessinateur et externe]` comptent les lignes de la
  variante « Tous les projets » (toutes les lignes valides, comme l'Excel) ; elles suivent les segments et clics sur
  les colonnes de `Lignes` (chef, projet, type) mais pas un clic sur une affaire (`REMOVEFILTERS ( Affaires )` : les
  colonnes d'`Affaires` diffèrent d'une variante à l'autre). Sans filtre : valeur d'Excel (AL24, CD27).
- **Lignes d'une ressource** : le tableau mêle des colonnes de `Lignes` et `Affaires[Affaire]` ;
  `[Heures de la ligne sur la période (h)]` = `IF ( NOT ( ISEMPTY ( Lignes ) ), [Charge répartie (h)] )` vaut 0 pour
  une ligne sans heure sur la période (sans date de début, 0 h, hors des mois choisis) et reste vide pour une
  combinaison ligne × affaire inexistante : toutes les lignes sont listées, comme l'onglet Dessinateurs.
- **Fiche affaire** : `[Charge répartie sans les mois vides (h)]` (sans « + 0 ») ne garde sur l'axe que les mois où
  l'affaire a des heures.
- **Ressources** : toutes les ressources des listes Data sont présentes (l'Excel n'en affiche que 16) ; les ressources
  saisies dans Charge mais absentes de Data apparaissent avec les réglages par défaut (l'Excel les regroupe dans
  « Autres »).
- **Groupe projet** : son rang et sa couleur dépendent de la variante de périmètre ; il n'a donc pas de colonne de tri
  (Power BI exige une seule valeur de tri par valeur). D'où les mesures `[Groupe 1 – …]` à `[Groupe 9 – …]` pour des
  couleurs fixes.
- **Numéros de ligne** : `Lignes[N° ligne]` suppose que la plage du nom `bi_Charge` est lue à partir de sa 1re ligne
  (ligne 4) ; il ne sert qu'à l'affichage et à l'ordre.
- **FORMAT** des mesures textes : culture fr-FR explicite ; le séparateur de milliers est une espace fine insécable.
- **Libellés de mois** des sous-titres : lus dans `Mois` sans tenir compte des filtres de la page
  (`CALCULATE ( LOOKUPVALUE ( … ), REMOVEFILTERS ( Mois ) )`).
- **Taille** : au plus 500 lignes de Charge et 200 lignes de Data, comme l'Excel.

## 10. Preuves

Power BI, le moteur M et le moteur DAX ne s'exécutent pas ici ; la justesse est établie ainsi :

1. **Syntaxe et références** (`check_model.py`) : les 34 expressions M passent l'analyseur officiel
   `@microsoft/powerquery-parser` et tous leurs identifiants sont résolus (portée locale, requêtes, bibliothèque
   standard connue) ; les 136 mesures DAX passent `pbi_parsers` (texte entièrement reconnu) avec des fonctions
   connues et un nombre d'arguments valide, et chaque colonne, mesure ou table citée existe. `pbi_parsers` couvre
   toutes les constructions utilisées (VAR / RETURN, y compris dans CALCULATE, `IN { … }`, FORMAT à 3 arguments,
   TOPN à plusieurs clés) mais tolère les arguments vides (`CALCULATE ( x, )`) : un contrôle textuel les refuse.
   Les témoins (expressions fausses rejetées, justes acceptées) et `fautes_check_model.py` (16 fautes injectées,
   toutes signalées) prouvent que ces contrôles détectent réellement les erreurs.
2. **Logique** (`test_mirror.py`) : le miroir Python reproduit pas à pas les requêtes M et les mesures DAX ; il est
   comparé, sans écart (tolérance 0,01 h, égalité stricte des valeurs arrondies à 0,01 h), aux valeurs calculées par
   le classeur : `build12/final.xlsx`, les 13 scénarios `build12/sc_<cas>.xlsx`, `build13/final.xlsx` (lu par les
   vrais noms bi_*), 11 scénarios supplémentaires (`scen_variantes.py` : variantes de réglages, 1er mois saisi,
   arrondis à la demi-unité, seuils du retard) et 14 croisements « variante de segment ↔ Excel recalculé avec ce
   réglage ». Les fonctions M sensibles (NB.JOURS.OUVRES sur 3 000 couples de dates, ARRONDI, SUPPRESPACE, clé
   d'affaire, conversion UTC → Paris de 2024 à 2031) ont en plus des tests unitaires.
3. **Sensibilité** (`mutations_mirror.py`) : 14 erreurs plausibles (samedi ouvré, fériés ignorés, arrondi au pair,
   périmètre, capacités propres ignorées, seuil de retard, date du jour en UTC, casse, espace insécable, moyenne
   pondérée, occupation non arrondie, 1er mois saisi ignoré, ressource K+L, répartition sur le mois du besoin) sont
   toutes détectées par `test_mirror.py`.
4. **Contrôle croisé** : `validate_bim.py` (outil de la mission « format », données du miroir en entrée) ne signale
   aucune erreur ; seul avertissement, voulu : la relation à double sens `Lignes` ↔ `Affaires`.

Journaux de la dernière exécution :

`check_model.txt` :

```
témoins M : 9 expressions, analyseur discriminant : oui
témoins DAX : 7 expressions, analyseur discriminant : oui
34 expression(s) M analysée(s), 0 erreur(s)
87 fonctions et constantes de la bibliothèque M, toutes connues : oui
34 expressions M reconnues par @microsoft/powerquery-parser
34 requêtes M, dépendances sans cycle : oui
133 colonnes déclarées, toutes produites par leur requête M et typées : oui
136 mesures DAX analysées par pbi_parsers, 417 références contrôlées, fonctions : ABS, ALLSELECTED, AVERAGEX, BLANK, CALCULATE, COALESCE, CONCATENATEX, COUNTROWS, DISTINCTCOUNT, DIVIDE, FALSE, FILTER, FORMAT, HASONEVALUE, IF, INT, ISBLANK, ISEMPTY, KEEPFILTERS, LOOKUPVALUE, MAX, MAXX, MIN, MINX, NOT, REMOVEFILTERS, ROUND, SELECTEDVALUE, SUM, SUMX, SWITCH, TOPN, TRUE, UPPER, VALUES
11 tables, 6 relations, 23 expressions partagées, 136 mesures
CONTRÔLES STATIQUES : 0 erreur
```

`fautes_check_model.txt` :

```
détectée     colonne DAX inconnue : ['ERREUR DAX [Charge identifiée (h)] : colonne inconnue Lignes[Heure (h)]']
détectée     mesure DAX inconnue : ['ERREUR DAX [ETP-mois] : mesure inconnue [Charge identifiee (h)]']
détectée     fonction DAX mal écrite : ['ERREUR DAX [Charge identifiée (h)] : fonction inconnue ou non autorisée SELECTEDVALEU']
détectée     argument DAX vide : ["ERREUR DAX [Nombre d'affaires] : argument vide"]
détectée     parenthèse DAX manquante : ['ERREUR DAX [Capacité projets par personne (h/mois)] : analyse impossible (AssertionError )']
détectée     syntaxe M : ["ERREUR syntaxe   RépartitionLignes : Expected to find a right parenthesis <')'>, but a keyword <'in'> was found instead (ligne 39, colonne 1)"]
détectée     colonne non produite : ['ERREUR colonne Lignes[Colonne fantôme] non produite par la requête M']
détectée     relation vers une colonne absente : ['ERREUR relation 7eb6d8d7-a090-51f8-9e85-aa319759187b : Mois[Mois du calendrier] inconnue']
détectée     mesure homonyme d'une colonne : ['ERREUR DAX [Sous-titre lignes de la ressource] : mesure inconnue [Nombre de lignes]']
détectée     double chemin de filtre : ["ERREUR plusieurs chemins de filtre Périmètre -> Charge mensuelle : [['Périmètre', 'Affaires', 'Lignes', 'Charge mensuelle'], ['Périmètre', 'Affaires', 'Charge mensuelle']]"]
détectée     deuxième source dans Classeur : ["ERREUR accès aux données : un seul Web.Contents attendu dans Classeur, trouvé [('Classeur', 'Web.Contents'), ('Classeur', 'Web.Contents')]"]
détectée     source dans une autre requête : ['ERREUR M : fonction ou constante de bibliothèque inconnue File.Contents']
détectée     requête M inconnue : ['ERREUR identifiant inconnu  DonnéesData : fnLireNoms']
détectée     variable M hors de portée : ['ERREUR identifiant inconnu  RépartitionLignes : DébutMois']
détectée     fonction M mal écrite : ['ERREUR M : fonction ou constante de bibliothèque inconnue Table.SelecRows']
détectée     texte sans accent : ['ERREUR DAX [Valeur Power BI] : mesure inconnue [Mois en surcharge]']
TOUTES LES FAUTES DÉTECTÉES
```

`test_mirror.txt` :

```
OK   18294 contrôles,   0 écart(s)  couverture des mesures et heure de Paris
OK    2809 contrôles,   0 écart(s)  build12/final.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_ext.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_double.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2946 contrôles,   0 écart(s)  build12/sc_tous.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2820 contrôles,   0 écart(s)  build12/sc_casse.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_joker.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2878 contrôles,   0 écart(s)  build12/sc_partiel.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_weekend.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_debut_we.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_retard.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2813 contrôles,   0 écart(s)  build12/sc_termine.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_surcharge.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2997 contrôles,   0 écart(s)  build12/sc_termine_tous.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build12/sc_cap_zero.xlsx (Comme dans Excel, aujourd'hui = 19/10/2026, noms bi_* de repli)
OK    2809 contrôles,   0 écart(s)  build13/final.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2968 contrôles,   0 écart(s)  scen_variantes_1/v_final_tous.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2809 contrôles,   0 écart(s)  scen_variantes_1/v_final_debut.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2809 contrôles,   0 écart(s)  scen_variantes_1/v_premier_mois.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2813 contrôles,   0 écart(s)  scen_variantes_1/v_termine_encours.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2968 contrôles,   0 écart(s)  scen_variantes_1/v_retard_tous_debut.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2809 contrôles,   0 écart(s)  scen_variantes_1/v_surcharge_debut.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2809 contrôles,   0 écart(s)  scen_variantes_1/v_weekend_debut.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2968 contrôles,   0 écart(s)  scen_variantes_1/v_joker_tous.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    3037 contrôles,   0 écart(s)  scen_variantes_1/v_partiel_tous.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2925 contrôles,   0 écart(s)  scen_variantes_3/v_arrondi.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    3252 contrôles,   0 écart(s)  scen_variantes_4/v_seuil.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2918 contrôles,   0 écart(s)  scen_lo/r_sans_debut.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2691 contrôles,   0 écart(s)  scen_lo/r_data_trou.xlsx (Comme dans Excel, aujourd'hui = 20/10/2026)
OK    2704 contrôles,   0 écart(s)  variante Projets en cours / Lissée jusqu'au besoin sur final.xlsx = Excel final.xlsx (aujourd'hui = 19/10/2026)
OK    2863 contrôles,   0 écart(s)  variante Tous les projets / Lissée jusqu'au besoin sur final.xlsx = Excel v_final_tous.xlsx (aujourd'hui = 20/10/2026)
OK    2704 contrôles,   0 écart(s)  variante Projets en cours / Tout sur le mois de début sur final.xlsx = Excel v_final_debut.xlsx (aujourd'hui = 20/10/2026)
OK    2841 contrôles,   0 écart(s)  variante Tous les projets / Tout sur le mois de début sur final.xlsx = Excel sc_tous.xlsx (aujourd'hui = 19/10/2026)
OK    2704 contrôles,   0 écart(s)  variante Projets en cours / Lissée jusqu'au besoin sur sc_tous.xlsx = Excel final.xlsx (aujourd'hui = 19/10/2026)
OK    2863 contrôles,   0 écart(s)  variante Tous les projets / Lissée jusqu'au besoin sur sc_tous.xlsx = Excel v_final_tous.xlsx (aujourd'hui = 20/10/2026)
OK    2708 contrôles,   0 écart(s)  variante Projets en cours / Lissée jusqu'au besoin sur sc_termine_tous.xlsx = Excel v_termine_encours.xlsx (aujourd'hui = 20/10/2026)
OK    2892 contrôles,   0 écart(s)  variante Tous les projets / Lissée jusqu'au besoin sur v_termine_encours.xlsx = Excel sc_termine_tous.xlsx (aujourd'hui = 19/10/2026)
OK    2863 contrôles,   0 écart(s)  variante Tous les projets / Tout sur le mois de début sur sc_retard.xlsx = Excel v_retard_tous_debut.xlsx (aujourd'hui = 20/10/2026)
OK    2704 contrôles,   0 écart(s)  variante Projets en cours / Tout sur le mois de début sur sc_surcharge.xlsx = Excel v_surcharge_debut.xlsx (aujourd'hui = 20/10/2026)
OK    2704 contrôles,   0 écart(s)  variante Projets en cours / Tout sur le mois de début sur sc_weekend.xlsx = Excel v_weekend_debut.xlsx (aujourd'hui = 20/10/2026)
OK    2863 contrôles,   0 écart(s)  variante Tous les projets / Lissée jusqu'au besoin sur sc_joker.xlsx = Excel v_joker_tous.xlsx (aujourd'hui = 20/10/2026)
OK    2932 contrôles,   0 écart(s)  variante Tous les projets / Lissée jusqu'au besoin sur sc_partiel.xlsx = Excel v_partiel_tous.xlsx (aujourd'hui = 20/10/2026)
OK    2704 contrôles,   0 écart(s)  variante Projets en cours / Tout sur le mois de début sur sc_debut_we.xlsx = Excel sc_debut_we.xlsx (aujourd'hui = 19/10/2026)
TOTAL : 137663 contrôles sur 28 classeurs et 14 variantes, 0 écart(s)
MIROIR = EXCEL : 0 écart
```

`mutations.txt` :

```
référence : 0 écart(s) sur 66830 contrôles
mut_samedi               détectée : 9288 écart(s) dans ['build12/final.xlsx', 'build12/sc_cap_zero.xlsx', 'build12/sc_casse.xlsx', 'build12/sc_double.xlsx', 'build12/sc_joker.xlsx', 'build12/sc_partiel.xlsx', 'build12/sc_retard.xlsx', 'build12/sc_surcharge.xlsx', 'build12/sc_termine.xlsx', 'build12/sc_tous.xlsx', 'build12/sc_weekend.xlsx', 'couverture des mesures et heure de Paris', 'final.xlsx = Excel v_final_debut.xlsx', 'final.xlsx = Excel v_final_tous.xlsx', 'sc_tous.xlsx = Excel v_final_tous.xlsx', 'scen_variantes_1/v_premier_mois.xlsx', 'scen_variantes_3/v_arrondi.xlsx', 'scen_variantes_4/v_seuil.xlsx']
mut_feries               détectée : 408 écart(s) dans ['build12/sc_weekend.xlsx']
mut_arrondi              détectée : 22 écart(s) dans ['couverture des mesures et heure de Paris', 'scen_variantes_3/v_arrondi.xlsx']
mut_perimetre            détectée (exception TypeError)
mut_capacite             détectée : 994 écart(s) dans ['build12/final.xlsx', 'build12/sc_cap_zero.xlsx', 'build12/sc_casse.xlsx', 'build12/sc_double.xlsx', 'build12/sc_joker.xlsx', 'build12/sc_partiel.xlsx', 'build12/sc_retard.xlsx', 'build12/sc_surcharge.xlsx', 'build12/sc_termine.xlsx', 'build12/sc_tous.xlsx', 'build12/sc_weekend.xlsx', 'final.xlsx = Excel v_final_debut.xlsx', 'final.xlsx = Excel v_final_tous.xlsx', 'sc_tous.xlsx = Excel v_final_tous.xlsx', 'scen_variantes_1/v_premier_mois.xlsx', 'scen_variantes_3/v_arrondi.xlsx', 'scen_variantes_4/v_seuil.xlsx']
mut_seuil_retard         détectée : 19 écart(s) dans ['couverture des mesures et heure de Paris', 'scen_variantes_4/v_seuil.xlsx']
mut_aujourdhui           détectée : 70 écart(s) dans ['build12/final.xlsx', 'build12/sc_cap_zero.xlsx', 'build12/sc_casse.xlsx', 'build12/sc_double.xlsx', 'build12/sc_joker.xlsx', 'build12/sc_partiel.xlsx', 'build12/sc_retard.xlsx', 'build12/sc_surcharge.xlsx', 'build12/sc_termine.xlsx', 'build12/sc_tous.xlsx', 'build12/sc_weekend.xlsx', 'final.xlsx = Excel v_final_debut.xlsx', 'final.xlsx = Excel v_final_tous.xlsx', 'sc_tous.xlsx = Excel v_final_tous.xlsx', 'scen_variantes_1/v_premier_mois.xlsx', 'scen_variantes_3/v_arrondi.xlsx', 'scen_variantes_4/v_seuil.xlsx']
mut_casse                détectée : 1012 écart(s) dans ['build12/final.xlsx', 'build12/sc_cap_zero.xlsx', 'build12/sc_casse.xlsx', 'build12/sc_double.xlsx', 'build12/sc_joker.xlsx', 'build12/sc_partiel.xlsx', 'build12/sc_retard.xlsx', 'build12/sc_surcharge.xlsx', 'build12/sc_termine.xlsx', 'build12/sc_tous.xlsx', 'build12/sc_weekend.xlsx', 'couverture des mesures et heure de Paris', 'final.xlsx = Excel v_final_debut.xlsx', 'final.xlsx = Excel v_final_tous.xlsx', 'sc_tous.xlsx = Excel v_final_tous.xlsx', 'scen_variantes_1/v_premier_mois.xlsx', 'scen_variantes_3/v_arrondi.xlsx', 'scen_variantes_4/v_seuil.xlsx']
mut_insecable            détectée : 3 écart(s) dans ['build12/sc_joker.xlsx', 'couverture des mesures et heure de Paris']
mut_pondere              détectée : 62 écart(s) dans ['build12/final.xlsx', 'build12/sc_cap_zero.xlsx', 'build12/sc_casse.xlsx', 'build12/sc_double.xlsx', 'build12/sc_joker.xlsx', 'build12/sc_partiel.xlsx', 'build12/sc_retard.xlsx', 'build12/sc_surcharge.xlsx', 'build12/sc_termine.xlsx', 'build12/sc_tous.xlsx', 'build12/sc_weekend.xlsx', 'final.xlsx = Excel v_final_debut.xlsx', 'final.xlsx = Excel v_final_tous.xlsx', 'sc_tous.xlsx = Excel v_final_tous.xlsx', 'scen_variantes_1/v_premier_mois.xlsx', 'scen_variantes_3/v_arrondi.xlsx', 'scen_variantes_4/v_seuil.xlsx']
mut_occupation           détectée : 1852 écart(s) dans ['build12/final.xlsx', 'build12/sc_cap_zero.xlsx', 'build12/sc_casse.xlsx', 'build12/sc_double.xlsx', 'build12/sc_joker.xlsx', 'build12/sc_partiel.xlsx', 'build12/sc_retard.xlsx', 'build12/sc_surcharge.xlsx', 'build12/sc_termine.xlsx', 'build12/sc_weekend.xlsx', 'final.xlsx = Excel v_final_tous.xlsx', 'sc_tous.xlsx = Excel v_final_tous.xlsx', 'scen_variantes_1/v_premier_mois.xlsx', 'scen_variantes_3/v_arrondi.xlsx', 'scen_variantes_4/v_seuil.xlsx']
mut_premier_mois         détectée : 105 écart(s) dans ['scen_variantes_1/v_premier_mois.xlsx']
mut_ressource_double     détectée (exception TypeError)
mut_debut_fin            détectée : 610 écart(s) dans ['build12/sc_tous.xlsx', 'final.xlsx = Excel v_final_debut.xlsx']
TOUTES LES MUTATIONS DÉTECTÉES
```

## Annexe — code DAX des mesures

**[Charge identifiée (h)]** (Charge)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SUM ( Lignes[Heures (h)] ), 'Périmètre'[Périmètre] = _p ) + 0
```

**[Réalisé (h)]** (Charge)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SUM ( Lignes[Réalisé de la ligne (h)] ), 'Périmètre'[Périmètre] = _p ) + 0
```

**[Reste à faire (h)]** (Charge)

```dax
[Charge identifiée (h)] - [Réalisé (h)]
```

**[Capacité projets par personne (h/mois)]** (Charge)

```dax
MAX ( Réglages[Capacité projets par personne (Excel)] )
```

**[ETP-mois]** (Charge)

```dax
VAR _c = [Capacité projets par personne (h/mois)]
RETURN
    IF ( _c <= 0, 0, DIVIDE ( [Charge identifiée (h)], _c ) )
```

**[Nombre de chefs de projet]** (Charge)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( DISTINCTCOUNT ( Lignes[Chef de projet] ), 'Périmètre'[Périmètre] = _p ) + 0
```

**[Nombre de projets]** (Charge)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( DISTINCTCOUNT ( Lignes[Projet] ), 'Périmètre'[Périmètre] = _p ) + 0
```

**[Nombre de types]** (Charge)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( DISTINCTCOUNT ( Lignes[Type] ), 'Périmètre'[Périmètre] = _p ) + 0
```

**[Nombre de lignes]** (Charge)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( COUNTROWS ( Lignes ), 'Périmètre'[Périmètre] = _p ) + 0
```

**[Part de la charge (%)]** (Charge)

```dax
DIVIDE ( [Charge identifiée (h)], CALCULATE ( [Charge identifiée (h)], ALLSELECTED () ) )
```

**[Nombre d'affaires]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( COUNTROWS ( Affaires ), 'Périmètre'[Périmètre] = _p ) + 0
```

**[Avancement moyen simple]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _n = COUNTROWS ( Affaires )
        RETURN IF ( _n = 0, 0, DIVIDE ( SUM ( Affaires[Avancement] ), _n ) ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Avancement moyen pondéré]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _h = SUM ( Affaires[Heures toutes lignes (h)] )
        RETURN IF ( _h = 0, [Avancement moyen simple], DIVIDE ( SUM ( Affaires[Réalisé toutes lignes (h)] ), _h ) ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Affaires en retard]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( COUNTROWS ( Affaires ), 'Périmètre'[Périmètre] = _p, KEEPFILTERS ( Affaires[Retard] > 0 ) ) + 0
```

**[Affaires hors délai]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( COUNTROWS ( Affaires ), 'Périmètre'[Périmètre] = _p, KEEPFILTERS ( Affaires[Retard] = 2 ) ) + 0
```

**[Affaires en écart]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( COUNTROWS ( Affaires ), 'Périmètre'[Périmètre] = _p, KEEPFILTERS ( Affaires[Retard] = 1 ) ) + 0
```

**[Plus gros reste à faire (h)]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    COALESCE ( CALCULATE ( MAX ( Affaires[Reste à faire de l'affaire (h)] ), 'Périmètre'[Périmètre] = _p ), 0 )
```

**[Plus gros reste à faire – affaire]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN CONCATENATEX ( _t, Affaires[Affaire] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – projet]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN CONCATENATEX ( _t, Affaires[Projet] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – sous-projet]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN CONCATENATEX ( _t, Affaires[Sous-projet] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – produit]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN CONCATENATEX ( _t, Affaires[Produit] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – chef de projet]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN CONCATENATEX ( _t, Affaires[Chef de projet] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – ressources]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN CONCATENATEX ( _t, Affaires[Ressources] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – total de l'affaire (h)]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN SUMX ( _t, Affaires[Heures toutes lignes (h)] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – début planifié]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN MAXX ( _t, Affaires[Début planifié] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – date de besoin]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN MAXX ( _t, Affaires[Date de besoin] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – avancement]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN MAXX ( _t, Affaires[Avancement] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Plus gros reste à faire – retard]** (Affaires)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE (
        VAR _t = TOPN ( 1, Affaires, Affaires[Reste à faire de l'affaire (h)], DESC, Affaires[N° 1re ligne], ASC )
        RETURN MAXX ( _t, Affaires[Retard] ),
        'Périmètre'[Périmètre] = _p
    )
```

**[Couleur avancement]** (Affaires)

```dax
IF ( SELECTEDVALUE ( Affaires[Retard], 0 ) > 0, "#B42318", "#1F2937" )
```

**[Lignes terminées]** (Affaires)

```dax
CALCULATE (
    COUNTROWS ( Lignes ),
    REMOVEFILTERS ( Affaires ),
    'Périmètre'[Périmètre] = "Tous les projets",
    KEEPFILTERS ( Lignes[Terminée] = TRUE () )
) + 0
```

**[Lignes avec dessinateur et externe]** (Affaires)

```dax
CALCULATE (
    COUNTROWS ( Lignes ),
    REMOVEFILTERS ( Affaires ),
    'Périmètre'[Périmètre] = "Tous les projets",
    KEEPFILTERS ( Lignes[Dessinateur et externe] = TRUE () )
) + 0
```

**[Charge répartie (h)]** (Mois et capacité)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
VAR _r = SELECTEDVALUE ( 'Répartition'[Répartition], "Comme dans Excel" )
RETURN
    CALCULATE ( SUM ( 'Charge mensuelle'[Heures (h)] ), 'Périmètre'[Périmètre] = _p, 'Répartition'[Répartition] = _r ) + 0
```

**[Charge répartie sans les mois vides (h)]** (Mois et capacité)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
VAR _r = SELECTEDVALUE ( 'Répartition'[Répartition], "Comme dans Excel" )
RETURN
    CALCULATE ( SUM ( 'Charge mensuelle'[Heures (h)] ), 'Périmètre'[Périmètre] = _p, 'Répartition'[Répartition] = _r )
```

**[Heures de la ligne sur la période (h)]** (Mois et capacité)

```dax
IF ( NOT ( ISEMPTY ( Lignes ) ), [Charge répartie (h)] )
```

**[Affaires actives]** (Mois et capacité)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
VAR _r = SELECTEDVALUE ( 'Répartition'[Répartition], "Comme dans Excel" )
RETURN
    CALCULATE ( DISTINCTCOUNT ( 'Charge mensuelle'[Clé affaire] ), 'Périmètre'[Périmètre] = _p, 'Répartition'[Répartition] = _r ) + 0
```

**[Charge hors période ou sans date de début (h)]** (Mois et capacité)

```dax
MAX ( 0, ROUND ( [Charge identifiée (h)] - [Charge répartie (h)], 1 ) )
```

**[Capacité projets de l'équipe (h/mois)]** (Mois et capacité)

```dax
MAX ( Réglages[Capacité projets de l'équipe (Excel)] )
```

**[Mois en surcharge]** (Mois et capacité)

```dax
VAR _cap = ROUND ( [Capacité projets de l'équipe (h/mois)], 0 )
RETURN
    COUNTROWS ( FILTER ( VALUES ( Mois[Mois] ), ROUND ( [Charge répartie (h)], 0 ) > _cap ) ) + 0
```

**[Premier mois en surcharge]** (Mois et capacité)

```dax
VAR _cap = ROUND ( [Capacité projets de l'équipe (h/mois)], 0 )
RETURN
    MINX ( FILTER ( VALUES ( Mois[Mois] ), ROUND ( [Charge répartie (h)], 0 ) > _cap ), Mois[Mois] )
```

**[Pic de charge (h)]** (Mois et capacité)

```dax
COALESCE ( MAXX ( VALUES ( Mois[Mois] ), [Charge répartie (h)] ), 0 )
```

**[Mois du pic]** (Mois et capacité)

```dax
VAR _pic = [Pic de charge (h)]
RETURN
    MINX ( FILTER ( VALUES ( Mois[Mois] ), [Charge répartie (h)] = _pic ), Mois[Mois] )
```

**[Affaires actives au pic]** (Mois et capacité)

```dax
VAR _m = [Mois du pic]
RETURN
    IF ( ISBLANK ( _m ), 0, CALCULATE ( [Affaires actives], Mois[Mois] = _m ) )
```

**[ETP au pic]** (Mois et capacité)

```dax
VAR _c = [Capacité projets par personne (h/mois)]
RETURN
    IF ( _c <= 0, 0, DIVIDE ( [Pic de charge (h)], _c ) )
```

**[Pic en % de la capacité]** (Mois et capacité)

```dax
VAR _cap = [Capacité projets de l'équipe (h/mois)]
RETURN
    IF ( _cap <= 0, BLANK (), DIVIDE ( [Pic de charge (h)], _cap ) )
```

**[Part du pic dans la période]** (Mois et capacité)

```dax
VAR _t = [Charge répartie (h)]
RETURN
    IF ( _t = 0, 0, DIVIDE ( [Pic de charge (h)], _t ) )
```

**[Premier mois affiché]** (Mois et capacité)

```dax
MIN ( Mois[Mois] )
```

**[Dernier mois affiché]** (Mois et capacité)

```dax
MAX ( Mois[Mois] )
```

**[Groupe 1 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 1 )
```

**[Groupe 1 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 1 )
```

**[Groupe 1 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 1 )
```

**[Groupe 2 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 2 )
```

**[Groupe 2 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 2 )
```

**[Groupe 2 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 2 )
```

**[Groupe 3 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 3 )
```

**[Groupe 3 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 3 )
```

**[Groupe 3 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 3 )
```

**[Groupe 4 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 4 )
```

**[Groupe 4 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 4 )
```

**[Groupe 4 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 4 )
```

**[Groupe 5 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 5 )
```

**[Groupe 5 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 5 )
```

**[Groupe 5 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 5 )
```

**[Groupe 6 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 6 )
```

**[Groupe 6 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 6 )
```

**[Groupe 6 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 6 )
```

**[Groupe 7 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 7 )
```

**[Groupe 7 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 7 )
```

**[Groupe 7 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 7 )
```

**[Groupe 8 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 8 )
```

**[Groupe 8 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 8 )
```

**[Groupe 8 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 8 )
```

**[Groupe 9 – libellé]** (Groupes de projets)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    CALCULATE ( SELECTEDVALUE ( Lignes[Groupe projet] ), 'Périmètre'[Périmètre] = _p, Lignes[Ordre groupe projet] = 9 )
```

**[Groupe 9 – charge (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge identifiée (h)], Lignes[Ordre groupe projet] = 9 )
```

**[Groupe 9 – charge répartie (h)]** (Groupes de projets)

```dax
CALCULATE ( [Charge répartie (h)], Lignes[Ordre groupe projet] = 9 )
```

**[Capacité de la ressource (h/mois)]** (Ressource)

```dax
IF ( HASONEVALUE ( Ressources[Clé ressource] ), MAX ( Ressources[Capacité (h/mois)] ) )
```

**[Charge fixe de la ressource (h/mois)]** (Ressource)

```dax
IF ( HASONEVALUE ( Ressources[Clé ressource] ), MAX ( Ressources[Charge fixe (h/mois)] ) )
```

**[Capacité projets de la ressource (h/mois)]** (Ressource)

```dax
IF ( HASONEVALUE ( Ressources[Clé ressource] ), MAX ( Ressources[Capacité projets (h/mois)] ) )
```

**[Charge fixe (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), _f * COUNTROWS ( VALUES ( Mois[Mois] ) ) )
```

**[Charge totale de la ressource (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), SUMX ( VALUES ( Mois[Mois] ), _f + [Charge répartie (h)] ) )
```

**[Projets dans la capacité (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _h = [Charge répartie (h)]
            VAR _marge = _c - ROUND ( _f + _h, 0 )
            RETURN IF ( _marge < 0, MAX ( 0, MIN ( _h, _c - _f ) ), _h )
        )
    )
```

**[Dépassement de capacité (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _h = [Charge répartie (h)]
            VAR _marge = _c - ROUND ( _f + _h, 0 )
            RETURN _h - IF ( _marge < 0, MAX ( 0, MIN ( _h, _c - _f ) ), _h )
        )
    )
```

**[Marge de la ressource (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), SUMX ( VALUES ( Mois[Mois] ), _c - ROUND ( _f + [Charge répartie (h)], 0 ) ) )
```

**[Dépassement cumulé (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _m = _c - ROUND ( _f + [Charge répartie (h)], 0 )
            RETURN IF ( _m < 0, - _m, 0 )
        )
    )
```

**[Heures encore disponibles (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        SUMX (
            VALUES ( Mois[Mois] ),
            VAR _m = _c - ROUND ( _f + [Charge répartie (h)], 0 )
            RETURN IF ( _m > 0, _m, 0 )
        )
    )
```

**[Occupation (%)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        AVERAGEX (
            VALUES ( Mois[Mois] ),
            VAR _t = ROUND ( _f + [Charge répartie (h)], 0 )
            RETURN IF ( _c <= 0, IF ( _t > 0, 9.99, 0 ), DIVIDE ( _t, _c ) )
        )
    )
```

**[Occupation maximale (%)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        MAXX (
            VALUES ( Mois[Mois] ),
            VAR _t = ROUND ( _f + [Charge répartie (h)], 0 )
            RETURN IF ( _c <= 0, IF ( _t > 0, 9.99, 0 ), DIVIDE ( _t, _c ) )
        )
    )
```

**[Couleur de fond (occupation)]** (Ressource)

```dax
VAR _o = [Occupation (%)]
RETURN
    SWITCH ( TRUE (), ISBLANK ( _o ), BLANK (), _o > 1, "#F8D7D7", _o >= 0.9, "#FFF1CC", "#E3F4E8" )
```

**[Couleur du texte (occupation)]** (Ressource)

```dax
VAR _o = [Occupation (%)]
RETURN
    SWITCH ( TRUE (), ISBLANK ( _o ), BLANK (), _o > 1, "#B42318", _o >= 0.9, "#8A5A00", "#1E7B34" )
```

**[Mois en surcharge de la ressource]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        COUNTROWS ( FILTER ( VALUES ( Mois[Mois] ), _c - ROUND ( _f + [Charge répartie (h)], 0 ) < 0 ) ) + 0
    )
```

**[Premier mois en surcharge de la ressource]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _c = [Capacité de la ressource (h/mois)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        MINX ( FILTER ( VALUES ( Mois[Mois] ), _c - ROUND ( _f + [Charge répartie (h)], 0 ) < 0 ), Mois[Mois] )
    )
```

**[Charge au pic de la ressource (h)]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), MAXX ( VALUES ( Mois[Mois] ), _f + [Charge répartie (h)] ) )
```

**[Mois le plus chargé de la ressource]** (Ressource)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
VAR _pic = [Charge au pic de la ressource (h)]
RETURN
    IF (
        NOT ( ISBLANK ( _f ) ),
        MINX ( FILTER ( VALUES ( Mois[Mois] ), _f + [Charge répartie (h)] = _pic ), Mois[Mois] )
    )
```

**[Occupation au pic de la ressource (%)]** (Ressource)

```dax
VAR _m = [Mois le plus chargé de la ressource]
RETURN
    IF ( NOT ( ISBLANK ( _m ) ), CALCULATE ( [Occupation (%)], Mois[Mois] = _m ) )
```

**[Périmètre en vigueur]** (Textes)

```dax
VAR _p = SELECTEDVALUE ( 'Périmètre'[Périmètre], "Comme dans Excel" )
RETURN
    LOOKUPVALUE ( 'Périmètre'[Périmètre appliqué], 'Périmètre'[Périmètre], _p )
```

**[Répartition en vigueur]** (Textes)

```dax
VAR _r = SELECTEDVALUE ( 'Répartition'[Répartition], "Comme dans Excel" )
RETURN
    LOOKUPVALUE ( 'Répartition'[Répartition appliquée], 'Répartition'[Répartition], _r )
```

**[Rappel des réglages]** (Textes)

```dax
[Périmètre en vigueur] & "  ·  "
    & IF ( [Répartition en vigueur] = "Tout sur le mois de début", "tout sur le mois de début",
        "lissage en jours ouvrés jusqu'au besoin" )
    & "  ·  " & SELECTEDVALUE ( Réglages[Rappel capacité] )
```

**[Données lues le]** (Textes)

```dax
SELECTEDVALUE ( Actualisation[Libellé] )
```

**[Titre du périmètre]** (Textes)

```dax
IF ( [Périmètre en vigueur] = "Tous les projets", "TOUS LES PROJETS", "PROJETS EN COURS" )
```

**[Sous-titre chefs]** (Textes)

```dax
IF ( [Périmètre en vigueur] = "Tous les projets", "sur l'ensemble des projets", "sur les projets en cours" )
```

**[Sous-titre charge identifiée]** (Textes)

```dax
"soit ≈ " & FORMAT ( [ETP-mois], "0.0", "fr-FR" ) & " ETP-mois"
```

**[Sous-titre capacité]** (Textes)

```dax
"Capacité projets : " & FORMAT ( [Capacité projets de l'équipe (h/mois)], "#,0", "fr-FR" ) & " h/mois"
```

**[Sous-titre surcharge]** (Textes)

```dax
VAR _m = [Premier mois en surcharge]
VAR _pc = [Pic en % de la capacité]
RETURN
    IF (
        NOT ( ISBLANK ( _m ) ),
        "Dès " & CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _m ), REMOVEFILTERS ( Mois ) ),
        IF ( ISBLANK ( _pc ), "", "Pic à " & FORMAT ( ROUND ( _pc * 100, 0 ), "0" ) & " % de la capacité" )
    )
```

**[Sous-titre pic (mois)]** (Textes)

```dax
VAR _m = [Mois du pic]
RETURN
    IF ( NOT ( ISBLANK ( _m ) ), "en " & CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _m ), REMOVEFILTERS ( Mois ) ) )
```

**[Sous-titre pic (affaires)]** (Textes)

```dax
VAR _e = ROUND ( [ETP au pic], 1 )
VAR _n = [Affaires actives au pic]
RETURN
    _n & " affaire(s) active(s)  ·  ≈ "
        & IF ( _e = INT ( _e ), FORMAT ( _e, "0" ), FORMAT ( _e, "0.0", "fr-FR" ) ) & " ETP"
```

**[Sous-titre hors délai]** (Textes)

```dax
"dont " & [Affaires hors délai] & " hors délai (besoin dépassé)"
```

**[Sous-titre écart]** (Textes)

```dax
[Affaires en écart] & " en retard de plus de 15 pts"
```

**[Sous-titre projets]** (Textes)

```dax
"réparties sur " & [Nombre de projets] & " projet(s)"
```

**[Sous-titre lignes terminées]** (Textes)

```dax
IF (
    [Périmètre en vigueur] = "Tous les projets",
    "dont " & [Lignes terminées] & " ligne(s) terminée(s)",
    [Lignes terminées] & " ligne(s) terminée(s) non comptée(s)"
)
```

**[Sous-titre avancement]** (Textes)

```dax
"Moyenne simple par affaire : " & FORMAT ( ROUND ( [Avancement moyen simple] * 100, 0 ), "0" ) & " %"
```

**[Période affichée]** (Textes)

```dax
VAR _d = MIN ( Mois[Mois] )
VAR _f = MAX ( Mois[Mois] )
RETURN
    IF ( NOT ( ISBLANK ( _d ) ), "Période " & CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _d ), REMOVEFILTERS ( Mois ) ) & " – " & CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _f ), REMOVEFILTERS ( Mois ) ) )
```

**[Texte hors période]** (Textes)

```dax
VAR _h = [Charge hors période ou sans date de début (h)]
RETURN
    IF ( _h > 0, "Hors période ou sans date de début : " & FORMAT ( _h, "#,0", "fr-FR" ) & " h",
        "Toute la charge est visible sur la période" )
```

**[Titre de la ressource]** (Textes)

```dax
IF (
    HASONEVALUE ( Ressources[Clé ressource] ),
    "CHARGE MENSUELLE DE " & UPPER ( SELECTEDVALUE ( Ressources[Ressource] ) ) & " (HEURES)",
    "CHARGE MENSUELLE (HEURES) — choisir une ressource"
)
```

**[Sous-titre graphique mensuel]** (Textes)

```dax
VAR _n = COUNTROWS ( VALUES ( Mois[Mois] ) )
VAR _m = MIN ( Mois[Mois] )
RETURN
    IF ( _n > 0, _n & " mois à partir de " & CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _m ), REMOVEFILTERS ( Mois ) ) )
```

**[Sous-titre plus gros reste]** (Textes)

```dax
"sur " & FORMAT ( [Plus gros reste à faire – total de l'affaire (h)] + 0, "#,0", "fr-FR" )
    & " h au total pour cette affaire"
```

**[Origine de la ressource]** (Textes)

```dax
SWITCH (
    SELECTEDVALUE ( Ressources[Type de ressource] ),
    "Interne", "Dessinateur interne (Data, colonne E)",
    "Externe", "Ressource externe (Data, colonne H)",
    "Hors liste", "Absent des listes Data : réglages par défaut",
    "Non affecté", "Lignes sans dessinateur ni ressource externe",
    BLANK ()
)
```

**[Sous-titre charge fixe de la ressource]** (Textes)

```dax
VAR _f = [Charge fixe de la ressource (h/mois)]
RETURN
    IF ( NOT ( ISBLANK ( _f ) ), "dont " & FORMAT ( _f, "#,0", "fr-FR" ) & " h de charge fixe" )
```

**[Sous-titre capacité de la ressource]** (Textes)

```dax
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
    )
```

**[Sous-titre affaires de la ressource]** (Textes)

```dax
"sur " & [Nombre d'affaires] & " affaire(s)"
```

**[Sous-titre charge sur la période]** (Textes)

```dax
"dont " & FORMAT ( [Charge répartie (h)], "#,0", "fr-FR" ) & " h sur la période"
```

**[Sous-titre lignes de la ressource]** (Textes)

```dax
VAR _n = [Nombre de lignes]
RETURN
    IF (
        _n = 0,
        "Aucune affaire affectée à cette ressource dans le périmètre.",
        _n & " ligne(s) dans le périmètre · heures sur la période = mois sélectionnés (0 sans date de début)"
    )
```

**[Titre info-bulle projet]** (Textes)

```dax
VAR _m = IF ( HASONEVALUE ( Mois[Mois] ), VALUES ( Mois[Mois] ) )
RETURN
    IF (
        ISBLANK ( _m ),
        "Projet survolé · période affichée",
        "Projet survolé · " & CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _m ), REMOVEFILTERS ( Mois ) )
    )
```

**[Mois le plus chargé de la ressource (texte)]** (Textes)

```dax
VAR _m = [Mois le plus chargé de la ressource]
RETURN
    IF (
        ISBLANK ( _m ) || [Charge répartie (h)] = 0,
        "—",
        CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _m ), REMOVEFILTERS ( Mois ) )
    )
```

**[Sous-titre pic de la ressource]** (Textes)

```dax
VAR _p = [Charge au pic de la ressource (h)]
RETURN
    IF ( NOT ( ISBLANK ( _p ) ), FORMAT ( _p, "#,0", "fr-FR" ) & " h au total (fixe + projets)" )
```

**[Sous-titre occupation au pic]** (Textes)

```dax
VAR _o = [Occupation au pic de la ressource (%)]
RETURN
    IF ( NOT ( ISBLANK ( _o ) ), FORMAT ( ROUND ( _o * 100, 0 ), "0" ) & " % de la capacité" )
```

**[Sous-titre dépassement cumulé]** (Textes)

```dax
VAR _d = [Dépassement cumulé (h)]
RETURN
    IF ( NOT ( ISBLANK ( _d ) ), "Dépassement cumulé : " & FORMAT ( _d, "#,0", "fr-FR" ) & " h" )
```

**[Sous-titre surcharge de la ressource]** (Textes)

```dax
VAR _n = [Mois en surcharge de la ressource]
VAR _m = [Premier mois en surcharge de la ressource]
RETURN
    IF (
        NOT ( ISBLANK ( _n ) ),
        IF ( _n = 0, "Aucun dépassement prévu", "Dès " & CALCULATE ( LOOKUPVALUE ( Mois[Mois (long)], Mois[Mois], _m ), REMOVEFILTERS ( Mois ) ) )
    )
```

**[Sous-titre heures disponibles]** (Textes)

```dax
VAR _d = [Heures encore disponibles (h)]
RETURN
    IF ( NOT ( ISBLANK ( _d ) ), FORMAT ( _d, "#,0", "fr-FR" ) & " h encore disponibles" )
```

**[Valeur Power BI]** (Contrôle)

```dax
VAR _code = SELECTEDVALUE ( 'Contrôle Excel'[Code] )
VAR _mois = SELECTEDVALUE ( 'Contrôle Excel'[Mois] )
RETURN
    CALCULATE (
        SWITCH (
            _code,
            "total", [Charge identifiée (h)],
            "nbAff", [Nombre d'affaires],
            "nbChefs", [Nombre de chefs de projet],
            "nbProj", [Nombre de projets],
            "nbTypes", [Nombre de types],
            "realise", [Réalisé (h)],
            "reste", [Reste à faire (h)],
            "avgSimple", [Avancement moyen simple],
            "avgPond", [Avancement moyen pondéré],
            "periode", [Charge répartie (h)],
            "hors", [Charge hors période ou sans date de début (h)],
            "pic", [Pic de charge (h)],
            "picActives", [Affaires actives au pic],
            "picETP", [ETP au pic],
            "totalETP", [ETP-mois],
            "nbTermines", [Lignes terminées],
            "picPart", [Part du pic dans la période],
            "capEquipe", [Capacité projets de l'équipe (h/mois)],
            "nbSurcharge", [Mois en surcharge],
            "picCap", [Pic en % de la capacité],
            "nbRetard", [Affaires en retard],
            "nbDepasse", [Affaires hors délai],
            "nbEcart", [Affaires en écart],
            "mois", CALCULATE ( [Charge répartie (h)], Mois[Mois] = _mois ),
            BLANK ()
        ),
        REMOVEFILTERS ( Lignes ),
        REMOVEFILTERS ( Affaires ),
        REMOVEFILTERS ( Ressources ),
        REMOVEFILTERS ( 'Charge mensuelle' ),
        REMOVEFILTERS ( Mois ),
        'Périmètre'[Périmètre] = "Comme dans Excel",
        'Répartition'[Répartition] = "Comme dans Excel",
        Mois[Période] = "12 mois du Dashboard Excel"
    )
```

**[Mois Power BI]** (Contrôle)

```dax
VAR _code = SELECTEDVALUE ( 'Contrôle Excel'[Code] )
RETURN
    CALCULATE (
        SWITCH (
            _code,
            "autoStart", MAX ( Réglages[1er mois automatique] ),
            "start", [Premier mois affiché],
            "picMois", [Mois du pic],
            "finPeriode", [Dernier mois affiché],
            "premSurcharge", [Premier mois en surcharge],
            BLANK ()
        ),
        REMOVEFILTERS ( Lignes ),
        REMOVEFILTERS ( Affaires ),
        REMOVEFILTERS ( Ressources ),
        REMOVEFILTERS ( 'Charge mensuelle' ),
        REMOVEFILTERS ( Mois ),
        'Périmètre'[Périmètre] = "Comme dans Excel",
        'Répartition'[Répartition] = "Comme dans Excel",
        Mois[Période] = "12 mois du Dashboard Excel"
    )
```

**[Texte Power BI]** (Contrôle)

```dax
VAR _code = SELECTEDVALUE ( 'Contrôle Excel'[Code] )
RETURN
    CALCULATE (
        SWITCH (
            _code,
            "periodeTxt", [Période affichée],
            "reglagesTxt", SELECTEDVALUE ( Réglages[Rappel des réglages (Excel)] ),
            BLANK ()
        ),
        REMOVEFILTERS ( Lignes ),
        REMOVEFILTERS ( Affaires ),
        REMOVEFILTERS ( Ressources ),
        REMOVEFILTERS ( 'Charge mensuelle' ),
        REMOVEFILTERS ( Mois ),
        'Périmètre'[Périmètre] = "Comme dans Excel",
        'Répartition'[Répartition] = "Comme dans Excel",
        Mois[Période] = "12 mois du Dashboard Excel"
    )
```

**[Valeur Excel]** (Contrôle)

```dax
SUM ( 'Contrôle Excel'[Valeur lue] )
```

**[Écart]** (Contrôle)

```dax
VAR _n = SELECTEDVALUE ( 'Contrôle Excel'[Nature] )
RETURN
    IF ( _n IN { "Heures", "Nombre", "Décimal", "Pourcentage" }, [Valeur Power BI] - [Valeur Excel] )
```

**[Excel (affiché)]** (Contrôle)

```dax
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
    )
```

**[Power BI (affiché)]** (Contrôle)

```dax
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
    )
```

**[Statut du contrôle]** (Contrôle)

```dax
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
    IF ( ISBLANK ( _n ), BLANK (), IF ( _ok, "✔", "✘" ) )
```

**[Contrôles en écart]** (Contrôle)

```dax
COUNTROWS ( FILTER ( VALUES ( 'Contrôle Excel'[Ordre] ), [Statut du contrôle] = "✘" ) ) + 0
```

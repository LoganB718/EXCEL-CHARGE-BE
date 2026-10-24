# Rapport Power BI « Plan de charge du bureau d'études » : mode d'emploi

Le rapport reprend les onglets **Dashboard** et **Dessinateurs** du classeur `Plan_de_charge_BE_Dashboard.xlsm`.
Toutes les données viennent de ce classeur, déposé dans Teams. Power BI ne modifie jamais le classeur : on continue
à saisir dans Excel (onglets **Charge** et **Data**), Power BI relit le fichier à chaque actualisation.

Contenu du dossier (à garder ensemble, sans rien renommer) :

| Élément | Rôle |
|---|---|
| `PlanDeCharge_BI.pbip` | le fichier à ouvrir dans Power BI Desktop |
| `PlanDeCharge_BI.Report/` | les pages du rapport |
| `PlanDeCharge_BI.SemanticModel/` | la lecture du classeur et les calculs |
| `.gitignore` | utile seulement si le dossier est suivi dans Git |

---

## 1. Préparer le classeur dans Teams

1. Déposer `Plan_de_charge_BE_Dashboard.xlsm` dans l'onglet **Fichiers** d'un canal Teams (par exemple le canal
   Général de l'équipe). Teams le range dans le site SharePoint de l'équipe.
2. Ne pas renommer le fichier, ni les onglets, ni les **noms définis** du classeur (`bi_Charge`, `bi_Data`,
   `bi_Indicateurs`, `bi_Mois`, `p_Perimetre`, `p_Repartition`, `p_PremierMois`, `p_HeuresETP`, `p_ChargeFixe`).
   Power BI lit les données par ces noms. Ne pas insérer de colonnes dans Charge ou Data.
3. Après chaque modification, **enregistrer le classeur** dans Excel. Power BI lit la dernière version
   enregistrée : une saisie non enregistrée n'apparaît pas. Si le fichier est ouvert dans Excel pour le Web,
   l'enregistrement est automatique.

## 2. Copier le chemin du fichier

1. Ouvrir le classeur dans **Excel (application de bureau)**, depuis Teams : « Ouvrir dans l'application de bureau ».
2. **Fichier > Informations > Copier le chemin d'accès**.
3. Coller le chemin dans le Bloc-notes et **supprimer `?web=1`** à la fin s'il est présent. Le chemin doit
   ressembler à :
   `https://votre-entreprise.sharepoint.com/sites/votre-equipe/Shared%20Documents/General/Plan_de_charge_BE_Dashboard.xlsm`

Ne conviennent **pas** : le lien obtenu par « Copier le lien » ou « Partager » (lien de partage du type
`…/:x:/s/…`), l'adresse de la barre du navigateur (`…/Doc.aspx?…`), un chemin `C:\…` ou celui d'un dossier
OneDrive synchronisé.

## 3. Ouvrir le rapport dans Power BI Desktop

1. Utiliser une version récente de Power BI Desktop (de juin 2025 ou plus récente ; Microsoft Store ou
   Aide > À propos pour vérifier). Ouvrir `PlanDeCharge_BI.pbip`.
   - Si Desktop signale un format inconnu ou n'affiche pas les pages : mettre Desktop à jour. Sur une version
     ancienne, cocher dans **Fichier > Options et paramètres > Options > Fonctionnalités en préversion**
     « Enregistrement de projet Power BI (.pbip) » et « Stocker les rapports au format de métadonnées amélioré
     (PBIR) », puis redémarrer Desktop.
2. Les visuels sont vides à la première ouverture (aucune donnée n'est livrée avec le rapport). Aller dans
   **Accueil > Transformer les données > Modifier les paramètres**, coller le chemin du § 2 dans **UrlFichier**,
   valider.
3. Cliquer sur **Appliquer les modifications** (ou **Accueil > Actualiser**).
4. À la demande de connexion « Accès au contenu web », choisir **Compte d'organisation**, cliquer sur
   **Se connecter** avec son compte Microsoft 365 professionnel, puis **Se connecter** en bas de la fenêtre.
   Niveau de confidentialité demandé : **Organisationnel**.
5. Vérifier la page **Contrôle** : le nombre de contrôles en écart doit être 0, sauf éventuellement les lignes
   « Affaires en retard » (et « 1er mois automatique ») si le classeur a été enregistré un autre jour que
   l'actualisation : ces chiffres dépendent de la date du jour (voir § 7).
6. Enregistrer (Ctrl + S). Le projet reste au format dossier ; « Fichier > Enregistrer sous » au format `.pbix`
   donne, si on le souhaite, un fichier unique.

## 4. Publier et actualiser automatiquement

1. Dans Desktop : **Accueil > Publier**, choisir l'espace de travail de l'équipe (pas « Mon espace de travail »
   si le rapport doit être partagé).
2. Dans le navigateur, sur `app.powerbi.com` : ouvrir l'espace de travail, survoler le **modèle sémantique**
   `PlanDeCharge_BI` > **…** > **Paramètres**.
3. **Informations d'identification de la source de données** > **Modifier les informations d'identification** :
   méthode **OAuth2**, niveau de confidentialité **Organisationnel**, **Se connecter** avec le compte
   professionnel.
4. **Actualisation planifiée** : activer, choisir la fréquence (par exemple chaque jour ouvré à 7 h et 13 h,
   fuseau « (UTC+01:00) Bruxelles, Copenhague, Madrid, Paris »), **Appliquer**. Aucune passerelle n'est nécessaire :
   le classeur est dans le cloud (SharePoint).
5. Pour une actualisation immédiate : **…** > **Actualiser maintenant**.

## 5. Consulter et partager

- Navigateur : `app.powerbi.com` > espace de travail > rapport `PlanDeCharge_BI`.
- Tablette : application **Microsoft Power BI** (même compte) ; l'affichage est complet.
- Téléphone : même application, mais les pages sont conçues pour un écran large (pas de disposition mobile) :
  tourner le téléphone à l'horizontale et zoomer (pincer) pour lire les tableaux et les petits textes.
- Partager au chef : bouton **Partager** du rapport (saisir son adresse), ou l'ajouter à l'espace de travail
  avec le rôle **Lecteur**. Selon la licence de l'entreprise, chaque lecteur peut avoir besoin d'une licence
  Power BI Pro (sauf espace de travail sur une capacité Premium / Fabric).
- Le chef voit les chiffres de la dernière actualisation ; la date de lecture est affichée en haut de chaque page
  (« Données lues le … »).

## 6. Utiliser le rapport

- **Tableau de bord** (onglet Dashboard) : cartes, indicateurs, charge par mois et par projet, liste des affaires.
- **Ressources** (onglet Dessinateurs) : en tête, la ressource choisie dans la liste « Ressource » et son
  détail ; plus bas, l'occupation de toute l'équipe mois par mois (vert < 90 %, orange 90-100 %, rouge > 100 %).
- **Réglages** : le rapport applique toujours les réglages enregistrés dans le Dashboard du classeur (périmètre,
  répartition mensuelle) et ses 12 mois. Ils ne sont pas affichés en segments, pour que le rapport reste lisible
  par quelqu'un qui ne connaît pas le classeur. Pour comparer sans toucher au classeur : ouvrir le volet
  **Filtres** (à droite, replié) et changer « Périmètre », « Répartition » (filtres de tout le rapport) ou
  « Période » (filtre de la page : cocher aussi « Avant » / « Après » pour voir d'autres mois).
- **Chef de projet, Projet, Type d'affaire** (en haut du tableau de bord) filtrent la page, et aussi les pages PDF.
- Cliquer sur une barre, une part ou une ligne filtre les autres visuels ; cliquer à nouveau pour annuler.
  Les numéros 1 à 9 des couleurs renvoient au tableau des projets ; survoler une colonne affiche le projet.
- **Fiche affaire** : clic droit sur une affaire (tableau des affaires ou lignes d'une ressource) >
  **Extraire** > **Fiche affaire**, ou sélectionner la ligne puis le bouton « Ouvrir la fiche de l'affaire ».
  Dans Power BI Desktop, l'onglet « Fiche affaire » reste visible en bas : ouverte par cet onglet, sans
  extraction, la fiche n'affiche qu'une consigne (aucune affaire choisie).
  Bouton **‹ Retour** pour revenir.
- **Boutons** (« Ressources › », « Contrôle › », « Ouvrir la fiche de l'affaire », « ‹ Retour ») : dans
  **Power BI Desktop**, maintenir **Ctrl** en cliquant (un clic simple ne fait que sélectionner le bouton) ; dans
  le navigateur et l'application, un clic suffit.
- **Pages PDF** (« PDF – Tableau de bord (1/2) et (2/2) », « PDF – Ressources (1/2) et (2/2) ») : pages au format
  A4 paysage faites pour l'impression. Elles reprennent la sélection faite à l'écran (chef, projet, type sur le
  tableau de bord ; ressource choisie sur la page Ressources) et l'indiquent dans leur bandeau. La liste des
  affaires du PDF s'arrête aux **30 premières** (par charge décroissante) ; le nombre total est rappelé en bas.
  - Dans le navigateur ou l'application : ouvrir la page PDF voulue, **Exporter > PDF**, cocher
    « Exporter uniquement la page actuelle » (une page à la fois), ou exporter tout le rapport puis ne garder que
    les pages PDF.
  - Dans Power BI Desktop : **Fichier > Exporter > Exporter au format PDF** exporte toutes les pages visibles ;
    garder les pages « PDF – … » (ou masquer temporairement les autres pages avant l'export).
- **Contrôle** : compare Power BI (réglages du classeur, 12 mois du Dashboard) aux valeurs de l'onglet Calculs.
  ✔ = identique, ✘ = écart.

## 7. En cas de problème

| Message ou symptôme | Que faire |
|---|---|
| « Accès refusé », « Informations d'identification non valides », erreur 401 / 403 | Desktop : **Fichier > Options et paramètres > Paramètres de la source de données**, sélectionner le site SharePoint, **Modifier les autorisations** > **Modifier** > Compte d'organisation > Se connecter (ou **Effacer les autorisations** puis actualiser). Service : refaire le § 4.3. Vérifier qu'on a bien accès au fichier dans Teams. |
| Erreur 404, « fichier introuvable », « DataSource.Error » | Le chemin `UrlFichier` est faux : refaire le § 2 (pas de lien de partage, pas de `?web=1`). Le fichier a peut-être été renommé ou déplacé. |
| « Le classeur ne contient pas le nom défini bi_… » | Le classeur déposé est une ancienne version, ou un nom défini a été supprimé : redéposer le classeur produit par le générateur à jour (`scripts/build_dashboard.py`), puis actualiser. |
| Valeurs différentes de l'Excel | Vérifier dans le volet Filtres que Périmètre et Répartition sont sur « Comme dans Excel » et Période sur « 12 mois du Dashboard Excel ». Enregistrer le classeur dans Excel, attendre la synchronisation, puis actualiser. Regarder la page **Contrôle**. |
| Affaires en retard différentes de l'Excel | Normal si le classeur a été enregistré un autre jour : Excel compte les retards à la date de son dernier enregistrement, Power BI à la date de l'actualisation (heure de Paris). |
| Le rapport ne bouge pas après une saisie | Le classeur n'a pas été enregistré, ou il n'y a pas encore eu d'actualisation : enregistrer, puis **Actualiser maintenant**. |
| Visuels vides à l'ouverture dans Desktop | Normal la première fois : renseigner `UrlFichier` et actualiser (§ 3). |
| Un bouton ne réagit pas dans Desktop | Maintenir **Ctrl** en cliquant (§ 6). |

**Limites connues du classeur** (écarts signalés par la page Contrôle, à corriger dans la saisie) :

- un nom de chef, de projet, de type ou de ressource contenant `*`, `?` ou `~` : Excel les traite comme des
  caractères génériques et regroupe ce nom avec d'autres (« VEG* » compté avec « VEGA ») ; Power BI les compare
  exactement. Éviter ces caractères dans les noms.
- une date de début saisie avec une heure (ex. « 30/11/2026 08:00 », fréquent après un copier-coller depuis un
  planning) : Excel peut perdre les heures du dernier jour du mois, voire toute la ligne ; Power BI ne garde que
  la date. Saisir la date seule (format jj/mm/aaaa).

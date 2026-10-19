Attribute VB_Name = "PlanDeCharge"
Option Explicit

' Macros des boutons (Dashboard, Dessinateurs, Charge, Data) et masquage des lignes vides.
' Les cellules utilisées sont repérées par des noms définis (Formules > Gestionnaire de noms) :
'   aide_<onglet>                 : Oui / Non, affichage des textes d'aide (onglet Calculs)
'   edite_Dashboard, edite_Dessinateurs : date d'édition du dernier PDF de l'onglet
'   lignes_masquables_<onglet>_1, _2... : cellules repères (une zone par nom) ; leur ligne est
'       masquée quand elles sont vides (lignes inutilisées des listes, aide masquée, alertes absentes)
'   macros_actives                : Oui quand les macros tournent (le voyant rouge des onglets disparaît)
'   lignes_pdf_<onglet>           : lignes non imprimées dans le PDF (barres de réglages et de choix)
' Les modules ThisWorkbook (ouverture, enregistrement) et des feuilles Dashboard et Dessinateurs
' (activation de l'onglet, modification d'un réglage) appellent MacrosActives, AjusterLignes, etc.

Private Const TITRE_PDF As String = "Éditer PDF"
Private Const TITRE_AIDE As String = "Afficher / masquer l'aide"

Public Sub Dashboard_Aide()
    BasculerAide "Dashboard"
End Sub

Public Sub Dashboard_PDF()
    EditerPDF "Dashboard"
End Sub

Public Sub Dessinateurs_Aide()
    BasculerAide "Dessinateurs"
End Sub

Public Sub Dessinateurs_PDF()
    EditerPDF "Dessinateurs"
End Sub

Public Sub Charge_Aide()
    BasculerAide "Charge"
End Sub

Public Sub Data_Aide()
    BasculerAide "Data"
End Sub

' Cellule désignée par un nom défini du classeur, ou Nothing (après un message clair)
' si le nom a été supprimé ou ne désigne plus une cellule (#REF!).
Private Function CelluleNommee(ByVal nom As String, ByVal titre As String) As Range
    On Error Resume Next
    Set CelluleNommee = ThisWorkbook.Names(nom).RefersToRange
    On Error GoTo 0
    If CelluleNommee Is Nothing Then
        MsgBox "Le nom défini « " & nom & " » est introuvable ou ne désigne plus une cellule." & _
               vbCrLf & "Vérifier Formules > Gestionnaire de noms.", vbExclamation, titre
    End If
End Function

' Affiche ou masque les textes d'aide de l'onglet.
Private Sub BasculerAide(ByVal cle As String)
    Dim c As Range, actuel As String
    Set c = CelluleNommee("aide_" & cle, TITRE_AIDE)
    If c Is Nothing Then Exit Sub

    On Error GoTo Erreur
    If Not IsError(c.Value) Then actuel = Trim$(CStr(c.Value))
    ' comme les formules d'aide (aide = "Oui"), sans tenir compte de la casse ; case vide -> Oui
    If StrComp(actuel, "Oui", vbTextCompare) = 0 Then c.Value = "Non" Else c.Value = "Oui"
    ' En calcul manuel, les formules d'aide ne se mettraient à jour qu'avec F9.
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculate
    AjusterLignes cle          ' lignes d'aide masquées ou réaffichées
    Exit Sub

Erreur:
    MsgBox "Impossible de modifier la cellule " & c.Address(False, False) & " de l'onglet " & _
           c.Worksheet.Name & " (feuille protégée ?) :" & vbCrLf & Err.Description, _
           vbExclamation, TITRE_AIDE
End Sub

' Masque les lignes dont la cellule repère (noms lignes_masquables_<onglet>_1, _2...) est vide et
' réaffiche les autres ; toutAfficher = True réaffiche tout (avant un enregistrement). En calcul
' manuel, le classeur est d'abord recalculé pour ne jamais masquer une ligne d'après des valeurs
' périmées. Sans effet si les noms sont absents ; une feuille protégée reste telle quelle.
' (Pas d'appel depuis Worksheet_Calculate : masquer des lignes relance un calcul, d'où des boucles.)
Public Sub AjusterLignes(ByVal cle As String, Optional ByVal toutAfficher As Boolean = False)
    Dim zone As Range, c As Range, vide As Boolean, i As Integer
    Dim aMasquer As Range, aAfficher As Range

    On Error Resume Next
    If Not toutAfficher And Application.Calculation <> xlCalculationAutomatic Then Application.Calculate
    For i = 1 To 20
        Set zone = Nothing
        Set zone = ThisWorkbook.Names("lignes_masquables_" & cle & "_" & i).RefersToRange
        If zone Is Nothing Then Exit For
        For Each c In zone.Cells
            vide = False
            If Not toutAfficher And Not IsError(c.Value) Then vide = (Len(CStr(c.Value)) = 0)
            If c.EntireRow.Hidden <> vide Then
                If vide Then
                    If aMasquer Is Nothing Then Set aMasquer = c Else Set aMasquer = Union(aMasquer, c)
                Else
                    If aAfficher Is Nothing Then Set aAfficher = c Else Set aAfficher = Union(aAfficher, c)
                End If
            End If
        Next c
    Next i
    ' deux opérations groupées : chaque masquage relance un calcul du classeur
    If Not aAfficher Is Nothing Then aAfficher.EntireRow.Hidden = False
    If Not aMasquer Is Nothing Then aMasquer.EntireRow.Hidden = True
End Sub

' Ajuste les lignes des quatre onglets.
Public Sub AjusterTout(Optional ByVal toutAfficher As Boolean = False)
    AjusterLignes "Dashboard", toutAfficher
    AjusterLignes "Dessinateurs", toutAfficher
    AjusterLignes "Charge", toutAfficher
    AjusterLignes "Data", toutAfficher
End Sub

' État des macros (nom macros_actives) : Oui fait disparaître le voyant rouge des onglets.
Private Sub SignalerMacros(ByVal actives As Boolean)
    On Error Resume Next
    ThisWorkbook.Names("macros_actives").RefersToRange.Value = IIf(actives, "Oui", "Non")
End Sub

' À l'ouverture (Workbook_Open) : voyant éteint, lignes vides masquées.
Public Sub MacrosActives()
    SignalerMacros True
    AjusterTout
End Sub

' Avant un enregistrement (Workbook_BeforeSave) : toutes les lignes réaffichées et voyant rallumé,
' pour qu'un classeur ouvert un jour sans macros ne cache aucune ligne remplie et le signale.
' Rien n'est fait avec l'enregistrement automatique (OneDrive / SharePoint), qui enregistre toutes
' les quelques secondes.
Public Sub PreparerEnregistrement()
    If EnregistrementAutomatique() Then Exit Sub
    SignalerMacros False
    AjusterTout True
End Sub

' Après l'enregistrement (Workbook_AfterSave) : voyant éteint, lignes vides de nouveau masquées.
Public Sub TerminerEnregistrement()
    SignalerMacros True
    AjusterTout
End Sub

' Enregistrement automatique actif ? (propriété absente des anciennes versions d'Excel : liaison
' tardive, et Faux en cas d'erreur)
Private Function EnregistrementAutomatique() As Boolean
    Dim classeur As Object
    Set classeur = ThisWorkbook
    On Error Resume Next
    EnregistrementAutomatique = classeur.AutoSaveOn
End Function

' Plage désignée par un nom défini, ou Nothing si le nom est absent (sans message).
Private Function PlageFacultative(ByVal nom As String) As Range
    On Error Resume Next
    Set PlageFacultative = ThisWorkbook.Names(nom).RefersToRange
End Function

' Exporte la zone d'impression de l'onglet en PDF, sans les textes d'aide ni les lignes vides,
' et inscrit la date du jour dans « Édité le ».
Private Sub EditerPDF(ByVal cle As String)
    Dim ws As Worksheet, cAide As Range, cEdite As Range, horsPDF As Range
    Dim ancienneAide As Variant, ancienneDate As Variant, fichier As Variant
    Dim nomDefaut As String, dossier As String, message As String

    Set cEdite = CelluleNommee("edite_" & cle, TITRE_PDF)
    If cEdite Is Nothing Then Exit Sub
    Set cAide = CelluleNommee("aide_" & cle, TITRE_PDF)
    If cAide Is Nothing Then Exit Sub
    Set ws = cEdite.Worksheet
    Set horsPDF = PlageFacultative("lignes_pdf_" & cle)

    ' Dossier proposé : celui du classeur s'il est enregistré sur un disque (local ou réseau).
    ' Classeur jamais enregistré : Path = "" ; ouvert depuis OneDrive / SharePoint : Path est
    ' une URL (https://...) inutilisable comme chemin de fichier. Dans ces deux cas, seul le
    ' nom est proposé et la boîte de dialogue s'ouvre sur son dossier courant.
    dossier = ThisWorkbook.Path
    If LCase$(Left$(dossier, 5)) = "http:" Or LCase$(Left$(dossier, 6)) = "https:" Then dossier = ""
    nomDefaut = "Plan de charge - " & cle & " - " & Format$(Date, "yyyy-mm-dd") & ".pdf"
    If Len(dossier) > 0 Then nomDefaut = dossier & Application.PathSeparator & nomDefaut
    fichier = Application.GetSaveAsFilename(InitialFileName:=nomDefaut, _
                                            FileFilter:="Fichier PDF (*.pdf), *.pdf", _
                                            Title:=TITRE_PDF)
    If VarType(fichier) = vbBoolean Then Exit Sub   ' enregistrement annulé
    If LCase$(Right$(fichier, 4)) <> ".pdf" Then fichier = fichier & ".pdf"

    ancienneAide = cAide.Value
    ancienneDate = cEdite.Value
    On Error GoTo Erreur
    Application.ScreenUpdating = False
    cAide.Value = "Non"        ' le PDF ne contient pas les textes d'aide
    cEdite.Value = Date        ' « Édité le » = date de l'édition du PDF
    Application.Calculate
    AjusterLignes cle          ' pas de lignes vides (ni d'aide) dans le PDF
    If Not horsPDF Is Nothing Then horsPDF.EntireRow.Hidden = True
    ws.Select                  ' seulement cet onglet, même si des onglets étaient groupés
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fichier, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    If Not horsPDF Is Nothing Then horsPDF.EntireRow.Hidden = False
    cAide.Value = ancienneAide
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculate
    AjusterLignes cle
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    message = Err.Description
    Resume Restaurer           ' quitte le gestionnaire d'erreur avant de restaurer

Restaurer:
    On Error Resume Next       ' une restauration impossible ne doit pas lever une 2e erreur
    If Not horsPDF Is Nothing Then horsPDF.EntireRow.Hidden = False
    cAide.Value = ancienneAide
    cEdite.Value = ancienneDate
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculate
    AjusterLignes cle
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Le PDF n'a pas pu être créé :" & vbCrLf & message & vbCrLf & vbCrLf & _
           "Fichier : " & fichier & vbCrLf & _
           "(fichier ouvert dans un lecteur PDF, dossier en lecture seule ou feuille protégée ?)", _
           vbExclamation, TITRE_PDF
End Sub

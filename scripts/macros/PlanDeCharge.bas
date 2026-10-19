Attribute VB_Name = "PlanDeCharge"
Option Explicit

' Macros des boutons des onglets Dashboard et Dessinateurs.
' Les cellules utilisées sont repérées par des noms définis (Formules > Gestionnaire de noms) :
'   aide_Dashboard, aide_Dessinateurs    : Oui / Non, affichage des textes d'aide de l'onglet
'   edite_Dashboard, edite_Dessinateurs  : date d'édition du dernier PDF de l'onglet
'   pdf_lignes_Dashboard, pdf_lignes_Dessinateurs : cellules dont la ligne n'est pas imprimée si elles sont vides

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
    If Application.Calculation <> xlCalculationAutomatic Then c.Worksheet.Calculate
    Exit Sub

Erreur:
    MsgBox "Impossible de modifier la cellule " & c.Address(False, False) & " de l'onglet " & _
           c.Worksheet.Name & " (feuille protégée ?) :" & vbCrLf & Err.Description, _
           vbExclamation, TITRE_AIDE
End Sub

' Masque les lignes des cellules vides du nom pdf_lignes_<onglet> (lignes inutilisées des listes,
' aide masquée, alertes absentes) pour éviter des pages blanches. Renvoie les cellules masquées
' (Nothing si aucune) pour pouvoir réafficher exactement ces lignes. Nom absent : rien n'est masqué.
Private Function MasquerLignesVides(ByVal cle As String) As Range
    Dim zone As Range, bloc As Range, c As Range, masquees As Range
    On Error Resume Next
    Set zone = ThisWorkbook.Names("pdf_lignes_" & cle).RefersToRange
    On Error GoTo 0
    If zone Is Nothing Then Exit Function
    For Each bloc In zone.Areas
        For Each c In bloc.Cells
            If Not c.EntireRow.Hidden And Not IsError(c.Value) Then
                If Len(CStr(c.Value)) = 0 Then
                    If masquees Is Nothing Then Set masquees = c Else Set masquees = Union(masquees, c)
                End If
            End If
        Next c
    Next bloc
    If Not masquees Is Nothing Then masquees.EntireRow.Hidden = True
    Set MasquerLignesVides = masquees
End Function

' Exporte la zone d'impression de l'onglet en PDF, sans les textes d'aide ni les lignes vides,
' et inscrit la date du jour dans « Édité le ».
Private Sub EditerPDF(ByVal cle As String)
    Dim ws As Worksheet, cAide As Range, cEdite As Range, masquees As Range
    Dim ancienneAide As Variant, ancienneDate As Variant, fichier As Variant
    Dim nomDefaut As String, dossier As String, message As String

    Set cEdite = CelluleNommee("edite_" & cle, TITRE_PDF)
    If cEdite Is Nothing Then Exit Sub
    Set cAide = CelluleNommee("aide_" & cle, TITRE_PDF)
    If cAide Is Nothing Then Exit Sub
    Set ws = cEdite.Worksheet

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
    Set masquees = MasquerLignesVides(cle)
    ws.Select                  ' seulement cet onglet, même si des onglets étaient groupés
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fichier, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    If Not masquees Is Nothing Then masquees.EntireRow.Hidden = False
    cAide.Value = ancienneAide
    If Application.Calculation <> xlCalculationAutomatic Then ws.Calculate
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    message = Err.Description
    Resume Restaurer           ' quitte le gestionnaire d'erreur avant de restaurer

Restaurer:
    On Error Resume Next       ' une restauration impossible ne doit pas lever une 2e erreur
    If Not masquees Is Nothing Then masquees.EntireRow.Hidden = False
    cAide.Value = ancienneAide
    cEdite.Value = ancienneDate
    If Application.Calculation <> xlCalculationAutomatic Then ws.Calculate
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Le PDF n'a pas pu être créé :" & vbCrLf & message & vbCrLf & vbCrLf & _
           "Fichier : " & fichier & vbCrLf & _
           "(fichier ouvert dans un lecteur PDF, dossier en lecture seule ou feuille protégée ?)", _
           vbExclamation, TITRE_PDF
End Sub

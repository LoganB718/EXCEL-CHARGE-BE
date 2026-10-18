Attribute VB_Name = "PlanDeCharge"
Option Explicit

' Macros des boutons des onglets Dashboard et Dessinateurs.
' Les cellules utilisées sont repérées par des noms définis (Formules > Gestionnaire de noms) :
'   aide_Dashboard, aide_Dessinateurs    : Oui / Non, affichage des textes d'aide de l'onglet
'   edite_Dashboard, edite_Dessinateurs  : date d'édition du dernier PDF de l'onglet

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

' Affiche ou masque les textes d'aide de l'onglet.
Private Sub BasculerAide(ByVal cle As String)
    Dim c As Range
    Set c = ThisWorkbook.Names("aide_" & cle).RefersToRange
    If CStr(c.Value) = "Non" Then
        c.Value = "Oui"
    Else
        c.Value = "Non"
    End If
End Sub

' Exporte la zone d'impression de l'onglet en PDF, sans les textes d'aide,
' et inscrit la date du jour dans « Édité le ».
Private Sub EditerPDF(ByVal cle As String)
    Dim ws As Worksheet, cAide As Range, cDate As Range
    Dim ancienneAide As Variant, ancienneDate As Variant, fichier As Variant, nomDefaut As String

    Set cDate = ThisWorkbook.Names("edite_" & cle).RefersToRange
    Set cAide = ThisWorkbook.Names("aide_" & cle).RefersToRange
    Set ws = cDate.Worksheet

    nomDefaut = "Plan de charge - " & cle & " - " & Format(Date, "yyyy-mm-dd") & ".pdf"
    If Len(ThisWorkbook.Path) > 0 Then nomDefaut = ThisWorkbook.Path & Application.PathSeparator & nomDefaut
    fichier = Application.GetSaveAsFilename(InitialFileName:=nomDefaut, _
                                            FileFilter:="Fichier PDF (*.pdf), *.pdf", _
                                            Title:="Éditer le PDF")
    If VarType(fichier) = vbBoolean Then Exit Sub   ' enregistrement annulé

    ancienneAide = cAide.Value
    ancienneDate = cDate.Value
    On Error GoTo Erreur
    Application.ScreenUpdating = False
    cAide.Value = "Non"        ' le PDF ne contient pas les textes d'aide
    cDate.Value = Date         ' « Édité le » = date de l'édition du PDF
    Application.Calculate
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fichier, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    cAide.Value = ancienneAide
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    cAide.Value = ancienneAide
    cDate.Value = ancienneDate
    Application.ScreenUpdating = True
    MsgBox "Le PDF n'a pas pu être créé :" & vbCrLf & Err.Description, vbExclamation, "Éditer le PDF"
End Sub

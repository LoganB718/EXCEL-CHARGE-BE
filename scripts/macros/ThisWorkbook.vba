' Code du module du classeur (ThisWorkbook).
' Les procédures appelées sont dans le module PlanDeCharge.
Option Explicit

' À l'ouverture : voyant « macros désactivées » éteint, lignes vides masquées.
Private Sub Workbook_Open()
    MacrosActives
End Sub

' Le classeur est enregistré avec toutes les lignes visibles : ouvert un jour sans macros, il ne
' cache aucune ligne remplie. Les lignes vides sont masquées de nouveau juste après.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    PreparerEnregistrement
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    TerminerEnregistrement
    If Success Then Me.Saved = True   ' le masquage ne compte pas comme une modification à enregistrer
End Sub

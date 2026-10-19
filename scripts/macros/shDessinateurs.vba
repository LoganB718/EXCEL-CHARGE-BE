' Code du module de la feuille Dessinateurs (codeName shDessinateurs).
' AjusterLignes est dans le module PlanDeCharge.
Option Explicit

' Lignes vides des deux tableaux et de l'aide masquées à l'arrivée sur l'onglet
' et après chaque changement de ressource.
Private Sub Worksheet_Activate()
    AjusterLignes "Dessinateurs"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    AjusterLignes "Dessinateurs"
End Sub

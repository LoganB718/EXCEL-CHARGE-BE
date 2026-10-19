' Code du module de la feuille Dashboard (codeName shDashboard).
' AjusterLignes est dans le module PlanDeCharge.
Option Explicit

' Lignes vides du tableau des affaires et de l'aide masquées à l'arrivée sur l'onglet
' et après chaque modification d'un réglage (périmètre, répartition, 1er mois...).
Private Sub Worksheet_Activate()
    AjusterLignes "Dashboard"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    AjusterLignes "Dashboard"
End Sub

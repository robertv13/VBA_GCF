Attribute VB_Name = "modGL_DataService"
Option Explicit

Function Fn_ConstruirePlanComptable() As Object '2025-06-01 @ 08:20

    Dim dictComptes As Object
    Set dictComptes = CreateObject("Scripting.Dictionary")
    
    Dim dataPlan As Variant
    dataPlan = Fn_PlanComptableTableau2D(2)
    
    Dim i As Long
    For i = 1 To UBound(dataPlan, 1)
        dictComptes(dataPlan(i, 1)) = dataPlan(i, 2)
    Next i
    
    Set Fn_ConstruirePlanComptable = dictComptes
    
End Function

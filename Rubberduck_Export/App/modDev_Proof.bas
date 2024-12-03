Attribute VB_Name = "modDev_Proof"
Option Explicit

Sub Get_Hours_Billed_By_Invoice()

    '1. Obtenir toutes les charges facturées
    Dim ws As Worksheet: Set ws = wshTEC_Local
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim s As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 3 To lastUsedRow
        If ws.Cells(i, 16).value >= "24-24609" Then
            s = ws.Cells(i, 16).value & "-" & Format$(ws.Cells(i, 2), "00")
            'Ajoute au sommaire par facture / par Prof_ID
            If dict.Exists(s) Then
                dict(s) = dict(s) + ws.Cells(i, 8).value
            Else
                dict.Add s, ws.Cells(i, 8).value
            End If
        End If
    Next i

    'Création/Initialisation d'une feuille
    Dim feuilleNom As String
    feuilleNom = "X_Heures_Facturées_Par_Facture"
    Call Erase_And_Create_Worksheet(feuilleNom)
    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Sheets(feuilleNom)
    wsOutput.Cells(1, 1).value = "NuméroFact"
    wsOutput.Cells(1, 2).value = "Prof"
    wsOutput.Cells(1, 3).value = "HeuresFact"
    
    Dim key As Variant
    Dim Prof As String, profID As Long, saveInvNo As String
    Dim t As Currency, st As Currency
    Dim r As Long: r = 1
    If dict.count <> 0 Then
        For Each key In Fn_Sort_Dictionary_By_Keys(dict, False) 'Sort dictionary by hours in ascending order
            profID = Mid(key, 10, Len(key) - 2)
            Prof = Fn_Get_Prof_From_ProfID(profID)
            If Left(key, 8) <> saveInvNo Then
                Call Sub_Total_Hours(wsOutput, saveInvNo, r, st)
            End If
            t = t + dict(key)
            st = st + dict(key)
            saveInvNo = Left(key, 8)
            r = r + 1
            wsOutput.Cells(r, 1).value = Left(key, 8)
            wsOutput.Cells(r, 2).value = Prof
            wsOutput.Cells(r, 3).value = dict(key)
            wsOutput.Cells(r, 3).NumberFormat = "##0.00"
        Next key
        Call Sub_Total_Hours(wsOutput, saveInvNo, r, st)
        
        r = r + 2
        wsOutput.Cells(r, 1).value = "* TOTAL *"
        wsOutput.Cells(r, 4).value = t
        wsOutput.Cells(r, 4).NumberFormat = "##0.00"
        wsOutput.Cells(r, 4).Font.Bold = True
        
    End If
    
End Sub

Sub Sub_Total_Hours(ws As Worksheet, saveInv As String, ByRef r As Long, ByRef st As Currency)

    If saveInv <> "" Then
        With ws.Cells(r, 3).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        r = r + 1
        ws.Cells(r, 4).value = st
        ws.Cells(r, 4).NumberFormat = "##0.00"
        st = 0
    End If

End Sub

Attribute VB_Name = "Functions"
Option Explicit

Function GetID_FromInitials(i As String)

    Dim cell As Range
    
    For Each cell In wshAdmin.Range("Prof_Initiales")
        If cell.Value2 = i Then
            GetID_FromInitials = cell.Offset(0, 1).value
        End If
    Next cell

End Function

Function RMV_TEST_GetID_FromClientName(ClientNom As String)

    Debug.Print "Nom de client à trouver - '" & ClientNom & "'"
    Dim BenchMark As Double
    BenchMark = Timer

    Dim LastRow As Long
    LastRow = wshClientDB.Range("A99999").End(xlUp).Row
    
    Dim i As Long
    For i = 1 To LastRow
        If wshClientDB.Cells(i, 2) = ClientNom Then
            Debug.Print "ID du client - '" & wshClientDB.Cells(i, 1).value & "'"
            RMV_TEST_GetID_FromClientName = wshClientDB.Cells(i, 1).value
        End If
    Next i

End Function

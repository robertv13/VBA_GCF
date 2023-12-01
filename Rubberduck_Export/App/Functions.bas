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

Function RMV_TEST_GetID_FromClientName(i As String)

    Dim cell As Range
    
    For Each cell In wshClientDB.Range("Client_Name")
        If cell.Value2 = i Then
            RMV_TEST_GetID_FromClientName = cell.Offset(0, -1).value
        End If
    Next cell

End Function

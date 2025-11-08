Attribute VB_Name = "modDEV_Securite"
Option Explicit

Public Function GetInitialesObligatoiresFromADMIN(ByVal utilisateurWindows As String) As String '2025-05-31 @ 16:08

    Dim initialesPermises As String

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = wsdADMIN
    Dim tblRange As Range
    Set tblRange = ws.ListObjects("tbl_WindowsUser_Initials").DataBodyRange

    Dim nomWindows As String
    Dim i As Long
    For i = 1 To tblRange.Rows.count
        nomWindows = Trim(tblRange.Cells(i, 1).Value)
        initialesPermises = Trim(tblRange.Cells(i, 3).Value)

        If nomWindows = vbNullString Then Exit For 'Arrêter à la première ligne vide
        If nomWindows = utilisateurWindows Then
            If initialesPermises = vbNullString Then
                GetInitialesObligatoiresFromADMIN = vbNullString 'Aucune restriction
            Else
                GetInitialesObligatoiresFromADMIN = initialesPermises
            End If
            Exit Function
        End If
    Next i

    GetInitialesObligatoiresFromADMIN = "INVALID"
    
    Exit Function

ErrHandler:
    GetInitialesObligatoiresFromADMIN = "INVALID"
    
End Function


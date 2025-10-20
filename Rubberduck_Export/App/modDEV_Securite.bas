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

Public Function GetInitialesAutorises(ByVal userName As String) As String '2025-05-31 @ 15:41

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    'Associer les utilisateurs Windows à leurs initiales autorisées
    dict.Add "vgervais", "VG"
    dict.Add "Vlad_Portable", "VG"
    dict.Add "User", "ML"
    dict.Add "Annie", "AR"
    dict.Add "Oli_Portable", "OB"
    dict.Add "MARIE_FRANCE", "MFP"
    
    'Les utilisateurs avec toutes les autorisations retourne ""
    Select Case userName
        Case "Guillaume", "GuillaumeCharron", "gchar", "RobertMV", "robertmv"
            GetInitialesAutorises = vbNullString 'Toutes les initiales sont permises
        Case Else
            If dict.Exists(userName) Then
                GetInitialesAutorises = dict(userName)
            Else
                GetInitialesAutorises = "INVALID"
            End If
    End Select
    
End Function


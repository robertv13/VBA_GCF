Attribute VB_Name = "modDataValidation"
Option Explicit

Public Sub AfficherMessageDateInvalide(location As String) '2024-06-13 @ 12:40

    MsgBox "La date saisie ne peut être acceptée tel qu'elle est entrée." & vbNewLine & vbNewLine & _
           "Elle doit être obligatoirement de format:" & vbNewLine & _
           "     'j', jj', " & vbNewLine & _
           "     'jj-mm', 'jj/mm' ou " & vbNewLine & _
           "     'j-m-aa', 'j-m-aaaa', 'jj-mm-aaaa'" & vbNewLine & vbNewLine & _
           "Veuillez saisir la date de nouveau SVP", _
           vbCritical, _
           "La date saisie est INVALIDE - " & location

End Sub



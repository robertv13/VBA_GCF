Attribute VB_Name = "modInactiviteUserForm"
Option Explicit

'Public colWrappers As Collection

Public Sub ConnectFormControls(frm As Object) '2025-05-30 @ 13:22

    Set colWrappers = New Collection
    ConnectControlsRecursive frm.Controls
    
End Sub

Private Sub ConnectControlsRecursive(ctrls As MSForms.Controls) '2025-05-30 @ 13:22

    Dim ctrl As MSForms.Control
    For Each ctrl In ctrls
        If TypeName(ctrl) <> "Label" Then
'            Debug.Print "Contrôle '" & ctrl.Name & "' de type '" & TypeName(ctrl)
            Select Case TypeName(ctrl)
                Case "Frame", "TabStrip"
                    ConnectControlsRecursive ctrl.Controls
                Case "MultiPage"
                    Dim i As Integer
                    For i = 0 To ctrl.Pages.count - 1
                        ConnectControlsRecursive ctrl.Pages(i).Controls
                    Next i
                Case Else
                    On Error Resume Next
                    Dim wrapper As New clsControlWrapper
                    Set wrapper.ctrl = ctrl
                    colWrappers.Add wrapper, ctrl.Name
                    On Error GoTo 0
            End Select
        End If
    Next ctrl
    
End Sub



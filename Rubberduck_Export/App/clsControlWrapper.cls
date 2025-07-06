VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'clsControlWrapper - 2025-05-30 @ 13:24

Public WithEvents ctrl As MSForms.Control
Attribute ctrl.VB_VarHelpID = -1

Private Sub ctrl_Change()

    Call RafraichirActivite(ctrl.Name & "-" & TypeName(ctrl) & " [" & Fn_NomFeuilleActive() & "]")
    
End Sub

Private Sub ctrl_Click()

    Call RafraichirActivite(ctrl.Name & "-" & TypeName(ctrl) & " [" & Fn_NomFeuilleActive() & "]")
    
End Sub

Private Sub ctrl_Enter()

    Call RafraichirActivite(ctrl.Name & "-" & TypeName(ctrl) & " [" & Fn_NomFeuilleActive() & "]")
    
End Sub

Private Sub ctrl_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Call RafraichirActivite(ctrl.Name & "-" & TypeName(ctrl) & " [" & Fn_NomFeuilleActive() & "]")
    
End Sub

Private Sub ctrl_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)

    Call RafraichirActivite(ctrl.Name & "-" & TypeName(ctrl) & " [" & Fn_NomFeuilleActive() & "]")
    
End Sub

Private Sub ctrl_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)

    Call RafraichirActivite(ctrl.Name & "-" & TypeName(ctrl) & " [" & Fn_NomFeuilleActive() & "]")
    
End Sub

Private Sub ctrl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Call RafraichirActivite(ctrl.Name & "-" & TypeName(ctrl) & " [" & Fn_NomFeuilleActive() & "]")
    
End Sub


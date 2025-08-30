VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSurveillanceActivite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'clsSurveillanceActivite '2025-08-29 @ 18:36

Public WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Public WithEvents cb As MSForms.ComboBox
Attribute cb.VB_VarHelpID = -1
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private Sub tb_Change()

    EnregistrerActivite "TextBox"
    
End Sub

Private Sub cb_Change()

    EnregistrerActivite "ComboBox"
    
End Sub

Private Sub btn_Click()

    EnregistrerActivite "Bouton"
    
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrapperTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1

Private Sub tb_Change()

    modSurveillance.EnregistrerActivite "TextBox: " & tb.Name
    
End Sub


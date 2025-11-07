VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrapperCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents chk As MSForms.checkBox
Attribute chk.VB_VarHelpID = -1

Private Sub chk_Click()

    modSurveillance.EnregistrerActivite "CheckBox: " & chk.Name
    
End Sub


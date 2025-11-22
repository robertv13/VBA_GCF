VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrapperComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmb As MSForms.ComboBox
Attribute cmb.VB_VarHelpID = -1

Private Sub cmb_Change()

    modSurveillance.EnregistrerActivite "ComboBox: " & cmb.Name
    
End Sub


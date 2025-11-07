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

Public WithEvents cb As MSForms.ComboBox
Attribute cb.VB_VarHelpID = -1

Private Sub cb_Change()

    modSurveillance.EnregistrerActivite "ComboBox: " & cb.Name
    
End Sub


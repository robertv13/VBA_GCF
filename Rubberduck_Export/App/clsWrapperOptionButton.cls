VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrapperOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1

Private Sub opt_Click()

    Call modSurveillance.EnregistrerActivite("OptionButton: " & opt.Name)
    
End Sub


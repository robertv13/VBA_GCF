VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrapperListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lb As MSForms.ListBox
Attribute lb.VB_VarHelpID = -1

Private Sub lb_Change()

    modSurveillance.EnregistrerActivite "ListBox: " & lb.Name
    
End Sub


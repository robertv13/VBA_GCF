VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdSaisieHeures_Click()

    Load frmSaisieHeures
    
    frmSaisieHeures.show vbModal
    
End Sub

Private Sub cmdExportHeures_Click()

    Load frmExportHeures
    
    frmExportHeures.show vbModal

End Sub

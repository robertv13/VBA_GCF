VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdShowForm_Click()

        'Application.Visible = False
        
        frmSaisieHeures.show

End Sub

Private Sub cmdExportHeures_Click()

    MsgBox "Le code pour exporter vers le fichier de Guillaume"

End Sub


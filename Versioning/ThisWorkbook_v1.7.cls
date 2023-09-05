VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Make sure the Menu worksheet is selected, when we open the Workbook
Private Sub Workbook_Open()
    
    'Come back to menu Worksheet (wsMenu)
    wsMenu.Select

End Sub

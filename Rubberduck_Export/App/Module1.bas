Attribute VB_Name = "Module1"
Option Explicit

Sub test_Input_Box()

'    Dim Name As String
'    Name = InputBox("What is your name ? ", "Name entry")
    
    Dim output As Variant
    output = Application.InputBox("Saisi d'une chaine de caractères", "Exemple de saisi", "String")
    
'    Dim r As String
'    r = InputBox("Prompt", "Title", "Default")
    
End Sub

Sub Get_Last_Used_Row()

    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    'Dernière ligne utilisée
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    '-OU-
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    'Dernière colonne utilisée Attenttion au AdvanvedFilter...
    Dim lastUsedCol As Long
    lastUsedCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub



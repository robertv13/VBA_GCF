Attribute VB_Name = "Module1"
Option Explicit

Sub TesterInputBox()

'    Dim Name As String
'    Name = InputBox("What is your name ? ", "Name entry")
    
    Dim output As Variant
    output = Application.InputBox("Saisi d'une chaine de caract�res", "Exemple de saisi", "String")
    
'    Dim r As String
'    r = InputBox("Prompt", "Title", "Default")
    
End Sub

Sub ObtenirDerni�reRang�eUtilis�e()

    Dim ws As Worksheet: Set ws = wshTEC_Local
    
    'Derni�re ligne utilis�e
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    '-OU-
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    'Derni�re colonne utilis�e Attenttion au AdvanvedFilter...
    Dim lastUsedCol As Long
    lastUsedCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    'Lib�rer la m�moire
    Set ws = Nothing
    
End Sub


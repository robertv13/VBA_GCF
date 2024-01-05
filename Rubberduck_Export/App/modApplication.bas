Attribute VB_Name = "modApplication"
Option Explicit

Sub BackToMainMenu()

    wshMenu.Activate
'    Dim ws As Worksheet 'TO-DO - Remove comments, hide all worksheets
'    For Each ws In ActiveWorkbook.Worksheets
'        If ws.Name <> ActiveSheet.Name Then ws.Visible = xlSheetHidden
'    Next ws
    wshMenu.Range("B1").Select

End Sub

Public Sub ClearImmediateWindow()
    Dim i As Integer
    For i = 1 To 5 ' Adjust the number of lines based on your preference
        Debug.Print ""
    Next i
End Sub

Sub GetShapeProperties() 'List Properties of all the shapes

    Dim sShapes As Shape, lLoop As Long
    'Add headings for our lists. Expand as needed
    ActiveSheet.Range("E2:K2") = Array("Type", "Name", "Macro", "Height", "Width", "Left", "Top")
    lLoop = 1
    'Loop through all shapes on active sheet
    For Each sShapes In ActiveSheet.Shapes
        'Increment Variable lLoop for row numbers
        lLoop = lLoop + 1
        With sShapes
            'Add shape properties
            ActiveSheet.Cells(lLoop + 1, 5) = .Type
            ActiveSheet.Cells(lLoop + 1, 6) = .Name
            ActiveSheet.Cells(lLoop + 1, 7) = .OnAction
            ActiveSheet.Cells(lLoop + 1, 8) = .Height
            ActiveSheet.Cells(lLoop + 1, 9) = .Width
            ActiveSheet.Cells(lLoop + 1, 10) = .Left
            ActiveSheet.Cells(lLoop + 1, 11) = .Top
            'Follow the same pattern for more
        End With
    Next sShapes
End Sub

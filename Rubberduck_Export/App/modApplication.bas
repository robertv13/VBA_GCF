Attribute VB_Name = "modApplication"
Option Explicit

Sub BackToMainMenu()

    ActiveSheet.Visible = xlSheetHidden
    wshMenu.Activate
    wshMenu.Range("B1").Select

End Sub

Sub GetShapeProperties() 'List Properties of all the shapes

    Dim sShapes As Shape, lLoop As Long
    'Add headings for our lists. Expand as needed
    ActiveSheet.Range("E2:J2") = Array("Type", "Name", "Height", "Width", "Left", "Top")
    lLoop = 1
    'Loop through all shapes on active sheet
    For Each sShapes In ActiveSheet.Shapes
        'Increment Variable lLoop for row numbers
        lLoop = lLoop + 1
        With sShapes
            'Add shape properties
            ActiveSheet.Cells(lLoop + 1, 5) = .Type
            ActiveSheet.Cells(lLoop + 1, 6) = .Name
            ActiveSheet.Cells(lLoop + 1, 7) = .Height
            ActiveSheet.Cells(lLoop + 1, 8) = .Width
            ActiveSheet.Cells(lLoop + 1, 9) = .Left
            ActiveSheet.Cells(lLoop + 1, 10) = .Top
            'Follow the same pattern for more
        End With
    Next sShapes
End Sub

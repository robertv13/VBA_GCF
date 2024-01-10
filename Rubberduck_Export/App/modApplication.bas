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

Sub BuildDate(r As Range) '2024-01-06 @ 18:29
        Dim d, m, Y As Integer
        Dim strDateConsruite As String, cell As String
        Dim dateValide As Boolean
        cell = Trim(r.value)
        dateValide = True

        cell = Replace(cell, "/", "")
        cell = Replace(cell, "-", "")

        'Utilisation de la date du jour
        d = Day(Now())
        m = Month(Now())
        Y = Year(Now())

        Select Case Len(cell)
            Case 0
                strDateConsruite = Format(d, "00") & "/" & Format(m, "00") & "/" & Format(Y, "0000")
            Case 1, 2
                strDateConsruite = Format(cell, "00") & "/" & Format(m, "00") & "/" & Format(Y, "0000")
            Case 3
                strDateConsruite = Format(Left(cell, 1), "00") & "/" & Format(Mid(cell, 2, 2), "00") & "/" & Format(Y, "0000")
            Case 4
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(Y, "0000")
            Case 6
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & "20" & Format(Mid(cell, 5, 2), "00")
            Case 8
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(Mid(cell, 5, 4), "0000")
            Case Else
                dateValide = False
        End Select
        dateValide = IsDate(strDateConsruite)

    If dateValide Then
        r.value = Format(strDateConsruite, "dd/mm/yyyy")
    Else
        MsgBox "La saisie est invalide...", vbInformation, "Il est impossible de construire une date"
    End If

End Sub

Sub GetShapeProperties() 'List Properties of all the shapes

    Dim sShapes As Shape, lLoop As Long
    'Add headings for our lists. Expand as needed
    ActiveSheet.Range("G2:M2") = Array("Type", "Name", "Macro", "Left", "Top", "Height", "Width")
    lLoop = 3
    'Loop through all shapes on active sheet
    For Each sShapes In ActiveSheet.Shapes
        'Increment Variable lLoop for row numbers
        With sShapes
            'Add shape properties
            ActiveSheet.Cells(lLoop, 7) = .Type
            ActiveSheet.Cells(lLoop, 8) = .Name
            ActiveSheet.Cells(lLoop, 9) = .OnAction
            ActiveSheet.Cells(lLoop, 10) = .Left
            ActiveSheet.Cells(lLoop, 11) = .Top
            ActiveSheet.Cells(lLoop, 12) = .Height
            ActiveSheet.Cells(lLoop, 13) = .Width
            lLoop = lLoop + 1
            'Follow the same pattern for more
        End With
    Next sShapes
    ActiveSheet.Cells(lLoop + 1, 8) = Now()
End Sub

Sub AdjustShape()

    Dim sShapes As Shape, lLoop As Long
    'Loop through all shapes on active sheet
    For Each sShapes In ActiveSheet.Shapes
        Debug.Print sShapes.Name
        If sShapes.Name = "icoParametres" Then
            If sShapes.Left = 2.5 Then
                Debug.Print Tab(5); "The shape '" & sShapes.Name & "' as been adjusted !"
                sShapes.Name = "icoEXIT"
            End If
        End If
    Next sShapes
End Sub


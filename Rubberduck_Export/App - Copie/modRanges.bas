Attribute VB_Name = "modRanges"
Option Explicit

Public Sub RemoveFormatting()

    Dim c As Range

    For Each c In Range("A:A")
        c.ClearFormats
    Next c

End Sub

Sub DeleteNamedRanges()
    Dim MyName As Name

    For Each MyName In Names
        ActiveWorkbook.Names(MyName.Name).Delete
    Next

End Sub

Sub UnionExample()

    Dim Rng1, Rng2, Rng3 As Range

    Set Rng1 = Range("A1,A3,A5,A7,A9,A11,A13,A15,A17,A19,A21")
    Set Rng2 = Range("C1,C3,C5,C7,C9,C11,C13,C15,C17,C19,C21")
    Set Rng3 = Range("E1,E3,E5,E7,E9,E11,E13,E15,E17,E19,E21")

    Union(Rng1, Rng2, Rng3).Select

End Sub

Sub SizeChart2Range()

    Dim MyChart As Chart
    Dim MyRange As Range

    Set MyChart = ActiveSheet.ChartObjects(1).Chart
    Set MyRange = Feuil1.Range("B2:D6")

    With MyChart.Parent
        .Left = MyRange.Left
        .Top = MyRange.Top
        .Width = MyRange.Width
        .Height = MyRange.Height
    End With

End Sub

Sub MySelectAll()

    Feuil1.Activate
    Feuil1.Cells.Select

End Sub

Sub TestIfRange()

    If TypeName(Selection) = "Range" Then
        MsgBox "You selected a Range"
    Else
        MsgBox "Woops! You selected a " & TypeName(Selection)
    End If

End Sub


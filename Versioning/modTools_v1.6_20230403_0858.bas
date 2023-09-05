Attribute VB_Name = "modTools"
Option Explicit

Sub GetAllShapes()

Dim shp As Shape

    'Loop through each shape on ActiveSheet
    For Each shp In ActiveSheet.Shapes
        Debug.Print shp.Name
        Debug.Print shp.Height
        Debug.Print shp.Visible
    Next shp

End Sub

Sub getShapeNames()

    Dim sh As Shape
    Dim x As Double
    x = 1
    ActiveSheet.Select

    For Each sh In ActiveSheet.Shapes
        ActiveSheet.Range("C" + Trim(Str(x + 1))).value = x & _
            " - " & sh.Name & " - " & sh.Left
        sh.Delete
        x = x + 1
    Next

End Sub

Sub ChangeShapeProperties()
    
    ActiveSheet.Shapes(16).Select
    Selection.Name = "lblSwipeInAll"

End Sub

Sub GetOnACtionProperties()

    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        If Left(sh.Name, 3) = "btn" Then
            Debug.Print sh.Name & " - " & sh.OnAction
        End If
    Next

End Sub

Sub SetOnACtionProperties()

    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        If sh.Name = "btnTEC" Then
            sh.OnAction = "shpTECClick"
        End If
        Debug.Print sh.Name & " - " & sh.OnAction
    Next

End Sub

Sub DisplayTotalsRowAddress()

    Dim wrksht As Worksheet
    Dim objListObj As ListObject
    
    Set wrksht = ActiveWorkbook.Worksheets("Heures")
    Set objListObj = wrksht.ListObjects(1)
    Debug.Print objListObj.DisplayName
    Debug.Print objListObj.HeaderRowRange.Address
    Debug.Print objListObj.DataBodyRange.Address
    Debug.Print objListObj.ListRows.count
    Debug.Print objListObj.Range.Address
    Debug.Print objListObj.ShowTotals
    
End Sub

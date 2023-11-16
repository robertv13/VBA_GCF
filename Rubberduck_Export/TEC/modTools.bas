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
    Dim X As Double
    X = 1
    ActiveSheet.Select

    For Each sh In ActiveSheet.Shapes
        ActiveSheet.Range("C" + Trim(Str(X + 1))).value = X & _
                                                          " - " & sh.Name & " - " & sh.Left
        sh.Delete
        X = X + 1
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



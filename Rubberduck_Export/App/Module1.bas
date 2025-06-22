Sub Exemple()
    Dim r As Range
    Set r = ActiveSheet.range("A1").Offset(1, 0)
    Debug.Print r.value
End Sub



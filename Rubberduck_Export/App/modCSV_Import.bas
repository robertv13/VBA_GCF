Attribute VB_Name = "modCSV_Import"
Option Explicit

Sub zz_ImporterFichierCSV(ws As Worksheet, path As String, fn As String)

    'Does the file to process exist ?
    Dim fullFileName As String
    fullFileName = path & Application.PathSeparator & fn
    If Dir(fullFileName) = vbNullString Then
        MsgBox "Le fichier n'a pas été trouvé", vbExclamation
        Exit Sub
    End If
    
    Dim lastUsedRow As Long, firstAvailRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    firstAvailRow = lastUsedRow + 1
    
    'Import data from external file into the worksheet
    With ws.QueryTables.Add(Connection:="TEXT;" & fullFileName, Destination:=ws.Range("A" & firstAvailRow))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileStartRow = 3
        .TextFileColumnDataTypes = Array(1, 1, 1, 3, 1, 1, 1, 2, 2, 1, 1, 1, 1, 2) 'Specify text type (1 = Text, 2 = Number, 3 = Date)
        .Refresh BackgroundQuery:=False
    End With
    
    'Correct all formats
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Call AjusterColonnesFichierCSV(ws, firstAvailRow, lastUsedRow)
    
    Debug.Print "#023 - " & fn & " a été importée avec succès, " & lastUsedRow - firstAvailRow & " lignes"

End Sub

Sub AjusterColonnesFichierCSV(ws As Worksheet, first As Long, last As Long)

    Dim i As Long
    For i = first To last
        'Column B - Change the date format
        ws.Range("D" & i).Value = Fn_FormatDateCorrige(ws.Range("D" & i).Value)
        ws.Range("D" & i).NumberFormat = "dd/mm/yyyy"
        'Column H - Change the numeric format
        ws.Range("H" & i).Value = Replace(ws.Range("H" & i).Value, ",", vbNullString)
        ws.Range("H" & i).Value = Replace(ws.Range("H" & i).Value, ".", ",")
        ws.Range("H" & i).Value = CDbl(ws.Range("H" & i).Value)
        'Column H - Change the numeric format
        ws.Range("I" & i).Value = Replace(ws.Range("I" & i).Value, ",", vbNullString)
        ws.Range("I" & i).Value = Replace(ws.Range("I" & i).Value, ".", ",")
        ws.Range("I" & i).Value = CDbl(ws.Range("I" & i).Value)
        'Column N - Change the balance amount format
        ws.Range("N" & i).Value = Replace(ws.Range("N" & i).Value, ",", vbNullString)
        ws.Range("N" & i).Value = Replace(ws.Range("N" & i).Value, ".", ",")
        ws.Range("N" & i).Value = CDbl(ws.Range("N" & i).Value)
        
    Next i
    
End Sub

Function Fn_FormatDateCorrige(wrongFormatDate As String) As Date

    Debug.Print "#024 - " & wrongFormatDate

    Dim arr() As String
    arr = Split(wrongFormatDate, "/")

    Dim year As Long, month As Long, day As Long
    year = Format$(arr(2), "0000")
    Select Case arr(1)
        Case "mai"
            month = 5
    End Select
    month = Format$(month, "00")
    day = Format$(arr(0), "00")

    Fn_FormatDateCorrige = DateSerial(year, month, day)
    
    Debug.Print "#025 - " & DateSerial(year, month, day)

End Function


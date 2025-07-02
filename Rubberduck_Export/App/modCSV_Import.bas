Attribute VB_Name = "modCSV_Import"
Option Explicit

Sub main()

    'Setup the receiving worksheet and clear the previous results
    Dim ws As Worksheet: Set ws = wshzCSV_File
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastUsedRow > 1 Then
        ws.Range("A2:P" & lastUsedRow).ClearContents
    End If
    
    'Setup path to Source files
    Dim pathSourceFile As String
    pathSourceFile = "C:\Users\RobertMV\Downloads"
    
    Dim fileName As String
    
    'First file
    fileName = "Releve.csv"
    Call Import_CSV_File(ws, pathSourceFile, fileName)
    
    'Fix columns width
    Call Set_Column_Width(ws)

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub Import_CSV_File(ws As Worksheet, path As String, fn As String)

    'Does the file to process exist ?
    Dim fullFileName As String
    fullFileName = path & Application.PathSeparator & fn
    If Dir(fullFileName) = "" Then
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
    Call Fix_Columns(ws, firstAvailRow, lastUsedRow)
    
    Debug.Print "#023 - " & fn & " a été importée avec succès, " & lastUsedRow - firstAvailRow & " lignes"

End Sub

Sub Fix_Columns(ws As Worksheet, first As Long, last As Long)

    Dim i As Long
    For i = first To last
        'Column B - Change the date format
        ws.Range("D" & i).Value = Fn_Correct_Date_Format(ws.Range("D" & i).Value)
        ws.Range("D" & i).NumberFormat = "dd/mm/yyyy"
        'Column H - Change the numeric format
        ws.Range("H" & i).Value = Replace(ws.Range("H" & i).Value, ",", "")
        ws.Range("H" & i).Value = Replace(ws.Range("H" & i).Value, ".", ",")
        ws.Range("H" & i).Value = CDbl(ws.Range("H" & i).Value)
        'Column H - Change the numeric format
        ws.Range("I" & i).Value = Replace(ws.Range("I" & i).Value, ",", "")
        ws.Range("I" & i).Value = Replace(ws.Range("I" & i).Value, ".", ",")
        ws.Range("I" & i).Value = CDbl(ws.Range("I" & i).Value)
        'Column N - Change the balance amount format
        ws.Range("N" & i).Value = Replace(ws.Range("N" & i).Value, ",", "")
        ws.Range("N" & i).Value = Replace(ws.Range("N" & i).Value, ".", ",")
        ws.Range("N" & i).Value = CDbl(ws.Range("N" & i).Value)
        
    Next i
    
End Sub

Sub Set_Column_Width(ws As Worksheet)

    ws.Range("H:H").NumberFormat = "#,##0.00"
    ws.Range("I:I").NumberFormat = "#,##0.00"
    ws.Range("N:N").NumberFormat = "#,##0.00"
    
    Dim col As Long, lastUsedColumn As Long
    lastUsedColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For col = 1 To lastUsedColumn
        ws.Columns(col).AutoFit
    Next col

End Sub

Function Fn_Correct_Date_Format(wrongFormatDate As String) As Date

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

    Fn_Correct_Date_Format = DateSerial(year, month, day)
    Debug.Print "#025 - " & DateSerial(year, month, day)

End Function



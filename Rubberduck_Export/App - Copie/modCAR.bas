Attribute VB_Name = "modCAR"
Option Explicit

Sub CAR_TdB_Update_All()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR:CAR_TdB_Update_All", 0)
    
    Call CAR_Update_TdB_Data
    Call CAR_Refresh_CAR_PivotTables
    
    Call Log_Record("modCAR:CAR_TdB_Update_All()", startTime)

End Sub

Sub CAR_Update_TdB_Data()

'    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR:CAR_Update_TdB_Data",0)

    Dim wsSource As Worksheet: Set wsSource = wshFAC_Comptes_Clients
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Cells(wsSource.rows.count, "A").End(xlUp).Row
    
    Dim wsTarget As Worksheet: Set wsTarget = wshCAR_TDB_Data
    Dim lastUsedRowTarget As Long
    lastUsedRowTarget = wsTarget.Cells(wsTarget.rows.count, "A").End(xlUp).Row
    wsTarget.Range("A2:F" & lastUsedRowTarget).ClearContents
    
    Dim arr() As Variant
    ReDim arr(1 To lastUsedRow - 2, 1 To 6) '2 rowsSource of Heading
    
    Dim i As Long
    For i = 3 To lastUsedRow
        With wsSource
'            If .Range("J" & i).Value <> 0 Then
                arr(i - 2, 1) = .Range("A" & i).Value 'Invoice_No
                arr(i - 2, 2) = .Range("B" & i).Value 'Invoice_Date
                arr(i - 2, 3) = .Range("C" & i).Value 'ClientsName
                arr(i - 2, 4) = .Range("D" & i).Value 'ClientsCode
                arr(i - 2, 5) = .Range("G" & i).Value 'DueDate
                arr(i - 2, 6) = .Range("J" & i).Value 'Balance
'            End If
        End With
    Next i

    Dim rngTarget As Range: Set rngTarget = wshCAR_TDB_Data.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2))
    rngTarget.Value = arr
    
    'Remove rows, when Balance = 0 $
    lastUsedRowTarget = wsTarget.Cells(wsTarget.rows.count, "A").End(xlUp).Row
    For i = lastUsedRowTarget To 2 Step -1
        If wsTarget.Cells(i, 6).Value = 0 Then
            wsTarget.rows(i).delete
        End If
    Next i
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rngTarget = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
    
'    Call Log_Record("modCAR:CAR_Update_TdB_Data()", startTime)

End Sub

Sub CAR_Refresh_CAR_PivotTables()

'    Dim startTime As Double: startTime = Timer: Call Log_Record("modCAR:CAR_Refresh_CAR_PivotTables",0)
    
    Dim pt As PivotTable
    For Each pt In wshCAR_TDB_PivotTable.PivotTables
        pt.RefreshTable
    Next pt

    'Cleaning memory - 2024-07-01 @ 09:34
    Set pt = Nothing
    
'    Call Log_Record("modCAR:CAR_Refresh_CAR_PivotTables()", startTime)

End Sub



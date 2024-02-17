Attribute VB_Name = "Dashboard_Macros"
Option Explicit
Dim LastRow As Long, LastResultRow As Long, SelRow As Long, SelCol As Long, DetailNumb As Long

Sub Dashboard_TabChange()
    
    Application.ScreenUpdating = False
    
    With Dashboard
        SelCol = .Range("AA3").value 'Selected Column
        'Hide All Shapes & Graphs
        .Range("4:1004").EntireRow.Hidden = True
        .Shapes("DashGrp").Visible = msoFalse       'Hide Dashboard Group
        .Shapes("DetailGrp").Visible = msoFalse     'Hide Details Group
        
        Select Case SelCol
            Case Is = 2                                 'Dashboard
                .Range("4:32").EntireRow.Hidden = False
                .Shapes("DashGrp").Visible = msoCTrue   'Display Dashboard Graphs/Buttons
                Call Dashboard_Refresh
            Case Is = 3                                 'Aging Summary
                .Range("33:502").EntireRow.Hidden = False
                Call Aging_Refresh                      'Run Macro To Refresh Aging
            Case Is = 5                                 'Aging Detail
                .Range("503:1004").EntireRow.Hidden = False
                .Shapes("DetailGrp").Visible = msoCTrue 'Show Detail group
                Call AgingDetail_Refresh                'Run Macro to refresh aging detail
        End Select
        .Range("A1").Select
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Sub Dashboard_Refresh()
    'Get Current Data
    With InvoiceList 'Get Aged listing of A/R @ Invoice List.Range("P3:W999999")
        LastRow = .Range("A999999").End(xlUp).row
        .Range("A3:J" & LastRow).ClearContents
        
        'Copy AR_Entête to Invoice List
        Dim sourceRange As Range, targetRange As Range, maxRow As Long
        LastRow = wshAR.Range("A999999").End(xlUp).row
        Set sourceRange = wshAR.Range("A3:J" & LastRow)
        Set targetRange = InvoiceList.Range("A3:J" & LastRow)
        'Copy values from source range to target range
        targetRange.value = sourceRange.value

        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        LastRow = .Range("A99999").End(xlUp).row
        If LastRow < 3 Then Exit Sub
        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:D" & LastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("L1:L2"), _
            CopyToRange:=.Range("P2"), _
            Unique:=True
        LastResultRow = .Range("P99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        .Range("Q3:W" & LastResultRow).formula = .Range("Q1:W1").formula
    End With
End Sub

Sub Dashboard_SelectAgingDetails()
    DetailNumb = Replace(Application.Caller, "Aging", "")
    Dashboard.Range("AA504").value = DetailNumb  'Set Detail level #
    Dashboard.Range("E2").Select                 'Select to trigger macro
End Sub

Sub Aging_Refresh()
    Dashboard.Range("B35:R499").ClearContents    'Clear Previous Results
    With InvoiceList
        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        LastRow = .Range("A99999").End(xlUp).row
        If LastRow < 3 Then Exit Sub
        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        'Very long to execute - 2024-02-16 @ 06:49
        .Range("A2:D" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("L1:L2"), CopyToRange:=.Range("P2"), Unique:=True
        LastResultRow = .Range("P99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        .Range("Q3:V" & LastResultRow).formula = .Range("Q1:W1").formula
        Dashboard.Range("B35:H" & LastResultRow + 32).value = .Range("P3:V" & LastResultRow).value 'Bring over Aging Data
    End With
End Sub

Sub Aging_ShowCustDetail()
    Dashboard.Range("J34:R499").ClearContents    'Clear Previous Results
    SelRow = Dashboard.Range("AA1").value        'Set Selected Row
    With InvoiceList
        'Clear Prior Results
        LastRow = .Range("AB9999").End(xlUp).row + 1
        .Range("AB3:AJ" & LastRow).ClearContents
        LastRow = .Range("A99999").End(xlUp).row
        If LastRow < 3 Then Exit Sub
        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:J" & LastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("Y1:Z2"), _
            CopyToRange:=.Range("AB2:AJ2"), _
            Unique:=True
        LastResultRow = .Range("AB99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        Dashboard.Range("J" & SelRow & ":R" & SelRow + LastResultRow - 1).value = .Range("AB1:AJ" & LastResultRow).value 'Bring over Customer Details
            Dashboard.Range("J" & SelRow & ":R" & SelRow).HorizontalAlignment = xlCenterAcrossSelection
    End With
End Sub

Sub Aging_GoToInvoice()
    With Dashboard
        SelRow = .Range("AA2").value             'Selected Row
        If SelRow = 0 Then Exit Sub
        If .Range("J" & SelRow).value = "" Then Exit Sub
        Invoice.Activate
        Invoice.Range("L1").value = .Range("J" & SelRow).value 'set Invoice #
    End With
End Sub

Sub AgingDetail_Refresh()
    Dashboard.Range("B507:J9999").ClearContents  'Clear Existing Data
    With InvoiceList
        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        LastRow = .Range("A99999").End(xlUp).row
        If LastRow < 3 Then Exit Sub
        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:ND" & LastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("AL1:AM2"), _
            CopyToRange:=.Range("AB2:AJ2"), _
            Unique:=True
        LastResultRow = .Range("AB99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        Dashboard.Range("B507:J" & LastResultRow + 504).value = .Range("AB3:AJ" & LastResultRow).value 'Bring over Aging Data
    End With

End Sub



Attribute VB_Name = "Dashboard_Macros"
Option Explicit
Dim lastRow As Long, lastResultRow As Long, selRow As Long, selCol As Long

Sub wshARDashboard_TabChange()
    
    Application.ScreenUpdating = False
    
    With wshARDashboard
        selCol = .Range("AA3").value 'Selected Column
        'Hide All Shapes & Graphs
        .Range("4:1004").EntireRow.Hidden = True
        .Shapes("DashGrp").Visible = msoFalse       'Hide Dashboard Group
        .Shapes("DetailGrp").Visible = msoFalse     'Hide Details Group
        
        Select Case selCol
            Case Is = 2                                 'Dashboard
                .Range("4:32").EntireRow.Hidden = False
                .Shapes("DashGrp").Visible = msoCTrue   'Display Dashboard Graphs/Buttons
                Call wshARDashboard_Refresh
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

Sub wshARDashboard_Refresh()
    'Get Current Data
    With wshInvoiceList 'Get Aged listing of A/R @ Invoice List.Range("P3:W999999")
        lastRow = .Range("A999999").End(xlUp).row
        .Range("A3:J" & lastRow).ClearContents
        
        'Copy AR_Entête to Invoice List
        Dim sourceRange As Range, targetRange As Range, maxRow As Long
        lastRow = wshAR.Range("A999999").End(xlUp).row
        Set sourceRange = wshAR.Range("A3:J" & lastRow)
        Set targetRange = wshInvoiceList.Range("A3:J" & lastRow)
        'Copy values from source range to target range
        targetRange.value = sourceRange.value

        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        lastRow = .Range("A99999").End(xlUp).row
        If lastRow < 3 Then Exit Sub
        .Range("H3:J" & lastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:D" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("L1:L2"), _
            CopyToRange:=.Range("P2"), _
            Unique:=True
        lastResultRow = .Range("P99999").End(xlUp).row
        If lastResultRow < 3 Then Exit Sub
        .Range("Q3:W" & lastResultRow).formula = .Range("Q1:W1").formula
        'Define a table - 2024-02-17 @ 08:15
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim rng As Range
    End With
    
        'Set the worksheet where you want to create the table
        Set ws = wshInvoiceList
    
        'Define the range for the table and create the table
        Set rng = ws.Range("P1:W" & lastResultRow)
        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    
        'Define table properties
        tbl.name = "AgingSummary"
'        tbl.TableStyle = "TableStyleMedium2" ' Style of the table

End Sub

Sub wshARDashboard_SelectAgingDetails()
    detailNumb = Replace(Application.Caller, "Aging", "")
    wshARDashboard.Range("AA504").value = detailNumb  'Set Detail level #
    wshARDashboard.Range("E2").Select                 'Select to trigger macro
End Sub

Sub Aging_Refresh()
    wshARDashboard.Range("B35:R499").ClearContents    'Clear Previous Results
    With wshInvoiceList
        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        lastRow = .Range("A99999").End(xlUp).row
        If lastRow < 3 Then Exit Sub
        .Range("H3:J" & lastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        'Very long to execute - 2024-02-16 @ 06:49
        .Range("A2:D" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("L1:L2"), CopyToRange:=.Range("P2"), Unique:=True
        lastResultRow = .Range("P99999").End(xlUp).row
        If lastResultRow < 3 Then Exit Sub
        .Range("Q3:V" & lastResultRow).formula = .Range("Q1:W1").formula
        wshARDashboard.Range("B35:H" & lastResultRow + 32).value = .Range("P3:V" & lastResultRow).value 'Bring over Aging Data
    End With
End Sub

Sub Aging_ShowCustDetail()
    wshARDashboard.Range("J34:R499").ClearContents    'Clear Previous Results
    selRow = wshARDashboard.Range("AA1").value        'Set Selected Row
    With wshInvoiceList
        'Clear Prior Results
        lastRow = .Range("AB9999").End(xlUp).row + 1
        .Range("AB3:AJ" & lastRow).ClearContents
        lastRow = .Range("A99999").End(xlUp).row
        If lastRow < 3 Then Exit Sub
        .Range("H3:J" & lastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:J" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("Y1:Z2"), _
            CopyToRange:=.Range("AB2:AJ2"), _
            Unique:=True
        lastResultRow = .Range("AB99999").End(xlUp).row
        If lastResultRow < 3 Then Exit Sub
        wshARDashboard.Range("J" & selRow & ":R" & selRow + lastResultRow - 1).value = .Range("AB1:AJ" & lastResultRow).value 'Bring over Customer Details
            wshARDashboard.Range("J" & selRow & ":R" & selRow).HorizontalAlignment = xlCenterAcrossSelection
    End With
End Sub

Sub Aging_GoToInvoice()
    With wshARDashboard
        selRow = .Range("AA2").value             'Selected Row
        If selRow = 0 Then Exit Sub
        If .Range("J" & selRow).value = "" Then Exit Sub
        Invoice.Activate
        Invoice.Range("L1").value = .Range("J" & selRow).value 'set Invoice #
    End With
End Sub

Sub AgingDetail_Refresh()
    wshARDashboard.Range("B507:J9999").ClearContents  'Clear Existing Data
    With wshInvoiceList
        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        lastRow = .Range("A99999").End(xlUp).row
        If lastRow < 3 Then Exit Sub
        .Range("H3:J" & lastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:ND" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("AL1:AM2"), _
            CopyToRange:=.Range("AB2:AJ2"), _
            Unique:=True
        lastResultRow = .Range("AB99999").End(xlUp).row
        If lastResultRow < 3 Then Exit Sub
        wshARDashboard.Range("B507:J" & lastResultRow + 504).value = .Range("AB3:AJ" & lastResultRow).value 'Bring over Aging Data
    End With

End Sub



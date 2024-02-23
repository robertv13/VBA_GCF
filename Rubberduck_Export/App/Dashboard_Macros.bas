Attribute VB_Name = "Dashboard_Macros"
Option Explicit
Dim lastrow As Long, LastResultRow As Long, SelRow As Long, SelCol As Long, DetailNumb As Long

Sub wshARDashboard_TabChange()
    
    Application.ScreenUpdating = False
    
    With wshARDashboard
        SelCol = .Range("AA3").value 'Selected Column
        'Hide All Shapes & Graphs
        .Range("4:1004").EntireRow.Hidden = True
        .Shapes("DashGrp").Visible = msoFalse       'Hide Dashboard Group
        .Shapes("DetailGrp").Visible = msoFalse     'Hide Details Group
        
        Select Case SelCol
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
        lastrow = .Range("A999999").End(xlUp).row
        .Range("A3:J" & lastrow).ClearContents
        
        'Copy AR_Entête to Invoice List
        Dim sourceRange As Range, targetRange As Range, maxRow As Long
        lastrow = wshAR.Range("A999999").End(xlUp).row
        Set sourceRange = wshAR.Range("A3:J" & lastrow)
        Set targetRange = wshInvoiceList.Range("A3:J" & lastrow)
        'Copy values from source range to target range
        targetRange.value = sourceRange.value

        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        lastrow = .Range("A99999").End(xlUp).row
        If lastrow < 3 Then Exit Sub
        .Range("H3:J" & lastrow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:D" & lastrow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("L1:L2"), _
            CopyToRange:=.Range("P2"), _
            Unique:=True
        LastResultRow = .Range("P99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        .Range("Q3:W" & LastResultRow).formula = .Range("Q1:W1").formula
        'Define a table - 2024-02-17 @ 08:15
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim Rng As Range
    End With
    
        'Set the worksheet where you want to create the table
        Set ws = wshInvoiceList
    
        'Define the range for the table and create the table
        Set Rng = ws.Range("P1:W" & LastResultRow)
        Set tbl = ws.ListObjects.Add(xlSrcRange, Rng, , xlYes)
    
        'Define table properties
        tbl.name = "AgingSummary"
'        tbl.TableStyle = "TableStyleMedium2" ' Style of the table

End Sub

Sub wshARDashboard_SelectAgingDetails()
    DetailNumb = Replace(Application.Caller, "Aging", "")
    wshARDashboard.Range("AA504").value = DetailNumb  'Set Detail level #
    wshARDashboard.Range("E2").Select                 'Select to trigger macro
End Sub

Sub Aging_Refresh()
    wshARDashboard.Range("B35:R499").ClearContents    'Clear Previous Results
    With wshInvoiceList
        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        lastrow = .Range("A99999").End(xlUp).row
        If lastrow < 3 Then Exit Sub
        .Range("H3:J" & lastrow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        'Very long to execute - 2024-02-16 @ 06:49
        .Range("A2:D" & lastrow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("L1:L2"), CopyToRange:=.Range("P2"), Unique:=True
        LastResultRow = .Range("P99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        .Range("Q3:V" & LastResultRow).formula = .Range("Q1:W1").formula
        wshARDashboard.Range("B35:H" & LastResultRow + 32).value = .Range("P3:V" & LastResultRow).value 'Bring over Aging Data
    End With
End Sub

Sub Aging_ShowCustDetail()
    wshARDashboard.Range("J34:R499").ClearContents    'Clear Previous Results
    SelRow = wshARDashboard.Range("AA1").value        'Set Selected Row
    With wshInvoiceList
        'Clear Prior Results
        lastrow = .Range("AB9999").End(xlUp).row + 1
        .Range("AB3:AJ" & lastrow).ClearContents
        lastrow = .Range("A99999").End(xlUp).row
        If lastrow < 3 Then Exit Sub
        .Range("H3:J" & lastrow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:J" & lastrow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("Y1:Z2"), _
            CopyToRange:=.Range("AB2:AJ2"), _
            Unique:=True
        LastResultRow = .Range("AB99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        wshARDashboard.Range("J" & SelRow & ":R" & SelRow + LastResultRow - 1).value = .Range("AB1:AJ" & LastResultRow).value 'Bring over Customer Details
            wshARDashboard.Range("J" & SelRow & ":R" & SelRow).HorizontalAlignment = xlCenterAcrossSelection
    End With
End Sub

Sub Aging_GoToInvoice()
    With wshARDashboard
        SelRow = .Range("AA2").value             'Selected Row
        If SelRow = 0 Then Exit Sub
        If .Range("J" & SelRow).value = "" Then Exit Sub
        Invoice.Activate
        Invoice.Range("L1").value = .Range("J" & SelRow).value 'set Invoice #
    End With
End Sub

Sub AgingDetail_Refresh()
    wshARDashboard.Range("B507:J9999").ClearContents  'Clear Existing Data
    With wshInvoiceList
        'Clear Prior Results
        .Range("AB3:AJ9999").ClearContents
        lastrow = .Range("A99999").End(xlUp).row
        If lastrow < 3 Then Exit Sub
        .Range("H3:J" & lastrow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
        .Range("A2:ND" & lastrow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("AL1:AM2"), _
            CopyToRange:=.Range("AB2:AJ2"), _
            Unique:=True
        LastResultRow = .Range("AB99999").End(xlUp).row
        If LastResultRow < 3 Then Exit Sub
        wshARDashboard.Range("B507:J" & LastResultRow + 504).value = .Range("AB3:AJ" & LastResultRow).value 'Bring over Aging Data
    End With

End Sub



Attribute VB_Name = "modzDashboard_Macros"
Option Explicit

'Dim LastRow As Long, lastResultRow As Long, selRow As Long, selCol As Long

'Sub wshCAR_Dashboard_TabChange()
'
'    Application.ScreenUpdating = False
'
'    With wshCAR_Dashboard
'        selCol = .Range("AA3").Value 'Selected Column
'        'Hide All Shapes & Graphs
'        .Range("4:1004").EntireRow.Hidden = True
'        .Shapes("DashGrp").Visible = msoFalse       'Hide CAR_Dashboard Group
'        .Shapes("DetailGrp").Visible = msoFalse     'Hide Details Group
'
'        Select Case selCol
'            Case Is = 2                                 'CAR_Dashboard
'                .Range("4:32").EntireRow.Hidden = False
'                .Shapes("DashGrp").Visible = msoCTrue   'Display CAR_Dashboard Graphs/Buttons
'                Call wshCAR_Dashboard_Refresh
'            Case Is = 3                                 'Aging Summary
'                .Range("33:502").EntireRow.Hidden = False
'                Call Aging_Refresh                      'Run Macro To Refresh Aging
'            Case Is = 5                                 'Aging Detail
'                .Range("503:1004").EntireRow.Hidden = False
'                .Shapes("DetailGrp").Visible = msoCTrue 'Show Detail group
'                Call AgingDetail_Refresh                'Run Macro to refresh aging detail
'        End Select
'        .Range("A1").Select
'    End With
    
'    Application.ScreenUpdating = True
'
'End Sub
'
'Sub wshCAR_Dashboard_Refresh()
'
'    'Get Current Data
'    With wshFAC_Invoice_List 'Get Aged listing of A/R @ Invoice List.Range("P3:W999999")
'        LastRow = .Range("A999999").End(xlUp).row
'        .Range("A3:J" & LastRow).ClearContents
'
'        'Copy AR_Ent�te to Invoice List
'        Dim maxRow As Long
'        LastRow = wshFAC_Comptes_Clients.Range("A999999").End(xlUp).row
'        Dim sourceRange As Range: Set sourceRange = wshFAC_Comptes_Clients.Range("A3:J" & LastRow)
'        Dim targetRange As Range: Set targetRange = wshFAC_Invoice_List.Range("A3:J" & LastRow)
'        'Copy values from source range to target range
'        targetRange.Value = sourceRange.Value
'
'        'Clear Prior Results
'        .Range("AB3:AJ9999").ClearContents
'        LastRow = .Range("A99999").End(xlUp).row
'        If LastRow < 3 Then Exit Sub
'        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
'        .Range("A2:D" & LastRow).AdvancedFilter xlFilterCopy, _
'            criteriaRange:=.Range("L1:L2"), _
'            CopyToRange:=.Range("P2"), _
'            Unique:=True
'        lastResultRow = .Range("P99999").End(xlUp).row
'        If lastResultRow < 3 Then Exit Sub
'        .Range("Q3:W" & lastResultRow).formula = .Range("Q1:W1").formula
'        'Define a table - 2024-02-17 @ 08:15
'    End With
'
'        'Set the worksheet where you want to create the table
'        Dim ws As Worksheet: Set ws = wshFAC_Invoice_List
'
'        'Define the range for the table and create the table
'        Dim rng As Range: Set rng = ws.Range("P1:W" & lastResultRow)
'        Dim tbl As ListObject: Set tbl = ws.ListObjects.add(xlSrcRange, rng, , xlYes)
'
'        'Define table properties
'        tbl.name = "AgingSummary"
''        tbl.TableStyle = "TableStyleMedium2" ' Style of the table
'
'    'Cleaning memory - 2024-07-01 @ 09:34
'    Set rng = Nothing
'    Set sourceRange = Nothing
'    Set targetRange = Nothing
'    Set tbl = Nothing
'    Set ws = Nothing
'
'End Sub
'
'Sub wshCAR_Dashboard_SelectAgingDetails()
'
'    Dim detailNumb As String
'    detailNumb = Replace(Application.Caller, "Aging", "")
'    wshCAR_Dashboard.Range("AA504").Value = detailNumb  'Set Detail level #
'    wshCAR_Dashboard.Range("E2").Select                 'Select to trigger macro
'
'End Sub
'
'Sub Aging_Refresh()
'
'    wshCAR_Dashboard.Range("B35:R499").ClearContents    'Clear Previous Results
'    With wshFAC_Invoice_List
'        'Clear Prior Results
'        .Range("AB3:AJ9999").ClearContents
'        LastRow = .Range("A99999").End(xlUp).row
'        If LastRow < 3 Then Exit Sub
'        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
'        'Very long to execute - 2024-02-16 @ 06:49
'        .Range("A2:D" & LastRow).AdvancedFilter xlFilterCopy, criteriaRange:=.Range("L1:L2"), CopyToRange:=.Range("P2"), Unique:=True
'        lastResultRow = .Range("P99999").End(xlUp).row
'        If lastResultRow < 3 Then Exit Sub
'        .Range("Q3:V" & lastResultRow).formula = .Range("Q1:W1").formula
'        wshCAR_Dashboard.Range("B35:H" & lastResultRow + 32).Value = .Range("P3:V" & lastResultRow).Value 'Bring over Aging Data
'    End With
'
'End Sub
'
'Sub Aging_ShowCustDetail()
'
'    wshCAR_Dashboard.Range("J34:R499").ClearContents    'Clear Previous Results
'    selRow = wshCAR_Dashboard.Range("AA1").Value        'Set Selected Row
'    With wshFAC_Invoice_List
'        'Clear Prior Results
'        LastRow = .Range("AB9999").End(xlUp).row + 1
'        .Range("AB3:AJ" & LastRow).ClearContents
'        LastRow = .Range("A99999").End(xlUp).row
'        If LastRow < 3 Then Exit Sub
'        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
'        .Range("A2:J" & LastRow).AdvancedFilter xlFilterCopy, _
'            criteriaRange:=.Range("Y1:Z2"), _
'            CopyToRange:=.Range("AB2:AJ2"), _
'            Unique:=True
'        lastResultRow = .Range("AB99999").End(xlUp).row
'        If lastResultRow < 3 Then Exit Sub
'        wshCAR_Dashboard.Range("J" & selRow & ":R" & selRow + lastResultRow - 1).Value = .Range("AB1:AJ" & lastResultRow).Value 'Bring over Customer Details
'            wshCAR_Dashboard.Range("J" & selRow & ":R" & selRow).HorizontalAlignment = xlCenterAcrossSelection
'    End With
'
'End Sub
'
'Sub AgingDetail_Refresh()
'
'    wshCAR_Dashboard.Range("B507:J9999").ClearContents  'Clear Existing Data
'    With wshFAC_Invoice_List
'        'Clear Prior Results
'        .Range("AB3:AJ9999").ClearContents
'        LastRow = .Range("A99999").End(xlUp).row
'        If LastRow < 3 Then Exit Sub
'        .Range("H3:J" & LastRow).formula = .Range("H1:J1").formula 'Bring Down Total Paid & Days Overdue Formulas
'        .Range("A2:ND" & LastRow).AdvancedFilter xlFilterCopy, _
'            criteriaRange:=.Range("AL1:AM2"), _
'            CopyToRange:=.Range("AB2:AJ2"), _
'            Unique:=True
'        lastResultRow = .Range("AB99999").End(xlUp).row
'        If lastResultRow < 3 Then Exit Sub
'        wshCAR_Dashboard.Range("B507:J" & lastResultRow + 504).Value = .Range("AB3:AJ" & lastResultRow).Value 'Bring over Aging Data
'    End With
'
'End Sub
'
'

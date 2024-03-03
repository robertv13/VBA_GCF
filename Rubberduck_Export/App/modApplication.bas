Attribute VB_Name = "modApplication"
Option Explicit

Sub BackToMainMenu()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    wshMenu.Activate
    wshMenu.Range("A1").Select

End Sub

'Sub Build_Date(r As Range) '2024-01-06 @ 18:29
'        Dim d, m, y As Integer
'        Dim strDateConsruite As String, cell As String
'        Dim dateValide As Boolean
'        cell = Trim(r.value)
'        dateValide = True
'
'        cell = Replace(cell, "/", "")
'        cell = Replace(cell, "-", "")
'
'        'Utilisation de la date du jour
'        d = Day(Now())
'        m = Month(Now())
'        y = Year(Now())
'
'        Select Case Len(cell)
'            Case 0
'                strDateConsruite = Format(d, "00") & "/" & Format(m, "00") & "/" & Format(y, "0000")
'            Case 1, 2
'                strDateConsruite = Format(cell, "00") & "/" & Format(m, "00") & "/" & Format(y, "0000")
'            Case 3
'                strDateConsruite = Format(Left(cell, 1), "00") & "/" & Format(Mid(cell, 2, 2), "00") & "/" & Format(y, "0000")
'            Case 4
'                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(y, "0000")
'            Case 6
'                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & "20" & Format(Mid(cell, 5, 2), "00")
'            Case 8
'                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(Mid(cell, 5, 4), "0000")
'            Case Else
'                dateValide = False
'        End Select
'        dateValide = IsDate(strDateConsruite)
'
'    If dateValide Then
'        r.value = Format(strDateConsruite, "dd/mm/yyyy")
'    Else
'        MsgBox "La saisie est invalide...", vbInformation, "Il est impossible de construire une date"
'    End If
'
'End Sub

Sub ChartOfAccount_Import_All() '2024-02-17 @ 07:21

    Dim timerStart As Double: timerStart = Timer
    
    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.Offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entr�e.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Entr�e.xlsx"
    sourceWorksheet = "PlanComptable"
    
    'ADODB connection
    Dim connStr As ADODB.Connection
    Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source = " & sourceWorkbook & ";" & _
                               "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset
    Set recSet = New ADODB.Recordset
    
    recSet.ActiveConnection = connStr
    recSet.source = "SELECT * FROM [" & sourceWorksheet & "$]"
    recSet.Open
    
    'Copy to wshAdmin workbook
    wshAdmin.Range("T11").CopyFromRecordset recSet
'    wshClientDB.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Call RedefineDynamicRange
    
'    MsgBox _
'        Prompt:="J'ai import� un total de " & _
'            Format(wshAdmin.Range("T10").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " comptes du Grand Livre", _
'        Title:="V�rification du nombre de comptes", _
'        Buttons:=vbInformation
        
    Call Output_Timer_Results("ChartOfAccount_Import_All()", timerStart)

End Sub

Sub RedefineDynamicRange() '2024-02-13 @ 13:30
    
    'Delete existing dynamic named range (assuming it exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptableDescription").delete
    On Error GoTo 0
    
    'Define a new dynamic named range
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create a new dynamic named range
    ThisWorkbook.Names.add name:="dnrPlanComptableDescription", RefersTo:=newRangeFormula
    
End Sub

Sub Hide_All_Worksheet_Except_Menu() '2024-02-20 @ 07:28
    
    Dim timerStart As Double: timerStart = Timer
    
    Dim wsh As Worksheet
    For Each wsh In ThisWorkbook.Worksheets
        If wsh.codeName <> "wshMenu" And _
            InStr(wsh.codeName, "wshzDoc") = 0 Then
                wsh.Visible = xlSheetHidden
        End If
    Next wsh
    
    Call Output_Timer_Results("Hide_All_Worksheet_Except_Menu()", timerStart)
    
End Sub

Sub LoopThroughRows()
    Dim i As Long, lastRow As Long
    Dim pctdone As Single
    lastRow = Range("A" & Rows.count).End(xlUp).row
    lastRow = 30

    '(Step 1) Display your Progress Bar
    ufProgress.LabelProgress.width = 0
    ufProgress.show
    For i = 1 To lastRow
        '(Step 2) Periodically update progress bar
        pctdone = i / lastRow
        With ufProgress
            .Caption = "�tape " & i & " of " & lastRow
            .LabelProgress.width = pctdone * (.FrameProgress.width)
        End With
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        '--------------------------------------
        'the rest of your macro goes below here
        '
        '
        '--------------------------------------
        '(Step 3) Close the progress bar when you're done
        If i = lastRow Then Unload ufProgress
    Next i
End Sub

Sub FractionComplete(pctdone As Single)
    With ufProgress
        .Caption = "Compl�t� � " & pctdone * 100 & "%"
        .LabelProgress.width = pctdone * (.FrameProgress.width)
    End With
    DoEvents
End Sub

Sub Fill_Or_Empty_Range_Background(rng As Range, fill As Boolean, Optional colorIndex As Variant = xlNone)
    If fill Then
        If IsMissing(colorIndex) Or colorIndex = xlNone Then
            rng.Interior.colorIndex = xlColorIndexNone ' Clear the background color
        Else
            rng.Interior.colorIndex = colorIndex ' Fill with specified color
        End If
    Else
        rng.Interior.colorIndex = xlColorIndexNone ' Clear the background color
    End If
End Sub

'Sub Tab_Order_Toggle_Mode()
'
'    isTab_Order_Activated = Not isTab_Order_Activated
'    isTab_Order_Activated = True
'
'End Sub

Sub Buttons_Enabled_True_Or_False(clear As Boolean, add As Boolean, _
                                  update As Boolean, delete As Boolean)
    With ufSaisieHeures
        .cmdClear.Enabled = clear
        .cmdAdd.Enabled = add
        .cmdUpdate.Enabled = update
        .cmdDelete.Enabled = delete
    End With

End Sub

Sub Invalid_Date_Message() '2024-03-03 @ 07:45

    If strDate = "" Then
    MsgBox Prompt:="La valeur saisie ne peut �tre utilis�e comme une date valide", _
        Title:="Validation de la date", _
        Buttons:=vbCritical

End Sub

Sub Pause_Application(s As Double)
    
    If s > 5 Then Stop
    
    Dim endTime As Double
    endTime = Timer + s 'Set end time to 's' seconds from now
    
    Do While Timer < endTime
        'Sleep
    Loop
    
End Sub


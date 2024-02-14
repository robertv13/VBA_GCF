Attribute VB_Name = "modApplication"
Option Explicit

Sub BackToMainMenu()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    wshMenu.Activate
    wshMenu.Range("A1").Select

End Sub

Sub Build_Date(r As Range) '2024-01-06 @ 18:29
        Dim d, m, Y As Integer
        Dim strDateConsruite As String, cell As String
        Dim dateValide As Boolean
        cell = Trim(r.value)
        dateValide = True

        cell = Replace(cell, "/", "")
        cell = Replace(cell, "-", "")

        'Utilisation de la date du jour
        d = Day(Now())
        m = Month(Now())
        Y = Year(Now())

        Select Case Len(cell)
            Case 0
                strDateConsruite = Format(d, "00") & "/" & Format(m, "00") & "/" & Format(Y, "0000")
            Case 1, 2
                strDateConsruite = Format(cell, "00") & "/" & Format(m, "00") & "/" & Format(Y, "0000")
            Case 3
                strDateConsruite = Format(Left(cell, 1), "00") & "/" & Format(Mid(cell, 2, 2), "00") & "/" & Format(Y, "0000")
            Case 4
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(Y, "0000")
            Case 6
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & "20" & Format(Mid(cell, 5, 2), "00")
            Case 8
                strDateConsruite = Format(Left(cell, 2), "00") & "/" & Format(Mid(cell, 3, 2), "00") & "/" & Format(Mid(cell, 5, 4), "0000")
            Case Else
                dateValide = False
        End Select
        dateValide = IsDate(strDateConsruite)

    If dateValide Then
        r.value = Format(strDateConsruite, "dd/mm/yyyy")
    Else
        MsgBox "La saisie est invalide...", vbInformation, "Il est impossible de construire une date"
    End If

End Sub

Sub ChartOfAccount_Import() '2024-02-13 @ 13:27
    
    'Clear all cells, but the headers, in the target worksheet
    wshAdmin.Range("T10").CurrentRegion.Offset(2, 0).ClearContents

    'Import Accounts List from 'GCF_BD_Entrée.xlsx, in order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                     "GCF_BD_Entrée.xlsx"
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
    recSet.Source = "SELECT * FROM [" & sourceWorksheet & "$]"
    recSet.Open
    
    'Copy to wshAdmin workbook
    wshAdmin.Range("T11").CopyFromRecordset recSet
'    wshClientDB.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    Call RedefineDynamicRange
    
'    MsgBox _
'        Prompt:="J'ai importé un total de " & _
'            Format(wshAdmin.Range("T10").CurrentRegion.Rows.count - 1, _
'            "## ##0") & " comptes du Grand Livre", _
'        Title:="Vérification du nombre de comptes", _
'        Buttons:=vbInformation
        
End Sub

Sub RedefineDynamicRange() '2024-02-13 @ 13:30
    
    'Delete existing dynamic named range (assuming it exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptableDescription").Delete
    On Error GoTo 0
    
    'Define a new dynamic named range
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create a new dynamic named range
    ThisWorkbook.Names.Add Name:="dnrPlanComptableDescription", RefersTo:=newRangeFormula
    
End Sub

Sub HideAllExceptActiveSheet() '2024-02-13 @ 13:23
    
    Dim wsh As Worksheet
    For Each wsh In ThisWorkbook.Worksheets
        If wsh.Name <> ActiveSheet.Name Then wsh.Visible = xlSheetHidden
    Next wsh
    
End Sub



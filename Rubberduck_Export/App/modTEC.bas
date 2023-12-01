Attribute VB_Name = "modTEC"
Option Explicit

Global Const rmv_modeInitial As Integer = 1
Global Const rmv_modeCreation As Integer = 2
Global Const rmv_modeAffichage As Integer = 3
Global Const rmv_modeModification As Integer = 4

Global rmv_state As Integer

Global savedClient As String
Global savedActivite As String
Global savedHeures As String
Global savedFacturable As String
Global savedCommNote As String

Global Const gAppVersion As String = "v0.1.4"

Sub ImportClientList()                                          '---------------- 2023-11-12 @ 07:28
    
    'Clear all cells, but the headers, in the worksheet
    wshClientDB.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    'Import Clients List from 'GCF_Clients.xlsx. In order to always have the LATEST version
    Dim sourceWorkbook As String, sourceWorksheet As String
    sourceWorkbook = ThisWorkbook.Path & Application.PathSeparator & _
                     "DataFiles" & Application.PathSeparator & _
                     "GCF_Clients.xlsx"
    sourceWorksheet = "Clients"
    
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
    
    'Copy to wshClientDB workbook
    wshClientDB.Range("A2").CopyFromRecordset recSet
    wshClientDB.Range("A:B").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
    MsgBox _
        Prompt:="J'ai importé un total de " & _
            Format(wshClientDB.Range("A1").CurrentRegion.Rows.count - 1, _
            "## ##0") & " clients", _
        Title:="Vérification du nombre de clients", _
        Buttons:=vbInformation
        
End Sub

Sub TEC_FilterAndSort()
    'You need the two Non Null Values to Filter
    If wshBaseHours.Range("R3").value = "" Or _
        wshBaseHours.Range("S3").value = "" Then
        Exit Sub
    End If
    
    With wshBaseHours
        Dim LastRow As Long, LastResultRow As Long, ResultRow As Long
        LastRow = .Range("A999999").End(xlUp).Row 'Last BaseHours Row
        If LastRow < 2 Then Exit Sub 'Nothing to filter
        'Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("A2:P" & LastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("R2:S3"), _
            CopyToRange:=.Range("U2:AJ2"), _
            Unique:=True
        LastResultRow = .Range("U99999").End(xlUp).Row
        If LastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wshBaseHours.Range("X3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add Key:=wshBaseHours.Range("U3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshBaseHours.Range("U3:AF" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
NoSort:
    End With
    'Application.ScreenUpdating = True
End Sub

'************************************************************** EffaceFormulaire
Sub EffaceFormulaire()

    'Empty the dynamic fields after reseting the form
    With frmSaisieHeures
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
        .cmbProfessionnel.Enabled = True
        .txtDate.Enabled = True
    End With
    
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdDelete.Enabled = False
        .cmdUpdate.Enabled = False
    End With
        
    frmSaisieHeures.txtClient.SetFocus
    
End Sub

'************************************************************* AjouteLigneDetail
Sub AjouteLigneDetail()

    'Validations first (one field at a time)
    If frmSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox Prompt:="Le professionnel est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.cmbProfessionnel.SetFocus
        Exit Sub
    End If

    If frmSaisieHeures.txtDate.value = "" Or _
        IsDate(frmSaisieHeures.txtDate.value) = False Then
            MsgBox Prompt:="La date est OBLIGATOIRE !", _
                   Title:="Vérification", _
                   Buttons:=vbCritical
            frmSaisieHeures.txtDate.SetFocus
            Exit Sub
    End If

    If frmSaisieHeures.txtClient.value = "" Then
        MsgBox Prompt:="Le client est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtClient.SetFocus
        Exit Sub
    End If
    
    If frmSaisieHeures.txtHeures.value = "" Or _
       IsNumeric(frmSaisieHeures.txtHeures.value) = False Then
        MsgBox Prompt:="Le nombre d'heures est OBLIGATOIRE !", _
               Title:="Vérification", _
               Buttons:=vbCritical
        frmSaisieHeures.txtHeures.SetFocus
        Exit Sub
    End If

    Dim LastRow As Long
    LastRow = wshBaseHours.Range("A999999").End(xlUp).Row

    'Load the cmb & txt into the 'HeuresBase' worksheet
    With wshBaseHours
        .Range("A" & LastRow + 1).value = LastRow
        .Range("B" & LastRow + 1).value = wshAdmin.Range("Prof_ID")
        .Range("C" & LastRow + 1).value = frmSaisieHeures.cmbProfessionnel.value
        .Range("D" & LastRow + 1).value = CDate(frmSaisieHeures.txtDate.value)
        .Range("E" & LastRow + 1).value = wshAdmin.Range("Client_ID_Admin")
        .Range("F" & LastRow + 1).value = frmSaisieHeures.txtClient.value
        .Range("G" & LastRow + 1).value = frmSaisieHeures.txtActivite.value
        .Range("H" & LastRow + 1).value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
        .Range("I" & LastRow + 1).value = frmSaisieHeures.txtCommNote.value
        .Range("J" & LastRow + 1).value = frmSaisieHeures.chbFacturable.value
        .Range("K" & LastRow + 1).value = Now
        .Range("L" & LastRow + 1).value = False
        .Range("M" & LastRow + 1).value = ""
        .Range("N" & LastRow + 1).value = False
        .Range("O" & LastRow + 1).value = gAppVersion
        .Range("P" & LastRow + 1).value = ""
    End With

    'Empty the fields after saving
    frmSaisieHeures.txtClient.value = ""
    frmSaisieHeures.txtActivite.value = ""
    frmSaisieHeures.txtHeures.value = ""
    frmSaisieHeures.txtCommNote.value = ""
        
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdUpdate.Enabled = False
    End With
    
    frmSaisieHeures.txtClient.SetFocus
    
End Sub

'Sub RemoveTotalRow()
'
'    Dim tbl As ListObject
'    Set tbl = wshBaseHours.ListObjects("tCharges")
'    If tbl.ShowTotals = True Then
'        tbl.TotalsRowRange.Delete
'    End If
'
'    'Resize
'    '    With tbl.Range
'    '        tbl.Resize .Resize(.CurrentRegion.Rows.count + 1)
'    '        .Cells(.CurrentRegion.Rows.count + 1, 1).value = ""
'    '    End With
'
'End Sub

'************************************************************ ModifieLigneDetail
Sub ModifieLigneDetail()

    If frmSaisieHeures.txtID.value = "" Then
        MsgBox _
        Prompt:="Vous devez choisir un enregistrement à modifier !", _
        Title:="", _
        Buttons:=vbCritical
        Exit Sub
    End If
    
    'Validations first (one field at a time)
    If frmSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox _
        Prompt:="Le professionnel est OBLIGATOIRE !", _
        Title:="Vérification", _
        Buttons:=vbCritical
        frmSaisieHeures.cmbProfessionnel.SetFocus
        Exit Sub
    End If

    If frmSaisieHeures.txtDate.value = "" Or _
       IsDate(frmSaisieHeures.txtDate.value) = False Then
        MsgBox _
        Prompt:="La date est OBLIGATOIRE !", _
        Title:="Vérification", _
        Buttons:=vbCritical
        frmSaisieHeures.txtDate.SetFocus
        Exit Sub
    End If

    If frmSaisieHeures.txtClient.value = "" Then
        MsgBox _
        Prompt:="Le client est OBLIGATOIRE !", _
        Title:="Vérification", _
        Buttons:=vbCritical
        frmSaisieHeures.txtClient.SetFocus
        Exit Sub
    End If
    
    If frmSaisieHeures.txtHeures.value = "" Or _
       IsNumeric(frmSaisieHeures.txtHeures.value) = False Then
        MsgBox _
        Prompt:="Le nombre d'heures est OBLIGATOIRE !", _
        Title:="Vérification", _
        Buttons:=vbCritical
        frmSaisieHeures.txtHeures.SetFocus
        Exit Sub
    End If

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("HeuresBase")

    Dim selectedRow As Long
    selectedRow = Application.WorksheetFunction.Match(CLng(frmSaisieHeures.txtID.value), _
                                                      sh.Range("A:A"), 0)
    
    With frmSaisieHeures
        sh.Range("B" & selectedRow).value = .cmbProfessionnel.value
        sh.Range("C" & selectedRow).value = CDate(.txtDate.value)
        sh.Range("D" & selectedRow).value = .txtClient.value
        sh.Range("E" & selectedRow).value = .txtActivite.value
        sh.Range("F" & selectedRow).value = Format(.txtHeures.value, "#0.00")
        sh.Range("G" & selectedRow).value = .txtCommNote.value
        sh.Range("H" & selectedRow).value = .chbFacturable.value
        sh.Range("I" & selectedRow).value = Now
        sh.Range("J" & selectedRow).value = False
        sh.Range("K" & selectedRow).value = ""
        sh.Range("L" & selectedRow).value = False
        
        frmSaisieHeures.txtClient.value = ""
        frmSaisieHeures.txtActivite.value = ""
        frmSaisieHeures.txtHeures.value = ""
        frmSaisieHeures.txtCommNote.value = ""
    End With
   
    frmSaisieHeures.cmbProfessionnel.Enabled = True
    frmSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation

    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHours
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

'************************************************************* EffaceLigneDetail
Sub EffaceLigneDetail()

    If frmSaisieHeures.txtID.value = "" Then
        MsgBox _
        Prompt:="Vous devez choisir un enregistrement à DÉTRUIRE !", _
        Title:="", _
        Buttons:=vbCritical
        Exit Sub
    End If
    
    Dim answerYesNo As Integer
    answerYesNo = MsgBox("Êtes-vous certain de vouloir DÉTRUIRE cet enregistrement ? ", _
                         vbYesNo + vbQuestion, "Confirmation de DESTRUCTION")
    If answerYesNo = vbNo Then
        MsgBox _
        Prompt:="Cet enregistrement ne sera PAS détruit ! ", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("HeuresBase")
    
    Dim selectedRow As Long
    selectedRow = Application.WorksheetFunction.Match(CLng(frmSaisieHeures.txtID.value), _
                                                      sh.Range("A:A"), 0)
    
    'Assign 'VRAI' to colomn 12, since it is deleted
    sh.Range("I" & selectedRow).value = Now
    sh.Range("L" & selectedRow).value = True
    
    'Empty the dynamic fields after deleting
    With frmSaisieHeures
        .txtClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
    End With
    
    MsgBox _
        Prompt:="L'enregistrement a été DÉTRUIT !", _
        Title:="Confirmation", _
        Buttons:=vbCritical
        
    frmSaisieHeures.cmbProfessionnel.Enabled = True
    frmSaisieHeures.txtDate.Enabled = True
    rmv_state = rmv_modeCreation
    
    Call TEC_FilterAndSort
    Call RefreshListBoxAndAddHoursAndAddHours
    
    frmSaisieHeures.txtClient.SetFocus

End Sub

'********************* Reload listBox from HeuresFiltered and reset the buttons
Sub RefreshListBoxAndAddHours()

    If wshAdmin.Range("B4").value = "" Or wshAdmin.Range("B5").value = "" Then
        GoTo EndOfProcedure
    End If
    
    frmSaisieHeures.txtTotalHeures.value = ""
    
    Dim shFiltered As Worksheet
    Set shFiltered = ThisWorkbook.Sheets("HeuresBase")
    shFiltered.Activate
    
    'Last Row used in column A
    Dim LastRow As Long
    LastRow = wshBaseHours.Range("T2:T9999").End(xlUp).Row - 1
    If LastRow = 0 Then Exit Sub
        
    With frmSaisieHeures.lstData
        .ColumnHeads = True
        .ColumnCount = 9
        .ColumnWidths = "22; 28; 52; 120; 190; 35; 80; 30; 75"
        
        If LastRow = 1 Then
            .RowSource = "HeuresBase!T3:Z3"
        Else
            .RowSource = "HeuresBase!T3:Z" & LastRow + 1
        End If
    End With

    'Add hours to totalHeures
    Dim nbrRows, i As Integer
    nbrRows = frmSaisieHeures.lstData.ListCount
    Dim totalHeures As Double
    
    If nbrRows > 0 Then
        For i = 0 To nbrRows - 1
            totalHeures = totalHeures + CCur(frmSaisieHeures.lstData.List(i, 3))
        Next
        frmSaisieHeures.txtTotalHeures.value = Format(totalHeures, "#0.00")
    End If

EndOfProcedure:
    frmSaisieHeures.cmdClear.Enabled = False
    frmSaisieHeures.cmdAdd.Enabled = False
    frmSaisieHeures.cmdUpdate.Enabled = False
    frmSaisieHeures.cmdDelete.Enabled = False

    'frmSaisieHeures.txtClient.SetFocus
    
End Sub



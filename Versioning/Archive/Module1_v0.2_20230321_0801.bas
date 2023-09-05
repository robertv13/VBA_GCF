Attribute VB_Name = "Module1"
Option Explicit

Sub FilterProfDate()

    'Minimum - Professionnel + Date
    If Trim(frmSaisieHeures.cmbProfessionnel.value) = "" Or _
        Trim(frmSaisieHeures.txtDate.value) = "" Then
        Exit Sub
    End If
    
    'Base worksheet contains all entries
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")
    sh.AutoFilterMode = False
    
    'Filtered worksheet 'Heures_Work'
    Dim shFiltered As Worksheet
    Set shFiltered = ThisWorkbook.Sheets("HeuresFiltered")
    shFiltered.UsedRange.Clear
    shFiltered.Activate

    sh.Activate
    sh.UsedRange.AutoFilter 2, frmSaisieHeures.cmbProfessionnel.value
    sh.UsedRange.AutoFilter 3, frmSaisieHeures.txtDate.value
    sh.UsedRange.Select
    
    sh.UsedRange.Copy shFiltered.Range("A1")
    shFiltered.Activate
    
    sh.AutoFilterMode = False
    'sh.ShowAllData

End Sub

Sub ImportClientsList()

    'Delete all cells, but the headers in the destination worksheet
    shImportedClients.Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    'Source workbook (closed Excel file) - MUST BE IN THE SAME DIRECTORY
    Dim sourceWorkbook, sourceWorksheet As String
    sourceWorkbook = ThisWorkbook.Path & Application.PathSeparator & _
                     "GCF_Clients.xlsx"
    sourceWorksheet = "Clients"
    
    'ADODB connection
    Dim connStr As ADODB.Connection
    Set connStr = New ADODB.Connection
    
    'Connection String specific to EXCEL
    connStr.ConnectionString = _
        "Provider = Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source = " & sourceWorkbook & ";" & _
        "Extended Properties = 'Excel 12.0 Xml; HDR = YES';"
    connStr.Open
    
    'Recordset
    Dim recSet As ADODB.Recordset
    Set recSet = New ADODB.Recordset
    
    recSet.ActiveConnection = connStr
    recSet.Source = "SELECT Nom FROM [" & sourceWorksheet & "$]"
        
    recSet.Open
    
    'Copy to destination workbook (actual) into the 'Top2000' worksheet
    shImportedClients.Range("A2").CopyFromRecordset recSet
    
    shImportedClients.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close resource
    recSet.Close
    connStr.Close
    
End Sub

'************************************************************** EffaceFormulaire
Sub EffaceFormulaire()

    'Empty the dynamic fields after reseting the form
    With frmSaisieHeures
        .cmbClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
    End With
    
    Call FilterProfDate
    Call RefreshListBox
    
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdDelete.Enabled = False
        .cmdUpdate.Enabled = False
    End With
        
    frmSaisieHeures.cmbClient.SetFocus
    
End Sub

'************************************************************* AjouteLigneDetail
Sub AjouteLigneDetail()

    'MsgBox "Temp - Sub cmdAdd_Click()"
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")
        
    Dim lastRow As Long
    lastRow = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    
    'Validations first (one field at a time)
    If frmSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox "Le professionnel est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.cmbProfessionnel.SetFocus
        Exit Sub
    End If

    If frmSaisieHeures.txtDate.value = "" Then
        MsgBox "La date est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.txtDate.SetFocus
        Exit Sub
    End If

    If frmSaisieHeures.cmbClient.value = "" Then
        MsgBox "Le nom du client est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.cmbClient.SetFocus
    Exit Sub
    End If
    
    If frmSaisieHeures.txtHeures.value = "" Then
        MsgBox "Le nombre d'heures est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.txtHeures.SetFocus
        Exit Sub
    End If

    'Load the cmb & txt into the 'Heures' worksheet
    With sh
        .Range("A" & lastRow + 1).value = "=row()-1"
        .Range("B" & lastRow + 1).value = frmSaisieHeures.cmbProfessionnel.value
        .Range("C" & lastRow + 1).value = Format(frmSaisieHeures.txtDate.value, "dd-mm-yyyy")
        .Range("D" & lastRow + 1).value = frmSaisieHeures.cmbClient.value
        .Range("E" & lastRow + 1).value = frmSaisieHeures.txtActivite.value
        .Range("F" & lastRow + 1).value = Format(frmSaisieHeures.txtHeures.value, "#0.00")
        .Range("G" & lastRow + 1).value = frmSaisieHeures.txtCommNote.value
        .Range("H" & lastRow + 1).value = frmSaisieHeures.chbFacturable.value
        .Range("I" & lastRow + 1).value = Now
        .Range("J" & lastRow + 1).value = False
        .Range("K" & lastRow + 1).value = ""
    End With

    'Empty the fields after saving
    frmSaisieHeures.cmbClient.value = ""
    frmSaisieHeures.txtActivite.value = ""
    frmSaisieHeures.txtHeures.value = ""
    frmSaisieHeures.txtCommNote.value = ""
        
    Call FilterProfDate
    Call RefreshListBox
    
    With frmSaisieHeures
        .cmdClear.Enabled = False
        .cmdAdd.Enabled = False
        .cmdUpdate.Enabled = False
    End With
    
    frmSaisieHeures.cmbClient.SetFocus
    
End Sub

'************************************************************ ModifieLigneDetail
Sub ModifieLigneDetail()

    If frmSaisieHeures.txtID.value = "" Then
        MsgBox "Vous devez choisir un enregistrement à modifier !"
        Exit Sub
    End If
    
    'Validations first (one field at a time)
    If frmSaisieHeures.cmbProfessionnel.value = "" Then
        MsgBox "Le professionnel est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.cmbProfessionnel.SetFocus
        Exit Sub
    End If

    If frmSaisieHeures.txtDate.value = "" Then
        MsgBox "La date est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.txtDate.SetFocus
        Exit Sub
    End If

    If frmSaisieHeures.cmbClient.value = "" Then
        MsgBox "Le nom du client est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.cmbClient.SetFocus
    Exit Sub
    End If
    
    If frmSaisieHeures.txtHeures.value = "" Then
        MsgBox "Le nombre d'heures est OBLIGATOIRE !", vbCritical
        frmSaisieHeures.txtHeures.SetFocus
        Exit Sub
    End If

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")

    Dim selectedRow As Long
    selectedRow = Application.WorksheetFunction.Match(CLng(frmSaisieHeures.txtID.value), _
                    sh.Range("A:A"), 0)
    
    With frmSaisieHeures
        sh.Range("B" & selectedRow).value = .cmbProfessionnel.value
        sh.Range("C" & selectedRow).value = Format(.txtDate.value, "dd-mm-yyyy")
        sh.Range("D" & selectedRow).value = .cmbClient.value
        sh.Range("E" & selectedRow).value = .txtActivite.value
        sh.Range("F" & selectedRow).value = Format(.txtHeures.value, "#0.00")
        sh.Range("G" & selectedRow).value = .txtCommNote.value
        sh.Range("H" & selectedRow).value = .chbFacturable.value
        sh.Range("I" & selectedRow).value = Now
        sh.Range("J" & selectedRow).value = False
        sh.Range("K" & selectedRow).value = ""
        
        frmSaisieHeures.cmbClient.value = ""
        frmSaisieHeures.txtActivite.value = ""
        frmSaisieHeures.txtHeures.value = ""
        frmSaisieHeures.txtCommNote.value = ""
   End With
   
    'Empty the fields after modifying
    
    Call FilterProfDate
    Call RefreshListBox
    
    frmSaisieHeures.cmbClient.SetFocus

End Sub

'************************************************************* EffaceLigneDetail
Sub EffaceLigneDetail()

    If frmSaisieHeures.txtID.value = "" Then
        MsgBox "Vous devez choisir un enregistrement à DÉTRUIRE !"
        Exit Sub
    End If
    
    Dim answerYesNo As Integer
    answerYesNo = MsgBox("Êtes-vous certain de vouloir DÉTRUIRE cet enregistrement ? ", _
                                vbYesNo + vbQuestion, "Confirmation de DESTRUCTION")
    If answerYesNo = vbNo Then
        MsgBox "Cet enregistrement ne sera PAS détruit ! "
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Heures")
    
    Dim selectedRow As Long
    selectedRow = Application.WorksheetFunction.Match(CLng(frmSaisieHeures.txtID.value), _
                              sh.Range("A:A"), 0)
    
    sh.Range("A" & selectedRow).EntireRow.Delete
    
    'Empty the dynamic fields after deleting
    With frmSaisieHeures
        .cmbClient.value = ""
        .txtActivite.value = ""
        .txtHeures.value = ""
        .txtCommNote.value = ""
    End With

    MsgBox "L'enregistrement a été DÉTRUIT !"

    Call FilterProfDate
    Call RefreshListBox
    
    frmSaisieHeures.cmbClient.SetFocus

End Sub

'********************* Reload listBox from HeuresFiltered and reset the buttons
Sub RefreshListBox()

    If Trim(frmSaisieHeures.cmbProfessionnel) = "" _
            Or Trim(frmSaisieHeures.txtDate) = "" Then
        GoTo EndOfProcedure
    End If
    
    frmSaisieHeures.txtTotalHeures.value = ""
    
    Dim shFiltered As Worksheet
    Set shFiltered = ThisWorkbook.Sheets("HeuresFiltered")
    shFiltered.Activate
    
    'Last Row used in column A
    Dim lastRow As Long
    lastRow = Application.WorksheetFunction.CountA(shFiltered.Range("A:A"))
        
    With frmSaisieHeures.lstData
        .ColumnHeads = True
        .ColumnCount = 9
        .ColumnWidths = "25; 30; 55; 120; 190; 30; 90; 35; 70"
        
        If lastRow = 1 Then
            .RowSource = "HeuresFiltered!A2:K2"
        Else
            .RowSource = "HeuresFiltered!A2:K" & lastRow
        End If
    End With

    'Add hours to totalHeures
    Dim nbrRows, i As Integer
    nbrRows = frmSaisieHeures.lstData.ListCount
    Dim totalHeures As Double
    
    If nbrRows > 0 Then
        For i = 0 To nbrRows - 1
            totalHeures = totalHeures + CCur(frmSaisieHeures.lstData.List(i, 5))
        Next
        frmSaisieHeures.txtTotalHeures.value = Format(totalHeures, "#0.00")
    End If

EndOfProcedure:
    frmSaisieHeures.cmdClear.Enabled = False
    frmSaisieHeures.cmdAdd.Enabled = False
    frmSaisieHeures.cmdUpdate.Enabled = False
    frmSaisieHeures.cmdDelete.Enabled = False

    'frmSaisieHeures.cmbClient.SetFocus
    
End Sub





Attribute VB_Name = "Module1"
Option Explicit

Sub Show_Form()
    
    frmForm.Show

End Sub

Sub Reset()

    Dim iRow As Long
    iRow = [Counta(Donn�es!A:A)] 'Identifying the last row
    
    With frmForm
        .txtCodeClient.Value = ""
        .txtNomClient.Value = ""
        .txtContactFact.Value = ""
        .txtTitreContact.Value = ""
        .txtCourrielFact.Value = ""
        .txtAdresse1.Value = ""
        .txtAdresse2.Value = ""
        .txtVille.Value = ""
        .txtProvince.Value = ""
        .txtCodePostal.Value = ""
        .txtPays.Value = ""
        .txtReferePar.Value = ""
        .txtFinAnnee.Value = ""
        .txtComptable.Value = ""
        .txtNotaireAvocat.Value = ""
        
        'Default Color
        .txtCodeClient.BackColor = vbWhite
        .txtNomClient.BackColor = vbWhite
        .txtContactFact.BackColor = vbWhite
        .txtTitreContact.BackColor = vbWhite
        .txtCourrielFact.BackColor = vbWhite
        .txtAdresse1.BackColor = vbWhite
        .txtAdresse2.BackColor = vbWhite
        .txtVille.BackColor = vbWhite
        .txtProvince.BackColor = vbWhite
        .txtCodePostal.BackColor = vbWhite
        .txtPays.BackColor = vbWhite
        .txtReferePar.BackColor = vbWhite
        .txtFinAnnee.BackColor = vbWhite
        .txtComptable.BackColor = vbWhite
        .txtNotaireAvocat.BackColor = vbWhite
        
        .txtRowNumber.Value = ""
        
        'Below code are associated with Search Feature - Part 3
        Call Add_SearchColumn
        ThisWorkbook.Sheets("Donn�es").AutoFilterMode = False
        ThisWorkbook.Sheets("Donn�esRecherche").AutoFilterMode = False
        ThisWorkbook.Sheets("Donn�esRecherche").Cells.Clear
        
        .lstDatabase.ColumnCount = 15
        .lstDatabase.ColumnHeads = True
        
        .lstDatabase.ColumnWidths = "200; 45; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 70; 105; 105"
        
        If iRow > 1 Then
            .lstDatabase.RowSource = "Donn�es!A2:O" & iRow
        Else
            .lstDatabase.RowSource = "Donn�es!A2:O2"
        End If
    End With

End Sub

Sub Submit()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Donn�es")
    
    Dim iRow As Long
    If frmForm.txtRowNumber.Value = "" Then
        iRow = [Counta(Donn�es!A:A)] + 1
    Else
        iRow = frmForm.txtRowNumber.Value
    End If
    
    With sh
        .Cells(iRow, 1) = frmForm.txtNomClient.Value
        .Cells(iRow, 2) = frmForm.txtCodeClient.Value
        .Cells(iRow, 3) = frmForm.txtContactFact.Value
        .Cells(iRow, 4) = frmForm.txtTitreContact.Value
        .Cells(iRow, 5) = frmForm.txtCourrielFact.Value
        .Cells(iRow, 6) = frmForm.txtAdresse1.Value
        .Cells(iRow, 7) = frmForm.txtAdresse2.Value
        .Cells(iRow, 8) = frmForm.txtVille.Value
        .Cells(iRow, 9) = frmForm.txtProvince.Value
        .Cells(iRow, 10) = frmForm.txtCodePostal.Value
        .Cells(iRow, 11) = frmForm.txtPays.Value
        .Cells(iRow, 12) = frmForm.txtReferePar.Value
        .Cells(iRow, 13) = frmForm.txtFinAnnee.Value
        .Cells(iRow, 14) = frmForm.txtComptable.Value
        .Cells(iRow, 15) = frmForm.txtNotaireAvocat.Value
'        .Cells(iRow, 8) = Application.UserName
'        .Cells(iRow, 9) = [Text(Now(), "DD-MM-YYYY HH:MM:SS")]
    End With

End Sub

Function Selected_List() As Long

    Selected_List = 0
    
    Dim i As Long
    For i = 0 To frmForm.lstDatabase.ListCount - 1
        If frmForm.lstDatabase.Selected(i) = True Then
            Selected_List = i + 1
            frmForm.cmdEdit.Enabled = True
            Exit For
        End If
        frmForm.cmdEdit.Enabled = False
    Next i

End Function

Sub Add_SearchColumn()

    frmForm.EnableEvents = False

    With frmForm.cmbSearchColumn
        .Clear
        .AddItem "Tous"
        
        .AddItem "ClientNom"
        .AddItem "Client_ID"
        .AddItem "ContactFacturation"
        .AddItem "TitreContactFacturation"
        .AddItem "CourrielFacturation"
        .AddItem "Adresse_1"
        .AddItem "Adresse_2"
        .AddItem "Ville"
        .AddItem "Province"
        .AddItem "CodePostal"
        .AddItem "Pays"
        .AddItem "R�f�r� par"
        .AddItem "Fin d'ann�e"
        .AddItem "Comptable"
        .AddItem "Notaire/Avocat"
        
        .Value = "Client_ID"
    End With
    
    frmForm.EnableEvents = True
    
    frmForm.txtSearch.Value = ""
    frmForm.txtSearch.Enabled = True
'    frmForm.txtSearch.Enabled = False
    frmForm.cmdSearch.Enabled = True
'    frmForm.cmdSearch.Enabled = False

End Sub

Sub Donn�esRecherche()

    Application.ScreenUpdating = False
    
    Dim iColumn As Integer 'To hold the selected column number in Donn�es sheet
    Dim iDonn�esRow As Long 'To store the last non-blank row number available in Donn�es sheet
    Dim iSearchRow As Long 'To hold the last non-blank row number available in SearachData sheet
    
    Dim sColumn As String 'To store the column selection
    Dim sValue As String 'To hold the search text value
    
    Dim wshDonn�es As Worksheet 'Donn�es sheet
    Set wshDonn�es = ThisWorkbook.Sheets("Donn�es")
    Dim wshSearchData As Worksheet 'Donn�esRecherche sheet
    Set wshSearchData = ThisWorkbook.Sheets("Donn�esRecherche")
    
    iDonn�esRow = ThisWorkbook.Sheets("Donn�es").Range("A" & Application.Rows.Count).End(xlUp).Row
    sColumn = frmForm.cmbSearchColumn.Value
    sValue = frmForm.txtSearch.Value
    iColumn = Application.WorksheetFunction.Match(sColumn, wshDonn�es.Range("A1:O1"), 0)
    
    'Remove filter from Donn�es worksheet
    If wshDonn�es.FilterMode = True Then
        wshDonn�es.AutoFilterMode = False
    End If

    'Apply filter on Donn�es worksheet
    If frmForm.cmbSearchColumn.Value = "Code Client" Then
        wshDonn�es.Range("A1:O" & iDonn�esRow).AutoFilter Field:=iColumn, Criteria1:=sValue
    Else
        wshDonn�es.Range("A1:O" & iDonn�esRow).AutoFilter Field:=iColumn, Criteria1:="*" & sValue & "*"
    End If
    
    Dim searchRowsFound As Long
    searchRowsFound = Application.WorksheetFunction.Subtotal(3, wshDonn�es.Range("A:A")) - 1 'Heading
    If searchRowsFound >= 1 Then
        'Code to remove the previous data from Donn�esRecherche worksheet
        wshSearchData.Cells.Clear
        wshDonn�es.AutoFilter.Range.Copy wshSearchData.Range("A1")
        Application.CutCopyMode = False
        iSearchRow = wshSearchData.Range("A" & Application.Rows.Count).End(xlUp).Row
        frmForm.lstDatabase.ColumnCount = 15
        frmForm.lstDatabase.ColumnWidths = "200; 45; 110; 110; 150; 130; 90; 95; 40; 55; 80; 100; 70; 105; 105"
        If iSearchRow > 1 Then
            frmForm.lstDatabase.RowSource = "Donn�esRecherche!A2:O" & iSearchRow
'            MsgBox "J'ai trouv� " & searchRowsFound & " enregistrements."
        End If
    Else
       MsgBox "Je n'ai trouv� AUCUN enregistrement."
    End If

    wshDonn�es.AutoFilterMode = False
    Application.ScreenUpdating = True

End Sub

Function ValidateEntries() As Boolean

    ValidateEntries = True
    
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Donn�es")

    Dim iCodeClient As Variant
    iCodeClient = frmForm.txtCodeClient.Value
    
    With frmForm
        'Default Color
        .txtCodeClient.BackColor = vbWhite
        .txtNomClient.BackColor = vbWhite
        .txtContactFact.BackColor = vbWhite
        .txtTitreContact.BackColor = vbWhite
        .txtCourrielFact.BackColor = vbWhite
        .txtAdresse1.BackColor = vbWhite
        .txtAdresse2.BackColor = vbWhite
        .txtVille.BackColor = vbWhite
        .txtProvince.BackColor = vbWhite
        .txtCodePostal.BackColor = vbWhite
        .txtPays.BackColor = vbWhite
        .txtReferePar.BackColor = vbWhite
        .txtFinAnnee.BackColor = vbWhite
        .txtComptable.BackColor = vbWhite
        .txtNotaireAvocat.BackColor = vbWhite
        
        'Valeur OBLIGATOIRE
        If Trim(.txtCodeClient.Value) = "" Then
            MsgBox "SVP, saisir un code de client.", vbOKOnly + vbInformation, "Code de client"
            ValidateEntries = False
            .txtCodeClient.BackColor = vbRed
            .txtCodeClient.Enabled = True
            .txtCodeClient.SetFocus
            Exit Function
        End If
    
'        'Validating Duplicate Entries
'        If Not sh.Range("B:B").Find(what:=iCodeClient, lookat:=xlWhole) Is Nothing Then
'            MsgBox "Ce code de client existe d�j�.", vbOKOnly + vbInformation, "Doublon de code de client"
'            ValidateEntries = False
'            .txtCodeClient.BackColor = vbRed
'            .txtCodeClient.SetFocus
'            Exit Function
'        End If
        
        'Valeur OBLIGATOIRE
        If Trim(.txtNomClient.Value) = "" Then
            MsgBox "SVP, saisir le nom du client.", vbOKOnly + vbInformation, "Nom de client"
            ValidateEntries = False
            .txtNomClient.BackColor = vbRed
            .txtNomClient.SetFocus
            Exit Function
        End If
        
    End With

End Function


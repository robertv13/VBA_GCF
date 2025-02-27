Attribute VB_Name = "modFAC_Finale"
Option Explicit

Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub shp_FAC_Finale_Save_Click()

    Call FAC_Finale_Save

End Sub

Sub FAC_Finale_Save() '2024-03-28 @ 07:19

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Save", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)

    With wshFAC_Brouillon
        'Check For Mandatory Fields - Client
        If .Range("B18").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
        
        'Check For Mandatory Fields - Date de facture
        If .Range("O3").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture AVANT de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
        
        'Check For Mandatory Fields - Date de facture
        If Len(Trim(.Range("O6").value)) <> 8 Then
            MsgBox "Il faut corriger le numéro de facture AVANT de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
    End With
            
    'Valid Invoice - Let's update it ******************************************
    
    Call FAC_Finale_Disable_Save_Button

    Call FAC_Finale_Add_Invoice_Header_to_DB
    Call FAC_Finale_Add_Invoice_Header_Locally
    
    Call FAC_Finale_Add_Invoice_Details_to_DB
    Call FAC_Finale_Add_Invoice_Details_Locally
    
    Call FAC_Finale_Add_Invoice_Somm_Taux_to_DB
    Call FAC_Finale_Add_Invoice_Somm_Taux_Locally
    
    Call FAC_Finale_Add_Comptes_Clients_to_DB
    Call FAC_Finale_Add_Comptes_Clients_Locally
    
    Dim lastResultRow As Long
    lastResultRow = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "AQ").End(xlUp).row
        
    If lastResultRow > 2 Then
        Call FAC_Finale_TEC_Update_As_Billed_To_DB(3, lastResultRow)
        Call FAC_Finale_TEC_Update_As_Billed_Locally(3, lastResultRow)
    End If
    
    'Update FAC_Projets_Entête & FAC_Projets_Détails, if necessary
    Dim projetID As Long
    projetID = wshFAC_Brouillon.Range("B52").value
    If projetID <> 0 Then
        Call FAC_Finale_Softdelete_Projets_Détails_To_DB(projetID)
        Call FAC_Finale_Softdelete_Projets_Détails_Locally(projetID)
        
        Call FAC_Finale_Softdelete_Projets_Entête_To_DB(projetID)
        Call FAC_Finale_Softdelete_Projets_Entête_Locally(projetID)
    End If
        
    'Save Invoice total amount
    Dim invoice_Total As Currency
    invoice_Total = wshFAC_Brouillon.Range("O51").value
        
    'GL stuff will occur at the confirmation level (later)
'    Call FAC_Finale_GL_Posting_Preparation
    
    'Update TEC_DashBoard
    Call ActualiserTEC_TDB '2024-03-21 @ 12:32

    Call FAC_Brouillon_Clear_All_TEC_Displayed
    
    Application.ScreenUpdating = True
    
    MsgBox "La facture '" & wshFAC_Brouillon.Range("O6").value & "' est enregistrée." & _
        vbNewLine & vbNewLine & "Le total de la facture est " & _
        Trim(Format$(invoice_Total, "### ##0.00 $")) & _
        " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    
    wshFAC_Brouillon.Select
    Application.Wait (Now + TimeValue("0:00:02"))
    wshFAC_Brouillon.Range("E3").value = "" 'Reset client to empty
    
    wshFAC_Brouillon.Range("B27").value = False
    
    Call FAC_Brouillon_New_Invoice '2024-03-12 @ 08:08 - Maybe ??
    
Fast_Exit_Sub:

    wshFAC_Brouillon.Select
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Save", "", startTime)
    
End Sub

Sub FAC_Finale_Add_Invoice_Header_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Entête$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Can only ADD to the file, no modification is allowed
    
    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields(fFacEInvNo - 1).value = .Range("E28").value
        rs.Fields(fFacEDateFacture - 1).value = Format$(wshFAC_Brouillon.Range("O3").value, "yyyy-mm-dd")
        rs.Fields(fFacEACouC - 1).value = "AC" 'Facture to be confirmed MANUALLY - 2024-08-16 @ 05:46
        rs.Fields(fFacECustID - 1).value = wshFAC_Brouillon.Range("B18").value
        rs.Fields(fFacEContact - 1).value = .Range("B23").value
        rs.Fields(fFacENomClient - 1).value = .Range("B24").value
        rs.Fields(fFacEAdresse1 - 1).value = .Range("B25").value
        rs.Fields(fFacEAdresse2 - 1).value = .Range("B26").value
        rs.Fields(fFacEAdresse3 - 1).value = .Range("B27").value
        
        rs.Fields(fFacEHonoraires - 1).value = Format$(.Range("E69").value, "0.00")
        
        rs.Fields(fFacEAF1Desc - 1).value = .Range("B70").value
        rs.Fields(fFacEAutresFrais1 - 1).value = Format$(wshFAC_Finale.Range("E70").value, "0.00")
        rs.Fields(fFacEAF2Desc - 1).value = .Range("B71").value
        rs.Fields(fFacEAutresFrais2 - 1).value = Format$(.Range("E71").value, "0.00")
        rs.Fields(fFacEAF3Desc - 1).value = .Range("B72").value
        rs.Fields(fFacEAutresFrais3 - 1).value = Format$(.Range("E72").value, "0.00")
        
        rs.Fields(fFacETauxTPS - 1).value = Format$(.Range("C74").value, "0.00")
        rs.Fields(fFacEMntTPS - 1).value = Format$(.Range("E74").value, "0.00")
        rs.Fields(fFacETauxTVQ - 1).value = Format$(.Range("C75").value, "0.00000") '2024-10-15 @ 05:49
        rs.Fields(fFacEMntTVQ - 1).value = Format$(.Range("E75").value, "0.00")
        
        rs.Fields(fFacEARTotal - 1).value = Format$(.Range("E77").value, "0.00")
        
        rs.Fields(fFacEDépôt - 1).value = Format$(.Range("E79").value, "0.00")
        rs.Fields(fFacETimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") '2025-01-25 @ 15:01
    End With
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set rs = Nothing
    Set conn = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_to_DB", "", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Header_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Entête.Cells(wshFAC_Entête.Rows.count, "A").End(xlUp).row + 1
    
    With wshFAC_Entête
        .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).value = Format$(wshFAC_Brouillon.Range("O3").value, "mm-dd-yyyy")
        .Range("C" & firstFreeRow).value = "AC"
        .Range("D" & firstFreeRow).value = wshFAC_Brouillon.Range("B18").value
        .Range("E" & firstFreeRow).value = wshFAC_Finale.Range("B23").value
        .Range("F" & firstFreeRow).value = wshFAC_Finale.Range("B24").value
        .Range("G" & firstFreeRow).value = wshFAC_Finale.Range("B25").value
        .Range("H" & firstFreeRow).value = wshFAC_Finale.Range("B26").value
        .Range("I" & firstFreeRow).value = wshFAC_Finale.Range("B27").value
        
        .Range("J" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E69").value, "0.00")
        
        .Range("K" & firstFreeRow).value = wshFAC_Finale.Range("B70").value
        .Range("L" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E70").value, "0.00")
        .Range("M" & firstFreeRow).value = wshFAC_Finale.Range("B71").value
        .Range("N" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E71").value, "0.00")
        .Range("O" & firstFreeRow).value = wshFAC_Finale.Range("B72").value
        .Range("P" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E72").value, "0.00")
        
        .Range("Q" & firstFreeRow).value = Format$(wshFAC_Finale.Range("C74").value, "0.00")
        .Range("R" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E74").value, "0.00")
        .Range("S" & firstFreeRow).value = Format$(wshFAC_Finale.Range("C75").value, "0.000")
        .Range("T" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E75").value, "0.00")
        
        .Range("U" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E77").value, "0.00")
        
        .Range("V" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E79").value, "0.00")
        .Range("W" & firstFreeRow).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") '2025-01-25 @ 15:01
    End With
    
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B11").value = firstFreeRow
    Application.EnableEvents = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally", "", startTime)

    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_Add_Invoice_Details_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    Dim rowLastService As Long
    rowLastService = wshFAC_Finale.Range("B64").End(xlUp).row
    If rowLastService < 34 Then GoTo nothing_to_update
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Détails$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").value
    Dim r As Long
    For r = 34 To rowLastService
        'Add fields to the recordset before updating it
        rs.AddNew
        With wshFAC_Finale
            rs.Fields(fFacDInvNo - 1).value = CStr(noFacture)
            rs.Fields(fFacDDescription - 1).value = .Range("B" & r).value
            If .Range("C" & r).value <> 0 And _
               .Range("D" & r).value <> 0 And _
               .Range("E" & r).value <> 0 Then
                    rs.Fields(fFacDHeures - 1).value = Format$(.Range("C" & r).value, "0.00")
                    rs.Fields(fFacDTaux - 1).value = Format$(.Range("D" & r).value, "0.00")
                    rs.Fields(fFacDHonoraires - 1).value = Format$(.Range("E" & r).value, "0.00")
            End If
            rs.Fields(fFacDInvRow - 1).value = wshFAC_Brouillon.Range("B11").value
            rs.Fields(fFacDTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            
        End With
    'Update the recordset (create the record)
    rs.Update
    Next r
    
    'Create Summary By Rates lines
    Dim i As Long
    For i = 25 To 34
        If wshFAC_Brouillon.Range("R" & i).value <> "" And _
            wshFAC_Brouillon.Range("S" & i).value <> 0 Then
                rs.AddNew
                With wshFAC_Brouillon
                    rs.Fields(fFacDInvNo - 1).value = noFacture
                    rs.Fields(fFacDDescription - 1).value = "*** - [Sommaire des TEC] pour la facture - " & _
                                                wshFAC_Brouillon.Range("R" & i).value
                    rs.Fields(fFacDHeures - 1).value = CDbl(Format$(.Range("S" & i).value, "0.00"))
                    rs.Fields(fFacDTaux - 1).value = CDbl(Format$(.Range("T" & i).value, "0.00"))
                    rs.Fields(fFacDHonoraires - 1).value = CDbl(Format$(.Range("S" & i).value * .Range("T" & i).value, "0.00"))
                    rs.Fields(fFacDInvRow - 1).value = wshFAC_Brouillon.Range("B11").value
                    rs.Fields(fFacDTimeStamp - 1).value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
                End With
                rs.Update
        End If
    Next i
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
nothing_to_update:

    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_to_DB", "", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Details_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_Locally", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Get the last entered service
    Dim lastEnteredService As Long
    lastEnteredService = wshFAC_Finale.Range("B64").End(xlUp).row
    If lastEnteredService < 34 Then GoTo nothing_to_update
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Détails.Cells(wshFAC_Détails.Rows.count, "A").End(xlUp).row + 1
   
    Dim i As Long
    For i = 34 To lastEnteredService
        With wshFAC_Détails
            .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
            .Range("B" & firstFreeRow).value = wshFAC_Finale.Range("B" & i).value
            .Range("C" & firstFreeRow).value = Format$(wshFAC_Finale.Range("C" & i).value, "0.00")
            .Range("D" & firstFreeRow).value = Format$(wshFAC_Finale.Range("D" & i).value, "0.00")
            .Range("E" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E" & i).value, "0.00")
            .Range("F" & firstFreeRow).value = i
            .Range("G" & firstFreeRow).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
            firstFreeRow = firstFreeRow + 1
        End With
    Next i

nothing_to_update:
    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_Locally", "", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Somm_Taux_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    'Fees summary from wshFAC_Brouillon
    Dim firstRow As Long, lastRow As Long
    firstRow = 44
    lastRow = 48
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Sommaire_Taux$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").value
    Dim seq As Long
    Dim r As Long
    For r = firstRow To lastRow
        'Add fields to the recordset before updating it
        If wshFAC_Brouillon.Range("R" & r).value <> "" Then
            rs.AddNew
            With wshFAC_Finale
                rs.Fields(fFacSTInvNo - 1).value = noFacture
                rs.Fields(fFacSTSéquence - 1).value = seq
                rs.Fields(fFacSTProf - 1).value = wshFAC_Brouillon.Range("R" & r).value
                rs.Fields(fFacSTHeures - 1).value = wshFAC_Brouillon.Range("S" & r).value
                rs.Fields(fFacSTTaux - 1).value = wshFAC_Brouillon.Range("T" & r).value
                rs.Fields(fFacSTTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
                seq = seq + 1
            End With
            'Update the recordset (create the record)
            rs.Update
        End If
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    conn.Close
    On Error GoTo 0
   
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_to_DB", "", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Somm_Taux_Locally()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_Locally", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Fees summary from wshFAC_Brouillon
    Dim firstRow As Long, lastRow As Long
    firstRow = 44
    lastRow = 48
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Sommaire_Taux.Cells(wshFAC_Sommaire_Taux.Rows.count, "A").End(xlUp).row + 1
   
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").value
    Dim seq As Long
    Dim i As Long
    For i = firstRow To lastRow
        If wshFAC_Brouillon.Range("R" & i).value <> "" Then
            With wshFAC_Sommaire_Taux
                .Cells(firstFreeRow, fFacSTInvNo).value = noFacture
                .Cells(firstFreeRow, fFacSTSéquence).value = seq
                .Cells(firstFreeRow, fFacSTProf).value = wshFAC_Brouillon.Range("R" & i).value
                .Cells(firstFreeRow, fFacSTHeures).value = CCur(wshFAC_Brouillon.Range("S" & i).value)
                .Cells(firstFreeRow, fFacSTHeures).NumberFormat = "#,##0.00"
                .Cells(firstFreeRow, fFacSTTaux).value = CCur(wshFAC_Brouillon.Range("T" & i).value)
                .Cells(firstFreeRow, fFacSTTaux).NumberFormat = "#,##0.00"
                firstFreeRow = firstFreeRow + 1
                seq = seq + 1
            End With
        End If
    Next i

    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_Locally", "", startTime)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    'Formule pour le solde des Comptes Clients
    Dim formula As String
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields(fFacCCInvNo - 1).value = .Range("E28").value
        rs.Fields(fFacCCInvoiceDate - 1).value = CDate(wshFAC_Brouillon.Range("O3").value)
        rs.Fields(fFacCCCustomer - 1).value = .Range("B24").value
        rs.Fields(fFacCCCodeClient - 1).value = wshFAC_Brouillon.Range("B18").value
        rs.Fields(fFacCCStatus - 1).value = "Unpaid"
        rs.Fields(fFacCCTerms - 1).value = "Net"
        rs.Fields(fFacCCDueDate - 1).value = CDate(wshFAC_Brouillon.Range("O3").value)
        rs.Fields(fFacCCTotal - 1).value = .Range("E77").value 'Le dépôt s'il y en a un n'est pas comptabilisé ici!
        rs.Fields(fFacCCTotalPaid - 1).value = 0
        rs.Fields(fFacCCTotalRegul - 1).value = 0
        rs.Fields(fFacCCBalance - 1).value = .Range("E77").value
        rs.Fields(fFacCCDaysOverdue - 1).value = 0
        rs.Fields(fFacCCTimeStamp - 1).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    End With
    
    'Update the recordset (create the record)
    rs.Update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB", "", startTime)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_Locally() '2024-03-11 @ 08:49 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_Locally", _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Comptes_Clients.Cells(wshFAC_Comptes_Clients.Rows.count, "A").End(xlUp).row + 1
   
    With wshFAC_Comptes_Clients
        .Cells(firstFreeRow, fFacCCInvNo).value = wshFAC_Finale.Range("E28")
        .Cells(firstFreeRow, fFacCCInvoiceDate).value = CDate(wshFAC_Brouillon.Range("O3").value)
        .Cells(firstFreeRow, fFacCCCustomer).value = wshFAC_Finale.Range("B24").value
        .Cells(firstFreeRow, fFacCCCodeClient).value = wshFAC_Brouillon.Range("B18").value
        .Cells(firstFreeRow, fFacCCStatus).value = "Unpaid"
        .Cells(firstFreeRow, fFacCCTerms).value = "Net"
        .Cells(firstFreeRow, fFacCCDueDate).value = CDate(wshFAC_Brouillon.Range("O3").value)
        .Cells(firstFreeRow, fFacCCTotal).value = wshFAC_Finale.Range("E81").value
        .Cells(firstFreeRow, fFacCCTotalPaid).value = 0
        .Cells(firstFreeRow, fFacCCTotalRegul).value = 0
        .Cells(firstFreeRow, fFacCCBalance).value = wshFAC_Finale.Range("E81").value
        .Cells(firstFreeRow, fFacCCDaysOverdue).value = 0
        .Cells(firstFreeRow, fFacCCTimeStamp).value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    End With

nothing_to_update:

    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_Locally", "", startTime)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_To_DB(firstRow As Long, lastRow As Long) 'Update Billed Status in DB

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB", firstRow & ", " & lastRow, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim r As Long, tecID As Long, SQL As String
    For r = firstRow To lastRow
        If wshTEC_Local.Range("BB" & r).value = "VRAI" Or _
            wshFAC_Brouillon.Range("C" & r + 4) <> True Then
            GoTo next_iteration
        End If
        tecID = wshTEC_Local.Range("AQ" & r).value
        
        'Open the recordset for the specified ID
        SQL = "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID
        rs.Open SQL, conn, 2, 3
        If Not rs.EOF Then
            'Update EstFacturee, DateFacturee & NoFacture
            rs.Fields(fTECEstFacturee - 1).value = "VRAI"
            rs.Fields(fTECDateFacturee - 1).value = Format$(Date, "yyyy-mm-dd")
            rs.Fields(fTECVersionApp - 1).value = ThisWorkbook.Name
            rs.Fields(fTECNoFacture - 1).value = wshFAC_Brouillon.Range("O6").value
            rs.Update
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TECID '" & r & "' ne peut être trouvé!", _
                vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
        'Update the recordset (create the record)
        rs.Update
        rs.Close
next_iteration:
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB", "", startTime)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_Locally(firstResultRow As Long, lastResultRow As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_Locally", firstResultRow & ", " & lastResultRow, 0)
    
    'Set the range to look for
    Dim lastTECRow As Long
    lastTECRow = wshTEC_Local.Cells(wshTEC_Local.Rows.count, "A").End(xlUp).row
    Dim lookupRange As Range: Set lookupRange = wshTEC_Local.Range("A3:A" & lastTECRow)
    
    Dim r As Long, rowToBeUpdated As Long, tecID As Long
    For r = firstResultRow To lastResultRow
        If wshTEC_Local.Range("BB" & r).value = "FAUX" And _
                wshFAC_Brouillon.Range("C" & r + 4) = True Then
            tecID = wshTEC_Local.Range("AQ" & r).value
            rowToBeUpdated = Fn_Find_Row_Number_TECID(tecID, lookupRange)
            wshTEC_Local.Range("L" & rowToBeUpdated).value = "VRAI"
            wshTEC_Local.Range("M" & rowToBeUpdated).value = Format$(Date, "yyyy-mm-dd")
            wshTEC_Local.Range("O" & rowToBeUpdated).value = ThisWorkbook.Name
            wshTEC_Local.Range("P" & rowToBeUpdated).value = wshFAC_Brouillon.Range("O6").value
        End If
    Next r
    
    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_Locally", "", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Détails_To_DB(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_To_DB", CStr(projetID), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Détails$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Build the query
    Dim strSQL As String
    strSQL = "UPDATE [" & destinationTab & "] SET estDetruite = -1 WHERE projetID = " & projetID
    
    'Execute the SQL query
    conn.Execute strSQL
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_To_DB", "", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Détails_Locally(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_Locally", CStr(projetID), 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    
    Dim projetIDColumn As String, isDétruiteColumn As String
    projetIDColumn = "A"
    isDétruiteColumn = "I"

    'Find the last used row
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    'Use Range.Find to locate the first cell with the projetID
    Dim cell As Range
    Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).Find(What:=projetID, LookIn:=xlValues, LookAt:=xlWhole)

    'Check if the projetID was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Do
            'Update the isDétruite column for the found projetID
            ws.Cells(cell.row, isDétruiteColumn).value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_Locally", "", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Entête_To_DB(projetID)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_To_DB", CStr(projetID), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Entête$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Build the query
    Dim strSQL As String
    strSQL = "UPDATE [" & destinationTab & "] SET estDetruite = True WHERE ProjetID = " & projetID

    'Execute the SQL query
    On Error GoTo eh
    conn.Execute strSQL
    On Error GoTo 0
    
    'Close recordset and connection
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    
    Application.ScreenUpdating = True

    'Libérer la mémoire (Normal)
    Set conn = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_To_DB", "", startTime)
    Exit Sub

eh:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error # APP-001"
    If Not conn Is Nothing Then
        On Error Resume Next
        conn.Close
        Set conn = Nothing
        On Error GoTo 0
    End If
    
End Sub

Sub FAC_Finale_Softdelete_Projets_Entête_Locally(projetID)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_Locally", CStr(projetID), 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    
    Dim projetIDColumn As String, isDétruiteColumn As String
    projetIDColumn = "A"
    isDétruiteColumn = "Z"

    'Find the last used row
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    'Use Range.Find to locate the first cell with the projetID
    Dim cell As Range
    Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).Find(What:=projetID, LookIn:=xlValues, LookAt:=xlWhole)

    'Check if the projetID was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Do
            'Update the isDétruite column for the found projetID
            ws.Cells(cell.row, isDétruiteColumn).value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_Locally", "", startTime)

End Sub

'Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16
'
'    Dim startTime as Double: startTime = Timer: Call Log_Record("modFAC_Finale:Invoice_Load", 0)
'
'    With wshFAC_Brouillon
'        If wshFAC_Brouillon.Range("B20").Value = Empty Then
'            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
'            GoTo NoItems
'        End If
'        'Could that invoice been cancelled (more than 1 row) ?
'        Call InvoiceGetAllTrans(wshFAC_Brouillon.Range("O6").Value)
'        Dim NbTrans As Long
'        NbTrans = .Range("B31").Value
'        If NbTrans = 0 Then
'            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
'            GoTo NoItems
'        Else
'            If NbTrans > 1 Then
'                MsgBox "Cette facture a été annulée! Veuillez saisir un numéro de facture VALIDE pour votre recherche"
'                GoTo NoItems
'            End If
'        End If
'        .Range("B24").Value = True 'Set Invoice Load to true
'        .Range("S2,E4:F4,K4:L6,O3,K11:O45,Q11:Q45").ClearContents
'        wshFAC_Finale.Range("B34:F68").ClearContents
'        Dim InvListRow As Long
'        InvListRow = wshFAC_Brouillon.Range("B20").Value 'InvListRow = Row associated with the invoice
'        'Get values from wshFAC_Entête (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
'        .Range("O3").Value = wshFAC_Entête.Range("B" & InvListRow).Value
'        .Range("K3").Value = wshFAC_Entête.Range("D" & InvListRow).Value
'        .Range("K4").Value = wshFAC_Entête.Range("E" & InvListRow).Value
'        .Range("K5").Value = wshFAC_Entête.Range("F" & InvListRow).Value
'        .Range("K6").Value = wshFAC_Entête.Range("G" & InvListRow).Value
'        'Get values from wshFAC_Entête (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
'        Dim dFact As Date
'        dFact = wshFAC_Entête.Range("B" & InvListRow).Value
'        wshFAC_Finale.Range("B21").Value = "Le " & Format$(dFact, "d") & " " & _
'                                            UCase(Format$(dFact, "mmmm")) & " " & _
'                                            Format$(dFact, "yyyy")
'        wshFAC_Finale.Range("B23").Value = wshFAC_Entête.Range("D" & InvListRow).Value
'        wshFAC_Finale.Range("B24").Value = wshFAC_Entête.Range("E" & InvListRow).Value
'        wshFAC_Finale.Range("B25").Value = wshFAC_Entête.Range("F" & InvListRow).Value
'        wshFAC_Finale.Range("B26").Value = wshFAC_Entête.Range("G" & InvListRow).Value
'        'Load Invoice Detail Items
'        With wshFAC_Détails
'            Dim lastRow As Long, lastResultRow As Long
'            lastRow = .Range("A999999").End(xlUp).row
'            If lastRow < 4 Then Exit Sub 'No Item Lines
'            .Range("I3").Value = wshFAC_Brouillon.Range("O6").Value
'            wshFAC_Finale.Range("F28").Value = wshFAC_Brouillon.Range("O6").Value 'Invoice #
'            'Advanced Filter to get items specific to ONE invoice
'            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, criteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
'            lastResultRow = .Range("O999").End(xlUp).row
'            If lastResultRow < 3 Then GoTo NoItems
'            For resultRow = 3 To lastResultRow
'                invitemRow = .Range("O" & resultRow).Value
'                wshFAC_Brouillon.Range("L" & invitemRow & ":O" & invitemRow).Value = .Range("K" & resultRow & ":N" & resultRow).Value 'Description, Hours, Rate & Value
'                wshFAC_Brouillon.Range("Q" & invitemRow).Value = .Range("P" & resultRow).Value  'Set Item DB Row
'                wshFAC_Finale.Range("C" & invitemRow + 23 & ":F" & invitemRow + 23).Value = .Range("K" & resultRow & ":N" & resultRow).Value 'Description, Hours, Rate & Value
'            Next resultRow
'        End With
'        'Proceed with trailer data (Misc. charges & Taxes)
'        .Range("M48").Value = wshFAC_Entête.Range("I" & InvListRow).Value
'        .Range("O48").Value = wshFAC_Entête.Range("J" & InvListRow).Value
'        .Range("M49").Value = wshFAC_Entête.Range("K" & InvListRow).Value
'        .Range("O49").Value = wshFAC_Entête.Range("L" & InvListRow).Value
'        .Range("M50").Value = wshFAC_Entête.Range("M" & InvListRow).Value
'        .Range("O50").Value = wshFAC_Entête.Range("N" & InvListRow).Value
'        .Range("O52").Value = wshFAC_Entête.Range("P" & InvListRow).Value
'        .Range("O53").Value = wshFAC_Entête.Range("R" & InvListRow).Value
'        .Range("O57").Value = wshFAC_Entête.Range("T" & InvListRow).Value
'
'NoItems:
'    .Range("B24").Value = False 'Set Invoice Load To false
'    End With
'
'    Call Log_Record("modFAC_Finale:Invoice_Load","", startTime)
'
'End Sub

'Fonction pour vérifier si un nom de feuille existe déjà dans un classeur
Function NomFeuilleExiste(nom As String) As Boolean
    
    On Error Resume Next
    NomFeuilleExiste = Not Worksheets(nom) Is Nothing
    On Error GoTo 0
    
End Function

'Sub InvoiceGetAllTrans(inv As String)
'
'    Dim startTime as Double: startTime = Timer: Call Log_Record("modFAC_Finale:InvoiceGetAllTrans", 0)
'
'    Application.ScreenUpdating = False
'
'    wshFAC_Brouillon.Range("B31").Value = 0
'
'    With wshFAC_Entête
'        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
'        lastRow = .Range("A999999").End(xlUp).row 'Last wshFAC_Entête Row
'        If lastRow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
'        On Error Resume Next
'        .Names("Criterial").Delete
'        On Error GoTo 0
'        .Range("V3").Value = wshFAC_Brouillon.Range("O6").Value
'        'Advanced Filter setup
'        .Range("A3:T" & lastRow).AdvancedFilter xlFilterCopy, _
'            criteriaRange:=.Range("V2:V3"), _
'            CopyToRange:=.Range("X2:AQ2"), _
'            Unique:=True
'        lastResultRow = .Range("X999").End(xlUp).row 'How many rows trans for that invoice
'        If lastResultRow < 3 Then
'            GoTo Done
'        End If
''        With .Sort
''            .SortFields.Clear
''            .SortFields.add Key:=wshFAC_Entête.Range("X2"), _
''                SortOn:=xlSortOnValues, _
''                Order:=xlAscending, _
''                DataOption:=xlSortNormal 'Sort Based Invoice Number
''            .SortFields.add Key:=wshGL_Trans.Range("Y3"), _
''                SortOn:=xlSortOnValues, _
''                Order:=xlAscending, _
''                DataOption:=xlSortNormal 'Sort Based On TECID
''            .SetRange wshFAC_Entête.Range("X2:AQ" & lastResultRow) 'Set Range
''            .Apply 'Apply Sort
''         End With
'         wshFAC_Brouillon.Range("B31").Value = lastResultRow - 2 'Remove Header rows from row count
'Done:
'    End With
'    Application.ScreenUpdating = True
'
'    Call Log_Record("modFAC_Finale:InvoiceGetAllTrans","", startTime)
'
'End Sub

Sub FAC_Finale_Setup_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Setup_All_Cells", "", 0)
    
    Application.EnableEvents = False
     
    With wshFAC_Finale
'        Dim j As String, m As String, y As String
'        j = Format$(FAC_Brouillon!O3, "j")
'        m = UCase(Format$(FAC_Brouillon!O3, "mmm"))
'        y = Format$(FAC_Brouillon!O3, "yyyy")
        
        .Range("B21").formula = "= ""Le "" & DAY(FAC_Brouillon!o3) & "" "" & UPPER(TEXT(FAC_Brouillon!O3, ""mmmm"")) & "" "" & YEAR(FAC_Brouillon!O3)"
'        .Range("B21").formula = "= ""Le "" & TEXT(FAC_Brouillon!O3, ""j mmmm aaaa"")"
        .Range("B23:B27").value = ""
        .Range("E28").value = "=" & wshFAC_Brouillon.Name & "!O6"    'Invoice number
        
'        .Range("C65").Value = "Heures"                               'Summary Heading
'        .Range("D65").Value = "Taux"                                 'Summary Heading
'        .Range("C66").formula = "=" & wshFAC_Brouillon.Name & "!M47" 'Hours summary
'        .Range("D66").formula = "=" & wshFAC_Brouillon.Name & "!N47" 'Hourly Rate
'
'        With .Range("C65:D66")
'            .Font.ThemeColor = xlThemeColorLight1
'            .Font.TintAndShade = 0
'        End With

        Call FAC_Brouillon_Set_Labels(.Range("B69"), "FAC_Label_SubTotal_1")
        Call FAC_Brouillon_Set_Labels(.Range("B73"), "FAC_Label_SubTotal_2")
        Call FAC_Brouillon_Set_Labels(.Range("B74"), "FAC_Label_TPS")
        Call FAC_Brouillon_Set_Labels(.Range("B75"), "FAC_Label_TVQ")
        Call FAC_Brouillon_Set_Labels(.Range("B77"), "FAC_Label_GrandTotal")
        Call FAC_Brouillon_Set_Labels(.Range("B79"), "FAC_Label_Deposit")
        Call FAC_Brouillon_Set_Labels(.Range("B81"), "FAC_Label_AmountDue")

        'Establish formulas
        .Range("E69").formula = "=" & wshFAC_Brouillon.Name & "!O47" 'Fees Sub-Total
        
        .Range("B70").formula = "=" & wshFAC_Brouillon.Name & "!M48" 'Misc. Amount # 1 - Description
        .Range("E70").formula = "=" & wshFAC_Brouillon.Name & "!O48" 'Misc. Amount # 1
        
        .Range("B71").formula = "=" & wshFAC_Brouillon.Name & "!M49" 'Misc. Amount # 2 - Description
        .Range("E71").formula = "=" & wshFAC_Brouillon.Name & "!O49" 'Misc. Amount # 2
        
        .Range("B72").formula = "=" & wshFAC_Brouillon.Name & "!M50" 'Misc. Amount # 3 - Description
        .Range("E72").formula = "=" & wshFAC_Brouillon.Name & "!O50" 'Misc. Amount # 3
        
        .Range("E73").formula = "=SUM(E69:E72)"                      'Invoice Sub-Total
        
        .Range("C74").formula = "=" & wshFAC_Brouillon.Name & "!N52" 'GST Rate
        .Range("E74").formula = "=round(E73*C74,2)"                  'GST Amount"
        .Range("C75").formula = "=" & wshFAC_Brouillon.Name & "!N53" 'PST Rate
        .Range("E75").formula = "=round(E73*C75,2)"                  'PST Amount
        
        .Range("E77").formula = "=SUM(E73:E75)"                        'Total including taxes
        .Range("E79").formula = "=" & wshFAC_Brouillon.Name & "!O57" 'Deposit Amount
        .Range("E81").formula = "=E77-E79"                             'Total due on that invoice
    End With
    
    Application.EnableEvents = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Setup_All_Cells", "", startTime)

End Sub

Sub shp_FAC_Finale_Preview_PDF_Click()

    Call FAC_Finale_Preview_PDF

End Sub

Sub FAC_Finale_Preview_PDF() '2024-03-02 @ 16:18

    Dim ws As Worksheet
    Set ws = wshFAC_Finale
    
    'Imprimante PDF à utiliser
    Dim imprimantePDF As String
    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        imprimantePDF = "Adobe PDF"
    Else
        imprimantePDF = "Microsoft Print to PDF"
    End If
    
    Dim imprimanteCourante As String
    'Vérifiez si l'imprimante existe
    On Error Resume Next
    If Len(Application.ActivePrinter) > 0 Then
        'Mémoriser l'imprimante actuelle pour la réinitialiser après
        imprimanteCourante = Application.ActivePrinter
    End If
    On Error GoTo 0
    Debug.Print "#083 - Imprimante actuelle : " & imprimanteCourante
    
    'On imprime la facture
    wshFAC_Finale.PrintOut , , 1, True, True, , , , False
   
    'Restaurer l'imprimante précédente après l'impression
    If imprimanteCourante <> "" Then
        On Error Resume Next
        Application.ActivePrinter = imprimanteCourante
        On Error GoTo 0
    End If
    
    Debug.Print "#084 - Imprimante restaurée : " & Application.ActivePrinter

End Sub

Sub FAC_Finale_Creation_PDF_Click()

    Call FAC_Finale_Creation_PDF
    
End Sub

Sub FAC_Finale_Creation_PDF() '2024-10-13 @ 10:15
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Creation_PDF", wshFAC_Finale.Range("E28").value, 0)
    
    flagEtapeFacture = 1
    
    'Étape 1 - Création du document PDF
    Call FAC_Finale_Create_PDF(wshFAC_Finale.Range("E28").value)
    
    'Étape 2 - Copie de la facture en format EXCEL
    Call FAC_Finale_Copie_Vers_Excel(wshFAC_Brouillon.Range("B18").value, _
                                          wshFAC_Finale.Range("L81").value, _
                                          wshFAC_Finale.Range("E28").value, _
                                          Format$(wshFAC_Brouillon.Range("O3").value, "yyyy-mm-dd"))
    flagEtapeFacture = 3
    
    'Étape 3 - Envoi du courriel
    DoEvents
    Call FAC_Finale_Creation_Courriel(wshFAC_Finale.Range("E28").value, wshFAC_Brouillon.Range("B18").value)
    flagEtapeFacture = 4
    
    'Étape 4 - Activation du bouton SAUVEGARDE
    Call FAC_Finale_Enable_Save_Button
    flagEtapeFacture = 5

    Call Log_Record("modFAC_Finale:FAC_Finale_Creation_PDF", "", startTime)

End Sub

Sub FAC_Finale_Create_PDF(noFacture As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF", noFacture, 0)
    
    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF
    Dim result As Boolean
    result = FAC_Finale_Create_PDF_Func(noFacture, "SaveOnly")
    
    If result = False Then
        MsgBox "ATTENTION... Impossible de sauvegarder la facture en format PDF", _
                vbCritical, _
                "Impossible de sauvegarder la facture en format PDF"
        flagEtapeFacture = -1
    End If

    Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF", "", startTime)

End Sub

Function FAC_Finale_Create_PDF_Func(noFacture As String, Optional action As String = "SaveOnly") As Boolean
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF_Func", noFacture & ", " & action, 0)
    
    Dim SaveAs As String

    Application.ScreenUpdating = False

    'Construct the SaveAs filename
    SaveAs = wshAdmin.Range("F5").value & FACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2023-12-19 @ 07:28

    'Check if the file already exists
    Dim fileExists As Boolean
    fileExists = Dir(SaveAs) <> ""
    
    'If the file exists, prompt the user for confirmation
    Dim reponse As VbMsgBoxResult
    If fileExists Then
        reponse = MsgBox("La facture (PDF) numéro '" & noFacture & "' existe déja." & _
                          "Voulez-vous la remplacer ?", vbYesNo + vbQuestion, _
                          "Cette facture existe déjà en formt PDF")
        If reponse = vbNo Then
            GoTo EndMacro
        End If
    End If

    'Set Print Quality
    On Error Resume Next
    ActiveSheet.PageSetup.PrintQuality = 600
    Err.Clear
    On Error GoTo 0

    'Adjust Document Properties - 2023-10-06 @ 09:54
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
    End With
    
    'Create the PDF file and Save It
    On Error GoTo RefLibError
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=SaveAs, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    On Error GoTo 0
    
SaveOnly:
    FAC_Finale_Create_PDF_Func = True 'Return value
'    FAC_Finale_Create_Email = True 'Return value
    GoTo EndMacro
    
RefLibError:
    MsgBox "Incapable de préparer le courriel. La librairie n'est pas disponible"
    FAC_Finale_Create_PDF_Func = False 'Function return value
'    FAC_Finale_Create_Email = False 'Function return value

EndMacro:
    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF_Func", "", startTime)

End Function

Sub FAC_Finale_Copie_Vers_Excel(clientID As String, clientName As String, invNo As String, invDate As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Copie_Vers_Excel", _
        clientID & ", " & clientName & ", " & invNo & ", " & invDate, 0)
    
    Dim clientNamePurged As String
    clientNamePurged = clientName
    
    Application.ScreenUpdating = False
    
    'Purge le nom du client
    Do While InStr(clientNamePurged, "[") > 0 And InStr(clientNamePurged, "]") > 0
        clientNamePurged = Fn_Strip_Contact_From_Client_Name(clientNamePurged)
    Loop
    
    'Définir le chemin complet du répertoire des fichiers Excel
    Dim ExcelFilesFullPath As String
    ExcelFilesFullPath = wshAdmin.Range("F5").value & FACT_EXCEL_PATH
    ChDir ExcelFilesFullPath
    
    'Définir la feuille source et la plage à copier
    Dim wbSource As Workbook: Set wbSource = ThisWorkbook
    Dim wsSource As Worksheet: Set wsSource = wshFAC_Finale
    Dim plageSource As Range: Set plageSource = wsSource.Range("A1:F88")

    'Désactiver les événements pour éviter Workbook_Activate
    Application.EnableEvents = False
    
    'Ouvrir un nouveau Workbook (ou choisir un workbook existant)
    On Error Resume Next
    Dim strCible As String
    strCible = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx") 'Sélectionner un classeur cible
    On Error GoTo 0
    
    'Si l'utilisateur annule la sélection du fichier ou il y a une erreur
    Dim wbCible As Workbook
    If strCible = "Faux" Or strCible = "False" Or strCible = "" Then
        'Créer un nouveau workbook
        Set wbCible = Workbooks.Add
        strCible = ""
    Else
        'Ouvrir le workbook sélectionné
        Set wbCible = Workbooks.Open(strCible)
    End If
    
'    Set wsCible = wbCible.Sheets.add(After:=wbCible.Sheets(wbCible.Sheets.count))
    Dim strName As String
    Dim strNameBase As String
    strNameBase = invDate & " - " & invNo
    strName = strNameBase
    
    'On vérifie si le nom de la nouvelle feuille à ajouter existe déjà
    Dim wsExist As Boolean
    wsExist = False
    On Error Resume Next
    wsExist = Not wbCible.Worksheets(strNameBase) Is Nothing
    On Error GoTo 0
    
    'Si le worksheet existe déjà avec ce nom, demander à l'utilisateur ce qu'il souhaite faire
    Dim wsCible As Worksheet
    Dim suffixe As Integer
    Dim reponse As String
    
    If wsExist Then
        reponse = MsgBox("La feuille '" & strNameBase & "' existe déjà dans ce fichier" & vbCrLf & vbCrLf & _
                         "Voulez-vous :" & vbCrLf & vbCrLf & _
                         "1. Remplacer l'onglet existant par la facture courante ?" & vbCrLf & vbCrLf & _
                         "2. Créer un nouvel onglet avec un suffixe ?" & vbCrLf & vbCrLf & _
                         "Cliquez sur Oui pour remplacer, ou sur Non pour créer un nouvel onglet.", _
                         vbYesNoCancel + vbQuestion, "Le nouvel onglet à créer existe déjà")

        Select Case reponse
            Case vbYes 'Remplacer l'onglet existant
                Application.DisplayAlerts = False ' Désactiver les alertes pour écraser sans confirmation
                wbCible.Worksheets(strNameBase).Delete
                Application.DisplayAlerts = True
                
                'Créer une nouvelle feuille avec le même nom
                Set wsCible = wbCible.Worksheets.Add(After:=wbCible.Sheets(wbCible.Sheets.count))
                wsCible.Name = strNameBase 'Attribuer le nom d'origine

            Case vbNo 'L'utilisateur souhaite créer une nouvelle feuille
                suffixe = 1
                'Boucle pour trouver un nom unique de feuille (worksheet)
                Do
                    strName = strNameBase & "_" & Format$(suffixe, "00")
                    On Error Resume Next
                    Set wsCible = wbCible.Sheets(strName)
                    On Error GoTo 0
                    If wsCible Is Nothing Then Exit Do 'Nous avons un nom unique pour la feuille
                    suffixe = suffixe + 1
                Loop
                
                'Créer une nouvelle feuille avec ce nom directement lors de la création
                Application.DisplayAlerts = False ' Désactiver les alertes pour éviter Feuil1
                Set wsCible = wbCible.Worksheets.Add(After:=wbCible.Sheets(wbCible.Sheets.count))
                wsCible.Name = strName ' Attribuer le nouveau nom avec suffixe
                Application.DisplayAlerts = True ' Réactiver les alertes après la création
        End Select
    Else
        'Si la feuille n'existe pas, on peut directement la créer
        Set wsCible = wbCible.Worksheets.Add(After:=wbCible.Sheets(wbCible.Sheets.count))
        wsCible.Name = strNameBase
    End If
    
'    wsCible.Name = strName 'Renommer la nouvelle feuille
    
    '1. Copier les valeurs uniquement
    plageSource.Copy
    wsCible.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    DoEvents
    Application.CutCopyMode = False
    
    '2. Copier les formats de cellules
    plageSource.Copy
    wsCible.Range("A1").PasteSpecial Paste:=xlPasteFormats

    DoEvents
    Application.CutCopyMode = False
    
    '3. Conserver la taille des colonnes
    Dim i As Integer
    For i = 1 To plageSource.Columns.count
        wsCible.Columns(i).ColumnWidth = plageSource.Columns(i).ColumnWidth
    Next i

    '4. Ajuster les hauteurs de lignes (optionnel si nécessaire)
    For i = 1 To plageSource.Rows.count
        wsCible.Rows(i).RowHeight = plageSource.Rows(i).RowHeight
    Next i

    '5. Copier l'entête de la facture
    Dim forme As Shape, newForme As Shape
    For Each forme In wsSource.Shapes
        If forme.Name = "GCF_Entête" Then
            'Copier & coller l'entête de la facture (logo)
'            DoEvents
            forme.Copy
'            DoEvents
'            wsCible.Activate
'            DoEvents
            wsCible.PasteSpecial Format:="Picture (Enhanced Metafile)"
'            wsCible.Paste
            'Récupérer la nouvelle forme
'            DoEvents
            Set newForme = wsCible.Shapes(wsCible.Shapes.count)
            'Ajuster la position et la taille de la forme
            With newForme
                .Top = forme.Top
                .Left = forme.Left
                .Height = 250
            End With
        End If
    Next forme
    
    DoEvents
    Application.CutCopyMode = False

    '6. Copier les paramètres d'impression
    With wsCible.PageSetup
        .Orientation = wsSource.PageSetup.Orientation
        On Error Resume Next '2024-10-15 @ 06:51
        .PaperSize = xlPaperLetter '2024-10-13 @ 07:45
        On Error GoTo 0
        .Zoom = wsSource.PageSetup.Zoom
        .FitToPagesWide = wsSource.PageSetup.FitToPagesWide
        .FitToPagesTall = wsSource.PageSetup.FitToPagesTall
        .LeftMargin = wsSource.PageSetup.LeftMargin
        .RightMargin = wsSource.PageSetup.RightMargin
        .TopMargin = wsSource.PageSetup.TopMargin
        .BottomMargin = wsSource.PageSetup.BottomMargin
        .HeaderMargin = wsSource.PageSetup.HeaderMargin
        .FooterMargin = wsSource.PageSetup.FooterMargin
        .PrintArea = wsSource.PageSetup.PrintArea
        .PrintTitleRows = wsSource.PageSetup.PrintTitleRows
        .PrintTitleColumns = wsSource.PageSetup.PrintTitleColumns
        .CenterHorizontally = wsSource.PageSetup.CenterHorizontally
        .CenterVertically = wsSource.PageSetup.CenterVertically
    End With
    
    'Désactiver le mode copier-coller pour libérer la mémoire
    Application.CutCopyMode = False
    
    'Optionnel : Sauvegarder le workbook cible sous un nouveau nom si nécessaire
    If strCible = "" Then
        wbCible.SaveAs ExcelFilesFullPath & Application.PathSeparator & clientID & " - " & clientNamePurged & ".xlsx"
        MsgBox "Un nouveau fichier Excel (" & clientID & " - " & clientNamePurged & ".xlsx" & ")" & vbNewLine & vbNewLine & _
                "A été créé pour sauvegarder la facture", vbInformation
    End If
    
    'Réactiver les événements après l'ouverture
    Application.EnableEvents = True
    
    'La facture a été sauvegardé en format EXCEL
    flagEtapeFacture = 3
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set forme = Nothing
    Set plageSource = Nothing
    Set newForme = Nothing
    Set wbCible = Nothing
    Set wbSource = Nothing
    Set wsCible = Nothing
    Set wsSource = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Copie_Vers_Excel", "", startTime)

End Sub

Sub FAC_Finale_Creation_Courriel(noFacture As String, clientID As String) '2024-10-13 @ 11:33

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Creation_Courriel", _
        noFacture & "," & clientID, 0)
    
    Dim fileExists As Boolean
    
    '1a. Chemin de la pièce jointe (Facture en format PDF)
    Dim attachmentFullPathName As String
    attachmentFullPathName = wshAdmin.Range("F5").value & FACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2024-09-03 @ 16:43
    
    '1b. Vérification de l'existence de la pièce jointe
    fileExists = Dir(attachmentFullPathName) <> ""
    If Not fileExists Then
        MsgBox "La pièce jointe (Facture en format PDF) n'existe pas" & _
                    "à l'emplacement spécifié, soit " & attachmentFullPathName, vbCritical
        GoTo Exit_Sub
    End If
    
    '2a. Chemin du template (.oft) de courriel
    Dim templateFullPathName As String
    templateFullPathName = Environ("appdata") & "\Microsoft\Templates\GCF_Facturation.oft"

    '2b. Vérification de l'existence du template
    fileExists = Dir(templateFullPathName) <> ""
    If Not fileExists Then
        MsgBox "Le gabarit 'GCF_Facturation.oft' est introuvable " & _
                    "à l'emplacement spécifié, soit " & Environ("appdata") & "\Microsoft\Templates", _
                    vbCritical
        GoTo Exit_Sub
    End If
    
    '3. Initialisation de l'application Outlook
    Dim OutlookApp As Object
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    '4. Création de l'email à partir du template
    Dim mailItem As Object
    Set mailItem = OutlookApp.CreateItemFromTemplate(templateFullPathName)

    '5. Ajout de la pièce jointe
    mailItem.Attachments.Add attachmentFullPathName

    '6. Obtenir l'adresse courriel pour le client
    Dim ws As Worksheet: Set ws = wshBD_Clients
    Dim eMailFacturation As String
    eMailFacturation = Fn_Get_Value_From_UniqueID(ws, clientID, 2, fClntFMCourrielFacturation)
    If eMailFacturation = "uniqueID introuvable" Then
        mailItem.To = ""
    Else
        If Fn_Valider_Courriel(eMailFacturation) = True Then
            mailItem.To = eMailFacturation
        Else
            MsgBox "Je ne peux utiliser l'adresse courriel de ce client" & vbNewLine & vbNewLine & _
                    "soit '" & eMailFacturation & "' !", vbExclamation
            mailItem.To = ""
        End If
    End If
    
    mailItem.Display
    'MailItem.Send 'Pour envoyer directement l'email

Exit_Sub:

    'Libérer la mémoire
    Set mailItem = Nothing
    Set OutlookApp = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Creation_Courriel", "", startTime)

End Sub

Sub Test_FAC_Finale_Creation_Courriel()

    Call FAC_Finale_Creation_Courriel("24-24524", "1793")

End Sub

Sub shp_FAC_Finale_Cacher_Heures_Click()

    Call FAC_Finale_Cacher_Heures
    
End Sub

Sub FAC_Finale_Cacher_Heures()

    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
    
End Sub

Sub shp_FAC_Finale_Montrer_Heures_Click()

    Call FAC_Finale_Montrer_Heures
    
End Sub

Sub FAC_Finale_Montrer_Heures()

    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
    
End Sub

Sub shp_FAC_Finale_Cacher_Sommaire_Taux_Click()

    Call FAC_Finale_Cacher_Sommaire_Taux
    
End Sub

Sub FAC_Finale_Cacher_Sommaire_Taux()

    'First determine how many rows there is in the Fees Summary
    Dim nbItems As Long
    Dim i As Long
    For i = 66 To 62 Step -1
        If wshFAC_Finale.Range("C" & i).value <> "" Then
            nbItems = nbItems + 1
        End If
    Next i
    
    If nbItems > 0 Then
        Dim rngFeesSummary As Range: Set rngFeesSummary = _
            wshFAC_Finale.Range("C" & (66 - nbItems) + 1 & ":D66")
        rngFeesSummary.ClearContents
    End If
    
    'Libérer la mémoire
    Set rngFeesSummary = Nothing
    
End Sub

Sub shp_FAC_Finale_Montrer_Sommaire_Taux_Click()

    Call FAC_Finale_Montrer_Sommaire_Taux

End Sub

Sub FAC_Finale_Montrer_Sommaire_Taux()

    'Épure le sommaire des honoraires
    Dim hres As Currency
    Dim taux As Currency
    Dim nbTaux As Integer
    Dim dictTaux As Object
    Set dictTaux = CreateObject("Scripting.Dictionary")
    Dim tauxHeures() As Variant
    ReDim tauxHeures(1 To 5, 1 To 2)
    Dim dernierIndex As Integer
    dernierIndex = UBound(tauxHeures)
    
    Dim i As Integer
    For i = 44 To 48
        taux = wshFAC_Brouillon.Range("T" & i).value
        hres = wshFAC_Brouillon.Range("S" & i).value
        If taux <> 0 Then
            If dictTaux.Exists(taux) Then
                dictTaux(taux) = dictTaux(taux) + hres
            Else
                dictTaux.Add taux, hres
                If hres <> 0 Then
                    nbTaux = nbTaux + 1
                End If
            End If
        End If
    Next i
    
    If nbTaux > 0 Then
        Dim rowFAC_Finale As Long
        rowFAC_Finale = 66 - nbTaux
        Dim rngFeesSummary As Range: Set rngFeesSummary = wshFAC_Finale.Range("C" & rowFAC_Finale & ":D66")
        wshFAC_Finale.Range("C" & rowFAC_Finale).value = "Heures"
        wshFAC_Finale.Range("C" & rowFAC_Finale).Font.Bold = True
        wshFAC_Finale.Range("C" & rowFAC_Finale).Font.underline = True
        wshFAC_Finale.Range("C" & rowFAC_Finale).HorizontalAlignment = xlCenter

        wshFAC_Finale.Range("D" & rowFAC_Finale).value = "Taux"
        wshFAC_Finale.Range("D" & rowFAC_Finale).Font.Bold = True
        wshFAC_Finale.Range("D" & rowFAC_Finale).Font.underline = True
        wshFAC_Finale.Range("D" & rowFAC_Finale).HorizontalAlignment = xlCenter

        Dim t As Variant
        i = rowFAC_Finale + 1
        For Each t In dictTaux.keys
            wshFAC_Finale.Range("C" & i & ":D" & i).Font.Color = RGB(0, 0, 0)
            wshFAC_Finale.Range("C" & i).NumberFormat = "##0.00"
            wshFAC_Finale.Range("C" & i).HorizontalAlignment = xlCenter
            wshFAC_Finale.Range("C" & i).Font.Bold = False
            wshFAC_Finale.Range("C" & i).Font.underline = False
            wshFAC_Finale.Range("C" & i).Font.Name = "Verdana"
            wshFAC_Finale.Range("C" & i).Font.size = 11
            wshFAC_Finale.Range("C" & i).value = dictTaux(t)
            wshFAC_Finale.Range("D" & i).Font.Bold = False
            wshFAC_Finale.Range("D" & i).NumberFormat = "#,##0.00 $"
            wshFAC_Finale.Range("D" & i).HorizontalAlignment = xlCenter
            wshFAC_Finale.Range("D" & i).Font.underline = False
            wshFAC_Finale.Range("D" & i).Font.Name = "Verdana"
            wshFAC_Finale.Range("D" & i).Font.size = 11
            wshFAC_Finale.Range("D" & i).value = t
            i = i + 1
        Next t
        
    End If
    
    'Libérer la mémoire
    Set dictTaux = Nothing
    Set rngFeesSummary = Nothing
    Set t = Nothing
    
End Sub

Sub shp_FAC_Finale_Goto_Onglet_Brouillon_CLick()

    Call FAC_Finale_Goto_Onglet_FAC_Brouillon

End Sub

Sub FAC_Finale_Goto_Onglet_FAC_Brouillon()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Goto_Onglet_FAC_Brouillon", "", 0)
   
    Application.ScreenUpdating = False
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Brouillon.Range("E4").Select

    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Goto_Onglet_FAC_Brouillon", "", startTime)

End Sub

Sub FAC_Finale_Enable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = True
    
    flagEtapeFacture = 3

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub FAC_Finale_Disable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = False

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub


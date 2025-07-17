Attribute VB_Name = "modFAC_Finale"
'@Folder("Saisie_Facture")

Option Explicit

Private invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Private lastRow As Long, lastResultRow As Long, resultRow As Long

'@Description ("Clic sur le bouton Sauvegarde")
Sub shp_FAC_Finale_Save_Click() '2025-06-21 @ 08:20

    Call FAC_Finale_Save
    
    Call RestaurerFeuilleFinaleIntact

End Sub

Sub FAC_Finale_Save() '2024-03-28 @ 07:19

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Save", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    With wshFAC_Brouillon
        'Check For Mandatory Fields - Client
        If .Range("B18").Value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
        
        'Check For Mandatory Fields - Date de facture
        If .Range("O3").Value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture AVANT de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
        
        'Check For Mandatory Fields - Date de facture
        If Len(Trim$(.Range("O6").Value)) <> 8 Then
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
    lastResultRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "AQ").End(xlUp).Row
        
    If lastResultRow > 2 Then
        Call FAC_Finale_TEC_Update_As_Billed_To_DB(3, lastResultRow)
        Call FAC_Finale_TEC_Update_As_Billed_Locally(3, lastResultRow)
    End If
    
    'Update FAC_Projets_Entête & FAC_Projets_Détails, if necessary
    Dim projetID As Long
    projetID = wshFAC_Brouillon.Range("B52").Value
    If projetID <> 0 Then
        Call FAC_Finale_Softdelete_Projets_Détails_To_DB(projetID)
        Call FAC_Finale_Softdelete_Projets_Détails_Locally(projetID)
        
        Call FAC_Finale_Softdelete_Projets_Entête_To_DB(projetID)
        Call FAC_Finale_Softdelete_Projets_Entête_Locally(projetID)
    End If
        
    'Save Invoice total amount
    Dim invoice_Total As Currency
    invoice_Total = wshFAC_Brouillon.Range("O51").Value
        
    'GL stuff will occur at the confirmation level (later)
'    Call FAC_Finale_GL_Posting_Preparation
    
    'Update TEC_DashBoard
    Call modTEC_TDB.ActualiserTECTableauDeBord '2024-03-21 @ 12:32

    Call FAC_Brouillon_Clear_All_TEC_Displayed
    
    Application.ScreenUpdating = True
    
    MsgBox "La facture '" & wshFAC_Brouillon.Range("O6").Value & "' est enregistrée." & _
        vbNewLine & vbNewLine & "Le total de la facture est " & _
        Trim$(Format$(invoice_Total, "### ##0.00 $")) & _
        " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    
    wshFAC_Brouillon.Select
    Application.Wait (Now + TimeValue("0:00:02"))
    wshFAC_Brouillon.Range("E3").Value = vbNullString 'Reset client to empty
    
    wshFAC_Brouillon.Range("B27").Value = False
    
    Call modFAC_Brouillon.FAC_Brouillon_New_Invoice '2024-03-12 @ 08:08 - Maybe ??
    
Fast_Exit_Sub:

    wshFAC_Brouillon.Select
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Save", vbNullString, startTime)
    
End Sub

Sub FAC_Finale_Add_Invoice_Header_to_DB()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Header_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Entête$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Can only ADD to the file, no modification is allowed
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields(fFacEInvNo - 1).Value = .Range("E28").Value
        rs.Fields(fFacEDateFacture - 1).Value = Format$(wshFAC_Brouillon.Range("O3").Value, "yyyy-mm-dd")
        rs.Fields(fFacEACouC - 1).Value = "AC" 'Facture to be confirmed MANUALLY - 2024-08-16 @ 05:46
        rs.Fields(fFacECustID - 1).Value = wshFAC_Brouillon.Range("B18").Value
        rs.Fields(fFacEContact - 1).Value = .Range("B23").Value
        rs.Fields(fFacENomClient - 1).Value = .Range("B24").Value
        rs.Fields(fFacEAdresse1 - 1).Value = .Range("B25").Value
        rs.Fields(fFacEAdresse2 - 1).Value = .Range("B26").Value
        rs.Fields(fFacEAdresse3 - 1).Value = .Range("B27").Value
        
        rs.Fields(fFacEHonoraires - 1).Value = .Range("E69").Value
        
        rs.Fields(fFacEAF1Desc - 1).Value = .Range("B70").Value
        rs.Fields(fFacEAutresFrais1 - 1).Value = Format$(wshFAC_Finale.Range("E70").Value, "0.00")
        rs.Fields(fFacEAF2Desc - 1).Value = .Range("B71").Value
        rs.Fields(fFacEAutresFrais2 - 1).Value = Format$(.Range("E71").Value, "0.00")
        rs.Fields(fFacEAF3Desc - 1).Value = .Range("B72").Value
        rs.Fields(fFacEAutresFrais3 - 1).Value = Format$(.Range("E72").Value, "0.00")
        
        rs.Fields(fFacETauxTPS - 1).Value = Format$(.Range("C74").Value, "0.00")
        rs.Fields(fFacEMntTPS - 1).Value = Format$(.Range("E74").Value, "0.00")
        rs.Fields(fFacETauxTVQ - 1).Value = Format$(.Range("C75").Value, "0.00000") '2024-10-15 @ 05:49
        rs.Fields(fFacEMntTVQ - 1).Value = Format$(.Range("E75").Value, "0.00")
        
        rs.Fields(fFacEARTotal - 1).Value = Format$(.Range("E77").Value, "0.00")
        
        rs.Fields(fFacEDépôt - 1).Value = Format$(.Range("E79").Value, "0.00")
        rs.Fields(fFacETimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss") '2025-01-25 @ 15:01
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Header_to_DB", vbNullString, startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Header_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wsdFAC_Entete.Cells(wsdFAC_Entete.Rows.count, "A").End(xlUp).Row + 1
    
    With wsdFAC_Entete
        .Range("A" & firstFreeRow).Value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).Value = Format$(wshFAC_Brouillon.Range("O3").Value, "mm-dd-yyyy")
        .Range("C" & firstFreeRow).Value = "AC"
        .Range("D" & firstFreeRow).Value = wshFAC_Brouillon.Range("B18").Value
        .Range("E" & firstFreeRow).Value = wshFAC_Finale.Range("B23").Value
        .Range("F" & firstFreeRow).Value = wshFAC_Finale.Range("B24").Value
        .Range("G" & firstFreeRow).Value = wshFAC_Finale.Range("B25").Value
        .Range("H" & firstFreeRow).Value = wshFAC_Finale.Range("B26").Value
        .Range("I" & firstFreeRow).Value = wshFAC_Finale.Range("B27").Value
        
        .Range("J" & firstFreeRow).Value = wshFAC_Finale.Range("E69").Value
        
        .Range("K" & firstFreeRow).Value = wshFAC_Finale.Range("B70").Value
        .Range("L" & firstFreeRow).Value = wshFAC_Finale.Range("E70").Value
        .Range("M" & firstFreeRow).Value = wshFAC_Finale.Range("B71").Value
        .Range("N" & firstFreeRow).Value = wshFAC_Finale.Range("E71").Value
        .Range("O" & firstFreeRow).Value = wshFAC_Finale.Range("B72").Value
        .Range("P" & firstFreeRow).Value = wshFAC_Finale.Range("E72").Value
        
        .Range("Q" & firstFreeRow).Value = wshFAC_Finale.Range("C74").Value
        .Range("R" & firstFreeRow).Value = wshFAC_Finale.Range("E74").Value
        .Range("S" & firstFreeRow).Value = wshFAC_Finale.Range("C75").Value
        .Range("T" & firstFreeRow).Value = wshFAC_Finale.Range("E75").Value
        
        .Range("U" & firstFreeRow).Value = wshFAC_Finale.Range("E77").Value
        
        .Range("V" & firstFreeRow).Value = wshFAC_Finale.Range("E79").Value
        .Range("W" & firstFreeRow).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss") '2025-01-25 @ 15:01
    End With
    
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B11").Value = firstFreeRow
    Application.EnableEvents = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally", vbNullString, startTime)

    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_Add_Invoice_Details_to_DB()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Details_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    Dim rowLastService As Long
    rowLastService = wshFAC_Finale.Range("B64").End(xlUp).Row
    If rowLastService < 34 Then GoTo nothing_to_update
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Détails$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").Value
    Dim r As Long
    For r = 34 To rowLastService
        'Add fields to the recordset before updating it
        rs.AddNew
        With wshFAC_Finale
            rs.Fields(fFacDInvNo - 1).Value = CStr(noFacture)
            rs.Fields(fFacDDescription - 1).Value = .Range("B" & r).Value
            If .Range("C" & r).Value <> 0 And _
               .Range("D" & r).Value <> 0 And _
               .Range("E" & r).Value <> 0 Then
                    rs.Fields(fFacDHeures - 1).Value = Format$(.Range("C" & r).Value, "0.00")
                    rs.Fields(fFacDTaux - 1).Value = Format$(.Range("D" & r).Value, "0.00")
                    rs.Fields(fFacDHonoraires - 1).Value = Format$(.Range("E" & r).Value, "0.00")
            End If
            rs.Fields(fFacDInvRow - 1).Value = wshFAC_Brouillon.Range("B11").Value
            rs.Fields(fFacDTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            
        End With
    'Update the recordset (create the record)
    rs.Update
    Next r
    
    'Create Summary By Rates lines
    Dim i As Long
    For i = 25 To 34
        If wshFAC_Brouillon.Range("R" & i).Value <> vbNullString And _
            wshFAC_Brouillon.Range("S" & i).Value <> 0 Then
                rs.AddNew
                With wshFAC_Brouillon
                    rs.Fields(fFacDInvNo - 1).Value = noFacture
                    rs.Fields(fFacDDescription - 1).Value = "*** - [Sommaire des TEC] pour la facture - " & _
                                                wshFAC_Brouillon.Range("R" & i).Value
                    rs.Fields(fFacDHeures - 1).Value = CDbl(Format$(.Range("S" & i).Value, "0.00"))
                    rs.Fields(fFacDTaux - 1).Value = CDbl(Format$(.Range("T" & i).Value, "0.00"))
                    rs.Fields(fFacDHonoraires - 1).Value = CDbl(Format$(.Range("S" & i).Value * .Range("T" & i).Value, "0.00"))
                    rs.Fields(fFacDInvRow - 1).Value = vbNullString
                    rs.Fields(fFacDTimeStamp - 1).Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Details_to_DB", vbNullString, startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Details_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Details_Locally", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Get the last entered service
    Dim lastEnteredService As Long
    lastEnteredService = wshFAC_Finale.Range("B64").End(xlUp).Row
    If lastEnteredService < 34 Then GoTo nothing_to_update
    
    Dim ws As Worksheet
    Set ws = wsdFAC_Details
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wsdFAC_Details.Cells(wsdFAC_Details.Rows.count, "A").End(xlUp).Row + 1
   
    Dim i As Long
    For i = 34 To lastEnteredService
        With ws
            .Range("A" & firstFreeRow).Value = wshFAC_Finale.Range("E28")
            .Range("B" & firstFreeRow).Value = wshFAC_Finale.Range("B" & i).Value
            .Range("C" & firstFreeRow).Value = Format$(wshFAC_Finale.Range("C" & i).Value, "0.00")
            .Range("D" & firstFreeRow).Value = Format$(wshFAC_Finale.Range("D" & i).Value, "0.00")
            .Range("E" & firstFreeRow).Value = Format$(wshFAC_Finale.Range("E" & i).Value, "0.00")
            .Range("F" & firstFreeRow).Value = vbNullString
            .Range("G" & firstFreeRow).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        End With
        firstFreeRow = firstFreeRow + 1
    Next i

    'Create Summary By Rates lines
    For i = 25 To 34
        If wshFAC_Brouillon.Range("R" & i).Value <> vbNullString And _
            wshFAC_Brouillon.Range("S" & i).Value <> 0 Then
                With wshFAC_Brouillon
                    ws.Range("A" & firstFreeRow).Value = wshFAC_Finale.Range("E28")
                    ws.Range("B" & firstFreeRow).Value = "*** - [Sommaire des TEC] pour la facture - " & _
                        wshFAC_Brouillon.Range("R" & i).Value
                    ws.Range("C" & firstFreeRow).Value = Format$(.Range("S" & i).Value, "0.00")
                    ws.Range("D" & firstFreeRow).Value = Format$(.Range("T" & i).Value, "0.00")
                    ws.Range("E" & firstFreeRow).Value = Format$(.Range("S" & i).Value * .Range("T" & i).Value, "0.00")
                    ws.Range("F" & firstFreeRow).Value = vbNullString
                    ws.Range("G" & firstFreeRow).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                End With
            firstFreeRow = firstFreeRow + 1
        End If
    Next i

nothing_to_update:
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Details_Locally", vbNullString, startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Somm_Taux_to_DB()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    'Fees summary from wshFAC_Brouillon
    Dim firstRow As Long, lastRow As Long
    firstRow = 44
    lastRow = 48
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Sommaire_Taux$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").Value
    Dim seq As Long
    Dim r As Long
    For r = firstRow To lastRow
        'Add fields to the recordset before updating it
        If wshFAC_Brouillon.Range("R" & r).Value <> vbNullString Then
            rs.AddNew
            With wshFAC_Finale
                rs.Fields(fFacSTInvNo - 1).Value = noFacture
                rs.Fields(fFacSTSéquence - 1).Value = seq
                rs.Fields(fFacSTProf - 1).Value = wshFAC_Brouillon.Range("R" & r).Value
                rs.Fields(fFacSTHeures - 1).Value = wshFAC_Brouillon.Range("S" & r).Value
                rs.Fields(fFacSTTaux - 1).Value = wshFAC_Brouillon.Range("T" & r).Value
                rs.Fields(fFacSTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_to_DB", vbNullString, startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Somm_Taux_Locally()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_Locally", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Fees summary from wshFAC_Brouillon
    Dim firstRow As Long, lastRow As Long
    firstRow = 44
    lastRow = 48
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wsdFAC_Sommaire_Taux.Cells(wsdFAC_Sommaire_Taux.Rows.count, "A").End(xlUp).Row + 1
   
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").Value
    Dim seq As Long
    Dim i As Long
    For i = firstRow To lastRow
        If wshFAC_Brouillon.Range("R" & i).Value <> vbNullString Then
            With wsdFAC_Sommaire_Taux
                .Cells(firstFreeRow, fFacSTInvNo).Value = noFacture
                .Cells(firstFreeRow, fFacSTSéquence).Value = seq
                .Cells(firstFreeRow, fFacSTProf).Value = wshFAC_Brouillon.Range("R" & i).Value
                .Cells(firstFreeRow, fFacSTHeures).Value = CCur(wshFAC_Brouillon.Range("S" & i).Value)
                .Cells(firstFreeRow, fFacSTHeures).NumberFormat = "#,##0.00"
                .Cells(firstFreeRow, fFacSTTaux).Value = CCur(wshFAC_Brouillon.Range("T" & i).Value)
                .Cells(firstFreeRow, fFacSTTaux).NumberFormat = "#,##0.00"
                .Cells(firstFreeRow, fFacSTTimeStamp) = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                firstFreeRow = firstFreeRow + 1
                seq = seq + 1
            End With
        End If
    Next i

    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_Locally", vbNullString, startTime)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_to_DB()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    'Formule pour le solde des Comptes Clients
    Dim formula As String
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields(fFacCCInvNo - 1).Value = .Range("E28").Value
        rs.Fields(fFacCCInvoiceDate - 1).Value = CDate(wshFAC_Brouillon.Range("O3").Value)
        rs.Fields(fFacCCCustomer - 1).Value = .Range("B24").Value
        rs.Fields(fFacCCCodeClient - 1).Value = wshFAC_Brouillon.Range("B18").Value
        rs.Fields(fFacCCStatus - 1).Value = "Unpaid"
        rs.Fields(fFacCCTerms - 1).Value = "Net"
        rs.Fields(fFacCCDueDate - 1).Value = CDate(wshFAC_Brouillon.Range("O3").Value)
        rs.Fields(fFacCCTotal - 1).Value = .Range("E77").Value 'Le dépôt s'il y en a un n'est pas comptabilisé ici!
        rs.Fields(fFacCCTotalPaid - 1).Value = 0
        rs.Fields(fFacCCTotalRegul - 1).Value = 0
        rs.Fields(fFacCCBalance - 1).Value = .Range("E77").Value
        rs.Fields(fFacCCDaysOverdue - 1).Value = 0
        rs.Fields(fFacCCTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB", vbNullString, startTime)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_Locally() '2024-03-11 @ 08:49 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_Locally", _
         "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wsdFAC_Comptes_Clients.Cells(wsdFAC_Comptes_Clients.Rows.count, "A").End(xlUp).Row + 1
   
    With wsdFAC_Comptes_Clients
        .Cells(firstFreeRow, fFacCCInvNo).Value = wshFAC_Finale.Range("E28")
        .Cells(firstFreeRow, fFacCCInvoiceDate).Value = CDate(wshFAC_Brouillon.Range("O3").Value)
        .Cells(firstFreeRow, fFacCCCustomer).Value = wshFAC_Finale.Range("B24").Value
        .Cells(firstFreeRow, fFacCCCodeClient).Value = wshFAC_Brouillon.Range("B18").Value
        .Cells(firstFreeRow, fFacCCStatus).Value = "Unpaid"
        .Cells(firstFreeRow, fFacCCTerms).Value = "Net"
        .Cells(firstFreeRow, fFacCCDueDate).Value = CDate(wshFAC_Brouillon.Range("O3").Value)
        .Cells(firstFreeRow, fFacCCTotal).Value = wshFAC_Finale.Range("E81").Value
        .Cells(firstFreeRow, fFacCCTotalPaid).Value = 0
        .Cells(firstFreeRow, fFacCCTotalRegul).Value = 0
        .Cells(firstFreeRow, fFacCCBalance).Value = wshFAC_Finale.Range("E81").Value
        .Cells(firstFreeRow, fFacCCDaysOverdue).Value = 0
        .Cells(firstFreeRow, fFacCCTimeStamp).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    End With

nothing_to_update:

    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_Locally", vbNullString, startTime)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_To_DB(firstRow As Long, lastRow As Long) 'Update Billed Status in DB

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB", firstRow & ", " & lastRow, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim r As Long, tecID As Long, sql As String
    For r = firstRow To lastRow
        If wsdTEC_Local.Range("BB" & r).Value = "VRAI" Or _
            wshFAC_Brouillon.Range("C" & r + 4) <> True Then
            GoTo next_iteration
        End If
        tecID = wsdTEC_Local.Range("AQ" & r).Value
        
        'Open the recordset for the specified ID
        sql = "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID
        rs.Open sql, conn, 2, 3
        If Not rs.EOF Then
            'Update EstFacturee, DateFacturee & NoFacture
            rs.Fields(fTECEstFacturee - 1).Value = "VRAI"
            rs.Fields(fTECDateFacturee - 1).Value = Format$(Date, "yyyy-mm-dd")
            rs.Fields(fTECVersionApp - 1).Value = ThisWorkbook.Name
            rs.Fields(fTECNoFacture - 1).Value = wshFAC_Brouillon.Range("O6").Value
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB", vbNullString, startTime)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_Locally(firstResultRow As Long, lastResultRow As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_Locally", firstResultRow & ", " & lastResultRow, 0)
    
    'Set the range to look for
    Dim lastTECRow As Long
    lastTECRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "A").End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = wsdTEC_Local.Range("A3:A" & lastTECRow)
    
    Dim r As Long, rowToBeUpdated As Long, tecID As Long
    For r = firstResultRow To lastResultRow
        If wsdTEC_Local.Range("BB" & r).Value = "FAUX" And _
                wshFAC_Brouillon.Range("C" & r + 4) = True Then
            tecID = wsdTEC_Local.Range("AQ" & r).Value
            rowToBeUpdated = Fn_Find_Row_Number_TECID(tecID, lookupRange)
            wsdTEC_Local.Range("L" & rowToBeUpdated).Value = "VRAI"
            wsdTEC_Local.Range("M" & rowToBeUpdated).Value = Format$(Date, "yyyy-mm-dd")
            wsdTEC_Local.Range("O" & rowToBeUpdated).Value = ThisWorkbook.Name
            wsdTEC_Local.Range("P" & rowToBeUpdated).Value = wshFAC_Brouillon.Range("O6").Value
        End If
    Next r
    
    'Libérer la mémoire
    Set lookupRange = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_Locally", vbNullString, startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Détails_To_DB(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_To_DB", CStr(projetID), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_To_DB", vbNullString, startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Détails_Locally(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_Locally", CStr(projetID), 0)
    
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Details
    
    Dim projetIDColumn As String, isDétruiteColumn As String
    projetIDColumn = "A"
    isDétruiteColumn = "I"

    'Find the last used row
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    'Use Range.Find to locate the first cell with the projetID
    Dim cell As Range
    Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).Find(What:=projetID, LookIn:=xlValues, LookAt:=xlWhole)

    'Check if the projetID was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Do
            'Update the isDétruite column for the found projetID
            ws.Cells(cell.row, isDétruiteColumn).Value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_Locally", vbNullString, startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Entête_To_DB(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_To_DB", CStr(projetID), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_To_DB", vbNullString, startTime)
    Exit Sub

eh:
    MsgBox "An error occurred: " & Err.description, vbCritical, "Error # APP-001"
    If Not conn Is Nothing Then
        On Error Resume Next
        conn.Close
        Set conn = Nothing
        On Error GoTo 0
    End If
    
End Sub

Sub FAC_Finale_Softdelete_Projets_Entête_Locally(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_Locally", CStr(projetID), 0)
    
    Dim ws As Worksheet: Set ws = wsdFAC_Projets_Entete
    
    Dim projetIDColumn As String, isDétruiteColumn As String
    projetIDColumn = "A"
    isDétruiteColumn = "Z"

    'Find the last used row
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    'Use Range.Find to locate the first cell with the projetID
    Dim cell As Range
    Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).Find(What:=projetID, LookIn:=xlValues, LookAt:=xlWhole)

    'Check if the projetID was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Do
            'Update the isDétruite column for the found projetID
            ws.Cells(cell.row, isDétruiteColumn).Value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Libérer la mémoire
    Set cell = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_Locally", vbNullString, startTime)

End Sub

'Fonction pour vérifier si un nom de feuille existe déjà dans un classeur
Function NomFeuilleExiste(nom As String) As Boolean
    
    On Error Resume Next
    NomFeuilleExiste = Not ActiveWorkbook.Worksheets(nom) Is Nothing
    On Error GoTo 0
    
End Function

Sub FAC_Finale_Setup_All_Cells()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Setup_All_Cells", vbNullString, 0)
    
    Application.EnableEvents = False
     
    With wshFAC_Finale
        .Range("B21").formula = "= ""Le "" & DAY(FAC_Brouillon!o3) & "" "" & UPPER(TEXT(FAC_Brouillon!O3, ""mmmm"")) & "" "" & YEAR(FAC_Brouillon!O3)"
        .Range("B23:B27").Value = vbNullString
        .Range("E28").Value = "=" & wshFAC_Brouillon.Name & "!O6"    'Invoice number

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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Setup_All_Cells", vbNullString, startTime)

End Sub

Sub shp_FAC_Finale_Preview_PDF_Click()

    Call FAC_Finale_Preview_PDF

End Sub

Sub FAC_Finale_Preview_PDF() '2024-03-02 @ 16:18

    Dim ws As Worksheet
    Set ws = wshFAC_Finale
    
    'Imprimante PDF à utiliser
    Dim imprimantePDF As String
    If GetNomUtilisateur() = "RobertMV" Or GetNomUtilisateur() = "robertmv" Then
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
    If imprimanteCourante <> vbNullString Then
        On Error Resume Next
        Application.ActivePrinter = imprimanteCourante
        On Error GoTo 0
    End If
    
    Debug.Print "#084 - Imprimante restaurée : " & Application.ActivePrinter

End Sub

Sub FAC_Finale_Creation_PDF_Click()

    Call FAC_Finale_Creation_PDF
    
End Sub

Sub FAC_Finale_Creation_PDF() '2025-05-06 @ 11:07

    Dim startTime As Double: startTime = Timer
    Dim codeFacture As String: codeFacture = wshFAC_Finale.Range("E28").Value
    Dim nomClient As String: nomClient = wshFAC_Brouillon.Range("B18").Value
    Dim nomFichier As String: nomFichier = wshFAC_Finale.Range("L81").Value
    Dim dateFacture As String: dateFacture = Format$(wshFAC_Brouillon.Range("O3").Value, "yyyy-mm-dd")
    
    'État initial
    gFlagEtapeFacture = 1
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Creation_PDF", codeFacture, 0)
    
    'Sécuriser l’environnement
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .CutCopyMode = False
    End With
    
    On Error GoTo GestionErreur

    'Étape 1 - Création du document PDF
    Call FAC_Finale_Create_PDF(codeFacture)
    DoEvents: Application.Wait Now + TimeValue("0:00:01")
    
    'Étape 2 - Copie vers fichier Excel client
    Call FAC_Finale_Copie_Vers_Excel(nomClient, nomFichier, codeFacture, dateFacture)
    DoEvents: Application.Wait Now + TimeValue("0:00:01")
    gFlagEtapeFacture = 3

    'Étape 3 - Création du courriel avec pièce jointe PDF
    Call FAC_Finale_Creation_Courriel(codeFacture, nomClient)
    DoEvents: Application.Wait Now + TimeValue("0:00:01")
    gFlagEtapeFacture = 4

    'Étape 4 - Activation du bouton Sauvegarde
    Call FAC_Finale_Enable_Save_Button
    gFlagEtapeFacture = 5

    GoTo fin

GestionErreur:
    MsgBox "Une erreur est survenue à l'étape " & gFlagEtapeFacture & "." & vbCrLf & _
           "Erreur: " & Err.Number & " - " & Err.description, vbCritical
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Creation_PDF", codeFacture & " ÉTAPE " & gFlagEtapeFacture & " > " & Err.description, startTime)

fin:
    'Restaurer l’environnement
    With Application
        .CutCopyMode = False
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Creation_PDF", vbNullString, startTime)
    
End Sub

Sub FAC_Finale_Create_PDF(noFacture As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Create_PDF", noFacture, 0)
    
    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF
    Dim result As Boolean
    result = FAC_Finale_Create_PDF_Func(noFacture, "SaveOnly")
    
    If result = False Then
        MsgBox "ATTENTION... Impossible de sauvegarder la facture en format PDF", _
                vbCritical, _
                "Impossible de sauvegarder la facture en format PDF"
        gFlagEtapeFacture = -1
    End If

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Create_PDF", vbNullString, startTime)

End Sub

Function FAC_Finale_Create_PDF_Func(noFacture As String, Optional action As String = "SaveOnly") As Boolean
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Create_PDF_Func", noFacture & ", " & action, 0)
    
    Dim SaveAs As String

    Application.ScreenUpdating = False

    'Construct the SaveAs filename
    SaveAs = wsdADMIN.Range("F5").Value & gFACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2023-12-19 @ 07:28

    'Check if the file already exists
    Dim fileExists As Boolean
    fileExists = Dir(SaveAs) <> vbNullString
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Create_PDF_Func", vbNullString, startTime)

End Function

Sub FAC_Finale_Copie_Vers_Excel(clientID As String, clientName As String, invNo As String, invDate As String)
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Copie_Vers_Excel", _
        clientID & " - " & clientName & " - " & invNo & " - " & invDate, 0)
    
    Dim clientNamePurged As String
    clientNamePurged = clientName
    
    Application.ScreenUpdating = False
    
    'Purge le nom du client
    Do While InStr(clientNamePurged, "[") > 0 And InStr(clientNamePurged, "]") > 0
        clientNamePurged = Fn_Strip_Contact_From_Client_Name(clientNamePurged)
    Loop
    
    'Définir le chemin complet du répertoire des fichiers Excel
    Dim ExcelFilesFullPath As String
    ExcelFilesFullPath = wsdADMIN.Range("F5").Value & gFACT_EXCEL_PATH
    ChDir ExcelFilesFullPath
    
    'Définir la feuille source et la plage à copier
    Dim wbSource As Workbook: Set wbSource = ThisWorkbook
    Dim wsSource As Worksheet: Set wsSource = wshFAC_Finale
    Dim plageSource As Range: Set plageSource = wsSource.Range("A1:F88")

    'Désactiver les événements pour éviter Workbook_Activate
    Application.EnableEvents = False
    
    'Ouvrir un nouveau Workbook (ou choisir un workbook existant)
    On Error Resume Next
    Dim strCible As Variant
    strCible = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx") 'Sélectionner un classeur cible
    On Error GoTo 0
    
    'Si l'utilisateur annule la sélection du fichier ou il y a une erreur
    Dim wbCible As Workbook
    If strCible = "Faux" Or strCible = "False" Or strCible = vbNullString Then
        'Créer un nouveau workbook
        Set wbCible = Workbooks.Add
        strCible = vbNullString
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

    '5. Copier l'entête de la facture (logo)
    Call CopierFormeEnteteEnTouteSecurite(wsSource, wsCible) '2025-05-06 @ 10:59

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
    If strCible = vbNullString Then
        wbCible.SaveAs ExcelFilesFullPath & Application.PathSeparator & clientID & " - " & clientNamePurged & ".xlsx"
        MsgBox "Un nouveau fichier Excel (" & clientID & " - " & clientNamePurged & ".xlsx" & ")" & vbNewLine & vbNewLine & _
                "A été créé pour sauvegarder la facture", _
                vbInformation
    End If
    
    'Réactiver les événements après l'ouverture
    Application.EnableEvents = True
    
    'La facture a été sauvegardé en format EXCEL
    gFlagEtapeFacture = 3
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set plageSource = Nothing
    Set wbCible = Nothing
    Set wbSource = Nothing
    Set wsCible = Nothing
    Set wsSource = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Copie_Vers_Excel", vbNullString, startTime)

End Sub

Sub CopierFormeEnteteEnTouteSecurite(wsSource As Worksheet, wsCible As Worksheet) '2025-05-06 @ 11:12

    Dim forme As Shape, newForme As Shape
    On Error Resume Next
    Set forme = wsSource.Shapes("GCF_Entête")
    On Error GoTo 0

    If Not forme Is Nothing Then
        'Mémoriser la taille et la position exacte de la forme source
        Dim topPos As Double, leftPos As Double, heightVal As Double, widthVal As Double
        topPos = forme.Top
        leftPos = forme.Left
        heightVal = forme.Height
        widthVal = forme.Width
        
        forme.Copy
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
        
        'Coller en tant qu'image (Enhanced Metafile pour plus de compatibilité)
        wsCible.PasteSpecial Format:="Picture (Enhanced Metafile)"
        DoEvents

        'Récupérer la dernière forme collée
        Set newForme = wsCible.Shapes(wsCible.Shapes.count)
        
        'Réappliquer taille et position exactes
        With newForme
            .Top = topPos
            .Left = leftPos
            .Height = heightVal
            .Width = widthVal
         End With

        Application.CutCopyMode = False
    Else
        Debug.Print "Forme 'GCF_Entête' introuvable sur la feuille source."
    End If
    
End Sub

Sub FAC_Finale_Creation_Courriel(noFacture As String, clientID As String) '2024-10-13 @ 11:33

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Creation_Courriel", _
        noFacture & "," & clientID, 0)
    
    Dim fileExists As Boolean
    
    '1a. Chemin de la pièce jointe (Facture en format PDF)
    Dim attachmentFullPathName As String
    attachmentFullPathName = wsdADMIN.Range("F5").Value & gFACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2024-09-03 @ 16:43
    
    '1b. Vérification de l'existence de la pièce jointe
    fileExists = Dir(attachmentFullPathName) <> vbNullString
    If Not fileExists Then
        MsgBox "La pièce jointe (Facture en format PDF) n'existe pas" & _
                    "à l'emplacement spécifié, soit " & attachmentFullPathName, vbCritical
        GoTo Exit_Sub
    End If
    
    '2a. Chemin du template (.oft) de courriel
    Dim templateFullPathName As String
    templateFullPathName = Environ$("appdata") & "\Microsoft\Templates\GCF_Facturation.oft"

    '2b. Vérification de l'existence du template
    fileExists = Dir(templateFullPathName) <> vbNullString
    If Not fileExists Then
        MsgBox "Le gabarit 'GCF_Facturation.oft' est introuvable " & _
                    "à l'emplacement spécifié, soit " & Environ$("appdata") & "\Microsoft\Templates", _
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
    Dim ws As Worksheet: Set ws = wsdBD_Clients
    Dim eMailFacturation As String
    eMailFacturation = Fn_Get_Value_From_UniqueID(ws, clientID, 2, fClntFMCourrielFacturation)
    If eMailFacturation = "uniqueID introuvable" Then
        mailItem.To = vbNullString
    Else
        Dim adresseEmail  As Variant
        adresseEmail = Split(eMailFacturation, "; ") '2025-03-02 @ 16:59
        Dim nbAdresseCourriel As Integer
        nbAdresseCourriel = UBound(adresseEmail)
        
        Select Case nbAdresseCourriel
            Case 0
                mailItem.To = adresseEmail(0)
            Case Is > 0
                mailItem.To = adresseEmail(0)
                mailItem.cc = adresseEmail(1)
            Case Else
        End Select
    End If
    
    mailItem.Display
    'MailItem.Send 'Pour envoyer directement l'email

Exit_Sub:

    'Libérer la mémoire
    Set mailItem = Nothing
    Set OutlookApp = Nothing
    Set ws = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Creation_Courriel", vbNullString, startTime)

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
        If wshFAC_Finale.Range("C" & i).Value <> vbNullString Then
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
        taux = wshFAC_Brouillon.Range("T" & i).Value
        hres = wshFAC_Brouillon.Range("S" & i).Value
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
        wshFAC_Finale.Range("C" & rowFAC_Finale).Value = "Heures"
        wshFAC_Finale.Range("C" & rowFAC_Finale).Font.Bold = True
        wshFAC_Finale.Range("C" & rowFAC_Finale).Font.underline = True
        wshFAC_Finale.Range("C" & rowFAC_Finale).HorizontalAlignment = xlCenter

        wshFAC_Finale.Range("D" & rowFAC_Finale).Value = "Taux"
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
            wshFAC_Finale.Range("C" & i).Value = dictTaux(t)
            wshFAC_Finale.Range("D" & i).Font.Bold = False
            wshFAC_Finale.Range("D" & i).NumberFormat = "#,##0.00 $"
            wshFAC_Finale.Range("D" & i).HorizontalAlignment = xlCenter
            wshFAC_Finale.Range("D" & i).Font.underline = False
            wshFAC_Finale.Range("D" & i).Font.Name = "Verdana"
            wshFAC_Finale.Range("D" & i).Font.size = 11
            wshFAC_Finale.Range("D" & i).Value = t
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

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Goto_Onglet_FAC_Brouillon", vbNullString, 0)
   
    Application.ScreenUpdating = False
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Brouillon.Range("E4").Select

    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:FAC_Finale_Goto_Onglet_FAC_Brouillon", vbNullString, startTime)

End Sub

Sub FAC_Finale_Enable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = True
    
    gFlagEtapeFacture = 3

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub FAC_Finale_Disable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = False

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub RestaurerFeuilleFinaleIntact() '2025-06-06 @ 18:38
    
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim shp As Shape

    Set wsSource = ThisWorkbook.Sheets("FAC_Finale_Intact")
    Set wsDest = ThisWorkbook.Sheets("FAC_Finale")

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    '1. Effacer toutes les cellules, formules, formats
    wsDest.Cells.Clear

    '2. Effacer toutes les formes
    For Each shp In wsDest.Shapes
        shp.Delete
    Next shp

    '3. Copier tout le contenu cellules + formats + formules
    wsSource.Cells.Copy
    wsDest.Cells.PasteSpecial xlPasteAll 'Tout copier (valeurs, formules, formats, etc.)

    '4. Copier chaque forme individuellement
    For Each shp In wsSource.Shapes
        shp.Copy
        wsDest.Paste
        'Optionnel : replacer la forme exactement
        With wsDest.Shapes(wsDest.Shapes.count)
            .Top = shp.Top
            .Left = shp.Left
            .Width = shp.Width
            .Height = shp.Height
        End With
    Next shp

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub



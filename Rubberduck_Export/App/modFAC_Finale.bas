Attribute VB_Name = "modFAC_Finale"
'@Folder("Saisie_Facture")

Option Explicit

Private invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Private lastRow As Long, lastResultRow As Long, resultRow As Long
Public gCheminPDF As String

'@Description ("Clic sur le bouton Sauvegarde")
Sub shpMettreAJourFAC_Click() '2025-06-21 @ 08:20

    Call SauvegarderFacture
    
    Application.StatusBar = "Réinitialisation de la feuille en cours..."
    Call Reinitialiser_FAC_Finale
    Application.StatusBar = False
    
    Call PreparerFAC_Brouillon '2025-10-15 @ 23:03

End Sub

Sub SauvegarderFacture() '2024-03-28 @ 07:19

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderFacture", _
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
    
    Call CacherBoutonSauvegarder

    Call AjouterFACEnteteBDMaster
    Call AjouterFACEnteteBDLocale
    
    Call AjouterFACDetailsBDMaster
    Call AjouterFACDetailsBDLocale
    
    Call AjouterSommTauxBDMaster
    Call AjouterSommTauxBDLocale
    
    Call AjouterTransComptesClientsBDMaster
    Call AjouterTransComptesClientsBDLocale
    
    Dim lastResultRow As Long
    lastResultRow = wsdTEC_Local.Cells(wsdTEC_Local.Rows.count, "AQ").End(xlUp).Row
        
    If lastResultRow > 2 Then
        Call MettreAJourTECEstFactureeBDMaster(3, lastResultRow)
        Call MettreAJourTECEstFactureeBDLocale(3, lastResultRow)
    End If
    
    'Update FAC_Projets_Entete & FAC_Projets_Details, if necessary
    Dim projetID As Long
    projetID = wshFAC_Brouillon.Range("B52").Value
    If projetID <> 0 Then
        Call DetruireLogiquementProjetsDetailsBDMaster(projetID)
        Call DetruireLogiquementProjetsDetailsBDLocale(projetID)
        
        Call DetruireLogiquementProjetsEnteteBDMaster(projetID)
        Call DetruireLogiquementProjetsEnteteBDLocale(projetID)
    End If
        
    'Save Invoice total amount
    Dim invoice_Total As Currency
    invoice_Total = wshFAC_Brouillon.Range("O51").Value
        
    MsgBox "La facture '" & wshFAC_Brouillon.Range("O6").Value & "' est enregistrée." & _
        vbNewLine & vbNewLine & "Le total de la facture est " & _
        Trim$(Format$(invoice_Total, "### ##0.00 $")) & _
        " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    
Fast_Exit_Sub:

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderFacture", vbNullString, startTime)
    
End Sub

Sub PreparerFAC_Brouillon() '2025-10-15 @ 23:01

    wshFAC_Brouillon.Range("FactureStatut").Value = "" '2025-07-19 @ 19:02
    
    'Update TEC_DashBoard
    Call modTEC_TDB.ActualiserTECTableauDeBord '2024-03-21 @ 12:32

    wshFAC_Brouillon.Select
    Call modFAC_Brouillon.EffacerTECAffiches
    
    Application.ScreenUpdating = True
    
    wshFAC_Brouillon.Select
    Application.Wait (Now + TimeValue("0:00:02"))
    wshFAC_Brouillon.Range("E3").Value = vbNullString 'Reset client to empty
    
    wshFAC_Brouillon.Range("B27").Value = False
    
    Call modFAC_Brouillon.CreerNouvelleFactureBrouillon '2024-03-12 @ 08:08 - Maybe ??

End Sub

Sub AjouterFACEnteteBDMaster()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACEnteteBDMaster", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Entete$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'Can only ADD to the file, no modification is allowed
    
    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    recSet.AddNew
    With wshFAC_Finale
        recSet.Fields(fFacEInvNo - 1).Value = .Range("E28").Value
        recSet.Fields(fFacEDateFacture - 1).Value = Format$(wshFAC_Brouillon.Range("O3").Value, "yyyy-mm-dd")
        recSet.Fields(fFacEACouC - 1).Value = "AC" 'Facture to be confirmed MANUALLY - 2024-08-16 @ 05:46
        recSet.Fields(fFacECustID - 1).Value = wshFAC_Brouillon.Range("B18").Value
        recSet.Fields(fFacEContact - 1).Value = .Range("B23").Value
        recSet.Fields(fFacENomClient - 1).Value = .Range("B24").Value
        recSet.Fields(fFacEAdresse1 - 1).Value = .Range("B25").Value
        recSet.Fields(fFacEAdresse2 - 1).Value = .Range("B26").Value
        recSet.Fields(fFacEAdresse3 - 1).Value = .Range("B27").Value
        
        recSet.Fields(fFacEHonoraires - 1).Value = .Range("E69").Value
        
        recSet.Fields(fFacEAF1Desc - 1).Value = .Range("B70").Value
        recSet.Fields(fFacEAutresFrais1 - 1).Value = Format$(wshFAC_Finale.Range("E70").Value, "0.00")
        recSet.Fields(fFacEAF2Desc - 1).Value = .Range("B71").Value
        recSet.Fields(fFacEAutresFrais2 - 1).Value = Format$(.Range("E71").Value, "0.00")
        recSet.Fields(fFacEAF3Desc - 1).Value = .Range("B72").Value
        recSet.Fields(fFacEAutresFrais3 - 1).Value = Format$(.Range("E72").Value, "0.00")
        
        recSet.Fields(fFacETauxTPS - 1).Value = Format$(.Range("C74").Value, "0.00")
        recSet.Fields(fFacEMntTPS - 1).Value = Format$(.Range("E74").Value, "0.00")
        recSet.Fields(fFacETauxTVQ - 1).Value = Format$(.Range("C75").Value, "0.00000") '2024-10-15 @ 05:49
        recSet.Fields(fFacEMntTVQ - 1).Value = Format$(.Range("E75").Value, "0.00")
        
        recSet.Fields(fFacEARTotal - 1).Value = Format$(.Range("E77").Value, "0.00")
        
        recSet.Fields(fFacEDépôt - 1).Value = Format$(.Range("E79").Value, "0.00")
        recSet.Fields(fFacETimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss") '2025-01-25 @ 15:01
    End With
    'Update the recordset (create the record)
    recSet.Update
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set recSet = Nothing
    Set conn = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACEnteteBDMaster", vbNullString, startTime)

End Sub

Sub AjouterFACEnteteBDLocale() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACEnteteBDLocale", _
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACEnteteBDLocale", vbNullString, startTime)

    Application.ScreenUpdating = True

End Sub

Sub AjouterFACDetailsBDMaster()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACDetailsBDMaster", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    Dim rowLastService As Long
    rowLastService = wshFAC_Finale.Range("B64").End(xlUp).Row
    If rowLastService < 34 Then GoTo nothing_to_update
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Details$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").Value
    Dim r As Long
    For r = 34 To rowLastService
        'Add fields to the recordset before updating it
        recSet.AddNew
        With wshFAC_Finale
            recSet.Fields(fFacDInvNo - 1).Value = CStr(noFacture)
            recSet.Fields(fFacDDescription - 1).Value = .Range("B" & r).Value
            If .Range("C" & r).Value <> 0 And _
               .Range("D" & r).Value <> 0 And _
               .Range("E" & r).Value <> 0 Then
                    recSet.Fields(fFacDHeures - 1).Value = Format$(.Range("C" & r).Value, "0.00")
                    recSet.Fields(fFacDTaux - 1).Value = Format$(.Range("D" & r).Value, "0.00")
                    recSet.Fields(fFacDHonoraires - 1).Value = Format$(.Range("E" & r).Value, "0.00")
            End If
            recSet.Fields(fFacDInvRow - 1).Value = wshFAC_Brouillon.Range("B11").Value
            recSet.Fields(fFacDTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
            
        End With
    'Update the recordset (create the record)
    recSet.Update
    Next r
    
    'Create Summary By Rates lines
    Dim i As Long
    For i = 25 To 34
        If wshFAC_Brouillon.Range("R" & i).Value <> vbNullString And _
            wshFAC_Brouillon.Range("S" & i).Value <> 0 Then
                recSet.AddNew
                With wshFAC_Brouillon
                    recSet.Fields(fFacDInvNo - 1).Value = noFacture
                    recSet.Fields(fFacDDescription - 1).Value = "*** - [Sommaire des TEC] pour la facture - " & _
                                                wshFAC_Brouillon.Range("R" & i).Value
                    recSet.Fields(fFacDHeures - 1).Value = CDbl(Format$(.Range("S" & i).Value, "0.00"))
                    recSet.Fields(fFacDTaux - 1).Value = CDbl(Format$(.Range("T" & i).Value, "0.00"))
                    recSet.Fields(fFacDHonoraires - 1).Value = CDbl(Format$(.Range("S" & i).Value * .Range("T" & i).Value, "0.00"))
                    recSet.Fields(fFacDInvRow - 1).Value = vbNullString
                    recSet.Fields(fFacDTimeStamp - 1).Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
                End With
                recSet.Update
        End If
    Next i
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
nothing_to_update:

    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACDetailsBDMaster", vbNullString, startTime)

End Sub

Sub AjouterFACDetailsBDLocale() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACDetailsBDLocale", _
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterFACDetailsBDLocale", vbNullString, startTime)

End Sub

Sub AjouterSommTauxBDMaster()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterSommTauxBDMaster", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    'Fees summary from wshFAC_Brouillon
    Dim firstRow As Long, lastRow As Long
    firstRow = 44
    lastRow = 48
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Sommaire_Taux$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").Value
    Dim seq As Long
    Dim r As Long
    For r = firstRow To lastRow
        'Add fields to the recordset before updating it
        If wshFAC_Brouillon.Range("R" & r).Value <> vbNullString Then
            recSet.AddNew
            With wshFAC_Finale
                recSet.Fields(fFacSTInvNo - 1).Value = noFacture
                recSet.Fields(fFacSTSéquence - 1).Value = seq
                recSet.Fields(fFacSTProf - 1).Value = wshFAC_Brouillon.Range("R" & r).Value
                recSet.Fields(fFacSTHeures - 1).Value = wshFAC_Brouillon.Range("S" & r).Value
                recSet.Fields(fFacSTTaux - 1).Value = wshFAC_Brouillon.Range("T" & r).Value
                recSet.Fields(fFacSTTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
                seq = seq + 1
            End With
            'Update the recordset (create the record)
            recSet.Update
        End If
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    conn.Close
    On Error GoTo 0
   
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterSommTauxBDMaster", vbNullString, startTime)

End Sub

Sub AjouterSommTauxBDLocale()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterSommTauxBDLocale", _
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterSommTauxBDLocale", vbNullString, startTime)

End Sub

Sub AjouterTransComptesClientsBDMaster()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterTransComptesClientsBDMaster", _
        "# = " & wshFAC_Finale.Range("E28").Value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    'Formule pour le solde des Comptes Clients
    Dim formula As String
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Comptes_Clients$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'timeStamnp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Create an empty recordset
    recSet.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    recSet.AddNew
    With wshFAC_Finale
        recSet.Fields(fFacCCInvNo - 1).Value = .Range("E28").Value
        recSet.Fields(fFacCCInvoiceDate - 1).Value = CDate(wshFAC_Brouillon.Range("O3").Value)
        recSet.Fields(fFacCCCustomer - 1).Value = .Range("B24").Value
        recSet.Fields(fFacCCCodeClient - 1).Value = wshFAC_Brouillon.Range("B18").Value
        recSet.Fields(fFacCCStatus - 1).Value = "Unpaid"
        recSet.Fields(fFacCCTerms - 1).Value = "Net"
        recSet.Fields(fFacCCDueDate - 1).Value = CDate(wshFAC_Brouillon.Range("O3").Value)
        recSet.Fields(fFacCCTotal - 1).Value = .Range("E77").Value 'Le dépôt s'il y en a un n'est pas comptabilisé ici!
        recSet.Fields(fFacCCTotalPaid - 1).Value = 0
        recSet.Fields(fFacCCTotalRegul - 1).Value = 0
        recSet.Fields(fFacCCBalance - 1).Value = .Range("E77").Value
        recSet.Fields(fFacCCDaysOverdue - 1).Value = 0
        recSet.Fields(fFacCCTimeStamp - 1).Value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
    End With
    
    'Update the recordset (create the record)
    recSet.Update
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterTransComptesClientsBDMaster", vbNullString, startTime)

End Sub

Sub AjouterTransComptesClientsBDLocale() '2024-03-11 @ 08:49 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterTransComptesClientsBDLocale", _
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:AjouterTransComptesClientsBDLocale", vbNullString, startTime)

End Sub

Sub MettreAJourTECEstFactureeBDMaster(firstRow As Long, lastRow As Long) 'Update Billed Status in DB

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:MettreAJourTECEstFactureeBDMaster", firstRow & ", " & lastRow, 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "TEC_Local$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    Dim r As Long, tecID As Long, sql As String
    For r = firstRow To lastRow
        If wsdTEC_Local.Range("BB" & r).Value = "VRAI" Or _
            wshFAC_Brouillon.Range("C" & r + 4) <> True Then
            GoTo next_iteration
        End If
        tecID = wsdTEC_Local.Range("AQ" & r).Value
        
        'Open the recordset for the specified ID
        sql = "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID
        recSet.Open sql, conn, 2, 3
        If Not recSet.EOF Then
            'Update EstFacturee, DateFacturee & NoFacture
            recSet.Fields(fTECEstFacturee - 1).Value = "VRAI"
            recSet.Fields(fTECDateFacturee - 1).Value = Format$(Date, "yyyy-mm-dd")
            recSet.Fields(fTECVersionApp - 1).Value = ThisWorkbook.Name
            recSet.Fields(fTECNoFacture - 1).Value = wshFAC_Brouillon.Range("O6").Value
            recSet.Update
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TECID '" & r & "' ne peut être trouvé!", _
                vbExclamation
            recSet.Close
            conn.Close
            Exit Sub
        End If
        'Update the recordset (create the record)
        recSet.Update
        recSet.Close
next_iteration:
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:MettreAJourTECEstFactureeBDMaster", vbNullString, startTime)

End Sub

Sub MettreAJourTECEstFactureeBDLocale(firstResultRow As Long, lastResultRow As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:MettreAJourTECEstFactureeBDLocale", firstResultRow & ", " & lastResultRow, 0)
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:MettreAJourTECEstFactureeBDLocale", vbNullString, startTime)

End Sub

Sub DetruireLogiquementProjetsDetailsBDMaster(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsDetailsBDMaster", CStr(projetID), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Projets_Details$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'Build the query
    Dim strSQL As String
    strSQL = "UPDATE [" & destinationTab & "] SET estDetruite = -1 WHERE projetID = " & projetID
    
    'Execute the SQL query
    conn.Execute strSQL
    
    'Close recordset and connection
    On Error Resume Next
    recSet.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsDetailsBDMaster", vbNullString, startTime)

End Sub

Sub DetruireLogiquementProjetsDetailsBDLocale(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsDetailsBDLocale", CStr(projetID), 0)
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsDetailsBDLocale", vbNullString, startTime)

End Sub

Sub DetruireLogiquementProjetsEnteteBDMaster(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsEnteteBDMaster", CStr(projetID), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "FAC_Projets_Entete$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"

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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsEnteteBDMaster", vbNullString, startTime)
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

Sub DetruireLogiquementProjetsEnteteBDLocale(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsEnteteBDLocale", CStr(projetID), 0)
    
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DetruireLogiquementProjetsEnteteBDLocale", vbNullString, startTime)

End Sub

'Fonction pour vérifier si un nom de feuille existe déjà dans un classeur
Function Fn_FeuilleExiste(nom As String) As Boolean
    
    On Error Resume Next
    Fn_FeuilleExiste = Not ActiveWorkbook.Worksheets(nom) Is Nothing
    On Error GoTo 0
    
End Function

Sub MettreEnPlaceToutesLesCellules()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:MettreEnPlaceToutesLesCellules", vbNullString, 0)
    
    Application.EnableEvents = False
     
    With wshFAC_Finale
        .Range("B21").formula = "= ""Le "" & DAY(FAC_Brouillon!o3) & "" "" & UPPER(TEXT(FAC_Brouillon!O3, ""mmmm"")) & "" "" & YEAR(FAC_Brouillon!O3)"
        .Range("B23:B27").Value = vbNullString
        .Range("E28").Value = "=" & wshFAC_Brouillon.Name & "!O6"    'Invoice number

        Call modFAC_Brouillon.AjusterLesLibellesFACBrouillon(.Range("B69"), "FAC_Label_SubTotal_1")
        Call modFAC_Brouillon.AjusterLesLibellesFACBrouillon(.Range("B73"), "FAC_Label_SubTotal_2")
        Call modFAC_Brouillon.AjusterLesLibellesFACBrouillon(.Range("B74"), "FAC_Label_TPS")
        Call modFAC_Brouillon.AjusterLesLibellesFACBrouillon(.Range("B75"), "FAC_Label_TVQ")
        Call modFAC_Brouillon.AjusterLesLibellesFACBrouillon(.Range("B77"), "FAC_Label_GrandTotal")
        Call modFAC_Brouillon.AjusterLesLibellesFACBrouillon(.Range("B79"), "FAC_Label_Deposit")
        Call modFAC_Brouillon.AjusterLesLibellesFACBrouillon(.Range("B81"), "FAC_Label_AmountDue")

        'Mettre en place les formules de la feuille
        .Range("E69").formula = "=ROUND(" & wshFAC_Brouillon.Name & "!O47, 2)" 'Fees Sub-Total
        
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:MettreEnPlaceToutesLesCellules", vbNullString, startTime)

End Sub

Sub shpPrevisualiserFacture_Click()

    Call PrevisualiserFacturePDF

End Sub

Sub PrevisualiserFacturePDF() '2024-03-02 @ 16:18

    Dim ws As Worksheet
    Set ws = wshFAC_Finale
    
    'Sauvegarder l'imprimante actuelle
    Dim imprimanteActuelle As String
    'Vérifiez si l'imprimante existe
    On Error Resume Next
    If Len(Application.ActivePrinter) > 0 Then
        'Mémoriser l'imprimante actuelle pour la réinitialiser après
        imprimanteActuelle = Application.ActivePrinter
    End If
    On Error GoTo 0
    Debug.Print "#083 - Imprimante actuelle : " & imprimanteActuelle
    
    'Imprimante PDF à utiliser
    Dim imprimantePDF As String
    imprimantePDF = Fn_ObtenirPortFonctionnelAdobePDF '2025-10-17 @ 13:44
    Application.ActivePrinter = imprimantePDF
    
    'On définit la zone d'impression '2025-10-15 @ 10:27
    wshFAC_Finale.PageSetup.PrintArea = "$A1:$F88"
    
    'On imprime la facture
    wshFAC_Finale.PrintOut , , 1, True, True, , , , False
   
    'Restaurer l'imprimante précédente après l'impression
    If imprimanteActuelle <> vbNullString Then
        On Error Resume Next
        Application.ActivePrinter = imprimanteActuelle
        On Error GoTo 0
    End If
    
    Debug.Print "#084 - Imprimante restaurée : " & Application.ActivePrinter

End Sub

Sub shpSauvegarderPDFSauvegarderExcelEnvoyerCourriel_Click()

    Call SauvegarderPDFSauvegarderExcelEnvoyerCourriel
    
End Sub

Sub SauvegarderPDFSauvegarderExcelEnvoyerCourriel() '2025-05-06 @ 11:07

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderPDFSauvegarderExcelEnvoyerCourriel", wshFAC_Finale.Range("E28").Value, 0)
    
    Dim startTime As Double: startTime = Timer
    Dim numeroFacture As String: numeroFacture = wshFAC_Finale.Range("E28").Value
    Dim nomClient As String: nomClient = wshFAC_Brouillon.Range("B18").Value
    Dim nomFichier As String: nomFichier = wshFAC_Finale.Range("L81").Value
    Dim dateFacture As String: dateFacture = Format$(wshFAC_Brouillon.Range("O3").Value, "yyyy-mm-dd")
    
    'État initial
    gFlagEtapeFacture = 1
    
    'Sécuriser l’environnement
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .CutCopyMode = False
    End With
    
'    On Error GoTo GestionErreur

    Call TracerEtape("Début traitement facture")
    
    'Étape 1 - Création du document PDF
    Call TracerEtape("     Début création PDF")
    wshFAC_Finale.PageSetup.PrintArea = "$A1:$F88" '2025-10-15 @ 23:51
    Call SauvegarderFactureFormatPDF(numeroFacture)
    If Dir(gCheminPDF) = "" Then
        Call TracerEtape("     Échec création PDF")
        Exit Sub
    End If
    Call TracerEtape("     PDF créé avec succès")
    gFlagEtapeFacture = 2
    
    'Étape 2 - Copie vers fichier Excel client
    Call TracerEtape("     Début sauvegarde Excel client")
    Call SauvegarderCopieFactureDansExcel(nomClient, nomFichier, numeroFacture, dateFacture)
    Call TracerEtape("     Sauvegarde Excel client réussi")
    gFlagEtapeFacture = 3
'    Call PauseActive(1)

    'Étape 3 - Création du courriel avec pièce jointe PDF
    If Dir(gCheminPDF) = "" Then
        MsgBox "Le fichier PDF est introuvable pour l’envoi.", vbCritical
        Call TracerEtape("     Échec envoi courriel : PDF manquant")
        Exit Sub
    End If
    Call TracerEtape("     Prêt pour envoi courriel")
    Call EnvoyerFactureParCourriel(numeroFacture, nomClient)
    Call TracerEtape("     La facture a été envoyée par courriel")
    gFlagEtapeFacture = 4
'    Call PauseActive(1)

    'Étape 4 - Activation du bouton Sauvegarde
    Call TracerEtape("     Activation bouton Sauvegarde")
    DoEvents
    Call AfficherBoutonSauvegarder
    gFlagEtapeFacture = 5
    
    wshFAC_Brouillon.Range("FactureStatut").Value = "En attente de mise à jour" '2025-07-19 @ 18:35
    Call TracerEtape("Fin traitement facture")
    GoTo fin

GestionErreur:
    MsgBox "Une erreur est survenue à l'étape " & gFlagEtapeFacture & "." & vbCrLf & vbCrLf & _
           "Erreur: " & Err.Number & " - " & Err.description, _
           vbCritical, _
           "Gestion d'erreur dans 'SauvegarderPDFSauvegarderExcelEnvoyerCourriel'"
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderPDFSauvegarderExcelEnvoyerCourriel", numeroFacture & " ÉTAPE " & gFlagEtapeFacture & " > " & Err.description, startTime)

fin:
    'Restaurer l’environnement
    With Application
        .CutCopyMode = False
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderPDFSauvegarderExcelEnvoyerCourriel", vbNullString, startTime)
    
End Sub

Sub SauvegarderFactureFormatPDF(noFacture As String)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderFactureFormatPDF", noFacture, 0)
    
    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF
    Dim cheminPDF As String
    cheminPDF = Fn_ExporterFactureEnPDF(noFacture)
    
    If cheminPDF = vbNullString Then
        MsgBox "ATTENTION... Impossible de sauvegarder la facture en format PDF", _
                vbCritical, _
                "Impossible de sauvegarder la facture en format PDF"
        gFlagEtapeFacture = -1
    End If

    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderFactureFormatPDF", vbNullString, startTime)

End Sub

Function Fn_ExporterFactureEnPDF(noFacture As String) As String
    
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:Fn_ExporterFactureEnPDF", noFacture, 0)

    Application.ScreenUpdating = False

    'Construct the gCheminPDF filename
    gCheminPDF = wsdADMIN.Range("PATH_DATA_FILES").Value & gFACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2023-12-19 @ 07:28

    'Check if the file already exists
    Dim fileExists As Boolean
    fileExists = Dir(gCheminPDF) <> vbNullString
    
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

    'Sauvegarder l'imprimante actuelle
    Dim imprimanteActuelle As String
    'Vérifiez si l'imprimante existe
    On Error Resume Next
    If Len(Application.ActivePrinter) > 0 Then
        'Mémoriser l'imprimante actuelle pour la réinitialiser après
        imprimanteActuelle = Application.ActivePrinter
    End If
    On Error GoTo 0
    Debug.Print "#0883 - Imprimante actuelle : " & imprimanteActuelle

    'Imprimante PDF à utiliser
    Dim imprimantePDF As String
    imprimantePDF = Fn_ObtenirPortFonctionnelAdobePDF '2025-10-20 @ 06:29
    Application.ActivePrinter = imprimantePDF

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
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=gCheminPDF, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    On Error GoTo 0
    
    'Restaurer l'imprimante précédente après l'impression
    If imprimanteActuelle <> vbNullString Then
        On Error Resume Next
        Application.ActivePrinter = imprimanteActuelle
        On Error GoTo 0
    End If

SaveOnly:
    Fn_ExporterFactureEnPDF = True 'Return value
    GoTo EndMacro
    
RefLibError:
    MsgBox "Incapable de préparer le courriel. La librairie n'est pas disponible"
    Fn_ExporterFactureEnPDF = False 'Function return value

EndMacro:
    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:Fn_ExporterFactureEnPDF", vbNullString, startTime)

End Function

Sub SauvegarderCopieFactureDansExcel(clientID As String, clientName As String, invNo As String, invDate As String)
    
   'Call SauvegarderCopieFactureDansExcel(nomClient, nomFichier, numeroFacture, dateFacture)
 
    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderCopieFactureDansExcel", _
        clientID & " - " & clientName & " - " & invNo & " - " & invDate, 0)
    
    Application.ScreenUpdating = False
    
    'Purge le nom du client
    Dim clientNamePurged As String
    clientNamePurged = clientName
    Do While InStr(clientNamePurged, "[") > 0 And InStr(clientNamePurged, "]") > 0
        clientNamePurged = Fn_Strip_Contact_From_Client_Name(clientNamePurged)
    Loop
    If Right(clientNamePurged, 1) = "." Then
        clientNamePurged = Left(clientNamePurged, Len(clientNamePurged) - 1)
    End If
    
    'Définir le chemin complet du répertoire des fichiers Excel
    Dim ExcelFilesFullPath As String
    ExcelFilesFullPath = wsdADMIN.Range("PATH_DATA_FILES").Value & gFACT_EXCEL_PATH
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
                         "Oui = Remplacer l'onglet existant par la facture courante ?" & vbCrLf & vbCrLf & _
                         "Non = Créer un nouvel onglet avec un suffixe ?" & vbCrLf & vbCrLf & _
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

    '5. Copier le logo de l'entreprise
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
        MsgBox "Un nouveau fichier Excel a été créé pour sauvegarder la facture" & vbNewLine & vbNewLine & _
                "'" & clientID & " - " & clientNamePurged & ".xlsx" & "'", _
                vbInformation, _
                "Première facture pour ce client"
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:SauvegarderCopieFactureDansExcel", vbNullString, startTime)

End Sub

Sub CopierFormeEnteteEnTouteSecurite(wsSource As Worksheet, wsCible As Worksheet) '2025-05-06 @ 11:12

    Application.ScreenUpdating = False
    
    Dim forme As Shape, newForme As Shape
    On Error Resume Next
    Set forme = wsSource.Shapes("shpGCFLogo")
    On Error GoTo 0

    If Not forme Is Nothing Then
        Dim limiteMaxTop As Double
        limiteMaxTop = wsCible.Rows(21).Top 'Ligne de la date
        'Mémoriser la taille et la position exacte de la forme source
        Dim topPos As Double, leftPos As Double, heightVal As Double, widthVal As Double
        topPos = forme.Top
        leftPos = forme.Left
        heightVal = forme.Height
        If topPos + heightVal > limiteMaxTop Then
            heightVal = limiteMaxTop - topPos
        End If
        widthVal = forme.Width
        
        forme.Copy
        DoEvents
        Call PauseActive(1)
        
        'Coller en tant qu'image (Enhanced Metafile pour plus de compatibilité)
        wsCible.PasteSpecial Format:="Picture (Enhanced Metafile)"
        DoEvents
        Call PauseActive(1)

        'Récupérer la dernière forme collée
        Set newForme = wsCible.Shapes(wsCible.Shapes.count)
        
        'Réappliquer taille et position exactes
        If Not newForme Is Nothing Then
            With newForme
                .LockAspectRatio = msoFalse ' Permet de modifier Height sans contrainte
                .Top = topPos
                Debug.Print "Top = " & topPos
                .Left = leftPos
                .Height = heightVal
                .Width = widthVal
            End With
        Else
            MsgBox "Erreur : le logo de GCF n’a pas été reconnu.", vbCritical
        End If

        Debug.Print "Hauteur de la nouvelle forme " & newForme.Height
        
        Application.CutCopyMode = False
    Else
        Debug.Print "Forme 'shpGCFLogo' introuvable sur la feuille source."
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub EnvoyerFactureParCourriel(noFacture As String, clientID As String) '2024-10-13 @ 11:33

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:EnvoyerFactureParCourriel", _
        noFacture & "," & clientID, 0)
    
    Dim fileExists As Boolean
    
    '1a. Chemin de la pièce jointe (Facture en format PDF)
    Dim attachmentFullPathName As String
    attachmentFullPathName = wsdADMIN.Range("PATH_DATA_FILES").Value & gFACT_PDF_PATH & Application.PathSeparator & _
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
    eMailFacturation = Fn_ValeurAPartirUniqueID(ws, clientID, 2, fClntFMCourrielFacturation)
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
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:EnvoyerFactureParCourriel", vbNullString, startTime)

End Sub

Sub shpCacherHeuresParLigne_Click()

    Call CacherHeuresParLigne
    
End Sub

Sub CacherHeuresParLigne()

    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
    
End Sub

Sub shpMontrerHeuresParLigne_Click()

    Call MontrerHeuresParLigne
    
End Sub

Sub MontrerHeuresParLigne()

    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
    
End Sub

Sub shpCacherSommaireTaux_Click()

    Call CacherSommaireTaux
    
End Sub

Sub CacherSommaireTaux()

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

Sub shpMontrerSommaireTaux_Click()

    Call MontrerSommaireTaux

End Sub

Sub MontrerSommaireTaux()

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

Sub shpDeplacerVersFeuilleBrouillon_Click()

    If wshFAC_Brouillon.Range("FactureStatut").Value = "En attente de mise à jour" Then '2025-07-19 @ 18:44
        MsgBox "Veuillez d'abord SAUVEGARDER la présente facture" & vbNewLine & vbNewLine & _
                "avant d'en créer une nouvelle.", vbExclamation
        Exit Sub
    End If
    
    Call DeplacerVersFeuilleBrouillon

End Sub

Sub DeplacerVersFeuilleBrouillon()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DeplacerVersFeuilleBrouillon", vbNullString, 0)
   
    Application.ScreenUpdating = False
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Brouillon.Range("E4").Select

    Application.ScreenUpdating = True
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:DeplacerVersFeuilleBrouillon", vbNullString, startTime)

End Sub

Sub AfficherBoutonSauvegarder()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpMettreAJour")
    shp.Visible = True
    
    gFlagEtapeFacture = 3

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub CacherBoutonSauvegarder()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpMettreAJour")
    shp.Visible = False

    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub Reinitialiser_FAC_Finale() '2025-09-05 @ 07:42

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:Reinitialiser_FAC_Finale", vbNullString, 0)
    
    Dim FeuilleSource As Worksheet
    Dim FeuilleCible As Worksheet
    Dim nm As Name, shp As Shape, lo As ListObject
    Dim c As Range
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    '--- Feuilles source et cible ---
    Set FeuilleSource = Worksheets("FAC_Finale_Intact")
    Set FeuilleCible = Worksheets("FAC_Finale")
    
    '--- 1. Nettoyer la feuille cible ---
    FeuilleCible.Cells.Clear
    On Error Resume Next
    FeuilleCible.Cells.Validation.Delete
    FeuilleCible.Cells.FormatConditions.Delete
    On Error GoTo 0
    
    'Supprimer toutes les formes (images, boutons, graphiques, etc.)
    For Each shp In FeuilleCible.Shapes
        shp.Delete
    Next shp
    
    'Supprimer les noms locaux liés à la feuille @TODO
    For Each nm In FeuilleCible.Parent.Names
        If nm.Name Like FeuilleCible.Name & "!*" Then
            Debug.Print "Noms locaux - " & nm.Name
            nm.Delete
        End If
    Next nm
    
    'Supprimer les tableaux structurés s'il y en a
    On Error Resume Next
    For Each lo In FeuilleCible.ListObjects
        lo.Unlist
        Debug.Print "Tableaux structurés - " & lo.Name
    Next lo
    On Error GoTo 0
    
    '--- 2. Copier contenu et formats ---
    FeuilleSource.Cells.Copy
    FeuilleCible.Cells.PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    '--- 2a. Copier les largeurs de colonnes sans utiliser le presse-papier --- 2025-09-29 @ 08:16
    Dim col As Long
    For col = 1 To FeuilleSource.usedRange.Columns.count
        FeuilleCible.Columns(col).ColumnWidth = FeuilleSource.Columns(col).ColumnWidth
        Debug.Print "Largeurs de colonnes - " & col & " = " & FeuilleSource.Columns(col).ColumnWidth
    Next col
    
    '--- 3. Copier les formes avec leur position et taille --- '2025-10-02 @ 15:47
    Dim shpNew As Shape
    
    Application.EnableEvents = False
    
    For Each shp In FeuilleSource.Shapes
        shp.Copy ' Copie la forme
        FeuilleCible.Paste ' Colle la forme sur la feuille cible
        Set shpNew = FeuilleCible.Shapes(FeuilleCible.Shapes.count) 'Récupère la dernière forme collée
        shpNew.Top = shp.Top
        shpNew.Left = shp.Left
        shpNew.Width = shp.Width
        shpNew.Height = shp.Height
        'Déplacer la forme dupliquée sur la feuille cible
        Debug.Print "Copie des formes - " & shp.Name
    Next shp
    
    Application.EnableEvents = True

    '--- 4. Copier les noms locaux ---
    For Each nm In FeuilleSource.Parent.Names
        If nm.Name Like FeuilleSource.Name & "!*" Then
            On Error Resume Next
            FeuilleCible.Parent.Names.Add _
                Name:=Replace(nm.Name, FeuilleSource.Name, FeuilleCible.Name), _
                RefersTo:=Replace(nm.RefersTo, FeuilleSource.Name, FeuilleCible.Name)
            On Error GoTo 0
        End If
    Next nm
    
    '--- 5. Copier la mise en page ---
    On Error Resume Next
    FeuilleCible.PageSetup = FeuilleSource.PageSetup
    On Error GoTo 0
    
    '--- 6. Copier zoom et FreezePanes ---
    With ActiveWindow
        .Zoom = 100
        FeuilleSource.Activate
        .Zoom = .Zoom
        .FreezePanes = False
        If FeuilleSource.Parent.Windows(1).FreezePanes Then
            FeuilleSource.Parent.Windows(1).SplitColumn = _
                FeuilleSource.Parent.Windows(1).SplitColumn
            FeuilleSource.Parent.Windows(1).SplitRow = _
                FeuilleSource.Parent.Windows(1).SplitRow
            FeuilleCible.Parent.Windows(1).FreezePanes = True
        End If
    End With
    
    '--- 7. Corriger les formules qui pointent vers FAC_Finale_Intact ---
    For Each c In FeuilleCible.usedRange.Cells
        If c.HasFormula Then
            c.formula = Replace(c.formula, "FAC_Finale_Intact", "FAC_Finale")
        End If
    Next c
    
    '--- 8. Corriger les hyperliens qui pointent vers FAC_Finale_Intact ---
    Dim hl As Hyperlink
    For Each hl In FeuilleCible.Hyperlinks
        If InStr(hl.Address, "FAC_Finale_Intact") > 0 Then
            hl.Address = Replace(hl.Address, "FAC_Finale_Intact", "FAC_Finale")
        End If
        If InStr(hl.SubAddress, "FAC_Finale_Intact") > 0 Then
            hl.SubAddress = Replace(hl.SubAddress, "FAC_Finale_Intact", "FAC_Finale")
        End If
    Next hl
    
    '--- Fin ---
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    DoEvents
    Application.Wait (Now + TimeValue("0:00:02"))
    DoEvents
    
'    MsgBox "La feuille 'FAC_Finale' a été réinitialisée à partir de 'FAC_Finale_Intact'.", vbInformation
    
    Call modDev_Utils.EnregistrerLogApplication("modFAC_Finale:Reinitialiser_FAC_Finale", vbNullString, startTime)

End Sub

Sub RestaurerFeuilleFinaleIntact() '2025-08-22 @ 16:07

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Set wsSource = ThisWorkbook.Sheets("FAC_Finale_Intact")
    Set wsDest = ThisWorkbook.Sheets("FAC_Finale")

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    '1. Nettoyer la feuille destination
    Call NettoyerFeuille(wsDest)

    '2. Copier le contenu des cellules
    Call SupprimerNomsLocaux(wsDest)
    wsSource.Cells.Copy
    wsDest.Cells.PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    '3. Copier les formes avec positionnement
    Call CopierFormesAvecActions(wsSource, wsDest)

    '4. Recréer les plages nommées dynamiques
    Call ReassignerPlagesNomées(wsSource, wsDest)

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "La feuille 'FAC_Finale' a été restaurée avec succès.", vbInformation

End Sub

Sub NettoyerFeuille(ws As Worksheet) '2025-08-22 @ 16:08

    On Error Resume Next
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0
    
End Sub

Sub SupprimerNomsLocaux(ws As Worksheet) '2025-08-22 @ 16:20

    Dim nom As Name
    For Each nom In ThisWorkbook.Names
        If nom.RefersTo Like "='" & ws.Name & "'!*" Then
            Debug.Print nom.RefersTo
            nom.Delete
        End If
    Next nom
    
End Sub

Sub CopierFormesAvecActions(wsSource As Worksheet, wsDest As Worksheet) '2025-08-22 @ 16:09

    Dim shpSource As Shape
    Dim shpDest As Shape

    For Each shpSource In wsSource.Shapes
        shpSource.Copy
        wsDest.Paste

        Set shpDest = wsDest.Shapes(wsDest.Shapes.count)

        With shpDest
            .Top = shpSource.Top
            .Left = shpSource.Left
            .Width = shpSource.Width
            .Height = shpSource.Height
            On Error Resume Next
            .OnAction = shpSource.OnAction
            On Error GoTo 0
        End With
    Next shpSource
    
End Sub

Sub ReassignerPlagesNomées(wsSource As Worksheet, wsDest As Worksheet) '2025-08-22 @ 16:09

    Dim nom As Name
    Dim nouveauNom As String
    Dim nouvelleRef As String

    For Each nom In ThisWorkbook.Names
        If InStr(1, nom.RefersTo, wsSource.Name, vbTextCompare) > 0 Then
            nouveauNom = nom.Name
            nouvelleRef = Replace(nom.RefersTo, wsSource.Name, wsDest.Name)

            On Error Resume Next
            ThisWorkbook.Names(nouveauNom).Delete 'Supprimer si déjà existant
            ThisWorkbook.Names.Add Name:=nouveauNom, RefersTo:=nouvelleRef
            On Error GoTo 0
        End If
    Next nom
    
End Sub

Sub TracerEtape(nomEtape As String)

    Debug.Print Format(Now, "hh:nn:ss") & " - " & nomEtape
    
End Sub

Sub PauseActive(seconde As Double)

    Dim t0 As Double: t0 = Timer
    Do While Timer < t0 + seconde
        DoEvents
    Loop
    
End Sub


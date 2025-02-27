Attribute VB_Name = "modFAC"
Option Explicit
Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub FAC_Brouillon_New_Invoice() 'Clear contents
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If wshFAC_Brouillon.Range("B27").Value = False Then
        With wshFAC_Brouillon
            .Range("B24").Value = True
            .Range("K3:L7,O3,O5").Clearcontents 'Clear cells for a new Invoice
            .Range("J8:Q46").Clearcontents
            .Range("O6").Value = .Range("FACNextInvoiceNumber").Value 'Paste Invoice ID
            .Range("FACNextInvoiceNumber").Value = .Range("FACNextInvoiceNumber").Value + 1 'Increment Next Invoice ID
            
            Call FAC_Brouillon_Setup_All_Cells
            
            .Range("B20").Value = ""
            .Range("B24").Value = False
            .Range("B26").Value = False
            .Range("B27").Value = True 'Set the value to TRUE
        End With
        
        With wshFAC_Finale
            .Range("B21,B23:C27,E28").Clearcontents
            .Range("A34:F68").Clearcontents
            .Range("E28").Value = wshFAC_Brouillon.Range("O6").Value 'Invoice #
            .Range("B69:F81").Clearcontents 'NOT the formulas
            
            Call FAC_Finale_Setup_All_Cells
        
        End With
        
        wshFAC_Brouillon.Range("B16").Value = False '2024-03-14 @ 08:41
        
        Call FAC_Brouillon_Clear_All_TEC_Displayed
        
        'Move on to CLient Name
        wshFAC_Brouillon.Range("E4:F4").Clearcontents
        With wshFAC_Brouillon.Range("E4:F4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        wshFAC_Brouillon.Select
        wshFAC_Brouillon.Range("E4").Select 'Start inputing values for a NEW invoice
    End If

    'Save button is disabled UNTIL the invoice is saved
    Call FAC_Finale_Disable_Save_Button
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Call Output_Timer_Results("FAC_Brouillon_New_Invoice()", timerStart)

End Sub

Sub Client_Change(ClientName As String)

    Dim timerStart As Double: timerStart = Timer
    
    Dim myInfo() As Variant
    Dim rng As Range: Set rng = wshBD_Clients.Range("dnrClients_Names_Only")
    
    myInfo = Lookup_Data_In_A_Range(rng, 1, ClientName, 3)
    
    If myInfo(1) = "" Then
        MsgBox "Je ne peux retrouver ce client dans ma liste", vbCritical
        GoTo Clean_Exit
    End If
        
    wshFAC_Brouillon.Range("B18").Value = wshBD_Clients.Cells(myInfo(2), 2)
    
    With wshFAC_Brouillon
        .Range("K3").Value = wshBD_Clients.Cells(myInfo(2), 3)
        .Range("K4").Value = ClientName
        .Range("K5").Value = wshBD_Clients.Cells(myInfo(2), 5) 'Adresse1
        If wshBD_Clients.Cells(myInfo(2), 6) <> "" Then
            .Range("K6").Value = wshBD_Clients.Cells(myInfo(2), 6) 'Adresse2
            .Range("K7").Value = wshBD_Clients.Cells(myInfo(2), 7) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 8) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 9) 'Ville, Province & Code postal
        Else
            .Range("K6").Value = wshBD_Clients.Cells(myInfo(2), 7) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 8) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 9) 'Ville, Province & Code postal
            .Range("K7").Value = ""
        End If
    End With
    
    With wshFAC_Finale
        .Range("B23").Value = wshBD_Clients.Cells(myInfo(2), 3)
        .Range("B24").Value = ClientName
        .Range("B25").Value = wshBD_Clients.Cells(myInfo(2), 5) 'Adresse1
        If wshBD_Clients.Cells(myInfo(2), 6) <> "" Then
            .Range("B26").Value = wshBD_Clients.Cells(myInfo(2), 6) 'Adresse2
            .Range("B27").Value = wshBD_Clients.Cells(myInfo(2), 7) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 8) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 9) 'Ville, Province & Code postal
        Else
            .Range("B26").Value = wshBD_Clients.Cells(myInfo(2), 7) & " " & _
                                wshBD_Clients.Cells(myInfo(2), 8) & "  " & _
                                wshBD_Clients.Cells(myInfo(2), 9) 'Ville, Province & Code postal
            .Range("B27").Value = ""
        End If
    End With
    
    Call FAC_Brouillon_Clear_All_TEC_Displayed
    
    wshFAC_Brouillon.Range("O3").Select 'Move on to Invoice Date

Clean_Exit:

    Set rng = Nothing
    
    Call Output_Timer_Results("Client_Change()", timerStart)
    
End Sub

Sub Date_Change(d As String)

    Application.EnableEvents = False
    
    If d = "" Then d = Now()
    
    If InStr(1, wshFAC_Brouillon.Range("O6").Value, "-") = 0 Then
        Dim y As String
        y = Right(Year(d), 2)
        wshFAC_Brouillon.Range("O6").Value = y & "-" & wshFAC_Brouillon.Range("O6").Value
        wshFAC_Finale.Range("E28").Value = wshFAC_Brouillon.Range("O6").Value
    End If
    
    wshFAC_Finale.Range("B21").Value = "Le " & Format(d, "d mmmm yyyy")
    
    'Must Get GST & PST rates and store them in wshFAC_Brouillon 'B' column
    Dim DateTaxRates As Date
    DateTaxRates = d
    wshFAC_Brouillon.Range("B29").Value = GetTaxRate(DateTaxRates, "F")
    wshFAC_Brouillon.Range("B30").Value = GetTaxRate(DateTaxRates, "P")
        
    Dim cutoffDate As Date
    cutoffDate = d
    Call Get_All_TEC_By_Client(cutoffDate, False)
    
    Dim rng As Range
    Set rng = wshFAC_Brouillon.Range("O3")
    Call Fill_Or_Empty_Range_Background(rng, False)
    
    Set rng = wshFAC_Brouillon.Range("L11")
'    Call Fill_Or_Empty_Range_Background(rng, True, 6)

    On Error Resume Next
    wshFAC_Brouillon.Range("L11").Select 'Move on to Services Entry
    On Error GoTo 0
    
    Application.EnableEvents = True
    
End Sub

Sub Inclure_TEC_Factures_Click()

    Dim cutoffDate As Date
    cutoffDate = wshFAC_Brouillon.Range("O3").Value
    
    If wshFAC_Brouillon.Range("B16").Value = True Then
        Call Get_All_TEC_By_Client(cutoffDate, True)
    Else
        Call Get_All_TEC_By_Client(cutoffDate, False)
    End If
    
End Sub

Sub FAC_Finale_Save() '2024-03-28 @ 07:19

    Dim timerStart As Double: timerStart = Timer

    With wshFAC_Brouillon
        'Check For Mandatory Fields - Client
        If .Range("B18").Value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
        'Check For Mandatory Fields - Date de facture
        If .Range("O3").Value = Empty Or Len(Trim(.Range("O6").Value)) <> 8 Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture AVANT de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
        
        'Valid Invoice - Let's update it ******************************************
        
        Call FAC_Finale_Disable_Save_Button
    
        Call FAC_Finale_Add_Invoice_Header_to_DB
        Call FAC_Finale_Add_Invoice_Header_Locally
        
        Call FAC_Finale_Add_Invoice_Details_to_DB
        Call FAC_Finale_Add_Invoice_Details_Locally
        
        Call FAC_Finale_Add_Comptes_Clients_to_DB
        Call FAC_Finale_Add_Comptes_Clients_Locally
        
        Dim lastResultRow As Integer
        lastResultRow = wshTEC_Local.Range("AT9999").End(xlUp).row
        
        If lastResultRow > 2 Then
            Call TEC_Record_Update_As_Billed_To_DB(3, lastResultRow)
            Call TEC_Record_Update_As_Billed_Locally(3, lastResultRow)
        End If
        
        Call FAC_Brouillon_Clear_All_TEC_Displayed
        
    End With
    
    'GL stuff
    Call FAC_Finale_GL_Posting_Preparation
    
    'Update TEC_DashBoard
    Call TEC_DB_Update_All '2024-03-21 @ 12:32

    MsgBox "La facture '" & wshFAC_Brouillon.Range("O6").Value & "' est enregistrée." & vbNewLine & vbNewLine & "Le total de la facture est " & Trim(Format(wshFAC_Brouillon.Range("O51").Value, "### ##0.00 $")) & " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    
    wshFAC_Brouillon.Range("B27").Value = False
    Call FAC_Brouillon_New_Invoice '2024-03-12 @ 08:08 - Maybe ??
    
Fast_Exit_Sub:

    Call Output_Timer_Results("FAC_Finale_Save()", timerStart)
    
    wshFAC_Brouillon.Select
'    Call Goto_Onglet_FAC_Brouillon
    
End Sub

Sub FAC_Finale_Add_Invoice_Header_to_DB()

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'Can only ADD to the file, no modification is allowed
    
    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields("Inv_No") = .Range("E28").Value
        rs.Fields("Date_Facture") = CDate(wshFAC_Brouillon.Range("O3").Value)
        rs.Fields("Cust_ID") = wshFAC_Brouillon.Range("B18").Value
        rs.Fields("Contact") = .Range("B23").Value
        rs.Fields("Nom_Client") = .Range("B24").Value
        rs.Fields("Adresse1") = .Range("B25").Value
        rs.Fields("Adresse2") = .Range("B26").Value
        rs.Fields("Adresse3") = .Range("B27").Value
        
        rs.Fields("Honoraires") = .Range("E69").Value
        
        rs.Fields("AF1_Desc") = .Range("B70").Value
        rs.Fields("AutresFrais_1") = wshFAC_Finale.Range("E70").Value
        rs.Fields("AF2_Desc") = .Range("B71").Value
        rs.Fields("AutresFrais_2") = .Range("E71").Value
        rs.Fields("AF3_Desc") = .Range("B72").Value
        rs.Fields("AutresFrais_3") = .Range("E72").Value
        
        rs.Fields("Taux_TPS") = .Range("C74").Value
        rs.Fields("Mnt_TPS") = .Range("E74").Value
        rs.Fields("Taux_TVQ") = .Range("C75").Value
        rs.Fields("Mnt_TVQ") = .Range("E75").Value
        
        rs.Fields("AR_Total") = .Range("E77").Value
        
        rs.Fields("Dépôt") = .Range("E79").Value
    End With
    'Update the recordset (create the record)
    rs.update
    
    Application.EnableEvents = False
    
    'Prepare GL Posting
    With wshFAC_Brouillon
        .Range("B33").Value = wshFAC_Finale.Range("E81").Value   'AR amount
        .Range("B34").Value = -wshFAC_Finale.Range("E69").Value  'Revenues
        .Range("B35").Value = -wshFAC_Finale.Range("E70").Value  'Misc $ - 1
        .Range("B36").Value = -wshFAC_Finale.Range("E71").Value  'Misc $ - 2
        .Range("B37").Value = -wshFAC_Finale.Range("E72").Value  'Misc $ - 3
        .Range("B38").Value = -wshFAC_Finale.Range("E74").Value  'GST $
        .Range("B39").Value = -wshFAC_Finale.Range("E75").Value  'PST $
        .Range("B40").Value = wshFAC_Finale.Range("E79").Value   'Deposit
    End With
    
    Application.EnableEvents = True
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Release objects from memory
    Set rs = Nothing
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("FAC_Finale_Add_Invoice_Header_to_DB()", timerStart)

End Sub

Sub FAC_Finale_Add_Invoice_Header_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Entête.Range("A9999").End(xlUp).row + 1
    
    With wshFAC_Entête
        .Range("A" & firstFreeRow).Value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).Value = Format(wshFAC_Brouillon.Range("O3").Value, "dd/mm/yyyy")
        .Range("C" & firstFreeRow).Value = wshFAC_Brouillon.Range("B18").Value
        .Range("D" & firstFreeRow).Value = wshFAC_Finale.Range("B23").Value
        .Range("E" & firstFreeRow).Value = wshFAC_Finale.Range("B24").Value
        .Range("F" & firstFreeRow).Value = wshFAC_Finale.Range("B25").Value
        .Range("G" & firstFreeRow).Value = wshFAC_Finale.Range("B26").Value
        .Range("H" & firstFreeRow).Value = wshFAC_Finale.Range("B27").Value
        
        .Range("I" & firstFreeRow).Value = wshFAC_Finale.Range("E69").Value
        
        .Range("J" & firstFreeRow).Value = wshFAC_Finale.Range("B70").Value
        .Range("K" & firstFreeRow).Value = wshFAC_Finale.Range("E70").Value
        .Range("L" & firstFreeRow).Value = wshFAC_Finale.Range("B71").Value
        .Range("M" & firstFreeRow).Value = wshFAC_Finale.Range("E71").Value
        .Range("N" & firstFreeRow).Value = wshFAC_Finale.Range("B72").Value
        .Range("O" & firstFreeRow).Value = wshFAC_Finale.Range("E72").Value
        
        .Range("P" & firstFreeRow).Value = wshFAC_Finale.Range("C74").Value
        .Range("Q" & firstFreeRow).Value = wshFAC_Finale.Range("E74").Value
        .Range("R" & firstFreeRow).Value = wshFAC_Finale.Range("C75").Value
        .Range("S" & firstFreeRow).Value = wshFAC_Finale.Range("E75").Value
        
        .Range("T" & firstFreeRow).Value = wshFAC_Finale.Range("E77").Value
        
        .Range("U" & firstFreeRow).Value = wshFAC_Finale.Range("E79").Value
    End With
    
    Call Output_Timer_Results("FAC_Finale_Add_Invoice_Header_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_Add_Invoice_Details_to_DB()

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim rowLastService As Long
    rowLastService = wshFAC_Finale.Range("B64").End(xlUp).row
    If rowLastService < 34 Then GoTo nothing_to_update
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Détails"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim r As Integer
    For r = 34 To rowLastService
        'Add fields to the recordset before updating it
        rs.AddNew
        With wshFAC_Finale
            rs.Fields("Inv_No") = .Range("E28").Value
            rs.Fields("Description") = .Range("B" & r).Value
            rs.Fields("Heures") = .Range("C" & r).Value
            rs.Fields("Taux") = .Range("D" & r).Value
            If .Range("E" & r).Value <> "" Then
                rs.Fields("Honoraires") = .Range("E" & r).Value
            End If
            rs.Fields("Inv_Row") = r
            'rs.Fields("Row") = ""
        End With
    'Update the recordset (create the record)
    rs.update
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Release objects from memory
    Set rs = Nothing
    Set conn = Nothing
    
nothing_to_update:

    Application.ScreenUpdating = True

    Call Output_Timer_Results("FAC_Finale_Add_Invoice_Details_to_DB()", timerStart)

End Sub

Sub FAC_Finale_Add_Invoice_Details_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Get the last entered service
    Dim lastEnteredService As Long
    lastEnteredService = wshFAC_Finale.Range("B64").End(xlUp).row
    If lastEnteredService < 34 Then GoTo nothing_to_update
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Détails.Range("A99999").End(xlUp).row + 1
   
    Dim i As Integer
    For i = 34 To lastEnteredService
        With wshFAC_Détails
            .Range("A" & firstFreeRow).Value = wshFAC_Finale.Range("E28")
            .Range("B" & firstFreeRow).Value = wshFAC_Finale.Range("B" & i).Value
            .Range("C" & firstFreeRow).Value = wshFAC_Finale.Range("C" & i).Value
            .Range("D" & firstFreeRow).Value = wshFAC_Finale.Range("D" & i).Value
            .Range("E" & firstFreeRow).Value = wshFAC_Finale.Range("E" & i).Value
            .Range("F" & firstFreeRow).Value = i
            firstFreeRow = firstFreeRow + 1
        End With
    Next i

nothing_to_update:

    Call Output_Timer_Results("FAC_Finale_Add_Invoice_Details_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_Add_Comptes_Clients_to_DB()

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Comptes_Clients"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields("Invoice_No") = .Range("E28").Value
        rs.Fields("Invoice_Date") = CDate(wshFAC_Brouillon.Range("O3").Value)
        rs.Fields("Customer") = .Range("B24").Value
        rs.Fields("Status") = "Unpaid"
        rs.Fields("Terms") = "Net 30"
        rs.Fields("Due_Date") = CDate(CDate(wshFAC_Brouillon.Range("O3").Value) + 30)
        rs.Fields("Total") = .Range("E81").Value
        'rs.Fields("Total_Paid") = ""
        'rs.Fields("Balance") = ""
        'rs.Fields("Days_Overdue") = ""
    End With
    
    'Update the recordset (create the record)
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Release objects from memory
    Set rs = Nothing
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("FAC_Finale_Add_Comptes_Clients_to_DB()", timerStart)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_Locally() '2024-03-11 @ 08:49 - Write records locally
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshCC.Range("A9999").End(xlUp).row + 1
   
    With wshCC
        .Range("A" & firstFreeRow).Value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).Value = wshFAC_Brouillon.Range("O3").Value
        .Range("C" & firstFreeRow).Value = wshFAC_Finale.Range("B24").Value
        .Range("D" & firstFreeRow).Value = "Unpaid"
        .Range("E" & firstFreeRow).Value = "Net 30"
        .Range("F" & firstFreeRow).Value = CDate(CDate(wshFAC_Brouillon.Range("O3").Value) + 30)
        .Range("G" & firstFreeRow).Value = wshFAC_Finale.Range("E81").Value
        .Range("H" & firstFreeRow).formula = ""
        .Range("I" & firstFreeRow).formula = "=G" & firstFreeRow & "-H" & firstFreeRow
        .Range("J" & firstFreeRow).formula = "=IF(H" & firstFreeRow & "<G" & firstFreeRow & ",NOW()-F" & firstFreeRow & ")"
    End With

nothing_to_update:

    Call Output_Timer_Results("FAC_Finale_Add_Comptes_Clients_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub TEC_Record_Update_As_Billed_To_DB(firstRow As Integer, lastRow As Integer) 'Update Billed Status in DB

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "TEC"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    Dim r As Integer, TEC_ID As Long, SQL As String
    For r = firstRow To lastRow
        If wshTEC_Local.Range("BD" & r).Value = True Or _
            wshFAC_Brouillon.Range("C" & r + 5) <> True Then
            GoTo next_iteration
        End If
        TEC_ID = wshTEC_Local.Range("AT" & r).Value
        
        'Open the recordset for the specified ID
        SQL = "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & TEC_ID
        rs.Open SQL, conn, 2, 3
        If Not rs.EOF Then
            'Update DateSaisie, EstFacturee, DateFacturee & NoFacture
            rs.Fields("DateSaisie").Value = Now
            rs.Fields("EstFacturee").Value = True
            rs.Fields("DateFacturee").Value = Now
            rs.Fields("VersionApp").Value = gAppVersion
            rs.Fields("NoFacture").Value = wshFAC_Brouillon.Range("O6").Value
            rs.update
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut être trouvé!", _
                vbExclamation
            rs.Close
            conn.Close
            Exit Sub
        End If
        'Update the recordset (create the record)
        rs.update
        rs.Close
next_iteration:
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Set rs = Nothing
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("TEC_Record_Update_As_Billed_To_DB()", timerStart)

End Sub

Sub TEC_Record_Update_As_Billed_Locally(firstResultRow As Integer, lastResultRow As Integer)

    Dim timerStart As Double: timerStart = Timer
    
    'Set the range to look for
    Dim lookupRange As Range, lastTECRow As Long
    lastTECRow = wshTEC_Local.Range("A99999").End(xlUp).row
    Set lookupRange = wshTEC_Local.Range("A3:A" & lastTECRow)
    
    Dim r As Integer, rowToBeUpdated As Long, tecID As Long
    For r = firstResultRow To lastResultRow
        If wshTEC_Local.Range("BD" & r).Value = False And _
            wshFAC_Brouillon.Range("C" & r + 5) = True Then
            tecID = wshTEC_Local.Range("AT" & r).Value
            rowToBeUpdated = Get_TEC_Row_Number_By_TEC_ID(tecID, lookupRange)
            wshTEC_Local.Range("K" & rowToBeUpdated).Value = Now()
            wshTEC_Local.Range("L" & rowToBeUpdated).Value = True
            wshTEC_Local.Range("M" & rowToBeUpdated).Value = Now()
            wshTEC_Local.Range("O" & rowToBeUpdated).Value = gAppVersion
            wshTEC_Local.Range("P" & rowToBeUpdated).Value = wshFAC_Brouillon.Range("O6").Value
        End If
    Next r
    
    Call Output_Timer_Results("TEC_Record_Update_As_Billed_Locally()", timerStart)

End Sub

Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16
    With wshFAC_Brouillon
        If wshFAC_Brouillon.Range("B20").Value = Empty Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
            GoTo NoItems
        End If
        'Could that invoice been cancelled (more than 1 row) ?
        Call InvoiceGetAllTrans(wshFAC_Brouillon.Range("O6").Value)
        Dim NbTrans As Integer
        NbTrans = .Range("B31").Value
        If NbTrans = 0 Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
            GoTo NoItems
        Else
            If NbTrans > 1 Then
                MsgBox "Cette facture a été annulée! Veuillez saisir un numéro de facture VALIDE pour votre recherche"
                GoTo NoItems
            End If
        End If
        .Range("B24").Value = True 'Set Invoice Load to true
        .Range("S2,E4:F4,K4:L6,O3,K11:O45,Q11:Q45").Clearcontents
        wshFAC_Finale.Range("C34:F69").Clearcontents
        Dim InvListRow As Long
        InvListRow = wshFAC_Brouillon.Range("B20").Value 'InvListRow = Row associated with the invoice
        'Get values from wshFAC_Entête (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
        .Range("O3").Value = wshFAC_Entête.Range("B" & InvListRow).Value
        .Range("K3").Value = wshFAC_Entête.Range("D" & InvListRow).Value
        .Range("K4").Value = wshFAC_Entête.Range("E" & InvListRow).Value
        .Range("K5").Value = wshFAC_Entête.Range("F" & InvListRow).Value
        .Range("K6").Value = wshFAC_Entête.Range("G" & InvListRow).Value
        'Get values from wshFAC_Entête (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
        wshFAC_Finale.Range("B21").Value = "Le " & Format(wshFAC_Entête.Range("B" & InvListRow).Value, "d mmmm yyyy")
        wshFAC_Finale.Range("B23").Value = wshFAC_Entête.Range("D" & InvListRow).Value
        wshFAC_Finale.Range("B24").Value = wshFAC_Entête.Range("E" & InvListRow).Value
        wshFAC_Finale.Range("B25").Value = wshFAC_Entête.Range("F" & InvListRow).Value
        wshFAC_Finale.Range("B26").Value = wshFAC_Entête.Range("G" & InvListRow).Value
        'Load Invoice Detail Items
        With wshFAC_Détails
            Dim lastRow As Long, lastResultRow As Long
            lastRow = .Range("A999999").End(xlUp).row
            If lastRow < 4 Then Exit Sub 'No Item Lines
            .Range("I3").Value = wshFAC_Brouillon.Range("O6").Value
            wshFAC_Finale.Range("F28").Value = wshFAC_Brouillon.Range("O6").Value 'Invoice #
            'Advanced Filter to get items specific to ONE invoice
            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
            lastResultRow = .Range("O999").End(xlUp).row
            If lastResultRow < 3 Then GoTo NoItems
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("O" & resultRow).Value
                wshFAC_Brouillon.Range("L" & invitemRow & ":O" & invitemRow).Value = .Range("K" & resultRow & ":N" & resultRow).Value 'Description, Hours, Rate & Value
                wshFAC_Brouillon.Range("Q" & invitemRow).Value = .Range("P" & resultRow).Value  'Set Item DB Row
                wshFAC_Finale.Range("C" & invitemRow + 23 & ":F" & invitemRow + 23).Value = .Range("K" & resultRow & ":N" & resultRow).Value 'Description, Hours, Rate & Value
            Next resultRow
        End With
        'Proceed with trailer data (Misc. charges & Taxes)
        .Range("M48").Value = wshFAC_Entête.Range("I" & InvListRow).Value
        .Range("O48").Value = wshFAC_Entête.Range("J" & InvListRow).Value
        .Range("M49").Value = wshFAC_Entête.Range("K" & InvListRow).Value
        .Range("O49").Value = wshFAC_Entête.Range("L" & InvListRow).Value
        .Range("M50").Value = wshFAC_Entête.Range("M" & InvListRow).Value
        .Range("O50").Value = wshFAC_Entête.Range("N" & InvListRow).Value
        .Range("O52").Value = wshFAC_Entête.Range("P" & InvListRow).Value
        .Range("O53").Value = wshFAC_Entête.Range("R" & InvListRow).Value
        .Range("O57").Value = wshFAC_Entête.Range("T" & InvListRow).Value

NoItems:
    .Range("B24").Value = False 'Set Invoice Load To false
    End With
End Sub

Sub InvoiceGetAllTrans(inv As String)

    Application.ScreenUpdating = False

    wshFAC_Brouillon.Range("B31").Value = 0

    With wshFAC_Entête
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last wshFAC_Entête Row
        If lastRow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
        On Error Resume Next
        .Names("Criterial").delete
        On Error GoTo 0
        .Range("V3").Value = wshFAC_Brouillon.Range("O6").Value
        'Advanced Filter setup
        .Range("A3:T" & lastRow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("V2:V3"), _
            CopyToRange:=.Range("X2:AQ2"), _
            Unique:=True
        lastResultRow = .Range("X999").End(xlUp).row 'How many rows trans for that invoice
        If lastResultRow < 3 Then
            GoTo Done
        End If
'        With .Sort
'            .SortFields.Clear
'            .SortFields.Add Key:=wshFAC_Entête.Range("X2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based Invoice Number
'            .SortFields.Add Key:=wshGL_Trans.Range("Y3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based On TEC_ID
'            .SetRange wshFAC_Entête.Range("X2:AQ" & lastResultRow) 'Set Range
'            .Apply 'Apply Sort
'         End With
         wshFAC_Brouillon.Range("B31").Value = lastResultRow - 2 'Remove Header rows from row count
Done:
    End With
    Application.ScreenUpdating = True

End Sub

Sub FAC_Brouillon_Setup_All_Cells()

    Application.EnableEvents = False
    
    With wshFAC_Brouillon
        .Range("J47:P60").Clearcontents
        
        Call SetLabels(.Range("K47"), "FAC_Label_SubTotal_1")
        Call SetLabels(.Range("K51"), "FAC_Label_SubTotal_2")
        Call SetLabels(.Range("K52"), "FAC_Label_TPS")
        Call SetLabels(.Range("K53"), "FAC_Label_TVQ")
        Call SetLabels(.Range("K55"), "FAC_Label_GrandTotal")
        Call SetLabels(.Range("K57"), "FAC_Label_Deposit")
        Call SetLabels(.Range("K59"), "FAC_Label_AmountDue")
        
        .Range("M47").formula = "=IF(SUM(M11:M45),SOMME(M11:M45),B19)"   'Total hours entered OR TEC selected"
        .Range("N47").formula = wshAdmin.Range("TauxHoraireFacturation") 'Rate per hour
        .Range("O47").formula = "=M47*N47"                               'Fees sub-total
        .Range("O47").Font.Bold = True
        
        .Range("M48").Value = wshAdmin.Range("FAC_Label_Frais_1").Value 'Misc. # 1 - Descr.
        .Range("O48").Value = ""                                        'Misc. # 1 - Amount
        .Range("M49").Value = wshAdmin.Range("FAC_Label_Frais_2").Value 'Misc. # 2 - Descr.
        .Range("O49").Value = ""                                        'Misc. # 2 - Amount
        .Range("M50").Value = wshAdmin.Range("FAC_Label_Frais_3").Value 'Misc. # 3 - Descr.
        .Range("O50").Value = ""                                        'Misc. # 3 - Amount
        
        .Range("O51").formula = "=sum(O47:O50)"                         'Sub-total
        .Range("O51").Font.Bold = True
        
        .Range("N52").Value = wshFAC_Brouillon.Range("B29").Value       'GST Rate
        .Range("N52").NumberFormat = "0.00%"
        .Range("O52").formula = "=round(o51*n52,2)"                     'GST Amnt
        .Range("N53").Value = wshFAC_Brouillon.Range("B30").Value       'PST Rate
        .Range("N53").NumberFormat = "0.000%"
        .Range("O53").formula = "=round(o51*n53,2)"                     'PST Amnt
        .Range("O55").formula = "=sum(o51:o54)"                         'Grand Total"
        .Range("O57").Value = ""
        .Range("O59").formula = "=O55-O57"                              'Deposit Amount
        
    End With
    
    Application.EnableEvents = True

End Sub

Sub FAC_Finale_Setup_All_Cells()

    Dim timerStart As Double: timerStart = Timer
    
    Application.EnableEvents = False
    
    With wshFAC_Finale
        .Range("B21").formula = "= ""Le "" & TEXT(FAC_Brouillon!O3, ""j MMMM aaaa"")"
        .Range("B23:B27").Value = ""
        .Range("E28").Value = "=" & wshFAC_Brouillon.name & "!O6"    'Invoice number
        
        .Range("C65").Value = "Heures"                               'Summary Heading
        .Range("D65").Value = "Taux"                                 'Summary Heading
        .Range("C66").formula = "=" & wshFAC_Brouillon.name & "!M47" 'Hours summary
        .Range("D66").formula = "=" & wshFAC_Brouillon.name & "!N47" 'Hourly Rate
        
        With .Range("C65:D66")
            .Font.ThemeColor = xlThemeColorLight1
            .Font.TintAndShade = 0
        End With

        Call SetLabels(.Range("B69"), "FAC_Label_SubTotal_1")
        Call SetLabels(.Range("B73"), "FAC_Label_SubTotal_2")
        Call SetLabels(.Range("B74"), "FAC_Label_TPS")
        Call SetLabels(.Range("B75"), "FAC_Label_TVQ")
        Call SetLabels(.Range("B77"), "FAC_Label_GrandTotal")
        Call SetLabels(.Range("B79"), "FAC_Label_Deposit")
        Call SetLabels(.Range("B81"), "FAC_Label_AmountDue")

        .Range("E69").formula = "=C66*D66"                           'Fees Sub-Total
        
        .Range("B70").Value = "='" & wshFAC_Brouillon.name & "'!M48" 'Misc. Amount # 1 - Description
        .Range("E70").Value = "='" & wshFAC_Brouillon.name & "'!O48" 'Misc. Amount # 1
        
        .Range("B71").Value = "='" & wshFAC_Brouillon.name & "'!M49" 'Misc. Amount # 2 - Description
        .Range("E71").Value = "='" & wshFAC_Brouillon.name & "'!O49" 'Misc. Amount # 2
        
        .Range("B72").Value = "='" & wshFAC_Brouillon.name & "'!M50" 'Misc. Amount # 3 - Description
        .Range("E72").Value = "='" & wshFAC_Brouillon.name & "'!O50" 'Misc. Amount # 3
        
        .Range("E73").formula = "=SUM(E69:E72)"                      'Invoice Sub-Total
        
        .Range("C74").Value = "='" & wshFAC_Brouillon.name & "'!N52" 'GST Rate
        .Range("E74").formula = "=round(E73*C74,2)"                  'GST Amount"
        .Range("C75").Value = "='" & wshFAC_Brouillon.name & "'!N53" 'PST Rate
        .Range("E75").formula = "=round(E73*C75,2)"                  'PST Amount
        
        .Range("E77").Value = "=SUM(E73:E75)"                        'Total including taxes
        .Range("E79").Value = "='" & wshFAC_Brouillon.name & "'!O57" 'Deposit Amount
        .Range("E81").Value = "=E77-E79"                             'Total due on that invoice
    End With
    
    Application.EnableEvents = True
    
    Call Output_Timer_Results("FAC_Finale_Setup_All_Cells()", timerStart)

End Sub

Sub SetLabels(r As Range, l As String)

    r.Value = wshAdmin.Range(l).Value
    If wshAdmin.Range(l & "_Bold").Value = "OUI" Then r.Font.Bold = True

End Sub

Sub FAC_Brouillon_Goto_Misc_Charges()
    
    ActiveWindow.SmallScroll Down:=6
    wshFAC_Brouillon.Range("M47").Select 'Hours Summary
    
End Sub

Sub FAC_Brouillon_Clear_All_TEC_Displayed()

    Dim timerStart As Double: timerStart = Timer
    
    Application.EnableEvents = False
    
    Dim lastRow As Long
    lastRow = wshFAC_Brouillon.Range("D999").End(xlUp).row
    If lastRow > 7 Then
        wshFAC_Brouillon.Range("D8:I" & lastRow + 2).Clearcontents
        Call FAC_Brouillon_TEC_Remove_Check_Box(lastRow)
    End If
    
    Application.EnableEvents = True

    Call Output_Timer_Results("FAC_Brouillon_Clear_All_TEC_Displayed()", timerStart)

End Sub

Sub Get_All_TEC_By_Client(d As Date, includeBilledTEC As Boolean)

    'Set all criteria before calling FAC_Brouillon_TEC_Advanced_Filter_And_Sort
    Dim c1 As Long, c2 As String, c3 As Boolean
    Dim c4 As Boolean, c5 As Boolean
    c1 = wshFAC_Brouillon.Range("B18").Value
    c2 = "<=" & Format(d, "mm-dd-yyyy")
    c3 = True
    If includeBilledTEC Then c4 = True Else c4 = False
    c5 = False

    Call FAC_Brouillon_Clear_All_TEC_Displayed
    Call FAC_Brouillon_TEC_Advanced_Filter_And_Sort(c1, c2, c3, c4, c5)
    Call TEC_Filtered_Entries_Copy_To_FAC_Brouillon
    
End Sub

Sub FAC_Brouillon_TEC_Advanced_Filter_And_Sort(clientID As Long, _
        cutoffDate As String, _
        isBillable As Boolean, _
        isInvoiced As Boolean, _
        isDeleted As Boolean)
    
    Application.ScreenUpdating = False

    With wshTEC_Local
        'Is there anything to filter ?
        Dim lastSourceRow As Long, lastResultRow As Long
        lastSourceRow = .Range("A99999").End(xlUp).row 'Last TEC Entry row
        If lastSourceRow < 3 Then Exit Sub 'Nothing to filter
        
        'Clear the filtered rows area
        lastResultRow = .Range("AT9999").End(xlUp).row
        If lastResultRow > 2 Then .Range("AT3:BH" & lastResultRow).Clearcontents
        
        Dim rngSource As Range, rngCriteria As Range, rngCopyToRange As Range
        Set rngSource = wshTEC_Local.Range("A2:P" & lastSourceRow)
        If clientID <> 0 Then .Range("AN3").Value = clientID
        .Range("AO3").Value = cutoffDate
        .Range("AP3").Value = isBillable
        If isInvoiced <> True Then
            .Range("AQ3").Value = isInvoiced
        Else
            .Range("AQ3").Value = ""
        End If
        .Range("AR3").Value = isDeleted
        Set rngCriteria = .Range("AN2:AR3")
        Set rngCopyToRange = .Range("AT2:BH2")
        
        rngSource.AdvancedFilter xlFilterCopy, rngCriteria, rngCopyToRange, Unique:=True
        
        lastResultRow = .Range("AT9999").End(xlUp).row
        If lastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.clear
            .SortFields.add Key:=wshTEC_Local.Range("AW3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.add Key:=wshTEC_Local.Range("AU3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof_ID
            .SortFields.add Key:=wshTEC_Local.Range("AT3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshTEC_Local.Range("AT3:BH" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
No_Sort_Required:
    End With
    
    Application.ScreenUpdating = True

End Sub

Sub TEC_Filtered_Entries_Copy_To_FAC_Brouillon() '2024-03-21 @ 07:10

    Dim timerStart As Double: timerStart = Timer

    Dim lastUsedRow As Long
    lastUsedRow = wshTEC_Local.Range("AT9999").End(xlUp).row
    If lastUsedRow < 3 Then Exit Sub 'No rows
    
    Application.ScreenUpdating = False
    
    Dim arr() As Variant, totalHres As Double
    ReDim arr(1 To (lastUsedRow - 2), 1 To 6) As Variant
    With wshTEC_Local
        Dim i As Integer
        For i = 3 To lastUsedRow
            arr(i - 2, 1) = .Range("AW" & i).Value 'Date
            arr(i - 2, 2) = .Range("AV" & i).Value 'Prof
            arr(i - 2, 3) = .Range("AY" & i).Value 'Description
            arr(i - 2, 4) = .Range("AZ" & i).Value 'Heures
            totalHres = totalHres + .Range("AZ" & i).Value
            arr(i - 2, 5) = .Range("BD" & i).Value 'Facturée ou pas
            arr(i - 2, 6) = .Range("AT" & i).Value 'TEC_ID
        Next i
        'Copy array to worksheet
        Dim rng As Range
        'Set rng = .Range("D8").Resize(UBound(arr, 1), UBound(arr, 2))
        Set rng = wshFAC_Brouillon.Range("D8").Resize(lastUsedRow - 2, UBound(arr, 2))
        rng.Value = arr
    End With
    
    With wshFAC_Brouillon
        .Range("D8:H" & lastRow + 7).Font.Color = vbBlack
        .Range("D8:H" & lastRow + 7).Font.Bold = False
        
        .Range("G" & lastUsedRow + 7).Value = totalHres
        .Range("G8:G" & lastUsedRow + 7).NumberFormat = "##0.00"
    End With
        
    Call FAC_Brouillon_TEC_Add_Check_Box(lastUsedRow)

    Application.ScreenUpdating = True

    Call Output_Timer_Results("TEC_Filtered_Entries_Copy_To_FAC_Brouillon()", timerStart)
    
End Sub
 
Sub Invoice_Delete()
    With wshFAC_Brouillon
        If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Invoice") = vbNo Then Exit Sub
        If .Range("B20").Value = Empty Then GoTo NotSaved
        invRow = .Range("B20").Value 'Set Invoice Row
        wshFAC_Entête.Range(invRow & ":" & invRow).EntireRow.delete
'        With InvItems
'            lastRow = .Range("A99999").End(xlUp).row
'            If lastRow < 4 Then Exit Sub
'            .Range("A3:J" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("N2:N3"), CopyToRange:=.Range("P2:W2"), Unique:=True
'            lastResultRow = .Range("V99999").End(xlUp).row
'            If lastResultRow < 3 Then GoTo NoItems
'    '        If lastResultRow < 4 Then GoTo SkipSort
'    '        'Sort Rows Descending
'    '         With .Sort
'    '         .SortFields.Clear
'    '         .SortFields.Add Key:=wshFAC_Détails.Range("W3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal  'Sort
'    '         .SetRange wshFAC_Détails.Range("P3:W" & lastResultRow) 'Set Range
'    '         .Apply 'Apply Sort
'    '         End With
'SkipSort:
'            For ResultRow = 3 To lastResultRow
'                itemDBRow = .Range("V" & ResultRow).Value 'Set Invoice Database Row
'                .Range("A" & itemDBRow & ":J" & itemDBRow).ClearContents 'Clear Fields (deleting creates issues with results
'            Next ResultRow
'            'Resort DB to remove spaces
'            With .Sort
'                .SortFields.Clear
'                .SortFields.Add Key:=wshFAC_Détails.Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal  'Sort
'                .SetRange wshFAC_Détails.Range("A4:J" & lastResultRow) 'Set Range
'                .Apply 'Apply Sort
'            End With
'        End With
NoItems:
NotSaved:
    Call FAC_Brouillon_New_Invoice 'Add New Invoice
    End With
End Sub

Sub FAC_BROUILLON_Prev_PDF() '2024-03-02 @ 16:18

    Call Goto_Onglet_FAC_Finale
    Call FAC_Finale_Preview_PDF
    Call Goto_Onglet_FAC_Brouillon
    
End Sub

Sub FAC_Finale_Preview_PDF() '2024-03-02 @ 16:18

    wshFAC_Finale.PrintOut , , 1, True, True, , , , False
'    wshFAC_Finale.PrintOut , , , True, True, , , , False
    
End Sub

Sub FAC_Finale_Creation_PDF_And_Email() 'RMV - 2023-12-17 @ 14:35
    
    Call FAC_Finale_Create_PDF_Email_Sub(wshFAC_Finale.Range("E28").Value)
    
    Call FAC_Finale_Enable_Save_Button

End Sub

Sub FAC_Finale_Create_PDF_Email_Sub(noFacture As String)

    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF et préparation du courriel pour envoyer la facture
    Dim result As Boolean
    result = FAC_Finale_Create_PDF_Email_Func(noFacture, "CreateEmail")

End Sub

Function FAC_Finale_Create_PDF_Email_Func(noFacture As String, Optional action As String = "SaveOnly") As Boolean
    
    Dim SaveAs As String

    Application.ScreenUpdating = False

    'Construct the SaveAs filename
    SaveAs = wshAdmin.Range("FolderPDFInvoice").Value & Application.PathSeparator & _
                     noFacture & ".pdf" '2023-12-19 @ 07:28

    'Set Print Quality
    On Error Resume Next
    ActiveSheet.PageSetup.PrintQuality = 600
    Err.clear
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
    
    'Construct & Display the Email, allowing the user to modify the Email
    If action = "CreateEmail" Then
        On Error GoTo SaveOnly
        
        Dim outlookApp As Outlook.Application
        Set outlookApp = New Outlook.Application
        
        'Where are the email templates ? - 2024-03-27 @ 07:28
        Dim FullTemplatePathAndFile As String
        If userName <> "Robert M. Vigneault" Then
            FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\Test_de_gabarit.oft"
        Else
            FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\Test_de_gabarit.oft"
        End If

        Dim myMail As Outlook.MailItem
        Set myMail = outlookApp.CreateItemFromTemplate(FullTemplatePathAndFile)
'        Set myMail = outlookApp.CreateItem(olMailItem)

        Dim source_file As String
        source_file = wshAdmin.Range("FolderPDFInvoice").Value & Application.PathSeparator & _
                      noFacture & ".pdf" '2023-12-19 @ 07:22
        
        With myMail
            .To = "robertv13@hotmail.com"
            .CC = "robertv13@me.com"
            .BCC = "robertv13@gmail.com"
            .Subject = "TEST - GC FISCALITÉ INC. - Facturation - TEST"
'            .Body = "Bonjour," & vbNewLine & vbNewLine & "Vous trouverez ci-joint notre note d'honoraires." & _
                vbNewLine & vbNewLine & "Merci" & vbNewLine & vbNewLine & vbNewLine & "Guillaume Charron, CPA, CA, M. Fisc." & _
                vbNewLine & "Président"
            .Attachments.add source_file
           
            .Display 'Affiche le courriel, ce qui permet de corriger AVANT l'envoi
            'myMail.Send
        End With
        
        Set outlookApp = Nothing
        Set myMail = Nothing

    End If
    
SaveOnly:
    FAC_Finale_Create_PDF_Email_Func = True 'Return value
    GoTo EndMacro
    
RefLibError:
    MsgBox "Incapable de préparer le courriel. La librairie n'est pas disponible"
    FAC_Finale_Create_PDF_Email_Func = False 'Return value

EndMacro:
    Application.ScreenUpdating = True
    
End Function

Sub Prev_Invoice() 'TO-DO-RMV 2023-12-17
    With wshFAC_Brouillon
        Dim MininvNumb As Long
        On Error Resume Next
        MininvNumb = Application.WorksheetFunction.Min(wshFAC_Entête.Range("Inv_ID"))
        On Error GoTo 0
        If MininvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").Value
        If invNumb = 0 Or .Range("B20").Value = Empty Then 'On New Invoice
            invRow = wshFAC_Entête.Range("A99999").End(xlUp).row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFAC_Entête.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).row - 1
        End If
        If .Range("N6").Value = 1 Or MininvNumb = 0 Or MininvNumb = .Range("N6").Value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").Value = wshFAC_Entête.Range("A" & invRow).Value 'Place Inv. ID inside cell
        Invoice_Load
    End With
End Sub

Sub Next_Invoice() 'TO-DO-RMV 2023-12-17
    With wshFAC_Brouillon
        Dim MaxinvNumb As Long
        On Error Resume Next
        MaxinvNumb = Application.WorksheetFunction.Max(wshFAC_Entête.Range("Inv_ID"))
        On Error GoTo 0
        If MaxinvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").Value
        If invNumb = 0 Or .Range("B20").Value = Empty Then 'On New Invoice
            invRow = wshFAC_Entête.Range("A4").Value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFAC_Entête.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).row + 1
        End If
        If .Range("N6").Value >= MaxinvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").Value = wshFAC_Entête.Range("A" & invRow).Value 'Place Inv. ID inside cell
        Invoice_Load
    End With
End Sub

Sub Cacher_Heures()
    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Montrer_Heures()
    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Cacher_Sommaire_Heures()
    With wshFAC_Finale.Range("C65:D66")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Montrer_Sommaire_Heures()
    With wshFAC_Finale.Range("C65:D66")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Goto_Onglet_FAC_Brouillon()

    Dim timerStart As Double: timerStart = Timer
   
    Application.ScreenUpdating = False
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Brouillon.Range("E4").Select

    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("Goto_Onglet_FAC_Brouillon()", timerStart)

End Sub

Sub Goto_Onglet_FAC_Finale()

    Dim timerStart As Double: timerStart = Timer
   
    Application.ScreenUpdating = False
    
    Call Cacher_Heures
    Call Cacher_Sommaire_Heures
    
    wshFAC_Finale.Visible = xlSheetVisible
    wshFAC_Finale.Activate
    wshFAC_Finale.Range("I50").Select
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("Goto_Onglet_FAC_Finale()", timerStart)

End Sub

Sub ExportAllFacInvList() '2023-12-21 @ 14:36
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceRange As Range

    Application.ScreenUpdating = False
    
    'Work with the source range
    Set wsSource = wshFAC_Entête
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Range("A99999").End(xlUp).row
    wsSource.Range("A4:T" & lastUsedRow).Copy

    'Open the target workbook
    Workbooks.Open fileName:=wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"

    'Set references to the target workbook and target worksheet
    Set wb = Workbooks("GCF_BD_Sortie.xlsx")
    Set wsTarget = wb.Sheets("FACTURES")

    'PasteSpecial directly to the target range
    wsTarget.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    wb.Close SaveChanges:=True
    
    Application.ScreenUpdating = True
    
End Sub

Sub FAC_Finale_GL_Posting_Preparation() '2024-02-14 @ 05:56

    Dim timerStart As Double: timerStart = Timer

    Dim montant As Double
    Dim dateFact As Date
    Dim descGL_Trans As String, source As String
    
    dateFact = wshFAC_Brouillon.Range("O3").Value
    descGL_Trans = wshFAC_Brouillon.Range("E4").Value
    source = "FACT-" & wshFAC_Brouillon.Range("O6").Value
    
    Dim myArray(1 To 8, 1 To 4) As String
    
    'AR amount (wshFAC_Brouillon.Range("B33"))
    montant = wshFAC_Brouillon.Range("B33").Value
    If montant Then
        myArray(1, 1) = "1100"
        myArray(1, 2) = "Comptes Clients"
        myArray(1, 3) = montant
        myArray(1, 4) = ""
    End If
    
    'Professional Fees (wshFAC_Brouillon.Range("B34"))
    montant = wshFAC_Brouillon.Range("B34").Value
    If montant Then
        myArray(2, 1) = "4000"
        myArray(2, 2) = "Revenus"
        myArray(2, 3) = montant
        myArray(2, 4) = ""
    End If
    
    'Miscellaneous Amount # 1 (wshFAC_Brouillon.Range("B35"))
    montant = wshFAC_Brouillon.Range("B35").Value
    If montant Then
        myArray(3, 1) = "5009"
        myArray(3, 2) = "Frais divers # 1"
        myArray(3, 3) = montant
        myArray(3, 4) = ""
    End If
    
    'Miscellaneous Amount # 2 (wshFAC_Brouillon.Range("B36"))
    montant = wshFAC_Brouillon.Range("B36").Value
    If montant Then
        myArray(4, 1) = "5008"
        myArray(4, 2) = "Frais divers # 2"
        myArray(4, 3) = montant
        myArray(4, 4) = ""
    End If
    
    'Miscellaneous Amount # 3 (wshFAC_Brouillon.Range("B37"))
    montant = wshFAC_Brouillon.Range("B37").Value
    If montant Then
        myArray(5, 1) = "5002"
        myArray(5, 2) = "Frais divers # 3"
        myArray(5, 3) = montant
        myArray(5, 4) = ""
    End If
    
    'GST to pay (wshFAC_Brouillon.Range("B38"))
    montant = wshFAC_Brouillon.Range("B38").Value
    If montant Then
        myArray(6, 1) = "2200"
        myArray(6, 2) = "TPS à payer"
        myArray(6, 3) = montant
        myArray(6, 4) = ""
    End If
    
    'PST to pay (wshFAC_Brouillon.Range("B39"))
    montant = wshFAC_Brouillon.Range("B39").Value
    If montant Then
        myArray(7, 1) = "2201"
        myArray(7, 2) = "TVQ à payer"
        myArray(7, 3) = montant
        myArray(7, 4) = ""
    End If
    
    'Deposit applied (wshFAC_Brouillon.Range("B40"))
    montant = wshFAC_Brouillon.Range("B40").Value
    If montant Then
        myArray(8, 1) = "1230"
        myArray(8, 2) = "Avance - Prêt GCP"
        myArray(8, 3) = montant
        myArray(8, 4) = ""
    End If
    
    Call FAC_Finale_GL_Posting_To_DB(dateFact, descGL_Trans, source, myArray)
    Call FAC_Finale_GL_Posting_Locally(dateFact, descGL_Trans, source, myArray)
    
    Call Output_Timer_Results("FAC_Finale_GL_Posting_Preparation()", timerStart)

End Sub

Sub FAC_Finale_GL_Posting_To_DB(df, desc, source, arr As Variant)

    Dim timerStart As Double: timerStart = Timer

    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").Value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "GL_Trans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & destinationTab & "$]"

    'Open recordset to find out the next JE number
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim maxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").Value
    End If
    
    'Calculate the new JE number
    Dim nextJENo As Long
    nextJENo = lastJE + 1
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B41").Value = nextJENo '2024-03-13 @ 08:31
    Application.EnableEvents = True

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim i As Integer, j As Integer
    'Loop through the array and post each row
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = "" Then GoTo Nothing_to_Post
            rs.AddNew
                rs.Fields("No_Entrée") = nextJENo
                rs.Fields("Date") = CDate(df)
                rs.Fields("Description") = desc
                rs.Fields("Source") = source
                rs.Fields("No_Compte") = arr(i, 1)
                rs.Fields("Compte") = arr(i, 2)
                If arr(i, 3) > 0 Then
                    rs.Fields("Débit") = arr(i, 3)
                Else
                    rs.Fields("Crédit") = -arr(i, 3)
                End If
                rs.Fields("AutreRemarque") = arr(i, 4)
                rs.Fields("TimeStamp") = Format(Now(), "dd-mm-yyyy hh:mm:ss")
            rs.update
Nothing_to_Post:
    Next i

    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("FAC_Finale_GL_Posting_To_DB()", timerStart)

End Sub

Sub FAC_Finale_GL_Posting_Locally(df, desc, source, arr As Variant) 'Write records locally
    
    Dim timerStart As Double: timerStart = Timer
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshFAC_Brouillon.Range("B41").Value
    
    'What is the last used row in GL_Trans ?
    Dim rowToBeUsed As Long
    rowToBeUsed = wshGL_Trans.Range("A99999").End(xlUp).row + 1
    
    Dim i As Integer, j As Integer
    'Loop through the array and post each row
    With wshGL_Trans
        For i = LBound(arr, 1) To UBound(arr, 1)
            If arr(i, 1) <> "" Then
                .Range("A" & rowToBeUsed).Value = JENo
                .Range("B" & rowToBeUsed).Value = CDate(df)
                .Range("C" & rowToBeUsed).Value = desc
                .Range("D" & rowToBeUsed).Value = source
                .Range("E" & rowToBeUsed).Value = arr(i, 1)
                .Range("F" & rowToBeUsed).Value = arr(i, 2)
                If arr(i, 3) > 0 Then
                     .Range("G" & rowToBeUsed).Value = arr(i, 3)
                Else
                     .Range("H" & rowToBeUsed).Value = -arr(i, 3)
                End If
                .Range("I" & rowToBeUsed).Value = arr(i, 4)
                .Range("J" & rowToBeUsed).Value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
                rowToBeUsed = rowToBeUsed + 1
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("FAC_Finale_GL_Posting_Locally()", timerStart)

End Sub

Sub Back_To_FAC_Menu()

    wshMenuFACT.Activate
    Call SlideIn_PrepFact
    Call SlideIn_SuiviCC
    Call SlideIn_Encaissement
    wshMenuFACT.Range("A1").Select
    
End Sub

Sub FAC_Finale_Enable_Save_Button()

    Dim shp As Shape
    Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = True

End Sub

Sub FAC_Finale_Disable_Save_Button()

    Dim shp As Shape
    Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = False

End Sub

Sub test_AF_TEC() '2024-03-13 @ 11:58

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim data As Range
    Dim cRng As Range
    Dim dRng As Range
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TEC_Local")
    Set data = ws.Range("A2:P361")
    Set cRng = ws.Range("R7:V8")
    Set dRng = ws.Range("Y2:AN2")
    
    'Clear prior results (if any)
    Dim lastResultRow As Long
    lastResultRow = ws.Range("Y999").End(xlUp).row
    If lastResultRow > 2 Then ws.Range("Y3:AN" & lastResultRow).clear
    
    data.AdvancedFilter xlFilterCopy, cRng, dRng

End Sub

Sub FAC_Brouillon_Input_Misc_Description(rowBrouillon As Long, rowFinale As Long)

    Dim service As String
    service = Application.InputBox("", Title:="Description / Commentaire", Type:=2)
    wshFAC_Brouillon.Range("L" & rowBrouillon).Value = "'- " & service
    wshFAC_Finale.Range("B" & rowFinale).Value = "'- " & service
    
End Sub

Sub FAC_Brouillon_TEC_Add_Check_Box(row As Long)

    Dim timerStart As Double: timerStart = Timer
    
    Dim chkBoxRange As Range
    Set chkBoxRange = wshFAC_Brouillon.Range("C8:C" & row + 5)
    
    Application.EnableEvents = False
    
    Dim cell As Range
    Dim cbx As CheckBox
        For Each cell In chkBoxRange
        ' Check if the cell is empty and doesn't have a checkbox already
        If Cells(cell.row, 8).Value = False Then 'IsInvoiced = False
            'Create a checkbox linked to the cell
            Set cbx = wshFAC_Brouillon.CheckBoxes.add(cell.Left + 5, cell.Top, cell.width, cell.Height)
            With cbx
                .name = "chkBox - " & cell.row
                .Value = True
                .text = ""
                .LinkedCell = cell.Address
                .Display3DShading = True
            End With
        End If
    Next cell

    With wshFAC_Brouillon
        .Range("D8:D" & row + 5).NumberFormat = "dd/mm/yyyy"
        .Range("D8:D" & row + 5).Font.Bold = False
        
        .Range("D" & row + 7).formula = "=SUMIF(C8:C" & row + 5 & ",True,G8:G" & row + 5 & ")"
        .Range("D" & row + 7).NumberFormat = "##0.00"
        .Range("D" & row + 7).Font.Bold = True
        
        .Range("B19").formula = "=SUMIF(C8:C" & row + 5 & ",True,G8:G" & row + 5 & ")"
    End With
    
    Application.EnableEvents = True

    Call Output_Timer_Results("FAC_Brouillon_TEC_Add_Check_Box()", timerStart)

End Sub

Sub FAC_Brouillon_TEC_Remove_Check_Box(row As Long)

    Dim timerStart As Double: timerStart = Timer
    
    Application.EnableEvents = False
    
    Dim cbx As Shape
    For Each cbx In wshFAC_Brouillon.Shapes
        If InStr(cbx.name, "chkBox - ") Then
            cbx.delete
        End If
    Next cbx
    wshFAC_Brouillon.Range("C7:C" & row).Value = "" 'Remove text left over
    wshFAC_Brouillon.Range("D" & row + 2).Value = "" 'Remove the total formula

    Application.EnableEvents = True

    Call Output_Timer_Results("FAC_Brouillon_TEC_Remove_Check_Box()", timerStart)

End Sub


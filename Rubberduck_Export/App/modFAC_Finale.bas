Attribute VB_Name = "modFAC_Finale"
Option Explicit

Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub FAC_Finale_Save() '2024-03-28 @ 07:19

    Dim timerStart As Double: timerStart = Timer

    With wshFAC_Brouillon
        'Check For Mandatory Fields - Client
        If .Range("B18").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            GoTo Fast_Exit_Sub
        End If
        'Check For Mandatory Fields - Date de facture
        If .Range("O3").value = Empty Or Len(Trim(.Range("O6").value)) <> 8 Then
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
            Call FAC_Finale_TEC_Update_As_Billed_To_DB(3, lastResultRow)
            Call FAC_Finale_TEC_Update_As_Billed_Locally(3, lastResultRow)
        End If
        
        'Save Invoice total amount
        Dim invoice_Total As Currency
        invoice_Total = wshFAC_Brouillon.Range("O51").value
        
        Call FAC_Brouillon_Clear_All_TEC_Displayed
        
    End With
    
    'GL stuff
    Call FAC_Finale_GL_Posting_Preparation
    
    'Update TEC_DashBoard
    Call TEC_DB_Update_All '2024-03-21 @ 12:32

    MsgBox "La facture '" & wshFAC_Brouillon.Range("O6").value & "' est enregistr�e." & _
        vbNewLine & vbNewLine & "Le total de la facture est " & _
        Trim(Format(invoice_Total, "### ##0.00 $")) & _
        " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    
    wshFAC_Brouillon.Range("B27").value = False
    Call FAC_Brouillon_New_Invoice '2024-03-12 @ 08:08 - Maybe ??
    
Fast_Exit_Sub:

    Call Output_Timer_Results("FAC_Finale_Save()", timerStart)
    
    wshFAC_Brouillon.Select
'    Call FAC_Finale_Goto_Onglet_FAC_Brouillon
    
End Sub

Sub FAC_Finale_Add_Invoice_Header_to_DB()

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_Ent�te"
    
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
        rs.Fields("Inv_No") = .Range("E28").value
        rs.Fields("Date_Facture") = CDate(wshFAC_Brouillon.Range("O3").value)
        rs.Fields("Cust_ID") = wshFAC_Brouillon.Range("B18").value
        rs.Fields("Contact") = .Range("B23").value
        rs.Fields("Nom_Client") = .Range("B24").value
        rs.Fields("Adresse1") = .Range("B25").value
        rs.Fields("Adresse2") = .Range("B26").value
        rs.Fields("Adresse3") = .Range("B27").value
        
        rs.Fields("Honoraires") = .Range("E69").value
        
        rs.Fields("AF1_Desc") = .Range("B70").value
        rs.Fields("AutresFrais_1") = wshFAC_Finale.Range("E70").value
        rs.Fields("AF2_Desc") = .Range("B71").value
        rs.Fields("AutresFrais_2") = .Range("E71").value
        rs.Fields("AF3_Desc") = .Range("B72").value
        rs.Fields("AutresFrais_3") = .Range("E72").value
        
        rs.Fields("Taux_TPS") = .Range("C74").value
        rs.Fields("Mnt_TPS") = .Range("E74").value
        rs.Fields("Taux_TVQ") = .Range("C75").value
        rs.Fields("Mnt_TVQ") = .Range("E75").value
        
        rs.Fields("AR_Total") = .Range("E77").value
        
        rs.Fields("D�p�t") = .Range("E79").value
    End With
    'Update the recordset (create the record)
    rs.update
    
    Application.EnableEvents = False
    
    'Prepare GL Posting
    With wshFAC_Brouillon
        .Range("B33").value = wshFAC_Finale.Range("E81").value   'AR amount
        .Range("B34").value = -wshFAC_Finale.Range("E69").value  'Revenues
        .Range("B35").value = -wshFAC_Finale.Range("E70").value  'Misc $ - 1
        .Range("B36").value = -wshFAC_Finale.Range("E71").value  'Misc $ - 2
        .Range("B37").value = -wshFAC_Finale.Range("E72").value  'Misc $ - 3
        .Range("B38").value = -wshFAC_Finale.Range("E74").value  'GST $
        .Range("B39").value = -wshFAC_Finale.Range("E75").value  'PST $
        .Range("B40").value = wshFAC_Finale.Range("E79").value   'Deposit
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
    firstFreeRow = wshFAC_Ent�te.Range("A9999").End(xlUp).row + 1
    
    With wshFAC_Ent�te
        .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).value = Format(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy")
        .Range("C" & firstFreeRow).value = wshFAC_Brouillon.Range("B18").value
        .Range("D" & firstFreeRow).value = wshFAC_Finale.Range("B23").value
        .Range("E" & firstFreeRow).value = wshFAC_Finale.Range("B24").value
        .Range("F" & firstFreeRow).value = wshFAC_Finale.Range("B25").value
        .Range("G" & firstFreeRow).value = wshFAC_Finale.Range("B26").value
        .Range("H" & firstFreeRow).value = wshFAC_Finale.Range("B27").value
        
        .Range("I" & firstFreeRow).value = wshFAC_Finale.Range("E69").value
        
        .Range("J" & firstFreeRow).value = wshFAC_Finale.Range("B70").value
        .Range("K" & firstFreeRow).value = wshFAC_Finale.Range("E70").value
        .Range("L" & firstFreeRow).value = wshFAC_Finale.Range("B71").value
        .Range("M" & firstFreeRow).value = wshFAC_Finale.Range("E71").value
        .Range("N" & firstFreeRow).value = wshFAC_Finale.Range("B72").value
        .Range("O" & firstFreeRow).value = wshFAC_Finale.Range("E72").value
        
        .Range("P" & firstFreeRow).value = wshFAC_Finale.Range("C74").value
        .Range("Q" & firstFreeRow).value = wshFAC_Finale.Range("E74").value
        .Range("R" & firstFreeRow).value = wshFAC_Finale.Range("C75").value
        .Range("S" & firstFreeRow).value = wshFAC_Finale.Range("E75").value
        
        .Range("T" & firstFreeRow).value = wshFAC_Finale.Range("E77").value
        
        .Range("U" & firstFreeRow).value = wshFAC_Finale.Range("E79").value
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
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                          "GCF_BD_Sortie.xlsx"
    destinationTab = "FAC_D�tails"
    
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
            rs.Fields("Inv_No") = .Range("E28").value
            rs.Fields("Description") = .Range("B" & r).value
            rs.Fields("Heures") = .Range("C" & r).value
            rs.Fields("Taux") = .Range("D" & r).value
            If .Range("E" & r).value <> "" Then
                rs.Fields("Honoraires") = .Range("E" & r).value
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
    firstFreeRow = wshFAC_D�tails.Range("A99999").End(xlUp).row + 1
   
    Dim i As Integer
    For i = 34 To lastEnteredService
        With wshFAC_D�tails
            .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
            .Range("B" & firstFreeRow).value = wshFAC_Finale.Range("B" & i).value
            .Range("C" & firstFreeRow).value = wshFAC_Finale.Range("C" & i).value
            .Range("D" & firstFreeRow).value = wshFAC_Finale.Range("D" & i).value
            .Range("E" & firstFreeRow).value = wshFAC_Finale.Range("E" & i).value
            .Range("F" & firstFreeRow).value = i
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
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
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
        rs.Fields("Invoice_No") = .Range("E28").value
        rs.Fields("Invoice_Date") = CDate(wshFAC_Brouillon.Range("O3").value)
        rs.Fields("Customer") = .Range("B24").value
        rs.Fields("Status") = "Unpaid"
        rs.Fields("Terms") = "Net 30"
        rs.Fields("Due_Date") = CDate(CDate(wshFAC_Brouillon.Range("O3").value) + 30)
        rs.Fields("Total") = .Range("E81").value
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
        .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).value = wshFAC_Brouillon.Range("O3").value
        .Range("C" & firstFreeRow).value = wshFAC_Finale.Range("B24").value
        .Range("D" & firstFreeRow).value = "Unpaid"
        .Range("E" & firstFreeRow).value = "Net 30"
        .Range("F" & firstFreeRow).value = CDate(CDate(wshFAC_Brouillon.Range("O3").value) + 30)
        .Range("G" & firstFreeRow).value = wshFAC_Finale.Range("E81").value
        .Range("H" & firstFreeRow).formula = ""
        .Range("I" & firstFreeRow).formula = "=G" & firstFreeRow & "-H" & firstFreeRow
        .Range("J" & firstFreeRow).formula = "=IF(H" & firstFreeRow & "<G" & firstFreeRow & ",NOW()-F" & firstFreeRow & ")"
    End With

nothing_to_update:

    Call Output_Timer_Results("FAC_Finale_Add_Comptes_Clients_Locally()", timerStart)

    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_To_DB(firstRow As Integer, lastRow As Integer) 'Update Billed Status in DB

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
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
        If wshTEC_Local.Range("BD" & r).value = True Or _
            wshFAC_Brouillon.Range("C" & r + 5) <> True Then
            GoTo next_iteration
        End If
        TEC_ID = wshTEC_Local.Range("AT" & r).value
        
        'Open the recordset for the specified ID
        SQL = "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & TEC_ID
        rs.Open SQL, conn, 2, 3
        If Not rs.EOF Then
            'Update DateSaisie, EstFacturee, DateFacturee & NoFacture
            rs.Fields("DateSaisie").value = Now
            rs.Fields("EstFacturee").value = True
            rs.Fields("DateFacturee").value = Now
            rs.Fields("VersionApp").value = gAppVersion
            rs.Fields("NoFacture").value = wshFAC_Brouillon.Range("O6").value
            rs.update
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement avec le TEC_ID '" & r & "' ne peut �tre trouv�!", _
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

    Call Output_Timer_Results("FAC_Finale_TEC_Update_As_Billed_To_DB()", timerStart)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_Locally(firstResultRow As Integer, lastResultRow As Integer)

    Dim timerStart As Double: timerStart = Timer
    
    'Set the range to look for
    Dim lookupRange As Range, lastTECRow As Long
    lastTECRow = wshTEC_Local.Range("A99999").End(xlUp).row
    Set lookupRange = wshTEC_Local.Range("A3:A" & lastTECRow)
    
    Dim r As Integer, rowToBeUpdated As Long, tecID As Long
    For r = firstResultRow To lastResultRow
        If wshTEC_Local.Range("BD" & r).value = False And _
            wshFAC_Brouillon.Range("C" & r + 5) = True Then
            tecID = wshTEC_Local.Range("AT" & r).value
            rowToBeUpdated = Fn_Get_TEC_Row_Number_By_TEC_ID(tecID, lookupRange)
            wshTEC_Local.Range("K" & rowToBeUpdated).value = Now()
            wshTEC_Local.Range("L" & rowToBeUpdated).value = True
            wshTEC_Local.Range("M" & rowToBeUpdated).value = Now()
            wshTEC_Local.Range("O" & rowToBeUpdated).value = gAppVersion
            wshTEC_Local.Range("P" & rowToBeUpdated).value = wshFAC_Brouillon.Range("O6").value
        End If
    Next r
    
    Call Output_Timer_Results("FAC_Finale_TEC_Update_As_Billed_Locally()", timerStart)

End Sub

Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16
    With wshFAC_Brouillon
        If wshFAC_Brouillon.Range("B20").value = Empty Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un num�ro de facture VALIDE pour votre recherche"
            GoTo NoItems
        End If
        'Could that invoice been cancelled (more than 1 row) ?
        Call InvoiceGetAllTrans(wshFAC_Brouillon.Range("O6").value)
        Dim NbTrans As Integer
        NbTrans = .Range("B31").value
        If NbTrans = 0 Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un num�ro de facture VALIDE pour votre recherche"
            GoTo NoItems
        Else
            If NbTrans > 1 Then
                MsgBox "Cette facture a �t� annul�e! Veuillez saisir un num�ro de facture VALIDE pour votre recherche"
                GoTo NoItems
            End If
        End If
        .Range("B24").value = True 'Set Invoice Load to true
        .Range("S2,E4:F4,K4:L6,O3,K11:O45,Q11:Q45").Clearcontents
        wshFAC_Finale.Range("C34:F69").Clearcontents
        Dim InvListRow As Long
        InvListRow = wshFAC_Brouillon.Range("B20").value 'InvListRow = Row associated with the invoice
        'Get values from wshFAC_Ent�te (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
        .Range("O3").value = wshFAC_Ent�te.Range("B" & InvListRow).value
        .Range("K3").value = wshFAC_Ent�te.Range("D" & InvListRow).value
        .Range("K4").value = wshFAC_Ent�te.Range("E" & InvListRow).value
        .Range("K5").value = wshFAC_Ent�te.Range("F" & InvListRow).value
        .Range("K6").value = wshFAC_Ent�te.Range("G" & InvListRow).value
        'Get values from wshFAC_Ent�te (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
        wshFAC_Finale.Range("B21").value = "Le " & Format(wshFAC_Ent�te.Range("B" & InvListRow).value, "d mmmm yyyy")
        wshFAC_Finale.Range("B23").value = wshFAC_Ent�te.Range("D" & InvListRow).value
        wshFAC_Finale.Range("B24").value = wshFAC_Ent�te.Range("E" & InvListRow).value
        wshFAC_Finale.Range("B25").value = wshFAC_Ent�te.Range("F" & InvListRow).value
        wshFAC_Finale.Range("B26").value = wshFAC_Ent�te.Range("G" & InvListRow).value
        'Load Invoice Detail Items
        With wshFAC_D�tails
            Dim lastRow As Long, lastResultRow As Long
            lastRow = .Range("A999999").End(xlUp).row
            If lastRow < 4 Then Exit Sub 'No Item Lines
            .Range("I3").value = wshFAC_Brouillon.Range("O6").value
            wshFAC_Finale.Range("F28").value = wshFAC_Brouillon.Range("O6").value 'Invoice #
            'Advanced Filter to get items specific to ONE invoice
            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
            lastResultRow = .Range("O999").End(xlUp).row
            If lastResultRow < 3 Then GoTo NoItems
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("O" & resultRow).value
                wshFAC_Brouillon.Range("L" & invitemRow & ":O" & invitemRow).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
                wshFAC_Brouillon.Range("Q" & invitemRow).value = .Range("P" & resultRow).value  'Set Item DB Row
                wshFAC_Finale.Range("C" & invitemRow + 23 & ":F" & invitemRow + 23).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
            Next resultRow
        End With
        'Proceed with trailer data (Misc. charges & Taxes)
        .Range("M48").value = wshFAC_Ent�te.Range("I" & InvListRow).value
        .Range("O48").value = wshFAC_Ent�te.Range("J" & InvListRow).value
        .Range("M49").value = wshFAC_Ent�te.Range("K" & InvListRow).value
        .Range("O49").value = wshFAC_Ent�te.Range("L" & InvListRow).value
        .Range("M50").value = wshFAC_Ent�te.Range("M" & InvListRow).value
        .Range("O50").value = wshFAC_Ent�te.Range("N" & InvListRow).value
        .Range("O52").value = wshFAC_Ent�te.Range("P" & InvListRow).value
        .Range("O53").value = wshFAC_Ent�te.Range("R" & InvListRow).value
        .Range("O57").value = wshFAC_Ent�te.Range("T" & InvListRow).value

NoItems:
    .Range("B24").value = False 'Set Invoice Load To false
    End With
End Sub

Sub InvoiceGetAllTrans(inv As String)

    Application.ScreenUpdating = False

    wshFAC_Brouillon.Range("B31").value = 0

    With wshFAC_Ent�te
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last wshFAC_Ent�te Row
        If lastRow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
        On Error Resume Next
        .Names("Criterial").delete
        On Error GoTo 0
        .Range("V3").value = wshFAC_Brouillon.Range("O6").value
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
'            .SortFields.Add Key:=wshFAC_Ent�te.Range("X2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based Invoice Number
'            .SortFields.Add Key:=wshGL_Trans.Range("Y3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based On TEC_ID
'            .SetRange wshFAC_Ent�te.Range("X2:AQ" & lastResultRow) 'Set Range
'            .Apply 'Apply Sort
'         End With
         wshFAC_Brouillon.Range("B31").value = lastResultRow - 2 'Remove Header rows from row count
Done:
    End With
    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_Setup_All_Cells()

    Dim timerStart As Double: timerStart = Timer
    
    Application.EnableEvents = False
    
    With wshFAC_Finale
        .Range("B21").formula = "= ""Le "" & TEXT(FAC_Brouillon!O3, ""j MMMM aaaa"")"
        .Range("B23:B27").value = ""
        .Range("E28").value = "=" & wshFAC_Brouillon.name & "!O6"    'Invoice number
        
        .Range("C65").value = "Heures"                               'Summary Heading
        .Range("D65").value = "Taux"                                 'Summary Heading
        .Range("C66").formula = "=" & wshFAC_Brouillon.name & "!M47" 'Hours summary
        .Range("D66").formula = "=" & wshFAC_Brouillon.name & "!N47" 'Hourly Rate
        
        With .Range("C65:D66")
            .Font.ThemeColor = xlThemeColorLight1
            .Font.TintAndShade = 0
        End With

        Call FAC_Brouillon_Set_Labels(.Range("B69"), "FAC_Label_SubTotal_1")
        Call FAC_Brouillon_Set_Labels(.Range("B73"), "FAC_Label_SubTotal_2")
        Call FAC_Brouillon_Set_Labels(.Range("B74"), "FAC_Label_TPS")
        Call FAC_Brouillon_Set_Labels(.Range("B75"), "FAC_Label_TVQ")
        Call FAC_Brouillon_Set_Labels(.Range("B77"), "FAC_Label_GrandTotal")
        Call FAC_Brouillon_Set_Labels(.Range("B79"), "FAC_Label_Deposit")
        Call FAC_Brouillon_Set_Labels(.Range("B81"), "FAC_Label_AmountDue")

        .Range("E69").formula = "=C66*D66"                           'Fees Sub-Total
        
        .Range("B70").value = "='" & wshFAC_Brouillon.name & "'!M48" 'Misc. Amount # 1 - Description
        .Range("E70").value = "='" & wshFAC_Brouillon.name & "'!O48" 'Misc. Amount # 1
        
        .Range("B71").value = "='" & wshFAC_Brouillon.name & "'!M49" 'Misc. Amount # 2 - Description
        .Range("E71").value = "='" & wshFAC_Brouillon.name & "'!O49" 'Misc. Amount # 2
        
        .Range("B72").value = "='" & wshFAC_Brouillon.name & "'!M50" 'Misc. Amount # 3 - Description
        .Range("E72").value = "='" & wshFAC_Brouillon.name & "'!O50" 'Misc. Amount # 3
        
        .Range("E73").formula = "=SUM(E69:E72)"                      'Invoice Sub-Total
        
        .Range("C74").value = "='" & wshFAC_Brouillon.name & "'!N52" 'GST Rate
        .Range("E74").formula = "=round(E73*C74,2)"                  'GST Amount"
        .Range("C75").value = "='" & wshFAC_Brouillon.name & "'!N53" 'PST Rate
        .Range("E75").formula = "=round(E73*C75,2)"                  'PST Amount
        
        .Range("E77").value = "=SUM(E73:E75)"                        'Total including taxes
        .Range("E79").value = "='" & wshFAC_Brouillon.name & "'!O57" 'Deposit Amount
        .Range("E81").value = "=E77-E79"                             'Total due on that invoice
    End With
    
    Application.EnableEvents = True
    
    Call Output_Timer_Results("FAC_Finale_Setup_All_Cells()", timerStart)

End Sub

Sub FAC_Finale_Preview_PDF() '2024-03-02 @ 16:18

    wshFAC_Finale.PrintOut , , 1, True, True, , , , False
'    wshFAC_Finale.PrintOut , , , True, True, , , , False
    
End Sub

Sub FAC_Finale_Creation_PDF_And_Email() 'RMV - 2023-12-17 @ 14:35
    
    Call FAC_Finale_Create_PDF_Email_Sub(wshFAC_Finale.Range("E28").value)
    
    Call FAC_Finale_Enable_Save_Button

End Sub

Sub FAC_Finale_Create_PDF_Email_Sub(noFacture As String)

    'Cr�ation du fichier (NoFacture).PDF dans le r�pertoire de factures PDF de GCF et pr�paration du courriel pour envoyer la facture
    Dim result As Boolean
    result = FAC_Finale_Create_PDF_Email_Func(noFacture, "CreateEmail")

End Sub

Function FAC_Finale_Create_PDF_Email_Func(noFacture As String, Optional action As String = "SaveOnly") As Boolean
    
    Dim SaveAs As String

    Application.ScreenUpdating = False

    'Construct the SaveAs filename
    SaveAs = wshAdmin.Range("FolderPDFInvoice").value & Application.PathSeparator & _
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
            FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\GCF_Facturation.oft"
        Else
            FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\GCF_Facturation.oft"
        End If

        Dim myMail As Outlook.MailItem
        Set myMail = outlookApp.CreateItemFromTemplate(FullTemplatePathAndFile)
'        Set myMail = outlookApp.CreateItem(olMailItem)

        Dim source_file As String
        source_file = wshAdmin.Range("FolderPDFInvoice").value & Application.PathSeparator & _
                      noFacture & ".pdf" '2023-12-19 @ 07:22
        
        With myMail
            .To = "robertv13@hotmail.com"
            .CC = "robertv13@me.com"
            .BCC = "robertv13@gmail.com"
            .Subject = "TEST - GC FISCALIT� INC. - Facturation - TEST"
'            .Body = "Bonjour," & vbNewLine & vbNewLine & "Vous trouverez ci-joint notre note d'honoraires." & _
                vbNewLine & vbNewLine & "Merci" & vbNewLine & vbNewLine & vbNewLine & "Guillaume Charron, CPA, CA, M. Fisc." & _
                vbNewLine & "Pr�sident"
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
    MsgBox "Incapable de pr�parer le courriel. La librairie n'est pas disponible"
    FAC_Finale_Create_PDF_Email_Func = False 'Return value

EndMacro:
    Application.ScreenUpdating = True
    
End Function

Sub Prev_Invoice() 'TO-DO-RMV 2023-12-17
    With wshFAC_Brouillon
        Dim MininvNumb As Long
        On Error Resume Next
        MininvNumb = Application.WorksheetFunction.Min(wshFAC_Ent�te.Range("Inv_ID"))
        On Error GoTo 0
        If MininvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").value
        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            invRow = wshFAC_Ent�te.Range("A99999").End(xlUp).row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFAC_Ent�te.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).row - 1
        End If
        If .Range("N6").value = 1 Or MininvNumb = 0 Or MininvNumb = .Range("N6").value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFAC_Ent�te.Range("A" & invRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
End Sub

Sub Next_Invoice() 'TO-DO-RMV 2023-12-17
    With wshFAC_Brouillon
        Dim MaxinvNumb As Long
        On Error Resume Next
        MaxinvNumb = Application.WorksheetFunction.Max(wshFAC_Ent�te.Range("Inv_ID"))
        On Error GoTo 0
        If MaxinvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").value
        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            invRow = wshFAC_Ent�te.Range("A4").value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFAC_Ent�te.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).row + 1
        End If
        If .Range("N6").value >= MaxinvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFAC_Ent�te.Range("A" & invRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
End Sub

Sub FAC_Finale_Cacher_Heures()
    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
End Sub

Sub FAC_Finale_Montrer_Heures()
    With wshFAC_Finale.Range("C34:E63")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
End Sub

Sub FAC_Finale_Cacher_Sommaire_Heures()
    With wshFAC_Finale.Range("C65:D66")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
End Sub

Sub FAC_Finale_Montrer_Sommaire_Heures()
    With wshFAC_Finale.Range("C65:D66")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
End Sub

Sub FAC_Finale_Goto_Onglet_FAC_Brouillon()

    Dim timerStart As Double: timerStart = Timer
   
    Application.ScreenUpdating = False
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Brouillon.Range("E4").Select

    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("FAC_Finale_Goto_Onglet_FAC_Brouillon()", timerStart)

End Sub

Sub FAC_Finale_GL_Posting_Preparation() '2024-02-14 @ 05:56

    Dim timerStart As Double: timerStart = Timer

    Dim montant As Double
    Dim dateFact As Date
    Dim descGL_Trans As String, source As String
    
    dateFact = wshFAC_Brouillon.Range("O3").value
    descGL_Trans = wshFAC_Brouillon.Range("E4").value
    source = "FACT-" & wshFAC_Brouillon.Range("O6").value
    
    Dim myArray(1 To 8, 1 To 4) As String
    
    'AR amount (wshFAC_Brouillon.Range("B33"))
    montant = wshFAC_Brouillon.Range("B33").value
    If montant Then
        myArray(1, 1) = "1100"
        myArray(1, 2) = "Comptes Clients"
        myArray(1, 3) = montant
        myArray(1, 4) = ""
    End If
    
    'Professional Fees (wshFAC_Brouillon.Range("B34"))
    montant = wshFAC_Brouillon.Range("B34").value
    If montant Then
        myArray(2, 1) = "4000"
        myArray(2, 2) = "Revenus"
        myArray(2, 3) = montant
        myArray(2, 4) = ""
    End If
    
    'Miscellaneous Amount # 1 (wshFAC_Brouillon.Range("B35"))
    montant = wshFAC_Brouillon.Range("B35").value
    If montant Then
        myArray(3, 1) = "5009"
        myArray(3, 2) = "Frais divers # 1"
        myArray(3, 3) = montant
        myArray(3, 4) = ""
    End If
    
    'Miscellaneous Amount # 2 (wshFAC_Brouillon.Range("B36"))
    montant = wshFAC_Brouillon.Range("B36").value
    If montant Then
        myArray(4, 1) = "5008"
        myArray(4, 2) = "Frais divers # 2"
        myArray(4, 3) = montant
        myArray(4, 4) = ""
    End If
    
    'Miscellaneous Amount # 3 (wshFAC_Brouillon.Range("B37"))
    montant = wshFAC_Brouillon.Range("B37").value
    If montant Then
        myArray(5, 1) = "5002"
        myArray(5, 2) = "Frais divers # 3"
        myArray(5, 3) = montant
        myArray(5, 4) = ""
    End If
    
    'GST to pay (wshFAC_Brouillon.Range("B38"))
    montant = wshFAC_Brouillon.Range("B38").value
    If montant Then
        myArray(6, 1) = "2200"
        myArray(6, 2) = "TPS � payer"
        myArray(6, 3) = montant
        myArray(6, 4) = ""
    End If
    
    'PST to pay (wshFAC_Brouillon.Range("B39"))
    montant = wshFAC_Brouillon.Range("B39").value
    If montant Then
        myArray(7, 1) = "2201"
        myArray(7, 2) = "TVQ � payer"
        myArray(7, 3) = montant
        myArray(7, 4) = ""
    End If
    
    'Deposit applied (wshFAC_Brouillon.Range("B40"))
    montant = wshFAC_Brouillon.Range("B40").value
    If montant Then
        myArray(8, 1) = "1230"
        myArray(8, 2) = "Avance - Pr�t GCP"
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
    destinationFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
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
    strSQL = "SELECT MAX(No_Entr�e) AS MaxEJNo FROM [" & destinationTab & "$]"

    'Open recordset to find out the next JE number
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim maxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new JE number
    Dim nextJENo As Long
    nextJENo = lastJE + 1
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B41").value = nextJENo '2024-03-13 @ 08:31
    Application.EnableEvents = True

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim i As Integer, j As Integer
    'Loop through the array and post each row
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = "" Then GoTo Nothing_to_Post
            rs.AddNew
                rs.Fields("No_Entr�e") = nextJENo
                rs.Fields("Date") = CDate(df)
                rs.Fields("Description") = desc
                rs.Fields("Source") = source
                rs.Fields("No_Compte") = arr(i, 1)
                rs.Fields("Compte") = arr(i, 2)
                If arr(i, 3) > 0 Then
                    rs.Fields("D�bit") = arr(i, 3)
                Else
                    rs.Fields("Cr�dit") = -arr(i, 3)
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
    JENo = wshFAC_Brouillon.Range("B41").value
    
    'What is the last used row in GL_Trans ?
    Dim rowToBeUsed As Long
    rowToBeUsed = wshGL_Trans.Range("A99999").End(xlUp).row + 1
    
    Dim i As Integer, j As Integer
    'Loop through the array and post each row
    With wshGL_Trans
        For i = LBound(arr, 1) To UBound(arr, 1)
            If arr(i, 1) <> "" Then
                .Range("A" & rowToBeUsed).value = JENo
                .Range("B" & rowToBeUsed).value = CDate(df)
                .Range("C" & rowToBeUsed).value = desc
                .Range("D" & rowToBeUsed).value = source
                .Range("E" & rowToBeUsed).value = arr(i, 1)
                .Range("F" & rowToBeUsed).value = arr(i, 2)
                If arr(i, 3) > 0 Then
                     .Range("G" & rowToBeUsed).value = arr(i, 3)
                Else
                     .Range("H" & rowToBeUsed).value = -arr(i, 3)
                End If
                .Range("I" & rowToBeUsed).value = arr(i, 4)
                .Range("J" & rowToBeUsed).value = Format(Now(), "dd-mm-yyyy hh:mm:ss")
                rowToBeUsed = rowToBeUsed + 1
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    Call Output_Timer_Results("FAC_Finale_GL_Posting_Locally()", timerStart)

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

'TO BE DELETED EVENTUALLY...

'Sub FAC_Brouillon_Input_Misc_Description(rowBrouillon As Long, rowFinale As Long) - 2024-03-28 @ 14:16
'
'    Dim service As String
'    service = Application.InputBox("", Title:="Description / Commentaire", Type:=2)
'    wshFAC_Brouillon.Range("L" & rowBrouillon).value = "'- " & service
'    wshFAC_Finale.Range("B" & rowFinale).value = "'- " & service
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------

'Sub ExportAllFacInvList() '2024-03-28 @ 14:22
'    Dim wb As Workbook
'    Dim wsSource As Worksheet
'    Dim wsTarget As Worksheet
'    Dim sourceRange As Range
'
'    Application.ScreenUpdating = False
'
'    'Work with the source range
'    Set wsSource = wshFAC_Ent�te
'    Dim lastUsedRow As Long
'    lastUsedRow = wsSource.Range("A99999").End(xlUp).row
'    wsSource.Range("A4:T" & lastUsedRow).Copy
'
'    'Open the target workbook
'    Workbooks.Open fileName:=wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
'                   "GCF_BD_Sortie.xlsx"
'
'    'Set references to the target workbook and target worksheet
'    Set wb = Workbooks("GCF_BD_Sortie.xlsx")
'    Set wsTarget = wb.Sheets("FACTURES")
'
'    'PasteSpecial directly to the target range
'    wsTarget.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'    Application.CutCopyMode = False
'
'    wb.Close SaveChanges:=True
'
'    Application.ScreenUpdating = True
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------

'Sub FAC_Brouillon_Prev_PDF() '2024-03-28 @ 14:49
'
'    Call FAC_Brouillon_Goto_Onglet_FAC_Finale
'    Call FAC_Finale_Preview_PDF
'    Call FAC_Finale_Goto_Onglet_FAC_Brouillon
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------

'Sub test_AF_TEC() '2024-03-28 @ 14:55
'
'    Dim wb As Workbook
'    Dim ws As Worksheet
'    Dim data As Range
'    Dim cRng As Range
'    Dim dRng As Range
'
'    Set wb = ThisWorkbook
'    Set ws = wb.Worksheets("TEC_Local")
'    Set data = ws.Range("A2:P361")
'    Set cRng = ws.Range("R7:V8")
'    Set dRng = ws.Range("Y2:AN2")
'
'    'Clear prior results (if any)
'    Dim lastResultRow As Long
'    lastResultRow = ws.Range("Y999").End(xlUp).row
'    If lastResultRow > 2 Then ws.Range("Y3:AN" & lastResultRow).clear
'
'    data.AdvancedFilter xlFilterCopy, cRng, dRng
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------


Attribute VB_Name = "modFAC"
Option Explicit
Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub FAC_Prep_Invoice_New() 'Clear contents
    
    Dim timerStart As Double: timerStart = Timer
    
    If wshFACPrep.Range("B27").value = False Then
        With wshFACPrep
            .Range("B24").value = True
            .Range("K3:L6,O3,O5").ClearContents 'Clear cells for a new Invoice
            .Range("J10:Q46").ClearContents
            .Range("O6").value = .Range("FACNextInvoiceNumber").value 'Paste Invoice ID
            .Range("FACNextInvoiceNumber").value = .Range("FACNextInvoiceNumber").value + 1 'Increment Next Invoice ID
            
            Call TEC_Clear_All_Cells
            Call FAC_Prep_Clear_And_Fix_Totals_Formula
            
            .Range("B20").value = ""
            .Range("B24").value = False
            .Range("B26").value = False
            .Range("B27").value = True 'Set the value to TRUE
        End With
        With wshFACFinale
            .Range("B21,B23:C26,F28").ClearContents
            .Range("A33:G62, D65, E65").ClearContents
            .Range("F28").value = wshFACPrep.Range("O6").value 'Invoice #
            .Range("B68:F80").ClearContents
            
            Call FAC_Prep_Clear_And_Fix_Formula_Finale
        
        End With
        Call TEC_Clear_All_Cells
        
        'Move on to CLient Name
        wshFACPrep.Range("E4:F4").ClearContents
        With wshFACPrep.Range("E4:F4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        wshFACPrep.Range("E4").Select 'Start inputing values for a NEW invoice
    End If
    If wshFACPrep.Range("B28").value Then Debug.Print vbNewLine & "Le numéro de facture '" & wshFACPrep.Range("O6").value & "' a été assignée"
    
    Call Output_Timer_Results("FAC_Prep_Invoice_New()", timerStart)

End Sub

Sub FAC_Prep_Save_And_Update() '2024-02-21 @ 10:11
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub FAC_Prep_Save_And_Update() @ " & Time
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "B18 (Cust. ID) = " & wshFACPrep.Range("B18").value
    With wshFACPrep
        'Check For Mandatory Fields - Client
        If .Range("B18").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que le nom de client n'est pas encore saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Check For Mandatory Fields - Date de facture
        If .Range("O3").value = Empty Or Len(Trim(.Range("O6").value)) <> 8 Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture AVANT de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUSÉE parce que la date de facture et/ou le numéro de facture n'ont pas encore été saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        
        wshFACPrep.usedRange.Calculate
        wshFACFinale.usedRange.Calculate
        
        'Valid Invoice - Let's update it
        'Determine the row number (InvListRow) for New Invoice -OR- use existing one
        If wshFACPrep.Range("B20").value = Empty Then 'New Invoice
            Call FAC_Prep_Add_Invoice_Header_to_DB(0)
            Call FAC_Prep_Add_Invoice_Details_to_DB
            Call FAC_Prep_Add_Comptes_Clients_to_DB
            Dim lastResultRow As Integer
            lastResultRow = wshBaseHours.Range("Y9999").End(xlUp).row
            If lastResultRow > 2 Then
                Call Update_TEC_As_Billed_In_DB(3, lastResultRow)
                Call FAC_Prep_TEC_As_Billed_Locally(3, lastResultRow)
        End If
    End With
NoItems:
    wshFACPrep.Range("B27").value = False
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Total de la facture '"; wshFACPrep.Range("O6") & "' (avant taxes) est de " & Format(wshFACPrep.Range("O51").value, "### ##0.00 $")
Fast_Exit_Sub:
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFAC] - Sub FAC_Prep_Save_And_Update()" & vbNewLine
    
    Call FAC_Prepare_GL_Posting
    
    MsgBox "La facture '" & wshFACPrep.Range("O6").value & "' est enregistrée." & vbNewLine & vbNewLine & "Le total de la facture est " & Trim(Format(wshFACPrep.Range("O51").value, "### ##0.00 $")) & " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"

End Sub

Sub FAC_Prep_Add_Invoice_Header_to_DB(r As Long)

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "Invoice_Header"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'r = 0 ---> New record, r > 0 ---> Update an existing record (r)
    If r = 0 Then
    
        'Create an empty recordset
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
        
        'Add fields to the recordset before updating it
        rs.AddNew
        With wshFACPrep
            rs.Fields("InvNo") = wshFACPrep.Range("O6")
            rs.Fields("DateFacture") = .Range("O3").value
            rs.Fields("CustID") = .Range("B18").value
            rs.Fields("Contact") = .Range("K3").value
            rs.Fields("NomClient") = .Range("K4").value
            rs.Fields("Adresse") = .Range("K5").value
            rs.Fields("VilleProvCP") = .Range("K6").value
        End With
        With wshFACFinale
            rs.Fields("Honoraires") = wshFACFinale.Range("F68").value
            rs.Fields("AF1Desc") = wshFACFinale.Range("C69").value
            rs.Fields("AutresFrais1") = wshFACFinale.Range("F69").value
            rs.Fields("AF2Desc") = wshFACFinale.Range("C70").value
            rs.Fields("AutresFrais2") = wshFACFinale.Range("F70").value
            rs.Fields("AF3Desc") = wshFACFinale.Range("C71").value
            rs.Fields("AutresFrais3") = wshFACFinale.Range("F71").value
            rs.Fields("TauxTPS") = wshFACFinale.Range("D73").value
            rs.Fields("MntTPS") = wshFACFinale.Range("F73").value
            rs.Fields("TauxTVQ") = wshFACFinale.Range("D74").value
            rs.Fields("MntTVQ") = wshFACFinale.Range("F74").value
            rs.Fields("AR_Total") = wshFACFinale.Range("F76").value
            rs.Fields("Depot") = wshFACFinale.Range("F78").value
        End With
    Else 'Update an existing record
        'Open the recordset for the specified ID
        rs.Open "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & r, conn, 2, 3
        If Not rs.EOF Then
            'Update fields for the existing record
            With wshFACPrep
                rs.Fields("InvNo") = .Range("O6")
                rs.Fields("DateFacture") = .Range("O3").value
                rs.Fields("CustID") = .Range("B18").value
                rs.Fields("Contact") = .Range("K3").value
                rs.Fields("NomClient") = .Range("K4").value
                rs.Fields("Adresse") = .Range("K5").value
                rs.Fields("VilleProvCP") = .Range("K6").value
            End With
            With wshFACFinale
                rs.Fields("Honoraires") = .Range("F68").value
                rs.Fields("AF1Desc") = .Range("C69").value
                rs.Fields("AutresFrais1") = .Range("F69").value
                rs.Fields("AF2Desc") = .Range("C70").value
                rs.Fields("AutresFrais2") = .Range("F70").value
                rs.Fields("AF3Desc") = .Range("C71").value
                rs.Fields("AutresFrais3") = .Range("F71").value
                rs.Fields("TauxTPS") = .Range("D73").value
                rs.Fields("MntTPS") = .Range("F73").value
                rs.Fields("TauxTVQ") = .Range("D74").value
                rs.Fields("MntTVQ") = .Range("F74").value
                rs.Fields("AR_Total") = .Range("F76").value
                rs.Fields("Depot") = .Range("F78").value
            End With
        Else
            'Handle the case where the specified ID is not found
            MsgBox "L'enregistrement # '" & r & "' ne peut être trouvé!", vbExclamation
            rs.Close
            conn.Close
            Set rs = Nothing
            Set conn = Nothing
            Exit Sub
        End If
    End If
    'Update the recordset (create the record)
    rs.update
    
    'Prepare GL Posting
    wshFACPrep.Range("B33").value = wshFACFinale.Range("F80").value 'AR amount
    wshFACPrep.Range("B34").value = -wshFACFinale.Range("F68").value 'Revenues
    wshFACPrep.Range("B35").value = -wshFACFinale.Range("F69").value 'Misc $ - 1
    wshFACPrep.Range("B36").value = -wshFACFinale.Range("F70").value 'Misc $ - 2
    wshFACPrep.Range("B37").value = -wshFACFinale.Range("F71").value 'Misc $ - 3
    wshFACPrep.Range("B38").value = -wshFACFinale.Range("F73").value 'GST $
    wshFACPrep.Range("B39").value = -wshFACFinale.Range("F74").value 'PST $
    wshFACPrep.Range("B40").value = wshFACFinale.Range("F78").value 'Deposit
    
'    wshFACPrep.Range("B20").value = InvListRow
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Release objects from memory
    Set rs = Nothing
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("FAC_Prep_Add_Invoice_Header_to_DB()", timerStart)

End Sub

Sub FAC_Prep_Add_Invoice_Details_to_DB()

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim rowLastService As Long
    rowLastService = wshFACPrep.Range("L46").End(xlUp).row
    If rowLastService < 11 Then GoTo Nothing_to_Update
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "Invoice_Details"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    
    Dim r As Integer
    For r = 11 To rowLastService
        'Add fields to the recordset before updating it
        rs.AddNew
        With wshFACPrep
            rs.Fields("InvNo") = .Range("O6").value
            rs.Fields("Description") = .Range("L" & r).value
            rs.Fields("Heures") = .Range("M" & r).value
            rs.Fields("Taux") = .Range("N" & r).value
            rs.Fields("Honoraires") = .Range("O" & r).value
            rs.Fields("invRow") = r
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
    
Nothing_to_Update:

    Application.ScreenUpdating = True

    '{T2BRs}FAC_Prep_Add_Invoice_Details_to_DB{T2BRe}

End Sub

Sub FAC_Prep_Add_Comptes_Clients_to_DB()

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "Comptes_Clients"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFACPrep
        rs.Fields("Invoice_No") = .Range("O6").value
        rs.Fields("Invoice_Date") = CDate(.Range("O3").value)
        rs.Fields("Customer") = .Range("K4").value
        rs.Fields("Status") = "Unpaid"
        rs.Fields("Terms") = "Net 30"
        rs.Fields("Due_Date") = CDate(.Range("O3").value + 30)
        rs.Fields("Total") = wshFACFinale.Range("F80").value
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

    Call Output_Timer_Results("FAC_Prep_Add_Comptes_Clients_to_DB()", timerStart)

End Sub

Sub Update_TEC_As_Billed_In_DB(firstRow As Integer, lastRow As Integer) 'Update Billed Status in DB

    Dim timerStart As Double: timerStart = Timer

    Application.ScreenUpdating = False
    
    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "TEC"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Set rs = CreateObject("ADODB.Recordset")

    Dim r As Integer, TEC_ID As Long, SQL As String
    For r = firstRow To lastRow
        TEC_ID = wshBaseHours.Range("Y" & r).value
        'Open the recordset for the specified ID
        SQL = "SELECT * FROM [" & sheetName & "$] WHERE TEC_ID=" & TEC_ID
        rs.Open SQL, conn, 2, 3
        If Not rs.EOF Then
            'Update DateSaisie, EstFacturee, DateFacturee & NoFacture
            rs.Fields("DateSaisie").value = Now
            rs.Fields("EstFacturee").value = True
            rs.Fields("DateFacturee").value = CDate(wshFACPrep.Range("O3").value)
            rs.Fields("VersionApp").value = gAppVersion
            rs.Fields("NoFacture").value = wshFACPrep.Range("O6").value
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
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Set rs = Nothing
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    Call Output_Timer_Results("Update_TEC_As_Billed_In_DB()", timerStart)

End Sub

Sub FAC_Prep_TEC_As_Billed_Locally(firstResultRow As Integer, lastResultRow As Integer)

    Dim timerStart As Double: timerStart = Timer
    
    'Set the range to look for
    Dim lookupRange As Range, lastTECRow As Long
    lastTECRow = wshBaseHours.Range("A99999").End(xlUp).row
    Set lookupRange = wshBaseHours.Range("A3:A" & lastTECRow)
    
    Dim r As Integer, rowToBeUpdated As Long
    For r = firstResultRow To lastResultRow
        Dim tecID As Long
        tecID = wshBaseHours.Range("Y" & r).value
        rowToBeUpdated = Get_TEC_Row_Number_By_TEC_ID(tecID, lookupRange)
        Debug.Print "Need to update locally the TEC_ID = " & tecID & " which is at row # " & rowToBeUpdated
        wshBaseHours.Range("K" & rowToBeUpdated).value = Now()
        wshBaseHours.Range("L" & rowToBeUpdated).value = True
        wshBaseHours.Range("M" & rowToBeUpdated).value = CDate(wshFACPrep.Range("O3").value)
        wshBaseHours.Range("O" & rowToBeUpdated).value = gAppVersion
        wshBaseHours.Range("P" & rowToBeUpdated).value = wshFACPrep.Range("O6").value
    Next r
    
    Call Output_Timer_Results("FAC_Prep_TEC_As_Billed_Locally()", timerStart)

End Sub

'Sub ExampleUsage()
'    Dim rowToBeUpdated As Long
'
'    ' Call the function to get the row number of the unique ID
'    rowToBeUpdated = GetrowToBeUpdatedByTEC_ID(TEC_ID, lookupRange)
'
'    ' Display the result
'    If rowToBeUpdated > 0 Then
'        MsgBox "The row number for Unique ID '" & TEC_ID & "' is: " & rowToBeUpdated
'    Else
'        MsgBox "Unique ID '" & TEC_ID & "' not found."
'    End If
'End Sub

Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16
    If wshFACPrep.Range("B28").value Then Debug.Print vbNewLine & "[modFAC] - Now entering Sub Invoice_Load() @ " & Time
    With wshFACPrep
        If wshFACPrep.Range("B20").value = Empty Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
            GoTo NoItems
        End If
        'Could that invoice been cancelled (more than 1 row) ?
        Call InvoiceGetAllTrans(wshFACPrep.Range("O6").value)
        Dim NbTrans As Integer
        NbTrans = .Range("B31").value
        If NbTrans = 0 Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
            GoTo NoItems
        Else
            If NbTrans > 1 Then
                MsgBox "Cette facture a été annulée! Veuillez saisir un numéro de facture VALIDE pour votre recherche"
                GoTo NoItems
            End If
        End If
        If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Loading info from InvList with row # = " & .Range("B20").value
        .Range("B24").value = True 'Set Invoice Load to true
        .Range("S2,E4:F4,K4:L6,O3,K11:O45,Q11:Q45").ClearContents
        wshFACFinale.Range("C34:F63").ClearContents
        Dim InvListRow As Long
        InvListRow = wshFACPrep.Range("B20").value 'InvListRow = Row associated with the invoice
        'Get values from wshFACInvList (header) and enter them in the wshFACPrep - 2023-12-19 @ 08:29
        .Range("O3").value = wshFACInvList.Range("B" & InvListRow).value
        .Range("K3").value = wshFACInvList.Range("D" & InvListRow).value
        .Range("K4").value = wshFACInvList.Range("E" & InvListRow).value
        .Range("K5").value = wshFACInvList.Range("F" & InvListRow).value
        .Range("K6").value = wshFACInvList.Range("G" & InvListRow).value
        'Get values from wshFACInvList (header) and enter them in the wshFACPrep - 2023-12-19 @ 08:29
        wshFACFinale.Range("B21").value = "Le " & Format(wshFACInvList.Range("B" & InvListRow).value, "d mmmm yyyy")
        wshFACFinale.Range("B23").value = wshFACInvList.Range("D" & InvListRow).value
        wshFACFinale.Range("B24").value = wshFACInvList.Range("E" & InvListRow).value
        wshFACFinale.Range("B25").value = wshFACInvList.Range("F" & InvListRow).value
        wshFACFinale.Range("B26").value = wshFACInvList.Range("G" & InvListRow).value
        'Load Invoice Detail Items
        With wshFACInvItems
            Dim lastRow As Long, lastResultRow As Long
            lastRow = .Range("A999999").End(xlUp).row
            If lastRow < 4 Then Exit Sub 'No Item Lines
            .Range("I3").value = wshFACPrep.Range("O6").value
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Invoice Items - From Range '" & "A3:G" & lastRow & "', Critère = '" & .Range("I3").value & "'"
            wshFACFinale.Range("F28").value = wshFACPrep.Range("O6").value 'Invoice #
            'Advanced Filter to get items specific to ONE invoice
            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
            lastResultRow = .Range("O999").End(xlUp).row
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Based on column 'O' (Inv. Row), the lastResultRow = " & lastResultRow
            If lastResultRow < 3 Then GoTo NoItems
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("O" & resultRow).value
                If wshFACPrep.Range("B28").value Then Debug.Print Tab(10); "Loop = " & resultRow & " - Desc = " & .Range("K" & resultRow).value & " - Hrs = " & .Range("L" & resultRow).value
                wshFACPrep.Range("L" & invitemRow & ":O" & invitemRow).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
                wshFACPrep.Range("Q" & invitemRow).value = .Range("P" & resultRow).value  'Set Item DB Row
                wshFACFinale.Range("C" & invitemRow + 23 & ":F" & invitemRow + 23).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
            Next resultRow
        End With
        'Proceed with trailer data (Misc. charges & Taxes)
        .Range("M48").value = wshFACInvList.Range("I" & InvListRow).value
        .Range("O48").value = wshFACInvList.Range("J" & InvListRow).value
        .Range("M49").value = wshFACInvList.Range("K" & InvListRow).value
        .Range("O49").value = wshFACInvList.Range("L" & InvListRow).value
        .Range("M50").value = wshFACInvList.Range("M" & InvListRow).value
        .Range("O50").value = wshFACInvList.Range("N" & InvListRow).value
        .Range("O52").value = wshFACInvList.Range("P" & InvListRow).value
        .Range("O53").value = wshFACInvList.Range("R" & InvListRow).value
        .Range("O57").value = wshFACInvList.Range("T" & InvListRow).value

NoItems:
    .Range("B24").value = False 'Set Invoice Load To false
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "[modFAC] - Now exiting Sub Invoice_Load()" & vbNewLine
End Sub

Sub InvoiceGetAllTrans(inv As String)

    Application.ScreenUpdating = False

    wshFACPrep.Range("B31").value = 0

    With wshFACInvList
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A999999").End(xlUp).row 'Last wshFACInvList Row
        If lastRow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
        On Error Resume Next
        .Names("Criterial").delete
        On Error GoTo 0
        .Range("V3").value = wshFACPrep.Range("O6").value
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
'            .SortFields.Add Key:=wshFACInvList.Range("X2"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based Invoice Number
'            .SortFields.Add Key:=wshGL_Trans.Range("Y3"), _
'                SortOn:=xlSortOnValues, _
'                Order:=xlAscending, _
'                DataOption:=xlSortNormal 'Sort Based On TEC_ID
'            .SetRange wshFACInvList.Range("X2:AQ" & lastResultRow) 'Set Range
'            .Apply 'Apply Sort
'         End With
         wshFACPrep.Range("B31").value = lastResultRow - 2 'Remove Header rows from row count
Done:
    End With
    Application.ScreenUpdating = True

End Sub

Sub FAC_Prep_Clear_And_Fix_Totals_Formula()

    Application.EnableEvents = False
    
    With wshFACPrep
        .Range("J46:P59").ClearContents
        
        Call SetLabels(.Range("K47"), "FAC_Label_SubTotal_1")
        Call SetLabels(.Range("K51"), "FAC_Label_SubTotal_2")
        Call SetLabels(.Range("K52"), "FAC_Label_TPS")
        Call SetLabels(.Range("K53"), "FAC_Label_TVQ")
        Call SetLabels(.Range("K55"), "FAC_Label_GrandTotal")
        Call SetLabels(.Range("K57"), "FAC_Label_Deposit")
        Call SetLabels(.Range("K59"), "FAC_Label_AmountDue")
        
        .Range("O47").formula = "=SUM(O11:O45)" 'Fees sub-total
        .Range("O47").Font.Bold = True
        
        .Range("M48").value = wshAdmin.Range("FAC_Label_Frais_1").value 'Misc. # 1 - Descr.
        .Range("O48").value = "" 'Misc. # 1 - Amount
        .Range("M49").value = wshAdmin.Range("FAC_Label_Frais_2").value 'Misc. # 2 - Descr.
        .Range("O49").value = "" 'Misc. # 2 - Amount
        .Range("M50").value = wshAdmin.Range("FAC_Label_Frais_3").value 'Misc. # 3 - Descr.
        .Range("O50").value = "" 'Misc. # 3 - Amount
        
        .Range("O51").formula = "=sum(O47:O50)" 'Sub-total
        .Range("O51").Font.Bold = True
        
        .Range("N52").value = wshFACPrep.Range("B29").value 'GST Rate
        .Range("N52").NumberFormat = "0.00%"
        .Range("O52").formula = "=round(o51*n52,2)" 'GST Amnt
        .Range("N53").value = wshFACPrep.Range("B30").value 'PST Rate
        .Range("N53").NumberFormat = "0.000%"
        .Range("O53").formula = "=round(o51*n53,2)" 'GST Amnt
        .Range("O55").formula = "=sum(o51:o54)" 'Grand Total"
        .Range("O57").value = "" 'Deposit Amount
        .Range("O59").formula = "=O55-O57" 'Deposit Amount
        
    End With
    
    Application.EnableEvents = True

End Sub

Sub FAC_Prep_Clear_And_Fix_Formula_Finale()

    Application.EnableEvents = False
    
    With wshFACFinale
    
        Call SetLabels(.Range("B68"), "FAC_Label_SubTotal_1")
        Call SetLabels(.Range("B72"), "FAC_Label_SubTotal_2")
        Call SetLabels(.Range("B73"), "FAC_Label_TPS")
        Call SetLabels(.Range("B74"), "FAC_Label_TVQ")
        Call SetLabels(.Range("B76"), "FAC_Label_GrandTotal")
        Call SetLabels(.Range("B78"), "FAC_Label_Deposit")
        Call SetLabels(.Range("B80"), "FAC_Label_AmountDue")

        'Fix formulas to calculate amounts & Copy cells from FACPrep
        .Range("B23").value = "='" & wshFACPrep.name & "'!k3"
        .Range("B24").value = "='" & wshFACPrep.name & "'!k4"
        .Range("B25").value = "='" & wshFACPrep.name & "'!k5"
        .Range("B26").value = "='" & wshFACPrep.name & "'!k6"
        .Range("F68").formula = "=SUM(F33:F62)" 'Fees Sub-Total
        .Range("C69").formula = "='" & wshFACPrep.name & "'!M48" 'Misc. Amount # 1 - Description
        .Range("F69").formula = "='" & wshFACPrep.name & "'!O48" 'Misc. Amount # 1
        .Range("C70").formula = "='" & wshFACPrep.name & "'!M49" 'Misc. Amount # 2 - Description
        .Range("F70").formula = "='" & wshFACPrep.name & "'!O49" 'Misc. Amount # 2
        .Range("C71").formula = "='" & wshFACPrep.name & "'!M50" 'Misc. Amount # 3 - Description
        .Range("F71").formula = "='" & wshFACPrep.name & "'!O50" 'Misc. Amount # 3
        .Range("F72").formula = "=F68+F69+F70+F71" 'Sub-Total
        .Range("D73").formula = "='" & wshFACPrep.name & "'!N52" 'GST Rate
        .Range("F73").formula = "='" & wshFACPrep.name & "'!O52" 'GST Amount
        .Range("D74").formula = "='" & wshFACPrep.name & "'!N53" 'PST Rate
        .Range("F74").formula = "='" & wshFACPrep.name & "'!O53" 'PST Amount
        .Range("F76").formula = "=F72+F73+F74" 'Total including taxes
        .Range("F78").formula = "='" & wshFACPrep.name & "'!O57" 'Deposit Amount
        .Range("F80").formula = "=F76-F78" 'Total due on that invoice
    End With
    
    Application.EnableEvents = True
    
End Sub

Sub SetLabels(r As Range, l As String)

    r.value = wshAdmin.Range(l).value
    If wshAdmin.Range(l & "_Bold").value = "OUI" Then r.Font.Bold = True

End Sub

Sub MiscCharges()

    ActiveWindow.SmallScroll Down:=14
    wshFACPrep.Range("O48").Select 'Misc Amount 1
    
End Sub

Sub Client_Change(ClientName As String)

    wshFACPrep.Range("B18").value = GetID_From_Client_Name(ClientName)
    
    With wshFACPrep 'TODO[] - 2024-02-22
        .Range("K3").value = "Monsieur Robert M. Vigneault"
        .Range("K4").value = ClientName
        .Range("K5").value = "15 chemin des Mésanges" 'Address 1
        .Range("K6").value = "Mansonville, QC  J0E 1X0" 'Ville, Province & Code postal
    End With
    
    wshFACFinale.Range("B21").value = "Le " & Format(wshFACPrep.Range("O3").value, "d mmmm yyyy")
    
    Dim rng As Range
    Set rng = wshFACPrep.Range("E4:F4")
    Call Fill_Or_Empty_Range_Background(rng, False)
    
    Set rng = wshFACPrep.Range("O3")
    Call Fill_Or_Empty_Range_Background(rng, True, 6)
    
    wshFACPrep.Range("O3").Select 'Move on to Invoice Date

End Sub

Sub Date_Change(d As String)

    If InStr(1, wshFACPrep.Range("O6").value, "-") = 0 Then
        Dim Y As String
        Y = Right(Year(d), 2)
        wshFACPrep.Range("O6").value = Y & "-" & wshFACPrep.Range("O6").value
        wshFACFinale.Range("F28").value = wshFACPrep.Range("O6").value
    End If
    wshFACFinale.Range("B21").value = "Le " & Format(d, "d mmmm yyyy")
    
    'Must Get GST & PST rates and store them in wshFACPrep 'B' column
    Dim DateTaxRates As Date
    DateTaxRates = d
    wshFACPrep.Range("B29").value = GetTaxRate(DateTaxRates, "F")
    wshFACPrep.Range("B30").value = GetTaxRate(DateTaxRates, "P")
        
    Dim cutoffDate As Date
    cutoffDate = d
    Call Get_All_TEC_By_Client(cutoffDate, False)
    
    Dim rng As Range
    Set rng = wshFACPrep.Range("O3")
    Call Fill_Or_Empty_Range_Background(rng, False)
    
    Set rng = wshFACPrep.Range("L11")
    Call Fill_Or_Empty_Range_Background(rng, True, 6)

    On Error Resume Next
    wshFACPrep.Range("L11").Select 'Move on to Services Entry
    On Error GoTo 0
    
End Sub

Sub TEC_Clear_All_Cells()

    Dim lastRow As Long
    lastRow = wshFACPrep.Range("D999").End(xlUp).row
    If lastRow > 7 Then wshFACPrep.Range("D8:I" & lastRow).ClearContents
    
End Sub

Sub Get_All_TEC_By_Client(d As Date, includeBilledTEC As Boolean)

    'Set all criteria before calling FACPrep_TEC_Advanced_Filter_And_Sort
    Dim c1 As Long, c2 As String, c3 As Long
    Dim c4 As Boolean, c5 As Boolean, c6 As Boolean
    c1 = 0 'All professionnals
    c2 = "<=" & Format(d, "mm-dd-yyyy")
    c3 = wshFACPrep.Range("B18").value
    c4 = "VRAI"
    If includeBilledTEC Then c5 = "VRAI" Else c5 = "FAUX"
    c6 = "FAUX"

    Call FACPrep_TEC_Advanced_Filter_And_Sort(c1, c2, c3, c4, c5, c6)
    Call Copy_Filtered_Entries_To_FACPrep
    Call Add_Total_To_Filtered_Entries
    
End Sub

Sub FACPrep_TEC_Advanced_Filter_And_Sort(profID As Long, _
        cutoffDate As String, _
        clientID As Long, _
        isBillable As Boolean, _
        isInvoiced As Boolean, _
        isDeleted As Boolean)
    
    Application.ScreenUpdating = False

    With wshBaseHours
        'Is there anything to filter ?
        Dim lastSourceRow As Long, lastResultRow As Long
        lastSourceRow = .Range("A99999").End(xlUp).row 'Last TEC Entry row
        If lastSourceRow < 3 Then Exit Sub 'Nothing to filter
        
        'Clear the filtered rows area
        lastResultRow = .Range("Y9999").End(xlUp).row
        If lastResultRow > 2 Then .Range("Y3:AN" & lastResultRow).ClearContents
        
        Application.ScreenUpdating = False
        
        Dim rngSource As Range, rngCriteria As Range, rngCopyToRange As Range
        Set rngSource = wshBaseHours.Range("A2:P" & lastSourceRow)
        If profID <> 0 Then .Range("R3").value = profID
        .Range("S3").value = cutoffDate
        If clientID <> 0 Then .Range("T3").value = clientID
        .Range("U3").value = isBillable
        .Range("V3").value = isInvoiced
        .Range("W3").value = isDeleted
        Set rngCriteria = .Range("R2:W3")
        Set rngCopyToRange = .Range("Y2:AN2")
        
        rngSource.AdvancedFilter xlFilterCopy, rngCriteria, rngCopyToRange, Unique:=True
        
        lastResultRow = .Range("Y9999").End(xlUp).row
        If lastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If lastResultRow < 4 Then GoTo No_Sort_Required
        With .Sort
            .SortFields.clear
            .SortFields.add Key:=wshBaseHours.Range("AB3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.add Key:=wshBaseHours.Range("Z3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Prof_ID
            .SortFields.add Key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshBaseHours.Range("Y3:AN" & lastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
No_Sort_Required:
    End With
    
    Application.ScreenUpdating = True

End Sub

Sub Copy_Filtered_Entries_To_FACPrep()

    Dim lastRow As Long
    lastRow = wshBaseHours.Range("Y9999").End(xlUp).row
    If lastRow < 3 Then Exit Sub
    Dim arr() As String
    ReDim arr(1 To (lastRow - 2), 1 To 6) As String
    With wshBaseHours
        Dim i As Integer
        For i = 3 To lastRow
            arr(i - 2, 1) = .Range("AB" & i).value 'Date
            arr(i - 2, 2) = .Range("AA" & i).value 'Prof
            arr(i - 2, 3) = .Range("AE" & i).value 'Description
            arr(i - 2, 4) = .Range("AF" & i).value 'Heures
            arr(i - 2, 5) = .Range("AJ" & i).value 'Facturée ou pas
            arr(i - 2, 6) = .Range("Y" & i).value 'TEC_ID
        Next i
        'Copy array to worksheet
        Dim rng As Range
        'Set rng = .Range("D8").Resize(UBound(arr, 1), UBound(arr, 2))
        Set rng = wshFACPrep.Range("D8").Resize(lastRow - 2, UBound(arr, 2))
        rng.value = arr
        .Range("G8:G" & lastRow + 5).NumberFormat = "##0.00"
    End With
End Sub

Sub Add_Total_To_Filtered_Entries()

    Dim lastRow As Integer
    lastRow = wshFACPrep.Range("D9999").End(xlUp).row
    If lastRow = 7 Then Exit Sub 'Nothing to add
    Dim i As Integer, totalHres As Currency
    For i = 8 To lastRow
        totalHres = totalHres + wshFACPrep.Range("G" & i).value
    Next i

    wshFACPrep.Range("G" & lastRow + 2).value = totalHres
    
End Sub
 
Sub Invoice_Delete()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub Invoice_Delete() @ " & Time
    With wshFACPrep
        If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Invoice") = vbNo Then Exit Sub
        If .Range("B20").value = Empty Then GoTo NotSaved
        invRow = .Range("B20").value 'Set Invoice Row
        wshFACInvList.Range(invRow & ":" & invRow).EntireRow.delete
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
'    '         .SortFields.Add Key:=wshFACInvItems.Range("W3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal  'Sort
'    '         .SetRange wshFACInvItems.Range("P3:W" & lastResultRow) 'Set Range
'    '         .Apply 'Apply Sort
'    '         End With
'SkipSort:
'            For ResultRow = 3 To lastResultRow
'                itemDBRow = .Range("V" & ResultRow).value 'Set Invoice Database Row
'                .Range("A" & itemDBRow & ":J" & itemDBRow).ClearContents 'Clear Fields (deleting creates issues with results
'            Next ResultRow
'            'Resort DB to remove spaces
'            With .Sort
'                .SortFields.Clear
'                .SortFields.Add Key:=wshFACInvItems.Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal  'Sort
'                .SetRange wshFACInvItems.Range("A4:J" & lastResultRow) 'Set Range
'                .Apply 'Apply Sort
'            End With
'        End With
NoItems:
NotSaved:
    Call FAC_Prep_Invoice_New 'Add New Invoice
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFAC] - Sub Invoice_Delete()" & vbNewLine
End Sub

Sub Previsualisation_PDF() 'RMV - 2023-12-17 @ 14:33

    wshFACFinale.PrintOut , , , True, True, , , , False
    
End Sub

Sub Creation_PDF_Email() 'RMV - 2023-12-17 @ 14:35
    
    Call Create_PDF_Email_Sub(wshFACPrep.Range("O6").value)

End Sub

Sub Create_PDF_Email_Sub(noFacture As String)
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Create_PDF_Email_Sub(NoFacture As String) @ " & Time
    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF et préparation du courriel pour envoyer la facture
    Dim result As Boolean
    result = Create_PDF_Email_Function(noFacture, "CreateEmail")
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFAC] - Create_PDF_Email_Sub(NoFacture As String)" & vbNewLine
End Sub

Function Create_PDF_Email_Function(noFacture As String, Optional action As String = "SaveOnly") As Boolean
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Now entering - [modFAC] - Function Create_PDF_Email_Function" & _
        "(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean @ " & Time
    Dim SaveAs As String

    Application.ScreenUpdating = False

    'Construct the SaveAs filename
    'NoFactFormate = Format(NoFacture, "000000")
    SaveAs = wshAdmin.Range("FolderPDFInvoice").value & Application.PathSeparator & _
                     noFacture & ".pdf" '2023-12-19 @ 07:28

    'Set Print Quality
    On Error Resume Next
    ActiveSheet.PageSetup.PrintQuality = 600
    Err.clear
    On Error GoTo 0

    'Adjust Document Properties - 2023-10-06 @ 09:54
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.1)
        .RightMargin = Application.InchesToPoints(0.1)
        .TopMargin = Application.InchesToPoints(0.1)
        .BottomMargin = Application.InchesToPoints(0.1)
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
        Dim myMail As Outlook.MailItem
        
        Set outlookApp = New Outlook.Application
        Set myMail = outlookApp.CreateItem(olMailItem)

        Dim source_file As String
        source_file = wshAdmin.Range("FolderPDFInvoice").value & Application.PathSeparator & _
                      noFacture & ".pdf" '2023-12-19 @ 07:22
        
        With myMail
            .To = "robertv13@hotmail.com"
            .CC = "robertv13@me.com"
            .BCC = "robertv13@gmail.com"
            .Subject = "TEST - GC FISCALITÉ INC. - Facturation - TEST"
            .Body = "Bonjour," & vbNewLine & vbNewLine & "Vous trouverez ci-joint notre note d'honoraires." & _
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
    Create_PDF_Email_Function = True 'Return value
    GoTo EndMacro
    
RefLibError:
    MsgBox "Incapable de préparer le courriel. La librairie n'est pas disponible"
    Create_PDF_Email_Function = False 'Return value

EndMacro:
    Application.ScreenUpdating = True
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Now exiting  - [modFAC] - Create_PDF_Email_Function(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean" & vbNewLine
End Function

Sub Prev_Invoice() 'TO-DO-RMV 2023-12-17
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub Prev_Invoice() @ " & Time
    With wshFACPrep
        Dim MininvNumb As Long
        On Error Resume Next
        MininvNumb = Application.WorksheetFunction.Min(wshFACInvList.Range("Inv_ID"))
        On Error GoTo 0
        If MininvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").value
        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            invRow = wshFACInvList.Range("A99999").End(xlUp).row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFACInvList.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).row - 1
        End If
        If .Range("N6").value = 1 Or MininvNumb = 0 Or MininvNumb = .Range("N6").value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFACInvList.Range("A" & invRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFAC] - Sub Prev_Invoice()" & vbNewLine
End Sub

Sub Next_Invoice() 'TO-DO-RMV 2023-12-17
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub Next_Invoice() @ " & Time
    With wshFACPrep
        Dim MaxinvNumb As Long
        On Error Resume Next
        MaxinvNumb = Application.WorksheetFunction.Max(wshFACInvList.Range("Inv_ID"))
        On Error GoTo 0
        If MaxinvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").value
        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            invRow = wshFACInvList.Range("A4").value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFACInvList.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).row + 1
        End If
        If .Range("N6").value >= MaxinvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFACInvList.Range("A" & invRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFAC] - Sub Next_Invoice()" & vbNewLine
End Sub

Sub Cacher_Heures()
    With wshFACFinale.Range("D34:F66")
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Montrer_Heures()
    With wshFACFinale.Range("D34:F66")
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
End Sub

Sub Goto_Onglet_Preparation_Facture()
    wshFACPrep.Visible = xlSheetVisible
    wshFACPrep.Activate
    wshFACPrep.usedRange.Columns("B:P").Calculate
    wshFACPrep.Range("C1").Select
End Sub

Sub Goto_Onglet_Facture_Finale()
    wshFACFinale.Visible = xlSheetVisible
    wshFACFinale.Activate
    wshFACFinale.usedRange.Columns("A:G").Calculate
    wshFACFinale.Range("I48").Select
End Sub

Sub ExportAllFacInvList() '2023-12-21 @ 14:36
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceRange As Range

    Application.ScreenUpdating = False
    
    'Work with the source range
    Set wsSource = wshFACInvList
    Dim lastUsedRow As Long
    lastUsedRow = wsSource.Range("A99999").End(xlUp).row
    wsSource.Range("A4:T" & lastUsedRow).Copy

    'Open the target workbook
    Workbooks.Open fileName:=wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
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

Sub FAC_Prepare_GL_Posting() '2024-02-14 @ 05:56

    Dim montant As Double
    Dim dateFact As Date
    Dim descGL_Trans As String, source As String
    
    dateFact = wshFACPrep.Range("O3").value
    descGL_Trans = wshFACPrep.Range("E4").value
    source = "FACT-" & wshFACPrep.Range("O6").value
    
    Dim myArray(1 To 8, 1 To 4) As String
    
    'AR amount (wshFacPrep.Range("B33"))
    montant = wshFACPrep.Range("B33").value
    If montant Then
        myArray(1, 1) = "1100"
        myArray(1, 2) = "Comptes Clients"
        myArray(1, 3) = montant
        myArray(1, 4) = ""
    End If
    
    'Professional Fees (wshFacPrep.Range("B34"))
    montant = wshFACPrep.Range("B34").value
    If montant Then
        myArray(2, 1) = "4000"
        myArray(2, 2) = "Revenus"
        myArray(2, 3) = montant
        myArray(2, 4) = ""
    End If
    
    'Miscellaneous Amount # 1 (wshFacPrep.Range("B35"))
    montant = wshFACPrep.Range("B35").value
    If montant Then
        myArray(3, 1) = "5009"
        myArray(3, 2) = "Frais divers # 1"
        myArray(3, 3) = montant
        myArray(3, 4) = ""
    End If
    
    'Miscellaneous Amount # 2 (wshFacPrep.Range("B36"))
    montant = wshFACPrep.Range("B36").value
    If montant Then
        myArray(4, 1) = "5008"
        myArray(4, 2) = "Frais divers # 2"
        myArray(4, 3) = montant
        myArray(4, 4) = ""
    End If
    
    'Miscellaneous Amount # 3 (wshFacPrep.Range("B37"))
    montant = wshFACPrep.Range("B37").value
    If montant Then
        myArray(5, 1) = "5002"
        myArray(5, 2) = "Frais divers # 3"
        myArray(5, 3) = montant
        myArray(5, 4) = ""
    End If
    
    'GST to pay (wshFacPrep.Range("B38"))
    montant = wshFACPrep.Range("B38").value
    If montant Then
        myArray(6, 1) = "2200"
        myArray(6, 2) = "TPS à payer"
        myArray(6, 3) = montant
        myArray(6, 4) = ""
    End If
    
    'PST to pay (wshFacPrep.Range("B39"))
    montant = wshFACPrep.Range("B39").value
    If montant Then
        myArray(7, 1) = "2201"
        myArray(7, 2) = "TVQ à payer"
        myArray(7, 3) = montant
        myArray(7, 4) = ""
    End If
    
    'Deposit applied (wshFacPrep.Range("B40"))
    montant = wshFACPrep.Range("B40").value
    If montant Then
        myArray(8, 1) = "1230"
        myArray(8, 2) = "Avance - Prêt GCP"
        myArray(8, 3) = montant
        myArray(8, 4) = ""
    End If
    
    Call FAC_GL_Posting(dateFact, descGL_Trans, source, myArray)
    
End Sub

Sub FAC_GL_Posting(df, desc, source, arr As Variant)

    Dim fullFileName As String, sheetName As String
    fullFileName = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                   "GCF_BD_Sortie.xlsx"
    sheetName = "GL_Trans"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Initialize recordset
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(No_Entrée) AS MaxEJNo FROM [" & sheetName & "$]"

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

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & sheetName & "$] WHERE 1=0", conn, 2, 3
    
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
            rs.update
Nothing_to_Post:
    Next i

    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

End Sub

Sub Back_To_FAC_Menu()

    wshMenuFACT.Activate
    wshMenuFACT.Range("A1").Select
    
End Sub




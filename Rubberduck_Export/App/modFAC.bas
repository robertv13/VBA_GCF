Attribute VB_Name = "modFAC"
Option Explicit
Dim InvRow As Long, InvCol As Long, ItemDBRow As Long, InvItemRow As Long, InvNumb As Long
Dim lastrow As Long, LastItemRow As Long, LastResultRow As Long, ResultRow As Long

Sub FAC_Prep_Invoice_New() 'Clear contents
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
    If wshFACPrep.Range("B28").value Then Debug.Print vbNewLine & "Le num�ro de facture '" & wshFACPrep.Range("O6").value & "' a �t� assign�e"
End Sub

Sub FAC_Prep_Save_And_Update() '2024-02-21 @ 10:11
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub FAC_Prep_Save_And_Update() @ " & Time
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "B18 (Cust. ID) = " & wshFACPrep.Range("B18").value
    With wshFACPrep
        'Check For Mandatory Fields - Client
        If .Range("B18").value = Empty Then
            MsgBox "Veuillez vous assurer d'avoir un client avant de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUS�E parce que le nom de client n'est pas encore saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        'Check For Mandatory Fields - Date de facture
        If .Range("O3").value = Empty Or Len(Trim(.Range("O6").value)) <> 8 Then
            MsgBox "Veuillez vous assurer d'avoir saisi la date de facture AVANT de sauvegarder la facture"
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Sauvegarde REFUS�E parce que la date de facture et/ou le num�ro de facture n'ont pas encore �t� saisi, sortie de la routine"
            GoTo Fast_Exit_Sub
        End If
        
        'Valid Invoice - Let's update it
        'Determine the row number (InvListRow) for New Invoice -OR- use existing one
        If wshFACPrep.Range("B20").value = Empty Then 'New Invoice
            Call FAC_Prep_Add_Invoice_Header_to_DB(0)
            Call FAC_Prep_Add_Invoice_Details_to_DB
        End If
    End With
NoItems:
    MsgBox "La facture '" & wshFACPrep.Range("O6").value & "' est enregistr�e." & vbNewLine & vbNewLine & "Le total de la facture est " & Trim(Format(wshFACPrep.Range("O51").value, "### ##0.00 $")) & " (avant les taxes)", vbOKOnly, "Confirmation d'enregistrement"
    wshFACPrep.Range("B27").value = False
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Total de la facture '"; wshFACPrep.Range("O6") & "' (avant taxes) est de " & Format(wshFACPrep.Range("O51").value, "### ##0.00 $")
Fast_Exit_Sub:
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFAC] - Sub FAC_Prep_Save_And_Update()" & vbNewLine
    
    Call FAC_Prepare_GL_Posting(0)
    
End Sub

Sub FAC_Prep_Add_Invoice_Header_to_DB(r As Long)

    Dim timerStart As Double 'Speed tests - 2024-02-21
    timerStart = Timer

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
            MsgBox "L'enregistrement # '" & r & "' ne peut �tre trouv�!", vbExclamation
            rs.Close
            conn.Close
            Set rs = Nothing
            Set conn = Nothing
            Exit Sub
        End If
    End If
    'Update the recordset (create the record)
    rs.Update
    
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

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "FAC_Prep_Add_Invoice_Header_to_DB() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")

End Sub

Sub FAC_Prep_Add_Invoice_Details_to_DB()

    Dim timerStart As Double 'Speed tests - 2024-02-21
    timerStart = Timer

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
            rs.Fields("InvRow") = r
            rs.Fields("Row") = ""
        End With
    'Update the recordset (create the record)
    rs.Update
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

    Debug.Print vbNewLine & String(45, "*") & vbNewLine & _
        "FAC_Prep_Add_Invoice_Header_to_DB() - Secondes = " & Timer - timerStart & _
        vbNewLine & String(45, "*")

End Sub

Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16
    If wshFACPrep.Range("B28").value Then Debug.Print vbNewLine & "[modFAC] - Now entering Sub Invoice_Load() @ " & Time
    With wshFACPrep
        If wshFACPrep.Range("B20").value = Empty Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un num�ro de facture VALIDE pour votre recherche"
            GoTo NoItems
        End If
        'Could that invoice been cancelled (more than 1 row) ?
        Call InvoiceGetAllTrans(wshFACPrep.Range("O6").value)
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
            Dim lastrow As Long, LastResultRow As Long
            lastrow = .Range("A999999").End(xlUp).row
            If lastrow < 4 Then Exit Sub 'No Item Lines
            .Range("I3").value = wshFACPrep.Range("O6").value
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Invoice Items - From Range '" & "A3:G" & lastrow & "', Crit�re = '" & .Range("I3").value & "'"
            wshFACFinale.Range("F28").value = wshFACPrep.Range("O6").value 'Invoice #
            'Advanced Filter to get items specific to ONE invoice
            .Range("A3:G" & lastrow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
            LastResultRow = .Range("O999").End(xlUp).row
            If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Based on column 'O' (Inv. Row), the LastResultRow = " & LastResultRow
            If LastResultRow < 3 Then GoTo NoItems
            For ResultRow = 3 To LastResultRow
                InvItemRow = .Range("O" & ResultRow).value
                If wshFACPrep.Range("B28").value Then Debug.Print Tab(10); "Loop = " & ResultRow & " - Desc = " & .Range("K" & ResultRow).value & " - Hrs = " & .Range("L" & ResultRow).value
                wshFACPrep.Range("L" & InvItemRow & ":O" & InvItemRow).value = .Range("K" & ResultRow & ":N" & ResultRow).value 'Description, Hours, Rate & Value
                wshFACPrep.Range("Q" & InvItemRow).value = .Range("P" & ResultRow).value  'Set Item DB Row
                wshFACFinale.Range("C" & InvItemRow + 23 & ":F" & InvItemRow + 23).value = .Range("K" & ResultRow & ":N" & ResultRow).value 'Description, Hours, Rate & Value
            Next ResultRow
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

'Sub InvoiceGetAllTrans(inv As String)
'
'    Application.ScreenUpdating = False
'
'    wshFACPrep.Range("B31").value = 0
'
'    With wshFACInvList
'        Dim lastrow As Long, LastResultRow As Long, ResultRow As Long
'        lastrow = .Range("A999999").End(xlUp).row 'Last wshFACInvList Row
'        If lastrow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
'        On Error Resume Next
'        .Names("Criterial").Delete
'        On Error GoTo 0
'        .Range("V3").value = wshFACPrep.Range("O6").value
'        'Advanced Filter setup
'        .Range("A3:T" & lastrow).AdvancedFilter xlFilterCopy, _
'            CriteriaRange:=.Range("V2:V3"), _
'            CopyToRange:=.Range("X2:AQ2"), _
'            Unique:=True
'        LastResultRow = .Range("X999").End(xlUp).row 'How many rows trans for that invoice
'        If LastResultRow < 3 Then
'            GoTo Done
'        End If
''        With .Sort
''            .SortFields.Clear
''            .SortFields.Add Key:=wshFACInvList.Range("X2"), _
''                SortOn:=xlSortOnValues, _
''                Order:=xlAscending, _
''                DataOption:=xlSortNormal 'Sort Based Invoice Number
''            .SortFields.Add Key:=wshGL_Trans.Range("Y3"), _
''                SortOn:=xlSortOnValues, _
''                Order:=xlAscending, _
''                DataOption:=xlSortNormal 'Sort Based On TEC_ID
''            .SetRange wshFACInvList.Range("X2:AQ" & lastResultRow) 'Set Range
''            .Apply 'Apply Sort
''         End With
'         wshFACPrep.Range("B31").value = LastResultRow - 2 'Remove Header rows from row count
'Done:
'    End With
'    Application.ScreenUpdating = True
'
'End Sub

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

        'Fix formulas to calculate amounts & Copy cells from FAC_Pr�paration
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
    
    With wshFACPrep
        .Range("K3").value = "Monsieur Robert M. Vigneault"
        .Range("K4").value = ClientName
        .Range("K5").value = "15 chemin des M�sanges" 'Address 1
        .Range("K6").value = "Mansonville, QC  J0E 1X0" 'Ville, Province & Code postal
    End With
    With wshFACFinale
        .Range("B21").value = "Le " & Format(wshFACPrep.Range("O3").value, "d mmmm yyyy")
        .Range("B23").value = wshFACPrep.Range("K3").value 'Contact from wshFACPrep
        .Range("B24").value = wshFACPrep.Range("K4").value 'Client from wshFACPrep
        .Range("B25").value = wshFACPrep.Range("K5").value 'Address 1 from wshFACPrep
        .Range("B26").value = wshFACPrep.Range("K6").value
    End With
    
    TEC_Load
    
    Range("E4:F4").Select
    
    With wshFACPrep.Range("E4:F4").Interior 'No filling
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With wshFACPrep.Range("O3").Interior 'Yellow filling
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    wshFACPrep.Range("O3").Select 'Move on to Invoice Date

End Sub

Sub DateChange(d As String)

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
        
    'Reset the cell (Date)
    With wshFACPrep.Range("O3").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    wshFACPrep.Range("L11").Select 'Move on to Services Entry
    
End Sub

Sub TEC_Clear_All_Cells()

    Dim lastrow As Long
    lastrow = wshFACPrep.Range("D999").End(xlUp).row
    If lastrow > 7 Then wshFACPrep.Range("D8:I" & lastrow).ClearContents
    
End Sub

Sub TEC_Load()

    'Set Criteria, before Filtering entries
    wshBaseHours.Range("T3").value = wshFACPrep.Range("B18").value
    TECByClient_FilterAndSort (wshFACPrep.Range("B18").value)
    
    'Reset Criteria, after Filtering entries
    wshBaseHours.Range("T3").value = ""
    
    CopyFromFilteredEntriesToFACPrep
    
    wshFACPrep.Range("O3").Select
    
End Sub

Sub TECByClient_FilterAndSort(id As Long) 'RMV-2023-12-21 @ 11:00
    
    Application.ScreenUpdating = False

    Call TEC_Import_All '2024-02-14 @ 06:20
    
    With wshBaseHours
        Dim lastrow As Long, LastResultRow As Long, ResultRow As Long
        lastrow = .Range("A999999").End(xlUp).row 'Last BaseHours Row
        If lastrow < 3 Then Exit Sub 'Nothing to filter
        Application.ScreenUpdating = False
        On Error Resume Next
        .Names("Criterial").Delete
        On Error GoTo 0
        .Range("A2:Q" & lastrow).AdvancedFilter xlFilterCopy, _
            CriteriaRange:=.Range("T2:W3"), _
            CopyToRange:=.Range("Y2:AL2"), _
            Unique:=True
        LastResultRow = .Range("Y999999").End(xlUp).row
        If LastResultRow < 3 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        If LastResultRow < 4 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            .SortFields.Add key:=wshBaseHours.Range("AA3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On Date
            .SortFields.Add key:=wshBaseHours.Range("Y3"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal 'Sort Based On TEC_ID
            .SetRange wshBaseHours.Range("Y3:AL" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
         End With
NoSort:
    End With
    Application.ScreenUpdating = True
End Sub

Sub CopyFromFilteredEntriesToFACPrep()

    Dim lastrow As Long
    lastrow = wshBaseHours.Range("Y99999").End(xlUp).row
    Dim row As Long
    row = 8
    
    Dim i As Integer
    With wshBaseHours
        For i = 3 To lastrow
            If .Range("AH" & i).value = False And .Range("AJ" & i).value = False Then
                wshFACPrep.Range("D" & row).value = .Range("AA" & i).value 'Date
                wshFACPrep.Range("E" & row).value = .Range("Z" & i).value 'Date
                wshFACPrep.Range("F" & row).value = .Range("AC" & i).value 'Description
                wshFACPrep.Range("G" & row).value = .Range("AD" & i).value 'Heures
                wshFACPrep.Range("H" & row).value = .Range("AH" & i).value 'Factur�e ou pas
                wshFACPrep.Range("I" & row).value = .Range("Y" & i).value 'TEC_ID
                row = row + 1
            End If
        Next i
    End With
End Sub

Sub Invoice_Delete()
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub Invoice_Delete() @ " & Time
    With wshFACPrep
        If MsgBox("Are you sure you want to delete this Invoice?", vbYesNo, "Delete Invoice") = vbNo Then Exit Sub
        If .Range("B20").value = Empty Then GoTo NotSaved
        InvRow = .Range("B20").value 'Set Invoice Row
        wshFACInvList.Range(InvRow & ":" & InvRow).EntireRow.Delete
        With InvItems
            lastrow = .Range("A99999").End(xlUp).row
            If lastrow < 4 Then Exit Sub
            .Range("A3:J" & lastrow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("N2:N3"), CopyToRange:=.Range("P2:W2"), Unique:=True
            LastResultRow = .Range("V99999").End(xlUp).row
            If LastResultRow < 3 Then GoTo NoItems
    '        If LastResultRow < 4 Then GoTo SkipSort
    '        'Sort Rows Descending
    '         With .Sort
    '         .SortFields.Clear
    '         .SortFields.Add Key:=wshFACInvItems.Range("W3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal  'Sort
    '         .SetRange wshFACInvItems.Range("P3:W" & LastResultRow) 'Set Range
    '         .Apply 'Apply Sort
    '         End With
SkipSort:
            For ResultRow = 3 To LastResultRow
                ItemDBRow = .Range("V" & ResultRow).value 'Set Invoice Database Row
                .Range("A" & ItemDBRow & ":J" & ItemDBRow).ClearContents 'Clear Fields (deleting creates issues with results
            Next ResultRow
            'Resort DB to remove spaces
            With .Sort
                .SortFields.Clear
                .SortFields.Add key:=wshFACInvItems.Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal  'Sort
                .SetRange wshFACInvItems.Range("A4:J" & LastResultRow) 'Set Range
                .Apply 'Apply Sort
            End With
        End With
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
    'Cr�ation du fichier (NoFacture).PDF dans le r�pertoire de factures PDF de GCF et pr�paration du courriel pour envoyer la facture
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
    Err.Clear
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
            .Subject = "TEST - GC FISCALIT� INC. - Facturation - TEST"
            .Body = "Bonjour," & vbNewLine & vbNewLine & "Vous trouverez ci-joint notre note d'honoraires." & _
                vbNewLine & vbNewLine & "Merci" & vbNewLine & vbNewLine & vbNewLine & "Guillaume Charron, CPA, CA, M. Fisc." & _
                vbNewLine & "Pr�sident"
            .Attachments.Add source_file
           
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
    MsgBox "Incapable de pr�parer le courriel. La librairie n'est pas disponible"
    Create_PDF_Email_Function = False 'Return value

EndMacro:
    Application.ScreenUpdating = True
    If wshFACPrep.Range("B28").value Then Debug.Print Tab(5); "Now exiting  - [modFAC] - Create_PDF_Email_Function(NoFacture As Long, Optional action As String = """"SaveOnly"""") As Boolean" & vbNewLine
End Function

Sub Prev_Invoice() 'TO-DO-RMV 2023-12-17
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub Prev_Invoice() @ " & Time
    With wshFACPrep
        Dim MinInvNumb As Long
        On Error Resume Next
        MinInvNumb = Application.WorksheetFunction.Min(wshFACInvList.Range("Inv_ID"))
        On Error GoTo 0
        If MinInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").value
        If InvNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            InvRow = wshFACInvList.Range("A99999").End(xlUp).row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            InvRow = wshFACInvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).row - 1
        End If
        If .Range("N6").value = 1 Or MinInvNumb = 0 Or MinInvNumb = .Range("N6").value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFACInvList.Range("A" & InvRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    If wshFACPrep.Range("B28").value Then Debug.Print "Now exiting  - [modFAC] - Sub Prev_Invoice()" & vbNewLine
End Sub

Sub Next_Invoice() 'TO-DO-RMV 2023-12-17
    If wshFACPrep.Range("B28").value Then Debug.Print "Now entering - [modFAC] - Sub Next_Invoice() @ " & Time
    With wshFACPrep
        Dim MaxInvNumb As Long
        On Error Resume Next
        MaxInvNumb = Application.WorksheetFunction.Max(wshFACInvList.Range("Inv_ID"))
        On Error GoTo 0
        If MaxInvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        InvNumb = .Range("N6").value
        If InvNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            InvRow = wshFACInvList.Range("A4").value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            InvRow = wshFACInvList.Range("Inv_ID").Find(InvNumb, , xlValues, xlWhole).row + 1
        End If
        If .Range("N6").value >= MaxInvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFACInvList.Range("A" & InvRow).value 'Place Inv. ID inside cell
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
    wshFACPrep.Range("C1").Select
End Sub

Sub Goto_Onglet_Facture_Finale()
    wshFACFinale.Visible = xlSheetVisible
    wshFACFinale.Activate
    wshFACFinale.Range("C1").Select
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

Sub FAC_Prepare_GL_Posting(r As Long) '2024-02-14 @ 05:56

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
        myArray(6, 2) = "TPS � payer"
        myArray(6, 3) = montant
        myArray(6, 4) = ""
    End If
    
    'PST to pay (wshFacPrep.Range("B39"))
    montant = wshFACPrep.Range("B39").value
    If montant Then
        myArray(7, 1) = "2201"
        myArray(7, 2) = "TVQ � payer"
        myArray(7, 3) = montant
        myArray(7, 4) = ""
    End If
    
    'Deposit applied (wshFacPrep.Range("B40"))
    montant = wshFACPrep.Range("B40").value
    If montant Then
        myArray(8, 1) = "1230"
        myArray(8, 2) = "Avance - Pr�t GCP"
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
    strSQL = "SELECT MAX(No_Entr�e) AS MaxEJNo FROM [" & sheetName & "$]"

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
            rs.Update
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


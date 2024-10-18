Attribute VB_Name = "modFAC_Finale"
Option Explicit

Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub FAC_Finale_Save() '2024-03-28 @ 07:19

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Save (" & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy") & ")", 0)

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
            MsgBox "Il faut corriger le num�ro de facture AVANT de sauvegarder la facture"
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
    lastResultRow = wshTEC_Local.Range("AT9999").End(xlUp).Row
        
    If lastResultRow > 2 Then
        Call FAC_Finale_TEC_Update_As_Billed_To_DB(3, lastResultRow)
        Call FAC_Finale_TEC_Update_As_Billed_Locally(3, lastResultRow)
    End If
    
    'Update FAC_Projets_Ent�te & FAC_Projets_D�tails, if necessary
    Dim projetID As Long
    projetID = wshFAC_Brouillon.Range("B52").value
    If projetID <> 0 Then
        Call FAC_Finale_Softdelete_Projets_D�tails_To_DB(projetID)
        Call FAC_Finale_Softdelete_Projets_D�tails_Locally(projetID)
        
        Call FAC_Finale_Softdelete_Projets_Ent�te_To_DB(projetID)
        Call FAC_Finale_Softdelete_Projets_Ent�te_Locally(projetID)
    End If
        
    'Save Invoice total amount
    Dim invoice_Total As Currency
    invoice_Total = wshFAC_Brouillon.Range("O51").value
        
    'GL stuff will occur at the confirmation level (later)
'    Call FAC_Finale_GL_Posting_Preparation
    
    'Update TEC_DashBoard
    Call TEC_TdB_Update_All '2024-03-21 @ 12:32

    Call FAC_Brouillon_Clear_All_TEC_Displayed
    
    Application.ScreenUpdating = True
    
    MsgBox "La facture '" & wshFAC_Brouillon.Range("O6").value & "' est enregistr�e." & _
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
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Save()", startTime)
    
End Sub

Sub FAC_Finale_Add_Invoice_Header_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_to_DB (" & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy") & ")", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Ent�te"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Can only ADD to the file, no modification is allowed
    
    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields("Inv_No") = .Range("E28").value
        rs.Fields("Date_Facture") = Format$(wshFAC_Brouillon.Range("O3").value, "mm-dd-yyyy")
        rs.Fields("AC_C") = "AC" 'Facture to be confirmed MANUALLY - 2024-08-16 @ 05:46
        rs.Fields("Cust_ID") = wshFAC_Brouillon.Range("B18").value
        rs.Fields("Contact") = .Range("B23").value
        rs.Fields("Nom_Client") = .Range("B24").value
        rs.Fields("Adresse1") = .Range("B25").value
        rs.Fields("Adresse2") = .Range("B26").value
        rs.Fields("Adresse3") = .Range("B27").value
        
        rs.Fields("Honoraires") = Format$(.Range("E69").value, "0.00")
        
        rs.Fields("AF1_Desc") = .Range("B70").value
        rs.Fields("AutresFrais_1") = Format$(wshFAC_Finale.Range("E70").value, "0.00")
        rs.Fields("AF2_Desc") = .Range("B71").value
        rs.Fields("AutresFrais_2") = Format$(.Range("E71").value, "0.00")
        rs.Fields("AF3_Desc") = .Range("B72").value
        rs.Fields("AutresFrais_3") = Format$(.Range("E72").value, "0.00")
        
        rs.Fields("Taux_TPS") = Format$(.Range("C74").value, "0.00")
        rs.Fields("Mnt_TPS") = Format$(.Range("E74").value, "0.00")
        rs.Fields("Taux_TVQ") = Format$(.Range("C75").value, "0.00000")  '2024-10-15 @ 05:49
        rs.Fields("Mnt_TVQ") = Format$(.Range("E75").value, "0.00")
        
        rs.Fields("AR_Total") = Format$(.Range("E77").value, "0.00")
        
        rs.Fields("D�p�t") = Format$(.Range("E79").value, "0.00")
    End With
    'Update the recordset (create the record)
    rs.update
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set rs = Nothing
    Set conn = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_to_DB()", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Header_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally (" & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy") & ")", 0)
    
    Application.ScreenUpdating = False
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Ent�te.Range("A9999").End(xlUp).Row + 1
    
    With wshFAC_Ent�te
        .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).value = Format$(wshFAC_Brouillon.Range("O3").value, "mm-dd-yyyy")
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
    End With
    
    Application.EnableEvents = False
    wshFAC_Brouillon.Range("B11").value = firstFreeRow
    Application.EnableEvents = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally()", startTime)

    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_Add_Invoice_Details_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_to_DB (" & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy") & ")", 0)

    Application.ScreenUpdating = False
    
    Dim rowLastService As Long
    rowLastService = wshFAC_Finale.Range("B64").End(xlUp).Row
    If rowLastService < 34 Then GoTo nothing_to_update
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_D�tails"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").value
    Dim r As Long
    For r = 34 To rowLastService
        'Add fields to the recordset before updating it
        rs.AddNew
        With wshFAC_Finale
            rs.Fields("Inv_No") = CStr(noFacture)
            rs.Fields("Description") = .Range("B" & r).value
            If .Range("C" & r).value <> 0 And _
               .Range("D" & r).value <> 0 And _
               .Range("E" & r).value <> 0 Then
                    rs.Fields("Heures") = Format$(.Range("C" & r).value, "0.00")
                    rs.Fields("Taux") = Format$(.Range("D" & r).value, "0.00")
                    rs.Fields("Honoraires") = Format$(.Range("E" & r).value, "0.00")
            End If
            rs.Fields("Inv_Row") = wshFAC_Brouillon.Range("B11").value
        End With
    'Update the recordset (create the record)
    rs.update
    Next r
    
    'Create Summary By Rates lines
    Dim i As Long
    For i = 25 To 34
        If wshFAC_Brouillon.Range("R" & i).value <> "" And _
            wshFAC_Brouillon.Range("S" & i).value <> 0 Then
                rs.AddNew
                With wshFAC_Brouillon
                    rs.Fields("Inv_No") = noFacture
                    rs.Fields("Description") = "*** - [Sommaire des TEC] pour la facture - " & _
                                                wshFAC_Brouillon.Range("R" & i).value
                    rs.Fields("Heures") = CDbl(Format$(.Range("S" & i).value, "0.00"))
                    rs.Fields("Taux") = CDbl(Format$(.Range("T" & i).value, "0.00"))
                    rs.Fields("Honoraires") = CDbl(Format$(.Range("S" & i).value * .Range("T" & i).value, "0.00"))
                    rs.Fields("Inv_Row") = wshFAC_Brouillon.Range("B11").value
                End With
                rs.update
        End If
    Next i
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
nothing_to_update:

    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_to_DB()", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Details_Locally() '2024-03-11 @ 08:19 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_Locally (" & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy") & ")", 0)
    
    Application.ScreenUpdating = False
    
    'Get the last entered service
    Dim lastEnteredService As Long
    lastEnteredService = wshFAC_Finale.Range("B64").End(xlUp).Row
    If lastEnteredService < 34 Then GoTo nothing_to_update
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_D�tails.Range("A99999").End(xlUp).Row + 1
   
    Dim i As Long
    For i = 34 To lastEnteredService
        With wshFAC_D�tails
            .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
            .Range("B" & firstFreeRow).value = wshFAC_Finale.Range("B" & i).value
            .Range("C" & firstFreeRow).value = Format$(wshFAC_Finale.Range("C" & i).value, "0.00")
            .Range("D" & firstFreeRow).value = Format$(wshFAC_Finale.Range("D" & i).value, "0.00")
            .Range("E" & firstFreeRow).value = Format$(wshFAC_Finale.Range("E" & i).value, "0.00")
            .Range("F" & firstFreeRow).value = i
            firstFreeRow = firstFreeRow + 1
        End With
    Next i

nothing_to_update:
    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_Locally()", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Somm_Taux_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_to_DB - " & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    'Fees summary from wshFAC_Brouillon
    Dim firstRow As Long, lastRow As Long
    firstRow = 44
    lastRow = 48
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Sommaire_Taux"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").value
    Dim seq As Long
    Dim r As Long
    For r = firstRow To lastRow
        'Add fields to the recordset before updating it
        If wshFAC_Brouillon.Range("R" & r).value <> "" Then
            rs.AddNew
            With wshFAC_Finale
                rs.Fields("Inv_No") = noFacture
                rs.Fields("S�quence") = seq
                rs.Fields("Prof") = wshFAC_Brouillon.Range("R" & r).value
                rs.Fields("Heures") = wshFAC_Brouillon.Range("S" & r).value
                rs.Fields("Taux") = wshFAC_Brouillon.Range("T" & r).value
                seq = seq + 1
            End With
            'Update the recordset (create the record)
            rs.update
        End If
    Next r
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    conn.Close
    On Error GoTo 0
   
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_to_DB()", startTime)

End Sub

Sub FAC_Finale_Add_Invoice_Somm_Taux_Locally()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_Locally - " & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Fees summary from wshFAC_Brouillon
    Dim firstRow As Long, lastRow As Long
    firstRow = 44
    lastRow = 48
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Sommaire_Taux.Range("A99999").End(xlUp).Row + 1
   
    Dim noFacture As String
    noFacture = wshFAC_Finale.Range("E28").value
    Dim seq As Long
    Dim i As Long
    For i = firstRow To lastRow
        If wshFAC_Brouillon.Range("R" & i).value <> "" Then
            With wshFAC_Sommaire_Taux
                .Range("A" & firstFreeRow).value = noFacture
                .Range("B" & firstFreeRow).value = seq
                .Range("C" & firstFreeRow).value = wshFAC_Brouillon.Range("R" & i).value
                .Range("D" & firstFreeRow).value = CCur(wshFAC_Brouillon.Range("S" & i).value)
                .Range("D" & firstFreeRow).NumberFormat = "#,##0.00"
                .Range("E" & firstFreeRow).value = CCur(wshFAC_Brouillon.Range("T" & i).value)
                .Range("E" & firstFreeRow).NumberFormat = "#,##0.00"
                firstFreeRow = firstFreeRow + 1
                seq = seq + 1
            End With
        End If
    Next i

    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_Locally()", startTime)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB (" & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Comptes_Clients"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Create an empty recordset
    rs.Open "SELECT * FROM [" & destinationTab & "$] WHERE 1=0", conn, 2, 3
    
    'Add fields to the recordset before updating it
    rs.AddNew
    With wshFAC_Finale
        rs.Fields("Invoice_No") = .Range("E28").value
        rs.Fields("Invoice_Date") = Format$(CDate(wshFAC_Brouillon.Range("O3").value), "mm-dd-yyyy")
        rs.Fields("Customer") = .Range("B24").value
        rs.Fields("CodeClient") = wshFAC_Brouillon.Range("B18").value
        rs.Fields("Status") = "Unpaid"
        rs.Fields("Terms") = "Net 30"
        rs.Fields("Due_Date") = CDate(wshFAC_Brouillon.Range("O3").value) + 30
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
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB()", startTime)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_Locally() '2024-03-11 @ 08:49 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_Locally (" & _
        "# = " & wshFAC_Finale.Range("E28").value & " - Date = " & Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy"), 0)
    
    Application.ScreenUpdating = False
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Comptes_Clients.Range("A9999").End(xlUp).Row + 1
   
    With wshFAC_Comptes_Clients
        .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).value = wshFAC_Brouillon.Range("O3").value
        .Range("C" & firstFreeRow).value = wshFAC_Finale.Range("B24").value
        .Range("D" & firstFreeRow).value = wshFAC_Brouillon.Range("B18").value
        .Range("E" & firstFreeRow).value = "Unpaid"
        .Range("F" & firstFreeRow).value = "Net 30"
        .Range("G" & firstFreeRow).value = CDate(wshFAC_Brouillon.Range("O3").value) + 30
        .Range("H" & firstFreeRow).value = wshFAC_Finale.Range("E81").value
        .Range("I" & firstFreeRow).formula = ""
        .Range("J" & firstFreeRow).formula = "=G" & firstFreeRow & "-H" & firstFreeRow
        .Range("K" & firstFreeRow).formula = "=IF(H" & firstFreeRow & "<G" & firstFreeRow & ",NOW()-F" & firstFreeRow & ")"
    End With

nothing_to_update:

    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_Locally()", startTime)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_To_DB(firstRow As Long, lastRow As Long) 'Update Billed Status in DB

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB(" & firstRow & ", " & lastRow & ")", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim r As Long, TEC_ID As Long, SQL As String
    For r = firstRow To lastRow
        If wshTEC_Local.Range("BA" & r).value = True Or _
            wshFAC_Brouillon.Range("C" & r + 4) <> True Then
            GoTo next_iteration
        End If
        TEC_ID = wshTEC_Local.Range("AQ" & r).value
        
        'Open the recordset for the specified ID
        SQL = "SELECT * FROM [" & destinationTab & "$] WHERE TEC_ID=" & TEC_ID
        rs.Open SQL, conn, 2, 3
        If Not rs.EOF Then
            'Update DateSaisie, EstFacturee, DateFacturee & NoFacture
'            rs.Fields("DateSaisie").value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
            rs.Fields("EstFacturee").value = "VRAI"
            rs.Fields("DateFacturee").value = Format$(Now(), "dd/mm/yyyy hh:nn:ss")
            rs.Fields("VersionApp").value = ThisWorkbook.name
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
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB()", startTime)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_Locally(firstResultRow As Long, lastResultRow As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_Locally(" & firstResultRow & ", " & lastResultRow & ")", 0)
    
    'Set the range to look for
    Dim lastTECRow As Long
    lastTECRow = wshTEC_Local.Range("A99999").End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = wshTEC_Local.Range("A3:A" & lastTECRow)
    
    Dim r As Long, rowToBeUpdated As Long, TECID As Long
    For r = firstResultRow To lastResultRow
        If wshTEC_Local.Range("BA" & r).value = False And _
                wshFAC_Brouillon.Range("C" & r + 4) = True Then
            TECID = wshTEC_Local.Range("AQ" & r).value
            rowToBeUpdated = Fn_Find_Row_Number_TEC_ID(TECID, lookupRange)
'            wshTEC_Local.Range("K" & rowToBeUpdated).value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
            wshTEC_Local.Range("L" & rowToBeUpdated).value = "VRAI"
            wshTEC_Local.Range("M" & rowToBeUpdated).value = Format$(Now(), "dd/mm/yyyy hh:mm:ss")
            wshTEC_Local.Range("O" & rowToBeUpdated).value = ThisWorkbook.name
            wshTEC_Local.Range("P" & rowToBeUpdated).value = wshFAC_Brouillon.Range("O6").value
        End If
    Next r
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set lookupRange = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_Locally()", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_D�tails_To_DB(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_D�tails_To_DB(" & projetID & ")", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_D�tails"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Build the query
    Dim strSQL As String
    strSQL = "UPDATE [" & destinationTab & "$] SET estDetruite = -1 WHERE projetID = " & projetID
    
    'Execute the SQL query
    conn.Execute strSQL
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_D�tails_To_DB()", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_D�tails_Locally(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_D�tails_Locally(" & projetID & ")", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_D�tails
    
    Dim projetIDColumn As String, isD�truiteColumn As String
    projetIDColumn = "A"
    isD�truiteColumn = "I"

    'Find the last used row
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    
    'Use Range.Find to locate the first cell with the projetID
    Dim cell As Range
    Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).Find(What:=projetID, LookIn:=xlValues, LookAt:=xlWhole)

    'Check if the projetID was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Do
            'Update the isD�truite column for the found projetID
            ws.Cells(cell.Row, isD�truiteColumn).value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_D�tails_Locally()", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Ent�te_To_DB(projetID)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Ent�te_To_DB(" & projetID & ")", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Ent�te"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Build the query
    Dim strSQL As String
    strSQL = "UPDATE [" & destinationTab & "$] SET estD�truite = True WHERE ProjetID = " & projetID

    'Execute the SQL query
    On Error GoTo eh
    conn.Execute strSQL
    On Error GoTo 0
    
    'Close recordset and connection
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-23 @ 15:32
    Set conn = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Ent�te_To_DB()", startTime)
    Exit Sub

eh:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If

End Sub

Sub FAC_Finale_Softdelete_Projets_Ent�te_Locally(projetID)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Ent�te_Locally(" & projetID & ")", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Ent�te
    
    Dim projetIDColumn As String, isD�truiteColumn As String
    projetIDColumn = "A"
    isD�truiteColumn = "Z"

    'Find the last used row
    Dim lastUsedRow As Long
    lastUsedRow = ws.Range("A99999").End(xlUp).Row
    
    'Use Range.Find to locate the first cell with the projetID
    Dim cell As Range
    Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).Find(What:=projetID, LookIn:=xlValues, LookAt:=xlWhole)

    'Check if the projetID was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Do
            'Update the isD�truite column for the found projetID
            ws.Cells(cell.Row, isD�truiteColumn).value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Ent�te_Locally()", startTime)

End Sub

'Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:Invoice_Load", 0)
'
'    With wshFAC_Brouillon
'        If wshFAC_Brouillon.Range("B20").value = Empty Then
'            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un num�ro de facture VALIDE pour votre recherche"
'            GoTo NoItems
'        End If
'        'Could that invoice been cancelled (more than 1 row) ?
'        Call InvoiceGetAllTrans(wshFAC_Brouillon.Range("O6").value)
'        Dim NbTrans As Long
'        NbTrans = .Range("B31").value
'        If NbTrans = 0 Then
'            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un num�ro de facture VALIDE pour votre recherche"
'            GoTo NoItems
'        Else
'            If NbTrans > 1 Then
'                MsgBox "Cette facture a �t� annul�e! Veuillez saisir un num�ro de facture VALIDE pour votre recherche"
'                GoTo NoItems
'            End If
'        End If
'        .Range("B24").value = True 'Set Invoice Load to true
'        .Range("S2,E4:F4,K4:L6,O3,K11:O45,Q11:Q45").ClearContents
'        wshFAC_Finale.Range("B34:F68").ClearContents
'        Dim InvListRow As Long
'        InvListRow = wshFAC_Brouillon.Range("B20").value 'InvListRow = Row associated with the invoice
'        'Get values from wshFAC_Ent�te (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
'        .Range("O3").value = wshFAC_Ent�te.Range("B" & InvListRow).value
'        .Range("K3").value = wshFAC_Ent�te.Range("D" & InvListRow).value
'        .Range("K4").value = wshFAC_Ent�te.Range("E" & InvListRow).value
'        .Range("K5").value = wshFAC_Ent�te.Range("F" & InvListRow).value
'        .Range("K6").value = wshFAC_Ent�te.Range("G" & InvListRow).value
'        'Get values from wshFAC_Ent�te (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
'        Dim dFact As Date
'        dFact = wshFAC_Ent�te.Range("B" & InvListRow).value
'        wshFAC_Finale.Range("B21").value = "Le " & Format$(dFact, "d") & " " & _
'                                            UCase(Format$(dFact, "mmmm")) & " " & _
'                                            Format$(dFact, "yyyy")
'        wshFAC_Finale.Range("B23").value = wshFAC_Ent�te.Range("D" & InvListRow).value
'        wshFAC_Finale.Range("B24").value = wshFAC_Ent�te.Range("E" & InvListRow).value
'        wshFAC_Finale.Range("B25").value = wshFAC_Ent�te.Range("F" & InvListRow).value
'        wshFAC_Finale.Range("B26").value = wshFAC_Ent�te.Range("G" & InvListRow).value
'        'Load Invoice Detail Items
'        With wshFAC_D�tails
'            Dim lastRow As Long, lastResultRow As Long
'            lastRow = .Range("A999999").End(xlUp).Row
'            If lastRow < 4 Then Exit Sub 'No Item Lines
'            .Range("I3").value = wshFAC_Brouillon.Range("O6").value
'            wshFAC_Finale.Range("F28").value = wshFAC_Brouillon.Range("O6").value 'Invoice #
'            'Advanced Filter to get items specific to ONE invoice
'            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, criteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
'            lastResultRow = .Range("O999").End(xlUp).Row
'            If lastResultRow < 3 Then GoTo NoItems
'            For resultRow = 3 To lastResultRow
'                invitemRow = .Range("O" & resultRow).value
'                wshFAC_Brouillon.Range("L" & invitemRow & ":O" & invitemRow).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
'                wshFAC_Brouillon.Range("Q" & invitemRow).value = .Range("P" & resultRow).value  'Set Item DB Row
'                wshFAC_Finale.Range("C" & invitemRow + 23 & ":F" & invitemRow + 23).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
'            Next resultRow
'        End With
'        'Proceed with trailer data (Misc. charges & Taxes)
'        .Range("M48").value = wshFAC_Ent�te.Range("I" & InvListRow).value
'        .Range("O48").value = wshFAC_Ent�te.Range("J" & InvListRow).value
'        .Range("M49").value = wshFAC_Ent�te.Range("K" & InvListRow).value
'        .Range("O49").value = wshFAC_Ent�te.Range("L" & InvListRow).value
'        .Range("M50").value = wshFAC_Ent�te.Range("M" & InvListRow).value
'        .Range("O50").value = wshFAC_Ent�te.Range("N" & InvListRow).value
'        .Range("O52").value = wshFAC_Ent�te.Range("P" & InvListRow).value
'        .Range("O53").value = wshFAC_Ent�te.Range("R" & InvListRow).value
'        .Range("O57").value = wshFAC_Ent�te.Range("T" & InvListRow).value
'
'NoItems:
'    .Range("B24").value = False 'Set Invoice Load To false
'    End With
'
'    Call Log_Record("modFAC_Finale:Invoice_Load()", startTime)
'
'End Sub

'Fonction pour v�rifier si un nom de feuille existe d�j� dans un classeur
Function NomFeuilleExiste(nom As String, wb As Workbook) As Boolean
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(nom)
    NomFeuilleExiste = Not ws Is Nothing
    On Error GoTo 0
    
End Function

'Sub InvoiceGetAllTrans(inv As String)
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:InvoiceGetAllTrans", 0)
'
'    Application.ScreenUpdating = False
'
'    wshFAC_Brouillon.Range("B31").value = 0
'
'    With wshFAC_Ent�te
'        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
'        lastRow = .Range("A999999").End(xlUp).Row 'Last wshFAC_Ent�te Row
'        If lastRow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
'        On Error Resume Next
'        .Names("Criterial").Delete
'        On Error GoTo 0
'        .Range("V3").value = wshFAC_Brouillon.Range("O6").value
'        'Advanced Filter setup
'        .Range("A3:T" & lastRow).AdvancedFilter xlFilterCopy, _
'            criteriaRange:=.Range("V2:V3"), _
'            CopyToRange:=.Range("X2:AQ2"), _
'            Unique:=True
'        lastResultRow = .Range("X999").End(xlUp).Row 'How many rows trans for that invoice
'        If lastResultRow < 3 Then
'            GoTo Done
'        End If
''        With .Sort
''            .SortFields.Clear
''            .SortFields.Add Key:=wshFAC_Ent�te.Range("X2"), _
''                SortOn:=xlSortOnValues, _
''                Order:=xlAscending, _
''                DataOption:=xlSortNormal 'Sort Based Invoice Number
''            .SortFields.Add Key:=wshGL_Trans.Range("Y3"), _
''                SortOn:=xlSortOnValues, _
''                Order:=xlAscending, _
''                DataOption:=xlSortNormal 'Sort Based On TEC_ID
''            .SetRange wshFAC_Ent�te.Range("X2:AQ" & lastResultRow) 'Set Range
''            .Apply 'Apply Sort
''         End With
'         wshFAC_Brouillon.Range("B31").value = lastResultRow - 2 'Remove Header rows from row count
'Done:
'    End With
'    Application.ScreenUpdating = True
'
'    Call Log_Record("modFAC_Finale:InvoiceGetAllTrans()", startTime)
'
'End Sub

Sub FAC_Finale_Setup_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Setup_All_Cells", 0)
    
    Application.EnableEvents = False
     
    With wshFAC_Finale
'        Dim j As String, m As String, y As String
'        j = Format$(FAC_Brouillon!O3, "j")
'        m = UCase(Format$(FAC_Brouillon!O3, "mmm"))
'        y = Format$(FAC_Brouillon!O3, "yyyy")
        
        .Range("B21").formula = "= ""Le "" & DAY(FAC_Brouillon!o3) & "" "" & UPPER(TEXT(FAC_Brouillon!O3, ""mmmm"")) & "" "" & YEAR(FAC_Brouillon!O3)"
'        .Range("B21").formula = "= ""Le "" & TEXT(FAC_Brouillon!O3, ""j mmmm aaaa"")"
        .Range("B23:B27").value = ""
        .Range("E28").value = "=" & wshFAC_Brouillon.name & "!O6"    'Invoice number
        
'        .Range("C65").value = "Heures"                               'Summary Heading
'        .Range("D65").value = "Taux"                                 'Summary Heading
'        .Range("C66").formula = "=" & wshFAC_Brouillon.name & "!M47" 'Hours summary
'        .Range("D66").formula = "=" & wshFAC_Brouillon.name & "!N47" 'Hourly Rate
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
        .Range("E69").formula = "=" & wshFAC_Brouillon.name & "!O47" 'Fees Sub-Total
        
        .Range("B70").formula = "=" & wshFAC_Brouillon.name & "!M48" 'Misc. Amount # 1 - Description
        .Range("E70").formula = "=" & wshFAC_Brouillon.name & "!O48" 'Misc. Amount # 1
        
        .Range("B71").formula = "=" & wshFAC_Brouillon.name & "!M49" 'Misc. Amount # 2 - Description
        .Range("E71").formula = "=" & wshFAC_Brouillon.name & "!O49" 'Misc. Amount # 2
        
        .Range("B72").formula = "=" & wshFAC_Brouillon.name & "!M50" 'Misc. Amount # 3 - Description
        .Range("E72").formula = "=" & wshFAC_Brouillon.name & "!O50" 'Misc. Amount # 3
        
        .Range("E73").formula = "=SUM(E69:E72)"                      'Invoice Sub-Total
        
        .Range("C74").formula = "=" & wshFAC_Brouillon.name & "!N52" 'GST Rate
        .Range("E74").formula = "=round(E73*C74,2)"                  'GST Amount"
        .Range("C75").formula = "=" & wshFAC_Brouillon.name & "!N53" 'PST Rate
        .Range("E75").formula = "=round(E73*C75,2)"                  'PST Amount
        
        .Range("E77").formula = "=SUM(E73:E75)"                        'Total including taxes
        .Range("E79").formula = "=" & wshFAC_Brouillon.name & "!O57" 'Deposit Amount
        .Range("E81").formula = "=E77-E79"                             'Total due on that invoice
    End With
    
    Application.EnableEvents = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Setup_All_Cells()", startTime)

End Sub

Sub FAC_Finale_Preview_PDF() '2024-03-02 @ 16:18

    wshFAC_Finale.PrintOut , , 1, True, True, , , , False
'    wshFAC_Finale.PrintOut , , , True, True, , , , False

End Sub

Sub FAC_Finale_Bouton_Creation_PDF() '2024-10-13 @ 10:15
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Bouton_Creation_PDF - '" & wshFAC_Finale.Range("E28").value & "'", 0)
    
    flagEtapeFacture = 1
    
    '�tape 1 - Cr�ation du document PDF
    Call FAC_Finale_Create_PDF(wshFAC_Finale.Range("E28").value)
    
    '�tape 2 - Copie de la facture en format EXCEL
    Call FAC_Finale_Copie_Vers_Excel(wshFAC_Brouillon.Range("B18").value, _
                                          wshFAC_Finale.Range("L81").value, _
                                          wshFAC_Finale.Range("E28").value, _
                                          Format$(wshFAC_Brouillon.Range("O3").value, "yyyy-mm-dd"))
    flagEtapeFacture = 3
    
    '�tape 3 - Envoi du courriel
    Call FAC_Finale_Creation_Courriel(wshFAC_Finale.Range("E28").value, wshFAC_Brouillon.Range("B18").value)
    flagEtapeFacture = 4
    
    '�tape 4 - Activation du bouton SAUVEGARDE
    Call FAC_Finale_Enable_Save_Button
    flagEtapeFacture = 5

    Call Log_Record("modFAC_Finale:FAC_Finale_Bouton_Creation_PDF", startTime)

End Sub

Sub FAC_Finale_Create_PDF(noFacture As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF(" & noFacture & ")", 0)
    
    'Cr�ation du fichier (NoFacture).PDF dans le r�pertoire de factures PDF de GCF
    Dim result As Boolean
    result = FAC_Finale_Create_PDF_Func(noFacture, "SaveOnly")
    
    If result = False Then
        MsgBox "ATTENTION... Impossible de sauvegarder la facture en format PDF", _
                vbCritical, _
                "Impossible de sauvegarder la facture en format PDF"
        flagEtapeFacture = -1
    End If

    Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF", startTime)

End Sub

Function FAC_Finale_Create_PDF_Func(noFacture As String, Optional action As String = "SaveOnly") As Boolean
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF_Func(" & noFacture & ", " & action & ")", 0)
    
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
        reponse = MsgBox("La facture (PDF) num�ro '" & noFacture & "' existe d�ja." & _
                          "Voulez-vous la remplacer ?", vbYesNo + vbQuestion, _
                          "Cette facture existe d�j� en formt PDF")
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
    MsgBox "Incapable de pr�parer le courriel. La librairie n'est pas disponible"
    FAC_Finale_Create_PDF_Func = False 'Function return value
'    FAC_Finale_Create_Email = False 'Function return value

EndMacro:
    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Create_PDF_Func", startTime)

End Function

Sub FAC_Finale_Copie_Vers_Excel(clientID As String, clientName As String, invNo As String, invDate As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Copie_Vers_Excel(" & _
        clientID & ", " & clientName & ", " & invNo & ", " & invDate & ")", 0)
    
    Dim clientNamePurged As String
    clientNamePurged = clientName
    
    Application.ScreenUpdating = False
    
    'Purge le nom du client
    Do While InStr(clientNamePurged, "[") > 0 And InStr(clientNamePurged, "]") > 0
        clientNamePurged = Fn_Strip_Contact_From_Client_Name(clientNamePurged)
    Loop
    
    'D�finir le chemin complet du r�pertoire des fichiers Excel
    Dim ExcelFilesFullPath As String
    ExcelFilesFullPath = wshAdmin.Range("F5").value & FACT_EXCEL_PATH
    ChDir ExcelFilesFullPath
    
    'D�finir la feuille source et la plage � copier
    Dim wbSource As Workbook: Set wbSource = ThisWorkbook
    Dim wsSource As Worksheet: Set wsSource = wshFAC_Finale
    Dim plageSource As Range: Set plageSource = wsSource.Range("A1:F88")

    'D�sactiver les �v�nements pour �viter Workbook_Activate
    Application.EnableEvents = False
    
    'Ouvrir un nouveau Workbook (ou choisir un workbook existant)
    On Error Resume Next
    Dim strCible As String
    strCible = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx") 'S�lectionner un classeur cible
    On Error GoTo 0
    
    'Si l'utilisateur annule la s�lection du fichier ou il y a une erreur
    Dim wbCible As Workbook
    If strCible = "Faux" Or strCible = "False" Or strCible = "" Then
        'Cr�er un nouveau workbook
        Set wbCible = Workbooks.Add
        strCible = ""
    Else
        'Ouvrir le workbook s�lectionn�
        Set wbCible = Workbooks.Open(strCible)
    End If
    
'    Set wsCible = wbCible.Sheets.Add(After:=wbCible.Sheets(wbCible.Sheets.count))
    Dim strName As String
    Dim strNameBase As String
    strNameBase = invDate & " - " & invNo
    strName = strNameBase
    
    'On v�rifie si le nom de la nouvelle feuille � ajouter existe d�j�
    Dim wsExist As Boolean
    wsExist = False
    On Error Resume Next
    wsExist = Not wbCible.Worksheets(strNameBase) Is Nothing
    On Error GoTo 0
    
    'Si le worksheet existe d�j� avec ce nom, demander � l'utilisateur ce qu'il souhaite faire
    Dim wsCible As Worksheet
    Dim suffixe As Integer
    Dim reponse As String
    
    If wsExist Then
        reponse = MsgBox("La feuille '" & strNameBase & "' existe d�j� dans ce fichier" & vbCrLf & vbCrLf & _
                         "Voulez-vous :" & vbCrLf & vbCrLf & _
                         "1. Remplacer l'onglet existant par la facture courante ?" & vbCrLf & vbCrLf & _
                         "2. Cr�er un nouvel onglet avec un suffixe ?" & vbCrLf & vbCrLf & _
                         "Cliquez sur Oui pour remplacer, ou sur Non pour cr�er un nouvel onglet.", _
                         vbYesNoCancel + vbQuestion, "Le nouvel onglet � cr�er existe d�j�")

        Select Case reponse
            Case vbYes 'Remplacer l'onglet existant
                Application.DisplayAlerts = False ' D�sactiver les alertes pour �craser sans confirmation
                wbCible.Worksheets(strNameBase).Delete
                Application.DisplayAlerts = True
                
                'Cr�er une nouvelle feuille avec le m�me nom
                Set wsCible = wbCible.Worksheets.Add(After:=wbCible.Sheets(wbCible.Sheets.count))
                wsCible.name = strNameBase 'Attribuer le nom d'origine

            Case vbNo 'L'utilisateur souhaite cr�er une nouvelle feuille
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
                
                'Cr�er une nouvelle feuille avec ce nom directement lors de la cr�ation
                Application.DisplayAlerts = False ' D�sactiver les alertes pour �viter Feuil1
                Set wsCible = wbCible.Worksheets.Add(After:=wbCible.Sheets(wbCible.Sheets.count))
                wsCible.name = strName ' Attribuer le nouveau nom avec suffixe
                Application.DisplayAlerts = True ' R�activer les alertes apr�s la cr�ation
        End Select
    Else
        'Si la feuille n'existe pas, on peut directement la cr�er
        Set wsCible = wbCible.Worksheets.Add(After:=wbCible.Sheets(wbCible.Sheets.count))
        wsCible.name = strNameBase
    End If
    
'    wsCible.name = strName 'Renommer la nouvelle feuille
    
    '1. Copier les valeurs uniquement
    plageSource.Copy
    wsCible.Range("A1").PasteSpecial Paste:=xlPasteValues

    '2. Copier les formats de cellules
    plageSource.Copy
    wsCible.Range("A1").PasteSpecial Paste:=xlPasteFormats

    '3. Conserver la taille des colonnes
    Dim i As Integer
    For i = 1 To plageSource.columns.count
        wsCible.columns(i).ColumnWidth = plageSource.columns(i).ColumnWidth
    Next i

    '4. Ajuster les hauteurs de lignes (optionnel si n�cessaire)
    For i = 1 To plageSource.rows.count
        wsCible.rows(i).RowHeight = plageSource.rows(i).RowHeight
    Next i

    '5. Copier l'ent�te de la facture
    Dim forme As Shape
    For Each forme In wsSource.Shapes
        If forme.name = "GCF_Ent�te" Then
            forme.Copy
            wsCible.Paste
            'Ajuster la position et la taille de la forme
            With wsCible.Shapes(wsCible.Shapes.count)
                .Top = forme.Top
                .Left = forme.Left
                .Height = 255.15
            End With
        End If
    Next forme

    '6. Copier les param�tres d'impression
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
    
    'D�sactiver le mode copier-coller pour lib�rer la m�moire
    Application.CutCopyMode = False
    
    'Optionnel : Sauvegarder le workbook cible sous un nouveau nom si n�cessaire
    If strCible = "" Then
        wbCible.SaveAs ExcelFilesFullPath & Application.PathSeparator & clientID & " - " & clientNamePurged & ".xlsx"
        MsgBox "Un nouveau fichier Excel (" & clientID & " - " & clientNamePurged & ".xlsx" & ")" & vbNewLine & vbNewLine & _
                "A �t� cr�� pour sauvegarder la facture", vbInformation
    End If
    
    'R�activer les �v�nements apr�s l'ouverture
    Application.EnableEvents = True
    
    'La facture a �t� sauvegard� en format EXCEL
    flagEtapeFacture = 3
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Copie_Vers_Excel", startTime)

End Sub

Sub FAC_Finale_Creation_Courriel(noFacture As String, clientID As String) '2024-10-13 @ 11:33

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Creation_Courriel(" & _
        noFacture & ", " & clientID & ")", 0)
    
    Dim fileExists As Boolean
    
    '1a. Chemin de la pi�ce jointe (Facture en format PDF)
    Dim attachmentFullPathName As String
    attachmentFullPathName = wshAdmin.Range("F5").value & FACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2024-09-03 @ 16:43
    
    '1b. V�rification de l'existence de la pi�ce jointe
    fileExists = Dir(attachmentFullPathName) <> ""
    If Not fileExists Then
        MsgBox "La pi�ce jointe (Facture en format PDF) n'existe pas" & _
                    "� l'emplacement sp�cifi�, soit " & attachmentFullPathName, vbCritical
        GoTo Exit_sub
    End If
    
    '2a. Chemin du template (.oft) de courriel
    Dim templateFullPathName As String
    templateFullPathName = Environ("appdata") & "\Microsoft\Templates\GCF_Facturation.oft"

    '2b. V�rification de l'existence du template
    fileExists = Dir(templateFullPathName) <> ""
    If Not fileExists Then
        MsgBox "Le gabarit 'GCF_Facturation.oft' est introuvable " & _
                    "� l'emplacement sp�cifi�, soit " & Environ("appdata") & "\Microsoft\Templates", _
                    vbCritical
        GoTo Exit_sub
    End If
    
    '3. Initialisation de l'application Outlook
    Dim OutlookApp As Object
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    '4. Cr�ation de l'email � partir du template
    Dim MailItem As Object
    Set MailItem = OutlookApp.CreateItemFromTemplate(templateFullPathName)

    '5. Ajout de la pi�ce jointe
    MailItem.Attachments.Add attachmentFullPathName

'    'Obtenir la signature par d�faut
'    Dim signaturePath As String
'    signaturePath = Fn_Get_Outlook_Signature_Path()
'    Dim Signature As String
'    Signature = Fn_Get_Outlook_Signature()
'
        
    '6. Obtenir l'adresse courriel pour le client
    Dim ws As Worksheet: Set ws = wshBD_Clients
    Dim eMailFacturation As String
    eMailFacturation = Fn_Get_Value_From_UniqueID(ws, clientID, 2, 5)
    If eMailFacturation = "uniqueID introuvable" Then
        MailItem.To = ""
    Else
        If Fn_Valider_Courriel(eMailFacturation) = True Then
            MailItem.To = eMailFacturation
        Else
            MsgBox "Je ne peux utiliser l'adresse courriel de ce client" & vbNewLine & vbNewLine & _
                    "soit '" & eMailFacturation & "' !", vbExclamation
            MailItem.To = ""
        End If
    End If
    
'    'Optionnel : Modifiez les �l�ments de l'email (comme les destinataires)
'    MailItem.To = "robertv13@me.com"

    'Afficher (.Display) ou envoyer (.Send) le courriel
    MailItem.Display
    'MailItem.Send 'Pour envoyer directement l'email

Exit_sub:
    'Nettoyage
    Set MailItem = Nothing
    Set OutlookApp = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Creation_Courriel", startTime)

End Sub

Sub Test_FAC_Finale_Creation_Courriel()

    Call FAC_Finale_Creation_Courriel("24-24524", "1793")

End Sub

'Function Fn_Get_Outlook_Signature_Path() As String
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Creation_Courriel(" & _
'        noFacture & ", " & clientID & ")", 0)
'
'    Dim sPath As String
'    Dim sSignaturePath As String
'    Dim fso As Object
'    Dim file As Object
'
'    'Chemin vers le dossier des signatures
'    sPath = Environ("appdata") & "\Microsoft\Signatures\"
'
'    'D�finir l'objet FileSystem
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    'Trouver le fichier de signature HTML
'    For Each file In fso.GetFolder(sPath).Files
'        If LCase(fso.GetExtensionName(file.name)) = "htm" Then
'            sSignaturePath = file.path
'            Exit For
'        End If
'    Next
'
'    'Retourner le chemin de la signature
'    Fn_Get_Outlook_Signature_Path = sSignaturePath
'
'    Set fso = Nothing
'    Set file = Nothing
'
'End Function

'Function Fn_Get_Outlook_Signature() As String
'
'    'Chemin vers le dossier des signatures
'    Dim sPath As String
'    sPath = Environ("appdata") & "\Microsoft\Signatures\"
'
'    'D�finir l'objet FileSystem
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    'Lire la signature par d�faut dans le fichier .htm
'    Dim sSignature As String
'    Dim ts As Object
'    If fso.fileExists(sPath & Dir(sPath & "*.htm")) Then
'        Set ts = fso.GetFile(sPath & Dir(sPath & "*.htm")).OpenAsTextStream(1, -2)
'        sSignature = ts.ReadAll
'        ts.Close
'    End If
'
'    'Retourner la signature
'    Fn_Get_Outlook_Signature = sSignature
'
'    Set fso = Nothing
'    Set ts = Nothing
'
'End Function

'Function ReadSignatureFile(filepath As String) As String
'    Dim fso As Object
'    Dim ts As Object
'    Dim sContent As String
'
'    ' D�finir l'objet FileSystem
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Lire le contenu de la signature
'    If fso.fileExists(filepath) Then
'        Set ts = fso.GetFile(filepath).OpenAsTextStream(1, -2)
'        sContent = ts.ReadAll
'        ts.Close
'    End If
'
'    ' Retourner le contenu de la signature
'    ReadSignatureFile = sContent
'
'    Set fso = Nothing
'    Set ts = Nothing
'End Function
'
'Sub EnvoyerEmailHTMLAvecSignature()
'
'    Dim OutlookApp As Object
'    Dim MailItem As Object
'    Dim Destinataire As String
'    Dim Objet As String
'    Dim CorpsMessage As String
'    Dim Signature As String
'    Dim PieceJointe As String
'    Dim CheminImageSignature As String
'
'    ' Initialiser l'application Outlook
'    On Error Resume Next
'    Set OutlookApp = GetObject(Class:="Outlook.Application")
'    If OutlookApp Is Nothing Then
'        Set OutlookApp = CreateObject(Class:="Outlook.Application")
'    End If
'    On Error GoTo 0
'
'    ' Cr�er un nouvel e-mail
'    Set MailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
'
'    ' D�finir les param�tres de l'e-mail
'    Destinataire = "destinataire@exemple.com"
'    Objet = "Objet de votre e-mail"
'    PieceJointe = "C:\VBA\GC_FISCALIT�\Factures_PDF\24-24524.pdf"
'
'    'Obtenir la signature par d�faut
'    Dim signaturePath As String
'    signaturePath = Fn_Get_Outlook_Signature_Path()
'    Dim Signature As String
'    Signature = Fn_Get_Outlook_Signature()
'
'    ' Construire le corps de l'e-mail avec HTML
'    CorpsMessage = "<p style='text-indent: 40px;'>Bonjour,</p>" & _
'                   "<p style='text-indent: 40px;'>Veuillez trouver ci-joint le document requis concernant votre demande.</p>" & _
'                   "<p style='text-indent: 40px;'>Cordialement,</p>"
'
'    ' Signature en HTML avec image
'    Signature = "<br><img src='cid:SignatureImage' width='100' height='100'><br>" & _
'                "<b>Votre Nom</b><br>" & _
'                "Votre Fonction<br>" & _
'                "Votre Soci�t�<br>" & _
'                "T�l�phone : +33 1 23 45 67 89<br>" & _
'                "Email : <a href='mailto:Votre.Email@exemple.com'>Votre.Email@exemple.com</a><br>" & _
'                "Adresse compl�te de l'entreprise"
'
'    ' Construire l'e-mail avec HTML et la pi�ce jointe
'    With MailItem
'        .To = Destinataire
'        .Subject = Objet
'        .HTMLBody = CorpsMessage & Signature
'
'        ' Ajouter une pi�ce jointe si elle existe
'        If Dir(PieceJointe) <> "" Then
'            .Attachments.Add PieceJointe
'        Else
'            MsgBox "La pi�ce jointe n'a pas �t� trouv�e.", vbExclamation
'            Exit Sub
'        End If
'
'        ' Ajouter l'image de la signature dans l'e-mail
'        If Dir(CheminImageSignature) <> "" Then
'            .Attachments.Add CheminImageSignature, 1, 0, "SignatureImage" ' L'argument "cid" est utilis� ici
'        Else
'            MsgBox "L'image de la signature n'a pas �t� trouv�e.", vbExclamation
'            Exit Sub
'        End If
'
'        .Display ' Utilisez .Send pour envoyer directement
'    End With
'
'    ' Lib�rer les objets
'    Set MailItem = Nothing
'    Set OutlookApp = Nothing
'
'End Sub
'
'Sub EnvoyerEmailAvecPieceJointe()
'
'    Dim OutlookApp As Object
'    Dim MailItem As Object
'    Dim Destinataire As String
'    Dim Objet As String
'    Dim CorpsMessage As String
'    Dim PieceJointe As String
'    Dim Signature As String
'
'    'Initialiser l'application Outlook
'    On Error Resume Next
'    Set OutlookApp = GetObject(Class:="Outlook.Application")
'    If OutlookApp Is Nothing Then
'        Set OutlookApp = CreateObject(Class:="Outlook.Application")
'    End If
'    On Error GoTo 0
'
'    'Cr�er un nouvel e-mail
'    Set MailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
'
'    'D�finir les param�tres de l'e-mail
'    Destinataire = "robertv13@me.com"
'    Objet = "GC FISCALIT� - Facturation"
'    CorpsMessage = "Bonjour," & vbCrLf & vbCrLf & _
'                   "Veuillez trouver ci-joint le document requis." & vbCrLf & vbCrLf & _
'                   "Cordialement," & vbCrLf & _
'                   "Votre Nom"
'
'    PieceJointe = "C:\VBA\GC_FISCALIT�\Factures_PDF\24-24524.pdf"
'
'    'Signature sp�cifique (vous pouvez ajuster cela ou utiliser la signature par d�faut d'Outlook)
'    Signature = vbCrLf & "Votre Nom" & vbCrLf & _
'                "Votre Fonction" & vbCrLf & _
'                "Votre Soci�t�" & vbCrLf & _
'                "T�l�phone : +33 1 23 45 67 89" & vbCrLf & _
'                "Email : votre.email@exemple.com"
'
'    'Construire l'e-mail
'    With MailItem
'        .To = Destinataire
'        .Subject = Objet
'        .Body = CorpsMessage & vbCrLf & vbCrLf & Signature
'        If Dir(PieceJointe) <> "" Then ' V�rifier si la pi�ce jointe existe
'            .Attachments.Add PieceJointe
'        Else
'            MsgBox "La pi�ce jointe n'a pas �t� trouv�e.", vbExclamation
'            Exit Sub
'        End If
'        .Display ' Utilisez .Send pour envoyer directement
'    End With
'
'    ' Lib�rer les objets
'    Set MailItem = Nothing
'    Set OutlookApp = Nothing
'
'End Sub
'
'Sub FAC_Finale_Create_Email(noFacture As String)
'
'    Application.ScreenUpdating = False
'
'    'Construct the attachmentFullPathName filename
'    Dim attachmentFullPathName As String
'    attachmentFullPathName = wshAdmin.Range("F5").value & FACT_PDF_PATH & Application.PathSeparator & _
'                     noFacture & ".pdf" '2024-09-03 @ 16:43
'
'    'Check if the file already exists
'    Dim fileExists As Boolean
'    fileExists = Dir(attachmentFullPathName) <> ""
'
'    'If the file exists, prompt the user for confirmation
'    Dim reponse As VbMsgBoxResult
'    If Not fileExists Then
'        reponse = MsgBox("La facture '" & noFacture & "' n'existe pas en format PDF" & _
'                          "Je ne peux donc pas l'envoyer par courriel !", vbInformation, _
'                          "Cette facture n'existe pas en format PDF")
'        GoTo EndMacro
'    End If
'
'    'Set Print Quality
'    On Error Resume Next
'    ActiveSheet.PageSetup.PrintQuality = 600
'    Err.Clear
'    On Error GoTo 0
'
'    'Adjust Document Properties - 2024-09-03 @ 16:46
'    With ActiveSheet.PageSetup
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0)
'        .BottomMargin = Application.InchesToPoints(0)
'    End With
'
'    'Construct & Display the Email, allowing the user to modify the Email
'    On Error GoTo SaveOnly
'
'    Dim OutlookApp As Outlook.Application: Set OutlookApp = New Outlook.Application
'
'    'Where are the email templates ? - 2024-03-27 @ 07:28
'    Dim FullTemplatePathAndFile As String
'    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
'        FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\GCF_Facturation.oft"
'    Else
'        FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\GCF_Facturation.oft"
'    End If
'
'    Dim myMail As Outlook.MailItem: Set myMail = OutlookApp.CreateItemFromTemplate(FullTemplatePathAndFile)
''        Set myMail = outlookApp.CreateItem(olMailItem)
'
'    With myMail
'        .To = "robertv13@me.com"
'        .Subject = "GC FISCALIT� INC. - Facturation"
'        .Attachments.Add attachmentFullPathName
'        .Display 'Affiche le courriel, ce qui permet de corriger AVANT l'envoi
'
'        'myMail.Send
'    End With
'
'SaveOnly:
'    GoTo EndMacro
'
'RefLibError:
'    MsgBox "Incapable de pr�parer le courriel. La librairie n'est pas disponible"
'
'EndMacro:
'    Application.ScreenUpdating = True
'
'    'Cleaning memory - 2024-07-01 @ 09:34
'    Set myMail = Nothing
'    Set OutlookApp = Nothing
'
'End Sub
'
'Sub Prev_Invoice() 'TO-DO-RMV 2023-12-17
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:Prev_Invoice", 0)
'
'    With wshFAC_Brouillon
'        Dim MininvNumb As Long
'        On Error Resume Next
'        MininvNumb = Application.WorksheetFunction.Min(wshFAC_Ent�te.Range("Inv_ID"))
'        On Error GoTo 0
'        If MininvNumb = 0 Then
'            MsgBox "Please create and save an Invoice first"
'            Exit Sub
'        End If
'        invNumb = .Range("N6").value
'        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
'            invRow = wshFAC_Ent�te.Range("A99999").End(xlUp).Row 'On Empty Invoice Go to last one created
'        Else 'On Existing Inv. find Previous one
'            invRow = wshFAC_Ent�te.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).Row - 1
'        End If
'        If .Range("N6").value = 1 Or MininvNumb = 0 Or MininvNumb = .Range("N6").value Then
'            MsgBox "You are at the first invoice"
'            Exit Sub
'        End If
'        .Range("N3").value = wshFAC_Ent�te.Range("A" & invRow).value 'Place Inv. ID inside cell
'        Invoice_Load
'    End With
'
'    Call Log_Record("modFAC_Finale:Prev_Invoice()", startTime)
'
'End Sub

'Sub Next_Invoice() 'TO-DO-RMV 2023-12-17
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:Next_Invoice", 0)
'
'    With wshFAC_Brouillon
'        Dim MaxinvNumb As Long
'        On Error Resume Next
'        MaxinvNumb = Application.WorksheetFunction.Max(wshFAC_Ent�te.Range("Inv_ID"))
'        On Error GoTo 0
'        If MaxinvNumb = 0 Then
'            MsgBox "Please create and save an Invoice first"
'            Exit Sub
'        End If
'        invNumb = .Range("N6").value
'        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
'            invRow = wshFAC_Ent�te.Range("A4").value  'On Empty Invoice Go to First one created
'        Else 'On Existing Inv. find Previous one
'            invRow = wshFAC_Ent�te.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).Row + 1
'        End If
'        If .Range("N6").value >= MaxinvNumb Then
'            MsgBox "You are at the last invoice"
'            Exit Sub
'        End If
'        .Range("N3").value = wshFAC_Ent�te.Range("A" & invRow).value 'Place Inv. ID inside cell
'        Invoice_Load
'    End With
'
'    Call Log_Record("modFAC_Finale:Next_Invoice()", startTime)
'
'End Sub

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
    
End Sub

Sub FAC_Finale_Montrer_Sommaire_Taux()

    '�pure le sommaire des honoraires
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
        wshFAC_Finale.Range("C" & rowFAC_Finale).Font.Underline = True
        wshFAC_Finale.Range("C" & rowFAC_Finale).HorizontalAlignment = xlCenter

        wshFAC_Finale.Range("D" & rowFAC_Finale).value = "Taux"
        wshFAC_Finale.Range("D" & rowFAC_Finale).Font.Bold = True
        wshFAC_Finale.Range("D" & rowFAC_Finale).Font.Underline = True
        wshFAC_Finale.Range("D" & rowFAC_Finale).HorizontalAlignment = xlCenter

        Dim t As Variant
        i = rowFAC_Finale + 1
        For Each t In dictTaux.keys
            wshFAC_Finale.Range("C" & i & ":D" & i).Font.Color = RGB(0, 0, 0)
            wshFAC_Finale.Range("C" & i).NumberFormat = "##0.00"
            wshFAC_Finale.Range("C" & i).HorizontalAlignment = xlCenter
            wshFAC_Finale.Range("C" & i).Font.Bold = False
            wshFAC_Finale.Range("C" & i).Font.Underline = False
            wshFAC_Finale.Range("C" & i).Font.name = "Verdana"
            wshFAC_Finale.Range("C" & i).Font.size = 11
            wshFAC_Finale.Range("C" & i).value = dictTaux(t)
            wshFAC_Finale.Range("D" & i).Font.Bold = False
            wshFAC_Finale.Range("D" & i).NumberFormat = "#,##0.00 $"
            wshFAC_Finale.Range("D" & i).HorizontalAlignment = xlCenter
            wshFAC_Finale.Range("D" & i).Font.Underline = False
            wshFAC_Finale.Range("D" & i).Font.name = "Verdana"
            wshFAC_Finale.Range("D" & i).Font.size = 11
            wshFAC_Finale.Range("D" & i).value = t
            i = i + 1
        Next t
        
    End If
    
End Sub

Sub FAC_Finale_Goto_Onglet_FAC_Brouillon()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Goto_Onglet_FAC_Brouillon", 0)
   
    Application.ScreenUpdating = False
    
    wshFAC_Brouillon.Visible = xlSheetVisible
    wshFAC_Brouillon.Activate
    wshFAC_Brouillon.Range("E4").Select

    Application.ScreenUpdating = True
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Goto_Onglet_FAC_Brouillon()", startTime)

End Sub

Sub FAC_Finale_Enable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = True
    
    flagEtapeFacture = 3

    'Cleaning memory - 2024-07-01 @ 09:34
    Set shp = Nothing
    
End Sub

Sub FAC_Finale_Disable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = False

    'Cleaning memory - 2024-07-01 @ 09:34
    Set shp = Nothing
    
End Sub

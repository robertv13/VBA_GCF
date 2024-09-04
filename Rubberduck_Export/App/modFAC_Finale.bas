Attribute VB_Name = "modFAC_Finale"
Option Explicit

Dim invRow As Long, itemDBRow As Long, invitemRow As Long, invNumb As Long
Dim lastRow As Long, lastResultRow As Long, resultRow As Long

Sub FAC_Finale_Save() '2024-03-28 @ 07:19

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Save", 0)

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
    lastResultRow = wshTEC_Local.Range("AT9999").End(xlUp).Row
        
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
    Call TEC_TdB_Update_All '2024-03-21 @ 12:32

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
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Save()", startTime)
    
End Sub

Sub FAC_Finale_Add_Invoice_Header_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_to_DB", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Entête"
    
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
        rs.Fields("Date_Facture") = CDate(wshFAC_Brouillon.Range("O3").value)
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
        rs.Fields("Taux_TVQ") = Format$(.Range("C75").value, "0.0000")
        rs.Fields("Mnt_TVQ") = Format$(.Range("E75").value, "0.00")
        
        rs.Fields("AR_Total") = Format$(.Range("E77").value, "0.00")
        
        rs.Fields("Dépôt") = Format$(.Range("E79").value, "0.00")
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
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Entête.Range("A9999").End(xlUp).Row + 1
    
    With wshFAC_Entête
        .Range("A" & firstFreeRow).value = wshFAC_Finale.Range("E28")
        .Range("B" & firstFreeRow).value = Format$(wshFAC_Brouillon.Range("O3").value, "dd/mm/yyyy")
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
    
    wshFAC_Brouillon.Range("B11").value = firstFreeRow
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Header_Locally()", startTime)

    Application.ScreenUpdating = True

End Sub

Sub FAC_Finale_Add_Invoice_Details_to_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_to_DB", 0)

    Application.ScreenUpdating = False
    
    Dim rowLastService As Long
    rowLastService = wshFAC_Finale.Range("B64").End(xlUp).Row
    If rowLastService < 34 Then GoTo nothing_to_update
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Détails"
    
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
            rs.Fields("Inv_No") = noFacture
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
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Details_Locally", 0)
    
    Application.ScreenUpdating = False
    
    'Get the last entered service
    Dim lastEnteredService As Long
    lastEnteredService = wshFAC_Finale.Range("B64").End(xlUp).Row
    If lastEnteredService < 34 Then GoTo nothing_to_update
    
    'Get the first free row
    Dim firstFreeRow As Long
    firstFreeRow = wshFAC_Détails.Range("A99999").End(xlUp).Row + 1
   
    Dim i As Long
    For i = 34 To lastEnteredService
        With wshFAC_Détails
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_to_DB", 0)

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
                rs.Fields("Séquence") = seq
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Invoice_Somm_Taux_Locally", 0)
    
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB", 0)

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
        rs.Fields("Invoice_Date") = CDate(wshFAC_Brouillon.Range("O3").value)
        rs.Fields("Customer") = .Range("B24").value
        rs.Fields("CodeClient") = wshFAC_Brouillon.Range("B18").value
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
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_to_DB()", startTime)

End Sub

Sub FAC_Finale_Add_Comptes_Clients_Locally() '2024-03-11 @ 08:49 - Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Add_Comptes_Clients_Locally", 0)
    
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
        .Range("G" & firstFreeRow).value = CDate(CDate(wshFAC_Brouillon.Range("O3").value) + 30)
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

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB", 0)

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
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_To_DB()", startTime)

End Sub

Sub FAC_Finale_TEC_Update_As_Billed_Locally(firstResultRow As Long, lastResultRow As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_TEC_Update_As_Billed_Locally", 0)
    
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

Sub FAC_Finale_Softdelete_Projets_Détails_To_DB(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_To_DB", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Détails"
    
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
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_To_DB()", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Détails_Locally(projetID As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_Locally", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Détails
    
    Dim projetIDColumn As String, isDétruiteColumn As String
    projetIDColumn = "A"
    isDétruiteColumn = "I"

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
            'Update the isDétruite column for the found projetID
            ws.Cells(cell.Row, isDétruiteColumn).value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Détails_Locally()", startTime)

End Sub

Sub FAC_Finale_Softdelete_Projets_Entête_To_DB(projetID)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_To_DB", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "FAC_Projets_Entête"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

    'Build the query
    Dim strSQL As String
    strSQL = "UPDATE [" & destinationTab & "$] SET estDétruite = True WHERE ProjetID = " & projetID

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
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_To_DB()", startTime)
    Exit Sub

eh:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If

End Sub

Sub FAC_Finale_Softdelete_Projets_Entête_Locally(projetID)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_Locally", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Projets_Entête
    
    Dim projetIDColumn As String, isDétruiteColumn As String
    projetIDColumn = "A"
    isDétruiteColumn = "Z"

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
            'Update the isDétruite column for the found projetID
            ws.Cells(cell.Row, isDétruiteColumn).value = "VRAI"
            'Find the next cell with the projetID
            Set cell = ws.Range(projetIDColumn & "2:" & projetIDColumn & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Finale:FAC_Finale_Softdelete_Projets_Entête_Locally()", startTime)

End Sub

Sub Invoice_Load() 'Retrieve an existing invoice - 2023-12-21 @ 10:16

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:Invoice_Load", 0)

    With wshFAC_Brouillon
        If wshFAC_Brouillon.Range("B20").value = Empty Then
            MsgBox "Impossible de retrouver cette facture. Veuillez saisir un numéro de facture VALIDE pour votre recherche"
            GoTo NoItems
        End If
        'Could that invoice been cancelled (more than 1 row) ?
        Call InvoiceGetAllTrans(wshFAC_Brouillon.Range("O6").value)
        Dim NbTrans As Long
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
        .Range("B24").value = True 'Set Invoice Load to true
        .Range("S2,E4:F4,K4:L6,O3,K11:O45,Q11:Q45").ClearContents
        wshFAC_Finale.Range("B34:F68").ClearContents
        Dim InvListRow As Long
        InvListRow = wshFAC_Brouillon.Range("B20").value 'InvListRow = Row associated with the invoice
        'Get values from wshFAC_Entête (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
        .Range("O3").value = wshFAC_Entête.Range("B" & InvListRow).value
        .Range("K3").value = wshFAC_Entête.Range("D" & InvListRow).value
        .Range("K4").value = wshFAC_Entête.Range("E" & InvListRow).value
        .Range("K5").value = wshFAC_Entête.Range("F" & InvListRow).value
        .Range("K6").value = wshFAC_Entête.Range("G" & InvListRow).value
        'Get values from wshFAC_Entête (header) and enter them in the wshFAC_Brouillon - 2023-12-19 @ 08:29
        Dim dFact As Date
        dFact = wshFAC_Entête.Range("B" & InvListRow).value
        wshFAC_Finale.Range("B21").value = "Le " & Format$(dFact, "d") & " " & _
                                            UCase(Format$(dFact, "mmmm")) & " " & _
                                            Format$(dFact, "yyyy")
        wshFAC_Finale.Range("B23").value = wshFAC_Entête.Range("D" & InvListRow).value
        wshFAC_Finale.Range("B24").value = Fn_Strip_Contact_From_Client_Name(wshFAC_Entête.Range("E" & InvListRow).value)
        wshFAC_Finale.Range("B25").value = wshFAC_Entête.Range("F" & InvListRow).value
        wshFAC_Finale.Range("B26").value = wshFAC_Entête.Range("G" & InvListRow).value
        'Load Invoice Detail Items
        With wshFAC_Détails
            Dim lastRow As Long, lastResultRow As Long
            lastRow = .Range("A999999").End(xlUp).Row
            If lastRow < 4 Then Exit Sub 'No Item Lines
            .Range("I3").value = wshFAC_Brouillon.Range("O6").value
            wshFAC_Finale.Range("F28").value = wshFAC_Brouillon.Range("O6").value 'Invoice #
            'Advanced Filter to get items specific to ONE invoice
            .Range("A3:G" & lastRow).AdvancedFilter xlFilterCopy, criteriaRange:=.Range("I2:I3"), CopyToRange:=.Range("K2:P2"), Unique:=True
            lastResultRow = .Range("O999").End(xlUp).Row
            If lastResultRow < 3 Then GoTo NoItems
            For resultRow = 3 To lastResultRow
                invitemRow = .Range("O" & resultRow).value
                wshFAC_Brouillon.Range("L" & invitemRow & ":O" & invitemRow).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
                wshFAC_Brouillon.Range("Q" & invitemRow).value = .Range("P" & resultRow).value  'Set Item DB Row
                wshFAC_Finale.Range("C" & invitemRow + 23 & ":F" & invitemRow + 23).value = .Range("K" & resultRow & ":N" & resultRow).value 'Description, Hours, Rate & Value
            Next resultRow
        End With
        'Proceed with trailer data (Misc. charges & Taxes)
        .Range("M48").value = wshFAC_Entête.Range("I" & InvListRow).value
        .Range("O48").value = wshFAC_Entête.Range("J" & InvListRow).value
        .Range("M49").value = wshFAC_Entête.Range("K" & InvListRow).value
        .Range("O49").value = wshFAC_Entête.Range("L" & InvListRow).value
        .Range("M50").value = wshFAC_Entête.Range("M" & InvListRow).value
        .Range("O50").value = wshFAC_Entête.Range("N" & InvListRow).value
        .Range("O52").value = wshFAC_Entête.Range("P" & InvListRow).value
        .Range("O53").value = wshFAC_Entête.Range("R" & InvListRow).value
        .Range("O57").value = wshFAC_Entête.Range("T" & InvListRow).value

NoItems:
    .Range("B24").value = False 'Set Invoice Load To false
    End With
    
    Call Log_Record("modFAC_Finale:Invoice_Load()", startTime)

End Sub

Sub Copier_Facture_Vers_Classeur_Ferme(invNo As String, clientID As String)

    Dim wsCopie As Worksheet
    Dim cheminDest As String
    Dim nomNouveau As String
    Dim nomFeuilleBase As String
    Dim compteur As Integer
    Dim nomFinal As String
    
    'Désactiver les mises à jour de l'écran pour améliorer les performances
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error GoTo GestionErreur
    
    'Définir le classeur source et la feuille source
    Dim wbSource As Workbook: Set wbSource = ThisWorkbook
    Dim wsSource As Worksheet: Set wsSource = wshFAC_Finale
    
    'Initialiser la boîte de dialogue de sélection de fichier
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Sauvegarde de la facture (format EXCEL)"
        .Filters.clear
        .Filters.add "Classeur Excel", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .AllowMultiSelect = False
        If .show <> -1 Then
            MsgBox "Aucun fichier sélectionné. Opération annulée.", vbExclamation
            GoTo Fin
        End If
        cheminDest = .selectedItems(1)
    End With
    
    'Ouvrir le classeur de destination
    Dim wbDest As Workbook: Set wbDest = Workbooks.Open(cheminDest)
    
    'Ajouter une nouvelle feuille dans le classeur de destination
    Set wsCopie = wbDest.Sheets.add(After:=wbDest.Sheets(wbDest.Sheets.count))
    
    'Copier les colonnes A à F de la feuille source vers la nouvelle feuille du classeur de destination
    wsSource.Range("A:F").Copy Destination:=wsCopie.Range("A1")

    'Supprimer les formules et les remplacer par des valeurs
    With wsCopie.usedRange
        .value = .value
    End With

'    wsSource.Copy After:=wbDest.Sheets(wbDest.Sheets.count)
'
'    'Définir la feuille copiée (la dernière feuille du classeur de destination)
'    Set wsCopie = wbDest.Sheets(wbDest.Sheets.count)
    
    'Définir la base du nom de la nouvelle feuille
    nomFeuilleBase = Format$(wshFAC_Brouillon.Range("O3").value, "yyyy-mm-dd") & " - " & invNo
    
    'Vérifier si le nom existe déjà et ajuster en conséquence
    compteur = 1
    nomFinal = nomFeuilleBase
    Do While NomFeuilleExiste(nomFinal, wbDest)
        nomFinal = nomFeuilleBase & " V:" & compteur
        compteur = compteur + 1
    Loop
    
    'Renommer la feuille copiée
    wsCopie.name = nomFinal
    
    'Enregistrer et fermer le classeur de destination
    wbDest.Save
    wbDest.Close
    
Fin:
    'Réactiver les mises à jour de l'écran et les alertes
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Sub
    
GestionErreur:

    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical
    Resume Fin
    
End Sub

' Fonction pour vérifier si un nom de feuille existe déjà dans un classeur
Function NomFeuilleExiste(nom As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(nom)
    NomFeuilleExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

Sub InvoiceGetAllTrans(inv As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:InvoiceGetAllTrans", 0)

    Application.ScreenUpdating = False

    wshFAC_Brouillon.Range("B31").value = 0

    With wshFAC_Entête
        Dim lastRow As Long, lastResultRow As Long, resultRow As Long
        lastRow = .Range("A999999").End(xlUp).Row 'Last wshFAC_Entête Row
        If lastRow < 4 Then GoTo Done '3 rows of Header - Nothing to search/filter
        On Error Resume Next
        .Names("Criterial").delete
        On Error GoTo 0
        .Range("V3").value = wshFAC_Brouillon.Range("O6").value
        'Advanced Filter setup
        .Range("A3:T" & lastRow).AdvancedFilter xlFilterCopy, _
            criteriaRange:=.Range("V2:V3"), _
            CopyToRange:=.Range("X2:AQ2"), _
            Unique:=True
        lastResultRow = .Range("X999").End(xlUp).Row 'How many rows trans for that invoice
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
         wshFAC_Brouillon.Range("B31").value = lastResultRow - 2 'Remove Header rows from row count
Done:
    End With
    Application.ScreenUpdating = True

    Call Log_Record("modFAC_Finale:InvoiceGetAllTrans()", startTime)

End Sub

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

Sub FAC_Finale_Creation_PDF() 'RMV - 2023-12-17 @ 14:35
    
    Call FAC_Finale_Create_PDF_Sub(wshFAC_Finale.Range("E28").value)
    
    DoEvents
    
    Call Copier_Facture_Vers_Classeur_Ferme(wshFAC_Finale.Range("E28").value, _
                                            wshFAC_Brouillon.Range("B18").value)
    
    DoEvents
    
    Call FAC_Finale_Enable_Save_Button

End Sub

Sub FAC_Finale_Create_PDF_Sub(noFacture As String)

    'Création du fichier (NoFacture).PDF dans le répertoire de factures PDF de GCF
    Dim result As Boolean
    result = FAC_Finale_Create_PDF_Func(noFacture, "SaveOnly")

End Sub

Function FAC_Finale_Create_PDF_Func(noFacture As String, Optional action As String = "SaveOnly") As Boolean
    
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
    
SaveOnly:
    FAC_Finale_Create_PDF_Func = True 'Return value
'    FAC_Finale_Create_PDF_Email_Func = True 'Return value
    GoTo EndMacro
    
RefLibError:
    MsgBox "Incapable de préparer le courriel. La librairie n'est pas disponible"
    FAC_Finale_Create_PDF_Func = False 'Function return value
'    FAC_Finale_Create_PDF_Email_Func = False 'Function return value

EndMacro:
    Application.ScreenUpdating = True
    
End Function

Sub Prev_Invoice() 'TO-DO-RMV 2023-12-17
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:Prev_Invoice", 0)
    
    With wshFAC_Brouillon
        Dim MininvNumb As Long
        On Error Resume Next
        MininvNumb = Application.WorksheetFunction.Min(wshFAC_Entête.Range("Inv_ID"))
        On Error GoTo 0
        If MininvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").value
        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            invRow = wshFAC_Entête.Range("A99999").End(xlUp).Row 'On Empty Invoice Go to last one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFAC_Entête.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).Row - 1
        End If
        If .Range("N6").value = 1 Or MininvNumb = 0 Or MininvNumb = .Range("N6").value Then
            MsgBox "You are at the first invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFAC_Entête.Range("A" & invRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With
    
    Call Log_Record("modFAC_Finale:Prev_Invoice()", startTime)

End Sub

Sub Next_Invoice() 'TO-DO-RMV 2023-12-17

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:Next_Invoice", 0)
    
    With wshFAC_Brouillon
        Dim MaxinvNumb As Long
        On Error Resume Next
        MaxinvNumb = Application.WorksheetFunction.Max(wshFAC_Entête.Range("Inv_ID"))
        On Error GoTo 0
        If MaxinvNumb = 0 Then
            MsgBox "Please create and save an Invoice first"
            Exit Sub
        End If
        invNumb = .Range("N6").value
        If invNumb = 0 Or .Range("B20").value = Empty Then 'On New Invoice
            invRow = wshFAC_Entête.Range("A4").value  'On Empty Invoice Go to First one created
        Else 'On Existing Inv. find Previous one
            invRow = wshFAC_Entête.Range("Inv_ID").Find(invNumb, , xlValues, xlWhole).Row + 1
        End If
        If .Range("N6").value >= MaxinvNumb Then
            MsgBox "You are at the last invoice"
            Exit Sub
        End If
        .Range("N3").value = wshFAC_Entête.Range("A" & invRow).value 'Place Inv. ID inside cell
        Invoice_Load
    End With

    Call Log_Record("modFAC_Finale:Next_Invoice()", startTime)

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

Sub FAC_Finale_Cacher_Sommaire_Taux()

    'First determine how many rows there is in the Fees Summary
    Dim nbItems As Long
    Dim i As Long
    For i = 67 To 63 Step -1
        If wshFAC_Finale.Range("C" & i).value <> "" Then
            nbItems = nbItems + 1
        End If
    Next i
    
    If nbItems > 0 Then
        Dim rngFeesSummary As Range: Set rngFeesSummary = _
            wshFAC_Finale.Range("C" & (67 - nbItems) + 1 & ":D67")
        rngFeesSummary.ClearContents
        
'        Call Fees_Summary_Borders_Invisible(rngFeesSummary)
        
        'Clear the contents of the 'Sommaire' cell
'        wshFAC_Finale.Range("B" & 67 - nbItems).ClearContents
    End If
    
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
                dictTaux.add taux, hres
                nbTaux = nbTaux + 1
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
            wshFAC_Finale.Range("C" & i).Font.Underline = False
            wshFAC_Finale.Range("C" & i).Font.name = "Verdana"
            wshFAC_Finale.Range("C" & i).Font.size = 11
            wshFAC_Finale.Range("C" & i).value = dictTaux(t)
            
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

'Sub FAC_Finale_GL_Posting_Preparation() '2024-06-06 @ 10:31
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Finale:modFAC_Finale", 0)
'
'    Dim montant As Double
'    Dim dateFact As Date
'    Dim descGL_Trans As String, source As String
'    Dim GL_TransNo As Long
'
'    dateFact = wshFAC_Brouillon.Range("O3").value
'    descGL_Trans = wshFAC_Brouillon.Range("E4").value
'    source = "FACT-" & wshFAC_Brouillon.Range("O6").value
'    GL_TransNo = wshFAC_Brouillon.Range("B41").value
'
'    Dim MyArray(1 To 8, 1 To 4) As String
'
'    'AR amount (wshFAC_Brouillon.Range("B33"))
'    montant = wshFAC_Brouillon.Range("B33").value
'    If montant Then
'        MyArray(1, 1) = "1100"
'        MyArray(1, 2) = "Comptes clients"
'        MyArray(1, 3) = montant
'        MyArray(1, 4) = ""
'    End If
'
'    'Professional Fees (wshFAC_Brouillon.Range("B34"))
'    montant = wshFAC_Brouillon.Range("B34").value
'    If montant Then
'        MyArray(2, 1) = "4000"
'        MyArray(2, 2) = "Revenus de consultation"
'        MyArray(2, 3) = montant
'        MyArray(2, 4) = ""
'    End If
'
'    'Miscellaneous Amount # 1 (wshFAC_Brouillon.Range("B35"))
'    montant = wshFAC_Brouillon.Range("B35").value
'    If montant Then
'        MyArray(3, 1) = "9999"
'        MyArray(3, 2) = "Frais divers # 1"
'        MyArray(3, 3) = montant
'        MyArray(3, 4) = ""
'    End If
'
'    'Miscellaneous Amount # 2 (wshFAC_Brouillon.Range("B36"))
'    montant = wshFAC_Brouillon.Range("B36").value
'    If montant Then
'        MyArray(4, 1) = "9999"
'        MyArray(4, 2) = "Frais divers # 2"
'        MyArray(4, 3) = montant
'        MyArray(4, 4) = ""
'    End If
'
'    'Miscellaneous Amount # 3 (wshFAC_Brouillon.Range("B37"))
'    montant = wshFAC_Brouillon.Range("B37").value
'    If montant Then
'        MyArray(5, 1) = "9999"
'        MyArray(5, 2) = "Frais divers # 3"
'        MyArray(5, 3) = montant
'        MyArray(5, 4) = ""
'    End If
'
'    'GST to pay (wshFAC_Brouillon.Range("B38"))
'    montant = wshFAC_Brouillon.Range("B38").value
'    If montant Then
'        MyArray(6, 1) = "1202"
'        MyArray(6, 2) = "TPS percues"
'        MyArray(6, 3) = montant
'        MyArray(6, 4) = ""
'    End If
'
'    'PST to pay (wshFAC_Brouillon.Range("B39"))
'    montant = wshFAC_Brouillon.Range("B39").value
'    If montant Then
'        MyArray(7, 1) = "1203"
'        MyArray(7, 2) = "TVQ percues"
'        MyArray(7, 3) = montant
'        MyArray(7, 4) = ""
'    End If
'
'    'Deposit applied (wshFAC_Brouillon.Range("B40"))
'    montant = wshFAC_Brouillon.Range("B40").value
'    If montant Then
'        MyArray(8, 1) = "2400"
'        MyArray(8, 2) = "Produit perçu d'avance"
'        MyArray(8, 3) = montant
'        MyArray(8, 4) = ""
'    End If
'
'    Call GL_Posting_To_DB(dateFact, descGL_Trans, source, MyArray)
'
'    Call GL_Posting_Locally(dateFact, descGL_Trans, source, GL_TransNo, MyArray)
'
'    Call Log_Record("modFAC_Finale:modFAC_Finale()", startTime)
'
'End Sub
'
Sub FAC_Finale_Enable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set shp = Nothing
    
End Sub

Sub FAC_Finale_Disable_Save_Button()

    Dim shp As Shape: Set shp = wshFAC_Finale.Shapes("shpSauvegarde")
    shp.Visible = False

    'Cleaning memory - 2024-07-01 @ 09:34
    Set shp = Nothing
    
End Sub

'Sub ExportAllFacInvList() '2024-03-28 @ 14:22
'    Dim wb As Workbook
'    Dim wsSource As Worksheet
'    Dim wsTarget As Worksheet
'    Dim sourceRange As Range
'
'    Application.ScreenUpdating = False
'
'    'Work with the source range
'    Set wsSource = wshFAC_Entête
'    Dim lastUsedRow As Long
'    lastUsedRow = wsSource.Range("A99999").End(xlUp).row
'    wsSource.Range("A4:T" & lastUsedRow).Copy
'
'    'Open the target workbook
'    Workbooks.Open fileName:=wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & _
'                   "GCF_BD_MASTER.xlsx"
'
'    'Set references to the target workbook and target worksheet
'    Set wb = Workbooks("GCF_BD_MASTER.xlsx")
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



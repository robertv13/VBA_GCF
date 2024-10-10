Attribute VB_Name = "modFAC_Confirmation"
Option Explicit

Public invNo As String

Sub Get_Invoice_Data(noFact As String)

    'Save original worksheet
    Dim oWorkSheet As Worksheet: Set oWorkSheet = ActiveSheet
    
    'Reference to A/R master file
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
    Dim result As Variant
    Dim rngToSearch As Range: Set rngToSearch = ws.Range("A1").CurrentRegion.Offset(0, 0).Resize(lastUsedRow, 1)
    result = Application.WorksheetFunction.XLookup(noFact, _
                                                   rngToSearch, _
                                                   rngToSearch, _
                                                   "Not Found", _
                                                   0, _
                                                   1)
    
    If result <> "Not Found" Then
        Dim matchedRow As Long
        matchedRow = Application.Match(noFact, rngToSearch, 0)
        
        Call Display_Invoice_info(ws, matchedRow)
        
        Call Insert_Big_PDF_Icon
        
        Dim resultArr As Variant
        resultArr = Fn_Get_TEC_Invoiced_By_This_Invoice(noFact)
        
        If Not IsEmpty(resultArr) Then
            Dim TECSummary() As Variant
            ReDim TECSummary(1 To 10, 1 To 3)
            Call Get_TEC_Summary_For_That_Invoice(resultArr, TECSummary)
            
            Dim FeesSummary() As Variant
            ReDim FeesSummary(1 To 5, 1 To 3)
            Call Get_Fees_Summary_For_That_Invoice(resultArr, FeesSummary)
        End If
        
'        Call FAC_Confirmation_Get_GL_Posting(noFact)
'
        oWorkSheet.Activate
        
    Else
        MsgBox "La facture n'existe pas"
        GoTo Clean_Exit
    End If
    
Clean_Exit:
    Set oWorkSheet = Nothing
    Set rngToSearch = Nothing
    Set ws = Nothing

End Sub

Sub Insert_Big_PDF_Icon()

    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Dim i As Long
    Dim iconPath As String
    iconPath = wshAdmin.Range("F5").value & Application.PathSeparator & "Resources\AdobeAcrobatReader.png"
    
    Dim pic As Picture
    Dim cell As Range
    
    'Loop through each row and insert the icon if there is data in column E
    Set cell = ws.Cells(7, 12) 'Set the cell where the icon should be inserted
            
    'Insert the icon
    Set pic = ws.Pictures.Insert(iconPath)
    With pic
        .Top = cell.Top + 10
        .Left = cell.Left + 10
        .Height = 50 'cell.Height
        .width = 50 'cell.width
        .Placement = xlMoveAndSize
        .OnAction = "FAC_Confirmation_Display_PDF_Invoice"
    End With
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set pic = Nothing
    Set ws = Nothing
    
End Sub

Sub FAC_Confirmation_Display_PDF_Invoice()

    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    'Assuming the invoice number is at 'F5'
    Dim fullPDFFileName As String
    fullPDFFileName = wshAdmin.Range("F5").value & FACT_PDF_PATH & _
        Application.PathSeparator & ws.Cells(5, 6).value & ".pdf"
    
    'Open the invoice using Adobe Acrobat Reader
    If fullPDFFileName <> "" Then
        Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe " & Chr(34) & fullPDFFileName & Chr(34), vbNormalFocus
    Else
        MsgBox "Je ne retrouve pas cette facture", vbExclamation
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Sub Display_Invoice_info(wsF As Worksheet, r As Long)

    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    'Display all fields from FAC_Entête
    With ws
        .Range("L5").value = wsF.Cells(r, 2).value
    
        ws.Range("F7").value = wsF.Cells(r, 5).value
        ws.Range("F8").value = Fn_Strip_Contact_From_Client_Name _
                                                (wsF.Cells(r, 6).value)
        ws.Range("F9").value = wsF.Cells(r, 7).value
        ws.Range("F10").value = wsF.Cells(r, 8).value
        ws.Range("F11").value = wsF.Cells(r, 9).value
        
        ws.Range("L13").value = wsF.Cells(r, 10).value
        ws.Range("L14").value = wsF.Cells(r, 12).value
        ws.Range("L15").value = wsF.Cells(r, 14).value
        ws.Range("L16").value = wsF.Cells(r, 16).value
        ws.Range("L17").formula = "=SUM(L13:L16)"
        
        ws.Range("L18").value = wsF.Cells(r, 18).value
        ws.Range("L19").value = wsF.Cells(r, 20).value
        ws.Range("L21").formula = "=SUM(L17:L19)"
        
        ws.Range("L23").value = wsF.Cells(r, 22).value
        ws.Range("L25").formula = "=L21 - L23"
        
    End With
    
    'Take care of invoice type (to be confirmed OR already confirmed)
    If wsF.Cells(r, 3).value = "AC" Then
        ws.Range("H5").value = "À CONFIRMER"
        ws.Shapes("btnFAC_Confirmation").Visible = True
    Else
        ws.Range("H5").value = ""
        ws.Shapes("btnFAC_Confirmation").Visible = False
    End If
    
    'Make OK button visible
    ws.Shapes("btnFAC_Confirmation_OK").Visible = True
    
    Application.EnableEvents = True

End Sub

Sub Show_Unconfirmed_Invoice()

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Clear contents or the area
    Dim lastUsedRow As Long
    lastUsedRow = wshFAC_Confirmation.Cells(wshFAC_Confirmation.rows.count, "P").End(xlUp).Row
    If lastUsedRow > 3 Then
        wshFAC_Confirmation.Range("P4:AA" & lastUsedRow).ClearContents
    End If

    'Set criteria for AvancedFilter
    ws.Range("AW3").value = "AC"
    
    Call FAC_Entête_AdvancedFilter_AC_C
    
    Dim lastUsedRowAF As Long
    lastUsedRowAF = ws.Cells(ws.rows.count, "AY").End(xlUp).Row
    If lastUsedRowAF < 3 Then
        GoTo Clean_Exit
    End If
    
    wshFAC_Confirmation.Unprotect
    
    Dim i As Integer
    For i = 3 To lastUsedRowAF
        With wshFAC_Confirmation
            wshFAC_Confirmation.rows(i + 1).Locked = False
            .Cells(i + 1, 16).value = ws.Cells(i, 51)
            .Cells(i + 1, 17).value = ws.Cells(i, 52)
            .Cells(i + 1, 18).value = ws.Cells(i, 55)
            .Cells(i + 1, 19).value = ws.Cells(i, 67)
            .Cells(i + 1, 20).value = ws.Cells(i, 56)
            .Cells(i + 1, 21).value = ws.Cells(i, 58)
            .Cells(i + 1, 22).value = ws.Cells(i, 60)
            .Cells(i + 1, 23).value = ws.Cells(i, 62)
            .Cells(i + 1, 24).value = ws.Cells(i, 64)
            .Cells(i + 1, 25).value = ws.Cells(i, 66)
            .Cells(i + 1, 26).value = ws.Cells(i, 68)
        End With
    Next i
    
    With wshFAC_Confirmation
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
Clean_Exit:
    Set ws = Nothing

End Sub

Sub Get_TEC_Summary_For_That_Invoice(arr As Variant, ByRef TECSummary As Variant)

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    
    'Setup a Dictionary to summarize the hours by Professionnal
    Dim dictHours As Object: Set dictHours = CreateObject("Scripting.Dictionary")

    Dim pro As String
    Dim hres As Double
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        pro = wsTEC.Cells(arr(i), 3).value
        hres = wsTEC.Cells(arr(i), 8).value
        If hres <> 0 Then
            If dictHours.Exists(pro) Then
                dictHours(pro) = dictHours(pro) + hres
            Else
                dictHours.add pro, hres
            End If
        End If
    Next i
    
    Dim profID As Long
    Dim rowInWorksheet As Long: rowInWorksheet = 13
    Dim prof As Variant
    For Each prof In Fn_Sort_Dictionary_By_Value(dictHours, True) 'Sort dictionary by hours in descending order
        Dim strProf As String
        strProf = prof
        profID = Fn_GetID_From_Initials(strProf)
        hres = dictHours(prof)
        Dim tauxHoraire As Currency
        tauxHoraire = Fn_Get_Hourly_Rate(profID, wshFAC_Confirmation.Range("L5").value)
        wshFAC_Confirmation.Cells(rowInWorksheet, 6) = strProf
        wshFAC_Confirmation.Cells(rowInWorksheet, 7) = _
                CDbl(Format$(hres, "0.00"))
        wshFAC_Confirmation.Cells(rowInWorksheet, 8) = _
                CDbl(Format$(tauxHoraire, "# ##0.00 $"))
        rowInWorksheet = rowInWorksheet + 1
'        Debug.Print "Summary : " & strProf & " = " & hres & " @ " & tauxHoraire
'        Cells(rowSelected, 14).FormulaR1C1 = "=RC[-2]*RC[-1]"
'        rowSelected = rowSelected + 1
    Next prof
    
    'Cleanup - 2024-07-25 @ 18:06
    Set dictHours = Nothing
    Set wsTEC = Nothing
    
End Sub

Sub Get_Fees_Summary_For_That_Invoice(arr As Variant, ByRef FeesSummary As Variant)

    Dim wsFees As Worksheet: Set wsFees = wshFAC_Sommaire_Taux
    
    'Determine the last used row
    Dim lastUsedRow As Long
    lastUsedRow = wsFees.Cells(wsFees.rows.count, "A").End(xlUp).Row
    
    'Get Invoice number
    Dim invNo As String
    invNo = Trim(wshFAC_Confirmation.Range("F5").value)
    
    'Use Range.Find to locate the first cell with the InvoiceNo
    Dim cell As Range
    Set cell = wsFees.Range("A2:A" & lastUsedRow).Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    Dim rowFeesSummary As Long: rowFeesSummary = 20
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Application.EnableEvents = False
        Do
            'Display values in the worksheet
            wshFAC_Confirmation.Range("F" & rowFeesSummary).value = wsFees.Cells(cell.Row, 3).value
            wshFAC_Confirmation.Range("G" & rowFeesSummary).value = _
                        CDbl(Format$(wsFees.Cells(cell.Row, 4).value, "##0.00"))
            wshFAC_Confirmation.Range("H" & rowFeesSummary).value = _
                        CDbl(Format$(wsFees.Cells(cell.Row, 5).value, "##,##0.00 $"))
            rowFeesSummary = rowFeesSummary + 1
            'Find the next cell with the invNo
            Set cell = wsFees.Range("A2:A" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If
    
    'Cleanup - 2024-07-25 @ 18:06
    Set cell = Nothing
    Set wsFees = Nothing
    
End Sub

Sub FAC_Confirmation_Clear_Cells_And_PDF_Icon()

    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    ws.Range("F5,H5,L5").ClearContents
    
    ws.Range("F7:I11").ClearContents
    
    ws.Range("L13:L19").ClearContents
    
    ws.Range("L21,L23,L25").ClearContents
    
    ws.Range("F13:H17").ClearContents
    
    ws.Range("F20:H24").ClearContents
    
    Dim pic As Picture
    For Each pic In ws.Pictures
        pic.delete
    Next pic
    
    'Hide both buttons
    ws.Shapes("btnFAC_Confirmation").Visible = False
    ws.Shapes("btnFAC_Confirmation_OK").Visible = False
    
    Call Show_Unconfirmed_Invoice
    
    'Cleaning memory - 2024-07-01 @ 09:34 memory - 2024-07-01 @ 09:34
    Set pic = Nothing
    Set ws = Nothing

    Application.EnableEvents = True
    
    On Error Resume Next
    wshFAC_Confirmation.Range("F5").Select
    On Error GoTo 0
    
End Sub

Sub FAC_Confirmation_OK_Button_Click()

    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Call FAC_Confirmation_Clear_Cells_And_PDF_Icon
    
    ws.Range("F5").Select
    
    'Cleanup - 2024-07-26 @ 00:55
    Set ws = Nothing
    
End Sub

Sub FAC_Confirmation_Button_Click()

    Dim ws As Worksheet: Set ws = wshFAC_Confirmation
    
    Dim invNo As String
    invNo = ws.Range("F5").value
    
    ws.Shapes("btnFAC_Confirmation").Visible = False
    
    Dim answerYesNo As Long
    answerYesNo = MsgBox("Êtes-vous certain de vouloir CONFIRMER cette facture ? ", _
                         vbYesNo + vbQuestion, "Confirmation de facture")
    If answerYesNo = vbNo Then
        MsgBox _
            Prompt:="Cette facture ne sera PAS CONFIRMÉE ! ", _
            Title:="Confirmation", _
            Buttons:=vbCritical
            GoTo Clean_Exit
    End If
    
    If answerYesNo = vbYes Then
    
        Call FAC_Confirmation_Facture(invNo)
        
    End If
    
Clean_Exit:

    Call FAC_Confirmation_Clear_Cells_And_PDF_Icon

    wshFAC_Confirmation.Range("F5").Select
    
    'Cleanup - 2024-07-26 @ 00:55
    Set ws = Nothing
    
End Sub

Sub FAC_Confirmation_Get_GL_Posting(invNo)

    Dim wsGL As Worksheet: Set wsGL = wshGL_Trans
    
    Dim lastUsedRow
    lastUsedRow = wsGL.Range("A99999").End(xlUp).Row
    Dim rngToSearch As Range: Set rngToSearch = wsGL.Range("D1:D" & lastUsedRow)
    
    'Use Range.Find to locate the first cell with the invNo
    Dim cell As Range
    Set cell = wsGL.Range("D2:D" & lastUsedRow).Find(What:="FACTURE:" & invNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Check if the invNo was found at all
    Dim firstAddress As String
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Dim r As Long
        r = 38
        Application.EnableEvents = False
        Do
            'Save the information for invoice deletion
            r = r + 1
            'Find the next cell with the invNo
            Set cell = wsGL.Range("D2:D" & lastUsedRow).FindNext(After:=cell)
        Loop While Not cell Is Nothing And cell.Address <> firstAddress
        Application.EnableEvents = True
    End If

End Sub

Sub FAC_Confirmation_Facture(invNo As String)

    'Update the type of invoice (Master)
    Call FAC_Confirmation_Update_BD_MASTER(invNo)
    
    'Update the type of invoice (Locally)
    Call FAC_Confirmation_Update_Locally(invNo)
    
    'Do the G/L posting
    Call FAC_Confirmation_GL_Posting(invNo)
    
    MsgBox "Cette facture a été confirmée avec succès", vbInformation

    'Clear the cells on the current Worksheet
    Call FAC_Confirmation_Clear_Cells_And_PDF_Icon
    
End Sub

Sub FAC_Confirmation_Update_BD_MASTER(invoice As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_BD_MASTER", 0)

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

    Dim SQL As String
    'Open the recordset for the specified invoice
    SQL = "SELECT * FROM [" & destinationTab & "$] WHERE Inv_No = '" & invoice & "'"
    rs.Open SQL, conn, 2, 3
    If Not rs.EOF Then
        'Update AC_ouC with 'C'
        rs.Fields("AC_C").value = "C"
        rs.update
    Else
        'Handle the case where the specified invoice is not found
        MsgBox "La facture '" & invoice & "' n'existe pas!", vbCritical
    End If
    
    'Close recordset and connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_BD_MASTER()", startTime)

End Sub

Sub FAC_Confirmation_Update_Locally(invoice As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_Locally", 0)
    
    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long, rowToBeUpdated As Long, TECID As Long
    If Not foundRange Is Nothing Then
        r = foundRange.Row
        ws.Cells(r, 3).value = "C"
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entête."
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set lookupRange = Nothing
    Set ws = Nothing
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_Update_Locally()", startTime)

End Sub

Sub FAC_Confirmation_GL_Posting(invoice As String) '2024-08-18 @17:15

    Dim startTime As Double: startTime = Timer: Call Log_Record("modFAC_Confirmation:FAC_Confirmation_GL_Posting", 0)

    Dim ws As Worksheet: Set ws = wshFAC_Entête
    
    'Set the range to look for
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim lookupRange As Range: Set lookupRange = ws.Range("A3:A" & lastUsedRow)
    
    Dim foundRange As Range
    Set foundRange = lookupRange.Find(What:=invoice, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim r As Long
    If Not foundRange Is Nothing Then
        r = foundRange.Row
        Dim dateFact As Date
        dateFact = ws.Cells(r, 2).value
        Dim hono As Currency
        hono = ws.Cells(r, 10).value
        Dim misc1 As Currency, misc2 As Currency, misc3 As Currency
        misc1 = ws.Cells(r, 12).value
        misc2 = ws.Cells(r, 14).value
        misc3 = ws.Cells(r, 16).value
        Dim tps As Currency, tvq As Currency
        tps = ws.Cells(r, 18).value
        tvq = ws.Cells(r, 20).value
        Dim depot As Currency
        depot = ws.Cells(r, 22).value
        
        Dim descGL_Trans As String, source As String
        descGL_Trans = ws.Cells(r, 6).value
        source = "FACTURE:" & invoice
        
        Dim MyArray(1 To 7, 1 To 4) As String
        
        'AR amount
        If hono + misc1 + misc2 + misc3 + tps + tvq Then
            MyArray(1, 1) = "1100"
            MyArray(1, 2) = "Comptes clients"
            MyArray(1, 3) = hono + misc1 + misc2 + misc3 + tps + tvq
            MyArray(1, 4) = ""
        End If
        
        'Professional Fees (hono)
        If hono Then
            MyArray(2, 1) = "4000"
            MyArray(2, 2) = "Revenus de consultation"
            MyArray(2, 3) = -hono
            MyArray(2, 4) = ""
        End If
        
        'Miscellaneous Amount # 1 (misc1)
        If misc1 Then
            MyArray(3, 1) = "4010"
            MyArray(3, 2) = "Revenus - Frais de poste"
            MyArray(3, 3) = -misc1
            MyArray(3, 4) = ""
        End If
        
        'Miscellaneous Amount # 2 (misc2)
        If misc2 Then
            MyArray(4, 1) = "4015"
            MyArray(4, 2) = "Revenus - Sous-traitants"
            MyArray(4, 3) = -misc2
            MyArray(4, 4) = ""
        End If
        
        'Miscellaneous Amount # 3 (misc3)
        If misc3 Then
            MyArray(5, 1) = "4020"
            MyArray(5, 2) = "Revenus - Autres Frais"
            MyArray(5, 3) = -misc3
            MyArray(5, 4) = ""
        End If
        
        'GST to pay (tps)
        If tps Then
            MyArray(6, 1) = "1202"
            MyArray(6, 2) = "TPS percues"
            MyArray(6, 3) = -tps
            MyArray(6, 4) = ""
        End If
        
        'PST to pay (tvq)
        If tvq Then
            MyArray(7, 1) = "1203"
            MyArray(7, 2) = "TVQ percues"
            MyArray(7, 3) = -tvq
            MyArray(7, 4) = ""
        End If
        
    '    'Deposit applied (depot)
    '    If depot Then
    '        MyArray(8, 1) = "2400"
    '        MyArray(8, 2) = "Produit perçu d'avance"
    '        MyArray(8, 3) = depot
    '        MyArray(8, 4) = ""
    '    End If
        
        Dim glEntryNo As Long
        Call GL_Posting_To_DB(dateFact, descGL_Trans, source, MyArray, glEntryNo)
        
        Call GL_Posting_Locally(dateFact, descGL_Trans, source, MyArray, glEntryNo)
        
    Else
        MsgBox "La facture '" & invoice & "' n'existe pas dans FAC_Entête.", vbCritical
    End If
    
    'Clean up
    On Error Resume Next
    Set foundRange = Nothing
    Set lookupRange = Nothing
    Set ws = Nothing
    On Error GoTo 0
    
    Call Log_Record("modFAC_Confirmation:FAC_Confirmation_GL_Posting()", startTime)

End Sub

Sub FAC_Finale_Create_PDF_Email_Func(noFacture As String)

    Application.ScreenUpdating = False

    'Construct the attachmentFullPathName filename
    Dim attachmentFullPathName As String
    attachmentFullPathName = wshAdmin.Range("F5").value & FACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2024-09-03 @ 16:43

    'Check if the file already exists
    Dim fileExists As Boolean
    fileExists = Dir(attachmentFullPathName) <> ""

    'If the file exists, prompt the user for confirmation
    Dim reponse As VbMsgBoxResult
    If Not fileExists Then
        reponse = MsgBox("La facture '" & noFacture & "' n'existe pas en format PDF" & _
                          "Je ne peux donc poas l'envoyer par courriel ?", vbInformation, _
                          "Cette facture n'existe pas en format PDF")
        GoTo EndMacro
    End If

    'Set Print Quality
    On Error Resume Next
    ActiveSheet.PageSetup.PrintQuality = 600
    Err.Clear
    On Error GoTo 0

    'Adjust Document Properties - 2024-09-03 @ 16:46
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
    End With

'    'Create the PDF file and Save It
'    On Error GoTo RefLibError
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=attachmentFullPathName, Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
'    On Error GoTo 0

    'Construct & Display the Email, allowing the user to modify the Email
    On Error GoTo SaveOnly

    Dim OutlookApp As Outlook.Application: Set OutlookApp = New Outlook.Application

    'Where are the email templates ? - 2024-03-27 @ 07:28
    Dim FullTemplatePathAndFile As String
    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
        FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\GCF_Facturation.oft"
    Else
        FullTemplatePathAndFile = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Templates\GCF_Facturation.oft"
    End If

    Dim myMail As Outlook.MailItem: Set myMail = OutlookApp.CreateItemFromTemplate(FullTemplatePathAndFile)
'        Set myMail = outlookApp.CreateItem(olMailItem)

    With myMail
        .To = "robertv13@me.com"
'            .cc = "robertv13@me.com"
'            .BCC = "robertv13@gmail.com"
        .Subject = "GC FISCALITÉ INC. - Facturation"
'            .Body = "Bonjour," & vbNewLine & vbNewLine & "Vous trouverez ci-joint notre note d'honoraires." & _
            vbNewLine & vbNewLine & "Merci" & vbNewLine & vbNewLine & vbNewLine & "GCFiscalite, CPA, CA, M. Fisc." & _
            vbNewLine & "Président"
        .Attachments.add attachmentFullPathName

        .Display 'Affiche le courriel, ce qui permet de corriger AVANT l'envoi
        'myMail.Send
    End With

SaveOnly:
    GoTo EndMacro

RefLibError:
    MsgBox "Incapable de préparer le courriel. La librairie n'est pas disponible"

EndMacro:

    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set myMail = Nothing
    Set OutlookApp = Nothing

End Sub

Sub Envoyer_Email_Avec_Template_Et_Piece_Jointe(noFacture As String)

    'Chemin de la pièce jointe
    Dim attachmentFullPathName As String
    attachmentFullPathName = wshAdmin.Range("F5").value & FACT_PDF_PATH & Application.PathSeparator & _
                     noFacture & ".pdf" '2024-09-03 @ 16:43
    
    'Vérification de l'existence de la pièce jointe
    Dim fileExists As Boolean
    fileExists = Dir(attachmentFullPathName) <> ""
    If Not fileExists Then
        MsgBox "La pièce jointe (Facture en format PDF) n'existe pas" & _
                    "à l'emplacement spécifié, soit " & attachmentFullPathName, vbCritical
        GoTo Exit_sub
    End If
    
    'Chemin du modèle (template) de courriel
    Dim templatePath As String
    templatePath = Environ("appdata") & "\Microsoft\Templates\GCF_Facturation.oft"

    'Initialisation de l'application Outlook
    Dim OutlookApp As Object
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    'Création de l'email à partir du modèle
    Dim MailItem As Object
    Set MailItem = OutlookApp.CreateItemFromTemplate(templatePath)

    'Ajout de la pièce jointe
    MailItem.Attachments.add attachmentFullPathName

    'Obtenir la signature par défaut
    Dim signaturePath As String
    signaturePath = Fn_Get_Outlook_Signature_Path()
    Dim signature As String
    signature = Fn_Get_Outlook_Email_Signature()
    
'    If signaturePath <> "" Then
'        'Ajouter la signature au corps de l'email sans affecter le contenu HTML
'        MailItem.HTMLBody = MailItem.HTMLBody & "<br><br>" & ReadSignatureFile(signaturePath)
'    End If
'
'    'Ajouter la signature au corps de l'email
'    Dim templateBody As String
'    templateBody = MailItem.HTMLBody
'    MailItem.HTMLBody = templateBody & signature
    
    'Optionnel : Modifiez les éléments de l'email (comme les destinataires)
    MailItem.To = "robertv13@me.com"
    MailItem.Subject = "GCF FISCALITÉ INC. - Facturation"
    MailItem.Body = Replace(MailItem.Body, "{nom_du_destinataire}", "John Doe")

    'Ajoutez le contenu du corps et la signature
    MailItem.HTMLBody = "<p>Contenu de l'email ici...</p>" & signature
     
    'Afficher ou envoyer l'email
    MailItem.Display ' Pour afficher l'email avant envoi
    ' MailItem.Send ' Pour envoyer directement l'email

Exit_sub:
    'Nettoyage
    Set MailItem = Nothing
    Set OutlookApp = Nothing
    
End Sub

Function Fn_Get_Outlook_Signature_Path() As String
    
    Dim sPath As String
    Dim sSignaturePath As String
    Dim fso As Object
    Dim file As Object

    'Chemin vers le dossier des signatures
    sPath = Environ("appdata") & "\Microsoft\Signatures\"
    
    'Définir l'objet FileSystem
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Trouver le fichier de signature HTML
    For Each file In fso.GetFolder(sPath).Files
        If LCase(fso.GetExtensionName(file.name)) = "htm" Then
            sSignaturePath = file.path
            Exit For
        End If
    Next
    
    'Retourner le chemin de la signature
    Fn_Get_Outlook_Signature_Path = sSignaturePath
    
    Set fso = Nothing
    Set file = Nothing
End Function

Function ReadSignatureFile(filepath As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim sContent As String

    ' Définir l'objet FileSystem
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Lire le contenu de la signature
    If fso.fileExists(filepath) Then
        Set ts = fso.GetFile(filepath).OpenAsTextStream(1, -2)
        sContent = ts.ReadAll
        ts.Close
    End If
    
    ' Retourner le contenu de la signature
    ReadSignatureFile = sContent
    
    Set fso = Nothing
    Set ts = Nothing
End Function

Function Fn_Get_Outlook_Email_Signature() As String

    'Chemin vers le dossier des signatures
    Dim sPath As String
    sPath = Environ("appdata") & "\Microsoft\Signatures\"
    
    'Définir l'objet FileSystem
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Lire la signature par défaut dans le fichier .htm
    Dim sSignature As String
    Dim ts As Object
    If fso.fileExists(sPath & Dir(sPath & "*.htm")) Then
        Set ts = fso.GetFile(sPath & Dir(sPath & "*.htm")).OpenAsTextStream(1, -2)
        sSignature = ts.ReadAll
        ts.Close
    End If
    
    'Retourner la signature
    Fn_Get_Outlook_Email_Signature = sSignature
    
    Set fso = Nothing
    Set ts = Nothing
    
End Function

Sub test_Envoyer_Email_Avec_Template_Et_Piece_Jointe()

    Call Envoyer_Email_Avec_Template_Et_Piece_Jointe("24-24480")

End Sub

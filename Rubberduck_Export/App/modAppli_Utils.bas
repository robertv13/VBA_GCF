Attribute VB_Name = "modAppli_Utils"
Option Explicit

Public Sub ConvertRangeBooleanToText(rng As Range)

    Dim cell As Range
    For Each cell In rng
        Select Case cell.value
            Case 0, "False" 'False
                cell.value = "FAUX"
            Case -1, "True" 'True
                cell.value = "VRAI"
            Case "VRAI", "FAUX"
                
            Case Else
                MsgBox cell.value & " est une valeur INVALIDE pour la cellule " & cell.Address & " de la feuille TEC_Local"
        End Select
    Next cell

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    
End Sub

Public Sub ProtectCells(rng As Range)

    'Lock the checkbox
    rng.Locked = True
    
    'Protect the worksheet
    rng.Parent.Protect UserInterfaceOnly:=True


End Sub

Public Sub UnprotectCells(rng As Range)

    'Lock the checkbox
    rng.Locked = False
    
    'Protect the worksheet
    rng.Parent.Protect UserInterfaceOnly:=True


End Sub

Sub Start_Routine(subName As String) '2024-06-06 @ 10:12

    Dim modeOper As Integer
    modeOper = 2
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Integer: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & "Entering: " & subName & _
            vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A99999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName & " - entering"
        End With
    End If

End Sub

Sub Output_Timer_Results(subName As String, t As Double)

    Dim modeOper As Integer
    modeOper = 2 '2024-03-29 @ 11:37
    
    'Allows message to be used - 2024-06-06 @ 11:05
    If InStr(subName, "message:") = 1 Then
        subName = Right(subName, Len(subName) - 8)
    Else
        subName = subName & " - exiting"
    End If
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Integer: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & subName & " = " _
        & Format(Timer - t, "##0.0000") & " secondes" & vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A9999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName
            If t Then
                .Range("C" & lastUsedRow).value = Format(Round(Timer - t, 4), "##0.0000")
            End If
        End With
    End If

End Sub

Public Sub ArrayToRange(ByRef data As Variant _
                        , ByVal outRange As Range _
                        , Optional ByVal clearExistingData As Boolean = True _
                        , Optional ByVal clearExistingHeaderSize As Long = 1)
                        
    If clearExistingData = True Then
        outRange.CurrentRegion.Offset(clearExistingHeaderSize).ClearContents
    End If
    
    Dim rows As Long, columns As Long
    rows = UBound(data, 1) - LBound(data, 1) + 1
    columns = UBound(data, 2) - LBound(data, 2) + 1
    outRange.Resize(rows, columns).value = data
    
End Sub

Sub CreateOrReplaceWorksheet(wsName As String)
    
    Dim ws As Worksheet
    Dim wsExists As Boolean
    wsExists = False
    
    'Check if the worksheet exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = wsName Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    'If the worksheet exists, delete it
    If wsExists Then
        Application.DisplayAlerts = False
        ws.delete
        Application.DisplayAlerts = True
    End If
    
    'Add the new worksheet
    Set ws = ThisWorkbook.Worksheets.add
    ws.name = wsName

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Private Sub IntegrityVerification()

    Application.ScreenUpdating = False
    
    Call Erase_And_Create_Worksheet("Analyse_Intégrité")

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    wsOutput.Range("A1").value = "Feuille"
    wsOutput.Range("B1").value = "Message"
    wsOutput.Range("C1").value = "TimeStamp"
    Call Make_It_As_Header(wsOutput.Range("A1:C1"))

    Dim lastUsedRow As Long, r As Long
    lastUsedRow = wsOutput.Range("A9999").End(xlUp).row
    r = lastUsedRow + 1
    
    'wshBD_Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "clientsImportés")
    r = r + 1
    
    Call Client_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importé du fichier BD_Sortie.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_Clients(r)

    'wshBD_Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "Fournisseurs")
    r = r + 1
    
    Call Fournisseur_List_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La feuille a été importé du fichier BD_Sortie.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_Fournisseurs(r)
    
    'wshGL_Trans
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "GL_Trans")
    r = r + 1
    
    Call GL_Trans_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "GL_Trans a été importé du fichier BD_Sortie.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_GL_Trans(r)
    
    'wshTEC_Local
    Call Add_Message_To_WorkSheet(wsOutput, r, 1, "TEC_Local")
    r = r + 1
    
    Call TEC_Import_All
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "TEC_Local a été importé du fichier BD_Sortie.xlsx")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call check_TEC(r)
    
    With wsOutput.Range("A2:C" & r).Font
        .name = "Courier New"
        .Size = 10
    End With
    
    wsOutput.Range("A1").CurrentRegion.EntireColumn.AutoFit

    'Cleaning memory - 2024-07-01 @ 09:34
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_Clients(ByRef r As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshBD_Clients
    Dim ws As Worksheet: Set ws = wshBD_Clients
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshBD_Clients'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim arr As Variant
    arr = wshBD_Clients.Range("A1").CurrentRegion.value
    Dim dict_code_client As New Dictionary
    Dim dict_nom_client As New Dictionary
    
    Dim i As Long, code As String, nom As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        nom = arr(i, 1)
        code = arr(i, 2)
        If dict_nom_client.Exists(nom) = False Then
            dict_nom_client.add nom, code
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_client.Exists(code) = False Then
            dict_code_client.add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le client '" & nom & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format(UBound(arr, 1) - 1, "##,##0") & " clients ont été analysés!")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de nom")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de code")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    r = r + 1
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_Fournisseurs(ByRef r As Long)
    
    Application.ScreenUpdating = False

    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshBD_fournisseurs
    Dim ws As Worksheet: Set ws = wshBD_Fournisseurs
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshBD_Fournisseurs'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim arr As Variant
    arr = wshBD_Fournisseurs.Range("A1").CurrentRegion.value
    Dim dict_code_fournisseur As New Dictionary
    Dim dict_nom_fournisseur As New Dictionary
    
    Dim i As Long, code As String, nom As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        nom = arr(i, 1)
        code = arr(i, 2)
        If dict_nom_fournisseur.Exists(nom) = False Then
            dict_nom_fournisseur.add nom, code
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le nom '" & nom & "' est un doublon pour le code '" & code & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_fournisseur.Exists(code) = False Then
            dict_code_fournisseur.add code, nom
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Le code '" & code & "' est un doublon pour le nom '" & nom & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format(UBound(arr, 1) - 1, "#,##0") & " fournisseurs ont été analysés!")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    If cas_doublon_nom = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de nom")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_nom & " cas de doublons pour les noms")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_doublon_code = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de code")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_code & " cas de doublons pour les codes")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    r = r + 1
    
    'Cleaning memory - 2024-07-04 @ 12:37
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_GL_Trans(ByRef r As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshGL_Trans
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshGL_Trans'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim arr As Variant
    arr = wshGL_Trans.Range("A1").CurrentRegion.value
    Dim dict_GL_Entry As New Dictionary
    Dim sum_arr() As Double
    ReDim sum_arr(1 To 5000, 1 To 3)
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long, GL_Entry_No As String, dt As Double, ct As Double
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        GL_Entry_No = arr(i, 1)
        dt = arr(i, 7)
        ct = arr(i, 8)
        If dict_GL_Entry.Exists(GL_Entry_No) = False Then
            dict_GL_Entry.add GL_Entry_No, row
            sum_arr(row, 1) = GL_Entry_No
            row = row + 1
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
    Next i
    
    Dim sum_dt As Currency, sum_ct As Currency
    Dim cas_hors_balance As Long
    Dim v As Variant
    For Each v In dict_GL_Entry.items()
        GL_Entry_No = sum_arr(v, 1)
        dt = Round(sum_arr(v, 2), 2)
        ct = Round(sum_arr(v, 3), 2)
        If dt <> ct Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Écriture # " & v & " ne balance pas... Dt = " & Format(dt, "###,###,##0.00") & " et Ct = " & Format(ct, "###,###,##0.00"))
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_hors_balance = cas_hors_balance + 1
        End If
        sum_dt = sum_dt + dt
        sum_ct = sum_ct + ct
    Next v
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format(UBound(arr, 1) - 1, "##,##0") & " lignes de transactions ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Un total de " & dict_GL_Entry.count & " écritures ont été analysées")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    If cas_hors_balance = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Chacune des écritures balancent au niveau de l'écriture")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hors_balance & " écriture(s) qui ne balance(nt) pas !!!")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Les totaux des transactions sont:")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Dt = " & Format(sum_dt, "###,###,##0.00 $"))
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Ct = " & Format(sum_ct, "###,###,##0.00 $"))
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    If sum_dt - sum_ct <> 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Hors-Balance de " & Format(sum_dt - sum_ct, "###,###,##0.00$"))
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    r = r + 1
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set v = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub check_TEC(ByRef r As Long)

    Application.ScreenUpdating = False
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Worksheets("Analyse_Intégrité")
    
    'wshTEC_Local
    Dim ws As Worksheet: Set ws = wshTEC_Local
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Analyse de '" & ws.name & "' ou 'wshTEC_Local'")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim arr As Variant
    arr = wshTEC_Local.Range("A1").CurrentRegion.Offset(2)
    Dim dict_TEC_ID As New Dictionary
    Dim dict_prof As New Dictionary
    
    Dim i As Long, TECID As Long, ProfID As String, Prof As String, dateTEC As Date, testDate As Boolean
    Dim code As String, nom As String, hres As Double, testHres As Boolean, estFacturable As Boolean
    Dim estFacturee As Boolean, estDetruit As Boolean
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    
    For i = LBound(arr, 1) To UBound(arr, 1) - 2
        TECID = arr(i, 1)
        ProfID = arr(i, 2)
        Prof = arr(i, 3)
        dateTEC = arr(i, 4)
        testDate = IsDate(dateTEC)
        If testDate = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "TEC_ID =" & TECID & " a une date INVALIDE '" & dateTEC & " !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_date_invalide = cas_date_invalide + 1
        End If
        code = arr(i, 5)
        nom = arr(i, 6)
        hres = arr(i, 8)
        testHres = IsNumeric(hres)
        If testHres = False Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur des heures est INVALIDE '" & hres & " !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_hres_invalide = cas_date_invalide + 1
        End If
        estFacturable = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturable' est INVALIDE '" & estFacturable & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 12)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'EstFacturee' est INVALIDE '" & estFacturee & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 14)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** TEC_ID = " & TECID & " la valeur de la colonne 'estDetruit' est INVALIDE '" & estDetruit & "' !!!")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        total_hres_inscrites = total_hres_inscrites + hres
        If estDetruit = "Vrai" Then total_hres_detruites = total_hres_detruites + hres
        
        If estDetruit = "Faux" And estFacturable = "Vrai" Then total_hres_facturable = total_hres_facturable + hres
        If estDetruit = "Faux" And estFacturable = "Faux" Then total_hres_non_facturable = total_hres_non_facturable + hres
        If estDetruit = "Faux" And estFacturee = "Vrai" Then total_hres_facturees = total_hres_facturees + hres
        
        'Dictionary
        If dict_TEC_ID.Exists(TECID) = False Then
            dict_TEC_ID.add TECID, 0
        Else
            Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Le TEC_ID '" & TECID & "' est un doublon pour la ligne '" & i & "'")
            Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
            r = r + 1
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(Prof & "-" & ProfID) = False Then
            dict_prof.add Prof & "-" & ProfID, 0
        End If
    Next i
    
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "Un total de " & Format(UBound(arr, 1) - 2, "##,##0") & " charges de temps ont été analysées!")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    If cas_doublon_TECID = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucun doublon de TEC_ID")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_date_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune date INVALIDE")
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_date_invalide & " cas de date INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_hres_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune heures INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_hres_invalide & " cas d'heures INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_estFacturable_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturable' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_estFacturee_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estFacturee' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    If cas_estDetruit_invalide = 0 Then
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Aucune valeur 'estDetruit' n'est INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    Else
        Call Add_Message_To_WorkSheet(wsOutput, r, 2, "**** Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE")
        Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
        r = r + 1
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "La somme des heures donne ce resultat:")
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    Dim formattedHours As String
    formattedHours = Format(total_hres_inscrites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures inscrites       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format(total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures détruites       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format(total_hres_inscrites - total_hres_detruites, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures restantes       : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format(total_hres_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures facturables     : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1
    
    formattedHours = Format(total_hres_non_facturable, "#,##0.00")
    If Len(formattedHours) < 10 Then
        formattedHours = String(10 - Len(formattedHours), " ") & formattedHours
    End If
    Call Add_Message_To_WorkSheet(wsOutput, r, 2, "     Heures non_facturables : " & formattedHours)
    Call Add_Message_To_WorkSheet(wsOutput, r, 3, Format(Now(), "dd/mm/yyyy hh:mm:ss"))
    r = r + 1

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

Sub ADMIN_DataFiles_Folder_Selection() '2024-03-28 @ 14:10

    Dim SharedFolder As FileDialog: Set SharedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With SharedFolder
        .Title = "Choisir le répertoire de données partagées, selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F5").value = .selectedItems(1)
        End If
    End With
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set SharedFolder = Nothing
    
End Sub

Sub Make_It_As_Header(r As Range)

    With r
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .Size = 10
            .Italic = True
            .Bold = True
        End With
    End With
    
End Sub

Sub Add_Message_To_WorkSheet(ws As Worksheet, r As Long, c As Long, m As String)

    ws.Cells(r, c).value = m

End Sub
Sub ADMIN_PDF_Folder_Selection() '2024-03-28 @ 14:10

    Dim PDFFolder As FileDialog: Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With PDFFolder
        .Title = "Choisir le répertoire des copies de facture (PDF), selon les instructions de l'Administrateur"
        .AllowMultiSelect = False
        If .show = -1 Then
            wshAdmin.Range("F6").value = .selectedItems(1)
        End If
    End With
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set PDFFolder = Nothing

End Sub



Attribute VB_Name = "modGL_EJ"
Option Explicit

Dim sauvegardesCaracteristiquesForme As Object

Sub shp_GL_EJ_Update_Click()

    Call GL_EJ_Update
    
End Sub

Sub GL_EJ_Update()

    If wshGL_EJ.Range("F4").Value = "Renversement" Then
        Call JE_Renversement_Update
        Exit Sub
    End If
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Update", "", 0)
    
    If Fn_Is_Date_Valide(wshGL_EJ.Range("K4").Value) = False Then Exit Sub
    
    If Fn_Is_Ecriture_Balance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E23").End(xlUp).row  'Last Used Row in wshGL_EJ
    If Fn_Is_JE_Valid(rowEJLast) = False Then Exit Sub
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Call GL_Trans_Add_Record_To_DB(rowEJLast)
    Call GL_Trans_Add_Record_Locally(rowEJLast)
    
    If wshGL_EJ.ckbRecurrente = True Then
        Call Save_EJ_Recurrente(rowEJLast)
    End If
    
    'Save Current JE number
    Dim strCurrentJE As String
    strCurrentJE = wshGL_EJ.Range("B1").Value
    
    'Increment Next JE number
    wshGL_EJ.Range("B1").Value = wshGL_EJ.Range("B1").Value + 1
        
    Call GL_EJ_Clear_All_Cells
        
    With wshGL_EJ
        .Activate
        .Range("F4").Select
        .Range("F4").Activate
    End With
    
    MsgBox "L'écriture numéro '" & strCurrentJE & "' a été reporté avec succès", vbInformation, "Confirmation de traitement"
    
    Call Log_Record("modGL_EJ:GL_EJ_Update", "", startTime)
    
End Sub

Sub JE_Renversement_Update()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:JE_Renversement_Update", "", 0)
    
    If Fn_Is_Ecriture_Balance = False Then
        MsgBox "L'écriture à renverser ne balance pas !!!", vbCritical
        Exit Sub
    End If
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E23").End(xlUp).row  'Last Used Row in wshGL_EJ
    If Fn_Is_JE_Valid(rowEJLast) = False Then Exit Sub
    
    'Renverser les montants (DT --> CT & CT ---> DT)
    Application.ScreenUpdating = False
    Dim i As Integer
    For i = 9 To rowEJLast
        If wshGL_EJ.Cells(i, 8).Value <> 0 Then
            wshGL_EJ.Cells(i, 9).Value = wshGL_EJ.Cells(i, 8).Value
            wshGL_EJ.Cells(i, 8).Value = ""
        Else
            wshGL_EJ.Cells(i, 8).Value = wshGL_EJ.Cells(i, 9).Value
            wshGL_EJ.Cells(i, 9).Value = ""
        End If
    Next i
    
    wshGL_EJ.Range("F4").Value = "RENVERSEMENT:" & wshGL_Trans.Range("AA3").Value
    Dim saveDescription As String
    saveDescription = wshGL_EJ.Range("F6").Value
    wshGL_EJ.Range("F6").Value = "RENV. - " & wshGL_EJ.Range("F6").Value
    
    'Transfert des données vers wshGL, entête d'abord puis une ligne à la fois
    Call GL_Trans_Add_Record_To_DB(rowEJLast)
    Call GL_Trans_Add_Record_Locally(rowEJLast)
    
    MsgBox "L'écriture numéro '" & wshGL_Trans.Range("AA3").Value & "' a été RENVERSÉ avec succès", vbInformation, "Confirmation de traitement"
    
    Application.ScreenUpdating = True
    DoEvents
    
    'Reorganise wshGL_EJ
    Application.ScreenUpdating = False
    Dim shp As Shape
    Set shp = wshGL_EJ.Shapes("btnUpdate")
    Call Restaurer_Forme(shp)
    
    'Renverser les montants (DT --> CT & CT ---> DT)
    For i = 9 To rowEJLast
        If wshGL_EJ.Cells(i, 8).Value <> 0 Then
            wshGL_EJ.Cells(i, 9).Value = wshGL_EJ.Cells(i, 8).Value
            wshGL_EJ.Cells(i, 8).Value = ""
        Else
            wshGL_EJ.Cells(i, 8).Value = wshGL_EJ.Cells(i, 9).Value
            wshGL_EJ.Cells(i, 9).Value = ""
        End If
    Next i
    
    wshGL_EJ.Range("F4, K4, F6:k6").Font.Color = vbBlack
    wshGL_EJ.Range("E9:K23").Font.Color = vbBlack

    'Retour à la source
    wshGL_EJ.Range("F4").Value = ""
    wshGL_EJ.Range("F6").Value = saveDescription
    wshGL_EJ.Range("F4").Select
    
    Application.ScreenUpdating = True
    DoEvents
    
    'Libérer la mémoire
    Set shp = Nothing
    
    Call Log_Record("modGL_EJ:JE_Renversement_Update", "", startTime)
    
End Sub

Sub Save_EJ_Recurrente(ll As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:Save_EJ_Recurrente", "", 0)
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Cells(wshGL_EJ.Rows.count, "E").End(xlUp).row  'Last Used Row in wshGL_EJ
    
    Call GL_EJ_Recurrente_Add_Record_To_DB(ll)
    Call GL_EJ_Recurrente_Add_Record_Locally(ll)
    
    Call Log_Record("modGL_EJ:Save_EJ_Recurrente", "", startTime)
    
End Sub

Sub Load_JEAuto_Into_JE(EJAutoDesc As String, NoEJAuto As Long)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:Load_JEAuto_Into_JE", "", 0)
    
    'On copie l'E/J automatique vers wshEJ
    Dim rowJEAuto, rowJE As Long
    rowJEAuto = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.Rows.count, 1).End(xlUp).row  'Last Row used in wshGL_EJRecuurente
    
    Call GL_EJ_Clear_All_Cells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wshGL_EJ_Recurrente.Range("A" & r).Value = NoEJAuto And wshGL_EJ_Recurrente.Range("C" & r).Value <> "" Then
            wshGL_EJ.Range("E" & rowJE).Value = wshGL_EJ_Recurrente.Range("D" & r).Value
            wshGL_EJ.Range("H" & rowJE).Value = wshGL_EJ_Recurrente.Range("E" & r).Value
            wshGL_EJ.Range("I" & rowJE).Value = wshGL_EJ_Recurrente.Range("F" & r).Value
            wshGL_EJ.Range("J" & rowJE).Value = wshGL_EJ_Recurrente.Range("G" & r).Value
            wshGL_EJ.Range("L" & rowJE).Value = wshGL_EJ_Recurrente.Range("C" & r).Value
            rowJE = rowJE + 1
        End If
    Next r
    wshGL_EJ.Range("F6").Value = "[Auto]-" & EJAutoDesc
    wshGL_EJ.Range("K4").Activate

    Call Log_Record("modGL_EJ:Load_JEAuto_Into_JE", "", startTime)
    
End Sub

Sub GL_EJ_Clear_All_Cells()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Clear_All_Cells", "", 0)
    
    'Efface toutes les cellules de la feuille
    Application.EnableEvents = False
    ActiveSheet.Unprotect
    With wshGL_EJ
        .Range("B6").ClearContents 'Code de client
        .Range("F4,F6:K6").ClearContents
        .Range("F4, K4, F6:K6").Font.Color = vbBlack
        .Range("E9:K23").ClearContents
        .Range("E9:K23").Font.Color = vbBlack
'        .Range("E9:G23,H9:H23,I9:I23,J9:L23").ClearContents
        .ckbRecurrente = False
        .Range("E6").Value = "Description:"
        Application.EnableEvents = True
        wshGL_EJ.Activate
        wshGL_EJ.Range("F4").Select
    End With
    
    'Envlève la validation sur la cellule description/client
    Dim cell As Range
    Set cell = wshGL_EJ.Range("F6:K6")
    Call AnnulerValidation(cell)
    
    With ActiveSheet
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    'Libérer la mémoire
    Set cell = Nothing
    
    Call Log_Record("modGL_EJ:GL_EJ_Clear_All_Cells", "", startTime)

End Sub

Sub GL_EJ_Construire_Remise_TPS_TVQ(r As Integer)

    Dim dateFin As Date
    dateFin = CDate(wshGL_EJ.Range("K4").Value)
    
    'Remplir la description, si elle est vide
    If wshGL_EJ.Range("F6").Value = "" Then
        wshGL_EJ.Range("F6").Value = "Déclaration TPS/TVQ - Du " & _
            Format$(Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arrière(dateFin), wshAdmin.Range("B1").Value) & " au " & _
            Format$(dateFin, wshAdmin.Range("B1").Value)
    End If
    
    Dim cases() As Currency
    ReDim cases(101 To 213)
    
    'Remplir le formulaire de déclaration
    wshGL_EJ.Range("T5").Value = "du " & Format$(Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arrière(dateFin), wshAdmin.Range("B1").Value)
    wshGL_EJ.Range("V5").Value = "du " & Format$(Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arrière(dateFin), wshAdmin.Range("B1").Value)
    wshGL_EJ.Range("T6").Value = "du " & Format$(dateFin, wshAdmin.Range("B1").Value)
    wshGL_EJ.Range("V6").Value = "du " & Format$(dateFin, wshAdmin.Range("B1").Value)
    
    Dim rngResultAF As Range
    Call GL_Get_Account_Trans_AF(ObtenirNoGlIndicateur("Revenus de consultation"), Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arrière(dateFin), dateFin, rngResultAF)
    cases(101) = -Application.WorksheetFunction.Sum(rngResultAF.Columns(7)) _
                    - Application.WorksheetFunction.Sum(rngResultAF.Columns(8))

    With wshGL_EJ.Range("P10")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .Value = -cases(101)
    End With
    
    'TPS percues
    cases(105) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TPS Facturée"), dateFin)
    wshGL_EJ.Range("E" & r).Value = "TPS percues"
    If cases(105) <= 0 Then
        wshGL_EJ.Range("H" & r).Value = -cases(105)
    Else
        wshGL_EJ.Range("I" & r).Value = cases(105)
    End If
    r = r + 1
    With wshGL_EJ.Range("T10")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .Value = -cases(105)
    End With
    
    'TVQ percues
    cases(205) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TVQ Facturée"), dateFin)
    wshGL_EJ.Range("E" & r).Value = "TVQ percues"
    If cases(205) <= 0 Then
        wshGL_EJ.Range("H" & r).Value = -cases(205)
    Else
        wshGL_EJ.Range("I" & r).Value = cases(205)
    End If
    r = r + 1
    With wshGL_EJ.Range("V10")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .Value = -cases(205)
    End With
    
    cases(108) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TPS Payée"), dateFin)
    wshGL_EJ.Range("E" & r).Value = "TPS payées"
    If cases(108) <= 0 Then
        wshGL_EJ.Range("H" & r).Value = -cases(108)
    Else
        wshGL_EJ.Range("I" & r).Value = cases(108)
    End If
    r = r + 1
    With wshGL_EJ.Range("T13")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .Value = cases(108)
    End With
    
    cases(208) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TVQ Payée"), dateFin)
    wshGL_EJ.Range("E" & r).Value = "TVQ payées"
    If cases(208) <= 0 Then
        wshGL_EJ.Range("H" & r).Value = -cases(208)
    Else
        wshGL_EJ.Range("I" & r).Value = cases(208)
    End If
    r = r + 1
    With wshGL_EJ.Range("V13")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .Value = cases(208)
    End With
    
    cases(113) = -cases(105) - cases(108)
    With wshGL_EJ.Range("T16")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .Value = cases(113)
    End With
    
    cases(213) = -cases(205) - cases(208)
    With wshGL_EJ.Range("V16")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .Value = cases(213)
    End With
    
    Dim net As Double
    If cases(113) + cases(213) > 0 Then
        With wshGL_EJ.Range("X14")
            .Font.Bold = True
            .Font.size = 12
            .NumberFormat = "###,##0.00 $"
            .HorizontalAlignment = xlRight
            .Value = cases(113) + cases(213)
        End With
        net = cases(113) + cases(213)
    Else
        With wshGL_EJ.Range("X10")
            .Font.Bold = True
            .Font.size = 12
            .NumberFormat = "###,##0.00 $"
            .HorizontalAlignment = xlRight
            .Value = -(cases(113) + cases(213))
        End With
        net = -(cases(113) + cases(213))
    End If
    
    'Encaisse
    wshGL_EJ.Range("E" & r).Value = "Encaisse"
    If net <= 0 Then
        wshGL_EJ.Range("H" & r).Value = -net
    Else
        wshGL_EJ.Range("I" & r).Value = net
    End If
    r = r + 1
    
    With wshGL_EJ
        .Unprotect
        .Range("N:Y").EntireColumn.Hidden = False
    End With

End Sub

Sub GL_EJ_Renverser_Ecriture()

    Dim ws As Worksheet: Set ws = wshGL_Trans
    
    '1. Demande le numéro d'écriture à partir d'un ListBox
    Call Prepare_Affiche_Liste_Ecriture
    Dim no_Ecriture As Long
    If Range("B3").Value <> -1 Then
        no_Ecriture = Range("B3").Value
    Else
        MsgBox "Vous n'avez sélectionné aucune écriture à renverser", vbInformation, "Sélection d'une écriture à renverser"
        wshGL_EJ.Range("F4").Value = ""
        wshGL_EJ.Range("F4").Select
        Exit Sub
    End If
    
    '2. Affiche l'écriture à renverser
    Call GL_Get_JE_Detail_Trans_AF(no_Ecriture)
    Dim lastUsedRowResult As Long
    lastUsedRowResult = ws.Cells(ws.Rows.count, "AC").End(xlUp).row
    If lastUsedRowResult < 2 Then
        MsgBox "Je ne retrouve pas l'écriture '" & no_Ecriture & "'" & vbNewLine & vbNewLine & _
                "Veuillez vérifier votre numéro et reessayez", vbInformation, "Numéro d'écriture invalide"
        Exit Sub
    End If
    Dim rngResult As Range
    Set rngResult = ws.Range("AC1").CurrentRegion.offset(1, 0)
    If InStr(rngResult.Cells(1, 4).Value, "ENCAISSEMENT:") <> 0 Or _
        InStr(rngResult.Cells(1, 4).Value, "DÉBOURSÉ:") <> 0 Or _
        InStr(rngResult.Cells(1, 4).Value, "FACTURE:") <> 0 Or _
        InStr(rngResult.Cells(1, 4).Value, "RENVERSEMENT:") <> 0 Then
        MsgBox "Je ne peux renverser ce type d'écriture '" & _
                Left(rngResult.Cells(1, 4).Value, InStr(rngResult.Cells(1, 4).Value, ":") - 1) & _
                "'" & vbNewLine & vbNewLine & _
                "Veuillez vérifier votre numéro et reessayez", _
                vbInformation, "Type d'écriture impossible à renverser"
        wshGL_EJ.Range("F4").Value = ""
        wshGL_EJ.Range("F4").Select
        Exit Sub
    End If
    
    'Cette écriture a-t-elle déjà été RENVERSÉE ?
    Dim rng As Range
    Set rng = ws.Columns("D")
    Dim trouve As Range
    Set trouve = rng.Find(What:="RENVERSEMENT:" & no_Ecriture, LookIn:=xlValues, LookAt:=xlWhole)
    If Not trouve Is Nothing Then
        MsgBox "Cette écriture a déjà été RENVERSÉE..." & vbNewLine & vbNewLine & _
               "Avec le numéro d'écriture '" & ws.Cells(trouve.row, 1).Value & "'" & vbNewLine & vbNewLine & _
               "En date du " & Format$(ws.Cells(trouve.row, 2).Value, wshAdmin.Range("B1").Value) & ".", vbInformation
        Exit Sub
    End If
    
    Application.EnableEvents = False
    wshGL_EJ.Range("K4").Value = Format$(rngResult.Cells(1, 2).Value, wshAdmin.Range("B1").Value)
    wshGL_EJ.Range("F6").Value = rngResult.Cells(1, 3).Value
    Dim ligne As Range
    Dim l As Long: l = 9
    For Each ligne In rngResult.Rows
        wshGL_EJ.Range("E" & l).Value = ligne.Cells(6).Value
        If ligne.Cells(7).Value <> 0 Then
            wshGL_EJ.Range("H" & l).Value = ligne.Cells(7).Value
        End If
        If ligne.Cells(8).Value <> 0 Then
            wshGL_EJ.Range("I" & l).Value = ligne.Cells(8).Value
        End If
        wshGL_EJ.Range("J" & l).Value = ligne.Cells(9).Value
        wshGL_EJ.Range("L" & l).Value = ligne.Cells(5).Value
        l = l + 1
    Next ligne
    Application.EnableEvents = True
    
    'On affiche l'écriture à renverser en rouge
    wshGL_EJ.Range("F4, K4, F6:k6").Font.Color = vbRed
    wshGL_EJ.Range("E9:K23").Font.Color = vbRed
    
    'Change le libellé du Bouton & caractéristiques
    Dim shp As Shape
    Set shp = wshGL_EJ.Shapes("btnUpdate")
    Call Modifier_Forme(shp)
    
    'Libérer la mémoire
    Set ligne = Nothing
    Set rngResult = Nothing
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub GL_EJ_Depot_Client()

    Dim ws As Worksheet: Set ws = wshGL_EJ
    
    'Ajuster le formulaire
    ws.Range("E6").Value = "Client:"
            
    'Ajouter la validation des données
    Dim cell As Range: Set cell = wshGL_EJ.Range("F6:K6")
    
    Dim condition As Boolean
    condition = (wshGL_EJ.Range("F4").Value = "Dépôt de client")
    
    Call GérerValidation(cell, "dnrClients_Search_Field_Only", condition)
    
    'Force l'écriture
    wshGL_EJ.Range("E9").Value = "Encaisse"
    wshGL_EJ.Range("E10").Value = "Produit perçu d'avance"
    
    'Saisie du montant du dépôt
    wshGL_EJ.Range("K4").Select

    'Libérer les objects
    Set cell = Nothing
    Set ws = Nothing
    
End Sub

Sub GérerValidation(cell As Range, nomPlage As String, condition As Boolean)
    
    If condition Then
        'Condition remplie, appliquer la validation de liste
        Call AjouterValidation(cell, nomPlage)
    Else
        'Condition non remplie, supprimer la validation
        Call AnnulerValidation(cell)
    End If
    
End Sub

Sub AjouterValidation(cell As Range, nomPlage As String)

    Dim ws As Worksheet: Set ws = wshGL_EJ
    
    Dim feuilleProtégée As Boolean
    feuilleProtégée = ws.ProtectContents
    
    If feuilleProtégée Then ws.Unprotect
    
    On Error Resume Next
    cell.Validation.Delete 'Supprimer toute validation existante
    On Error GoTo 0
    
    'Ajouter la validation de données
    cell.Validation.Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, _
                        Formula1:="=" & ThisWorkbook.Names(nomPlage).Name

    'Configurer les propriétés de la validation de données
    If Not cell.Validation Is Nothing Then
        cell.Validation.IgnoreBlank = True
        cell.Validation.InCellDropdown = True
        cell.Validation.ShowInput = True
        cell.Validation.ShowError = True
    End If
    
    If feuilleProtégée Then
        With ws
            .Protect UserInterfaceOnly:=True
            .EnableSelection = xlUnlockedCells
        End With
    End If
    
    'Libérer la mémoire
    Set ws = Nothing
    
End Sub

Sub AnnulerValidation(cell As Range)

    cell.Validation.Delete
    
End Sub

Sub GL_EJ_Recurrente_Build_Summary()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Build_Summary", "", 0)
    
    'Build the summary at column K & L
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.Rows.count, 1).End(xlUp).row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.Rows.count, "J").End(xlUp).row
    If lastUsedRow2 > 1 Then
        wshGL_EJ_Recurrente.Range("J2:K" & lastUsedRow2).Clear
    End If
    
    With wshGL_EJ_Recurrente
        Dim i As Long, k As Long, oldEntry As String
        k = 2
        For i = 2 To lastUsedRow1
            If .Range("A" & i).Value <> oldEntry Then
                .Range("J" & k).Value = .Range("B" & i).Value
                .Range("K" & k).Value = "'" & Fn_Pad_A_String(.Range("A" & i).Value, " ", 5, "L")
                oldEntry = .Range("A" & i).Value
                k = k + 1
            End If
        Next i
    End With

    Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Build_Summary", "", startTime)

End Sub

Sub GL_Get_JE_Detail_Trans_AF(noEJ As Long) '2024-11-17 @ 12:08

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_Get_JE_Detail_Trans_AF", "", 0)

    Dim ws As Worksheet: Set ws = wshGL_Trans
    
    'wshGL_Trans_AF#2

    'Effacer les données de la dernière utilisation
    ws.Range("AA6:AA10").ClearContents
    ws.Range("AA6").Value = "Dernière utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'Définir le range pour la source des données en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
    ws.Range("AA7").Value = rngData.Address
    
    'Définir le range des critères
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("AA2:AA3")
    ws.Range("AA3").Value = noEJ
    ws.Range("AA8").Value = rngCriteria.Address
    
    'Définir le range des résultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("AC1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("AC1:AK1")
    ws.Range("AA9").Value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les résultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AC").End(xlUp).row
    ws.Range("AA10").Value = lastUsedRow - 1 & " lignes"

    'On tri les résultats par noGL / par date?
    If lastUsedRow > 2 Then
        With ws.Sort 'Sort - ID, Date, TecID
            .SortFields.Clear
            'First sort On noGL
            .SortFields.Add key:=ws.Range("AG2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On Date
            .SortFields.Add key:=ws.Range("AD2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wshTEC_Local.Range("AC2:AK" & lastUsedRow)
            .Apply 'Apply Sort
         End With
    End If
    
    'Libérer la mémoire
    Set rngCriteria = Nothing
    Set rngData = Nothing
    Set rngResult = Nothing
    Set ws = Nothing
    
    Call Log_Record("modGL_EJ:GL_Get_JE_Detail_Trans_AF", "", startTime)

End Sub

Sub GL_Trans_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_Trans_Add_Record_To_DB", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(NoEntrée) AS MaxEJNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim MaxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").Value
    End If
    
    'Calculate the new JE number
    Dim nextJENo As Long
    nextJENo = lastJE + 1
    wshGL_EJ.Range("B1").Value = nextJENo
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Read all line from Journal Entry
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields(fGlTNoEntrée - 1).Value = nextJENo
            rs.Fields(fGlTDate - 1).Value = Format$(CDate(wshGL_EJ.Range("K4").Value), "yyyy-mm-dd")
            If wshGL_EJ.Range("F4").Value <> "Dépôt de client" Then
                rs.Fields(fGlTDescription - 1).Value = wshGL_EJ.Range("F6").Value
                rs.Fields(fGlTSource - 1).Value = wshGL_EJ.Range("F4").Value
            Else
                rs.Fields(fGlTDescription - 1).Value = "Client:" & wshGL_EJ.Range("B6").Value & " - " & wshGL_EJ.Range("F6").Value
                rs.Fields(fGlTSource - 1).Value = UCase(wshGL_EJ.Range("F4").Value)
            End If
            rs.Fields(fGlTNoCompte - 1).Value = wshGL_EJ.Range("L" & l).Value
            rs.Fields(fGlTCompte - 1).Value = wshGL_EJ.Range("E" & l).Value
            If wshGL_EJ.Range("H" & l).Value <> "" <> 0 Then
                rs.Fields(fGlTDébit - 1).Value = CDbl(Replace(wshGL_EJ.Range("H" & l).Value, ".", ","))
            End If
            If wshGL_EJ.Range("I" & l).Value <> "" Then
                rs.Fields(fGlTCrédit - 1).Value = CDbl(Replace(wshGL_EJ.Range("I" & l).Value, ".", ","))
            End If
            rs.Fields(fGlTAutreRemarque - 1).Value = wshGL_EJ.Range("J" & l).Value
            rs.Fields(fGlTTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_To_DB", "", startTime)

End Sub

Sub GL_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshGL_EJ.Range("B1").Value
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_Trans.Cells(wshGL_Trans.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally - r = " & r, -1)
    
    Dim i As Long
    For i = 9 To r
        wshGL_Trans.Range("A" & rowToBeUsed).Value = JENo
        wshGL_Trans.Range("B" & rowToBeUsed).Value = CDate(wshGL_EJ.Range("K4").Value)
        If wshGL_EJ.Range("F4").Value <> "Dépôt de client" Then
            wshGL_Trans.Range("C" & rowToBeUsed).Value = wshGL_EJ.Range("F6").Value
            wshGL_Trans.Range("D" & rowToBeUsed).Value = wshGL_EJ.Range("F4").Value
        Else
            wshGL_Trans.Range("C" & rowToBeUsed) = "Client:" & wshGL_EJ.Range("B6").Value & " - " & wshGL_EJ.Range("F6").Value
            wshGL_Trans.Range("D" & rowToBeUsed).Value = UCase(wshGL_EJ.Range("F4").Value)
        End If
        wshGL_Trans.Range("E" & rowToBeUsed).Value = wshGL_EJ.Range("L" & i).Value
        wshGL_Trans.Range("F" & rowToBeUsed).Value = wshGL_EJ.Range("E" & i).Value
        If wshGL_EJ.Range("H" & i).Value <> "" Then
            wshGL_Trans.Range("G" & rowToBeUsed).Value = wshGL_EJ.Range("H" & i).Value
        End If
        If wshGL_EJ.Range("I" & i).Value <> "" Then
            wshGL_Trans.Range("H" & rowToBeUsed).Value = wshGL_EJ.Range("I" & i).Value
        End If
        wshGL_Trans.Range("I" & rowToBeUsed).Value = wshGL_EJ.Range("J" & i).Value
        wshGL_Trans.Range("J" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
        
        Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally - i = " & i, -1)

    Next i
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally", "", startTime)

End Sub

Sub GL_EJ_Recurrente_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Add_Record_To_DB", "", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_EJ_Récurrente$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String, MaxEJANo As Long
    strSQL = "SELECT MAX(NoEjR) AS MaxEJANo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim lastEJA As Long, nextEJANo As Long
    If IsNull(rs.Fields("MaxEJANo").Value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastEJA = 1
    Else
        lastEJA = rs.Fields("MaxEJANo").Value
    End If
    
    'Calculate the new ID
    nextEJANo = lastEJA + 1
    wshGL_EJ_Recurrente.Range("B2").Value = nextEJANo

    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields(fGlEjRNoEjR - 1).Value = nextEJANo
            rs.Fields(fGlEjRDescription - 1).Value = Replace(wshGL_EJ.Range("F6").Value, "[Auto]-", "")
            rs.Fields(fGlEjRNoCompte - 1).Value = wshGL_EJ.Range("L" & l).Value
            rs.Fields(fGlEjRCompte - 1).Value = wshGL_EJ.Range("E" & l).Value
            If wshGL_EJ.Range("H" & l).Value <> "" Then
                rs.Fields(fGlEjRDébit - 1).Value = CDbl(Replace(wshGL_EJ.Range("H" & l).Value, ".", ","))
            End If
            If wshGL_EJ.Range("I" & l).Value <> "" Then
                rs.Fields(fGlEjRCrédit - 1).Value = CDbl(Replace(wshGL_EJ.Range("I" & l).Value, ".", ","))
            End If
            rs.Fields(fGlEjRAutreRemarque - 1).Value = wshGL_EJ.Range("J" & l).Value
            rs.Fields(fGlEjRTimeStamp - 1).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        rs.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Add_Record_To_DB", "", startTime)

End Sub

Sub GL_EJ_Recurrente_Add_Record_Locally(r As Long) 'Write records to local file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Add_Record_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshGL_EJ_Recurrente.Range("B2").Value
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wshGL_EJ_Recurrente.Cells(wshGL_EJ_Recurrente.Rows.count, "C").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wshGL_EJ_Recurrente.Range("A" & rowToBeUsed).Value = JENo
        wshGL_EJ_Recurrente.Range("B" & rowToBeUsed).Value = Replace(wshGL_EJ.Range("F6").Value, "[Auto]-", "")
        wshGL_EJ_Recurrente.Range("C" & rowToBeUsed).Value = wshGL_EJ.Range("L" & i).Value
        wshGL_EJ_Recurrente.Range("D" & rowToBeUsed).Value = wshGL_EJ.Range("E" & i).Value
        If wshGL_EJ.Range("H" & i).Value <> "" Then
            wshGL_EJ_Recurrente.Range("E" & rowToBeUsed).Value = wshGL_EJ.Range("H" & i).Value
        End If
        If wshGL_EJ.Range("I" & i).Value <> "" Then
            wshGL_EJ_Recurrente.Range("F" & rowToBeUsed).Value = wshGL_EJ.Range("I" & i).Value
        End If
        wshGL_EJ_Recurrente.Range("G" & rowToBeUsed).Value = wshGL_EJ.Range("J" & i).Value
        wshGL_EJ_Recurrente.Range("H" & rowToBeUsed).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        
        rowToBeUsed = rowToBeUsed + 1
    Next i
    
    Call GL_EJ_Recurrente_Build_Summary '2024-03-14 @ 07:40
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Add_Record_Locally", "", startTime)
    
End Sub

Sub shp_EJ_Exit_Click()

    Call GL_EJ_Back_To_Menu

End Sub

Sub GL_EJ_Back_To_Menu()
    
    Dim shp As Shape
    Set shp = wshGL_EJ.Shapes("btnUpdate")
    Call Restaurer_Forme(shp)

    'Nouvelle façon de faire
    wshGL_EJ.Visible = xlSheetVeryHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
    'Libérer la mémoire
    Set shp = Nothing
    
End Sub

Sub Sauvegarder_Forme(forme As Shape)

    'Initialiser le Dictionary pour sauvegarder les caractéristiques
    Set sauvegardesCaracteristiquesForme = CreateObject("Scripting.Dictionary")

    'Définir la feuille et la forme
    Dim ws As Worksheet: Set ws = wshGL_EJ

    'Sauvegarder les caractéristiques originales de la forme
    sauvegardesCaracteristiquesForme("Left") = forme.Left
    sauvegardesCaracteristiquesForme("Width") = forme.Width
    sauvegardesCaracteristiquesForme("Height") = forme.Height
    sauvegardesCaracteristiquesForme("FillColor") = forme.Fill.ForeColor.RGB
    sauvegardesCaracteristiquesForme("LineColor") = forme.Line.ForeColor.RGB
    sauvegardesCaracteristiquesForme("Text") = forme.TextFrame2.TextRange.Text
    sauvegardesCaracteristiquesForme("TextColor") = forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    
    'Libérer la mémoire
'    Set sauvegardesCaracteristiquesForme = Nothing
    Set ws = Nothing
    
End Sub

Sub Modifier_Forme(forme As Shape)

    'Appliquer des modifications à la forme
    Application.ScreenUpdating = True
    forme.Left = 470
    forme.Width = 175
    forme.Height = 27
    forme.Fill.ForeColor.RGB = RGB(255, 0, 0)  ' Rouge
    forme.Line.ForeColor.RGB = RGB(255, 255, 255) ' Noir
    forme.TextFrame2.TextRange.Text = "Renversement"
    forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    DoEvents
    Application.ScreenUpdating = False
    
End Sub

Sub Restaurer_Forme(forme As Shape)

    'Vérifiez si les caractéristiques originales sont sauvegardées
    If sauvegardesCaracteristiquesForme Is Nothing Then
        Exit Sub
    End If

    'Restaurer les caractéristiques de la forme
    forme.Left = sauvegardesCaracteristiquesForme("Left")
    forme.Width = sauvegardesCaracteristiquesForme("Width")
    forme.Height = sauvegardesCaracteristiquesForme("Height")
    forme.Fill.ForeColor.RGB = sauvegardesCaracteristiquesForme("FillColor")
    forme.Line.ForeColor.RGB = sauvegardesCaracteristiquesForme("LineColor")
    forme.TextFrame2.TextRange.Text = sauvegardesCaracteristiquesForme("Text")
    forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = sauvegardesCaracteristiquesForme("TextColor")

End Sub

Sub Prepare_Affiche_Liste_Ecriture()

    'Prepare la liste des écritures au G/L
    Dim ws As Worksheet: Set ws = wshGL_Trans
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    'Initialiser le tableau des résultats
    Dim resultats() As Variant
    Dim compteur As Long
    ReDim resultats(1 To lastUsedRow, 1 To 4)
    
    Dim strDejaVu As String
    
    Dim cell As Range
    For Each cell In ws.Range("D2:D" & lastUsedRow)
        If cell.Value = "" And InStr(strDejaVu, ws.Cells(cell.row, 1).Value) = 0 Then
            compteur = compteur + 1
            resultats(compteur, 1) = ws.Cells(cell.row, 1).Value
            resultats(compteur, 2) = Format$(ws.Cells(cell.row, 2).Value, wshAdmin.Range("B1").Value)
            resultats(compteur, 3) = ws.Cells(cell.row, 3).Value
            resultats(compteur, 4) = Format$(ws.Cells(cell.row, 10).Value, wshAdmin.Range("B1").Value & " hh:mm:ss")
            strDejaVu = strDejaVu & ws.Cells(cell.row, 1).Value
        End If
    Next cell
    
    'Est-ce que nous avons des résultats
    If compteur = 0 Then
        MsgBox "Aucune écriture à renverser.", vbInformation
        Exit Sub
    End If
   
    'Réduire la taille du tableau resultats
    Call Array_2D_Resizer(resultats, compteur, UBound(resultats, 2))
    
    'Charger les résultats dans la ListBox
    With ufListeÉcritureGL.lbListeÉcritureGL
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "35;62;310;92"
        .List = resultats
    End With
    
    Dim i As Long
    i = 1
    Do While i <= compteur
        'Ajouter chaque ligne au ListBox
        ufListeÉcritureGL.lbListeÉcritureGL.AddItem resultats(i, 1)
        ufListeÉcritureGL.lbListeÉcritureGL.List(ufListeÉcritureGL.lbListeÉcritureGL.ListCount - 1, 1) = resultats(i, 2)
        ufListeÉcritureGL.lbListeÉcritureGL.List(ufListeÉcritureGL.lbListeÉcritureGL.ListCount - 1, 2) = resultats(i, 3)
        ufListeÉcritureGL.lbListeÉcritureGL.List(ufListeÉcritureGL.lbListeÉcritureGL.ListCount - 1, 3) = resultats(i, 4)
        i = i + 1
    Loop

    'Déplacer le focus sur la dernière ligne
    If ufListeÉcritureGL.lbListeÉcritureGL.ListCount > 0 Then
        ufListeÉcritureGL.lbListeÉcritureGL.ListIndex = ufListeÉcritureGL.lbListeÉcritureGL.ListCount - 1
    End If
    
    'Afficher le UserForm
    ufListeÉcritureGL.show
    
End Sub

Sub ckbRecurrente_Click()

    If wshGL_EJ.ckbRecurrente.Value = True Then
        wshGL_EJ.ckbRecurrente.BackColor = COULEUR_SAISIE
    Else
        wshGL_EJ.ckbRecurrente.BackColor = RGB(217, 217, 217)
    End If

End Sub

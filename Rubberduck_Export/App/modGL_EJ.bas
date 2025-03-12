Attribute VB_Name = "modGL_EJ"
Option Explicit

Private gSauvegardesCaracteristiquesForme As Object
Private gNumeroEcritureARenverser As Long

Sub shp_GL_EJ_Update_Click()

    Call GL_EJ_Update
    
End Sub

Sub GL_EJ_Update()

    If wshGL_EJ.Range("F4").value = "Renversement" Then
        Call JE_Renversement_Update
        Exit Sub
    End If
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Update", "", 0)
    
    If Fn_Is_Date_Valide(wshGL_EJ.Range("K4").value) = False Then Exit Sub
    
    If Fn_Is_Ecriture_Balance = False Then Exit Sub
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E23").End(xlUp).row  'Last Used Row in wshGL_EJ
    If Fn_Is_JE_Valid(rowEJLast) = False Then Exit Sub
    
    'Transfert des donn�es vers wshGL, ent�te d'abord puis une ligne � la fois
    Call GL_Trans_Add_Record_To_DB(rowEJLast)
    Call GL_Trans_Add_Record_Locally(rowEJLast)
    
    If wshGL_EJ.ckbRecurrente = True Then
        Call Save_EJ_Recurrente(rowEJLast)
    End If
    
    'Save Current JE number
    Dim strCurrentJE As String
    strCurrentJE = wshGL_EJ.Range("B1").value
    
    'Increment Next JE number
    wshGL_EJ.Range("B1").value = wshGL_EJ.Range("B1").value + 1
        
    Call GL_EJ_Clear_All_Cells
        
    With wshGL_EJ
        .Activate
        .Range("F4").Select
        .Range("F4").Activate
    End With
    
    MsgBox "L'�criture num�ro '" & strCurrentJE & "' a �t� report� avec succ�s", vbInformation, "Confirmation de traitement"
    
    Call Log_Record("modGL_EJ:GL_EJ_Update", "", startTime)
    
End Sub

Sub JE_Renversement_Update()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:JE_Renversement_Update", "", 0)
    
    If Fn_Is_Ecriture_Balance = False Then
        MsgBox "L'�criture � renverser ne balance pas !!!", vbCritical
        Exit Sub
    End If
    
    Dim rowEJLast As Long
    rowEJLast = wshGL_EJ.Range("E23").End(xlUp).row  'Last Used Row in wshGL_EJ
    If Fn_Is_JE_Valid(rowEJLast) = False Then Exit Sub
    
    'Renverser les montants (DT --> CT & CT ---> DT)
    Application.ScreenUpdating = False
    Dim i As Integer
    For i = 9 To rowEJLast
        If wshGL_EJ.Cells(i, 8).value <> 0 Then
            wshGL_EJ.Cells(i, 9).value = wshGL_EJ.Cells(i, 8).value
            wshGL_EJ.Cells(i, 8).value = ""
        Else
            wshGL_EJ.Cells(i, 8).value = wshGL_EJ.Cells(i, 9).value
            wshGL_EJ.Cells(i, 9).value = ""
        End If
    Next i
    
    gNumeroEcritureARenverser = wsdGL_Trans.Range("AA3").value
    
    wshGL_EJ.Range("F4").value = "RENVERSEMENT:" & gNumeroEcritureARenverser
    Dim saveDescription As String
    saveDescription = wshGL_EJ.Range("F6").value
    wshGL_EJ.Range("F6").value = "RENV. - " & wshGL_EJ.Range("F6").value
    
    'Transfert des donn�es vers wshGL, ent�te d'abord puis une ligne � la fois
    Call GL_Trans_Add_Record_To_DB(rowEJLast)
    Call GL_Trans_Add_Record_Locally(rowEJLast)
    
    'Indiquer dans l'�criture originale qu'elle a �t� renvers�e par
    Call EJ_Trans_Update_Ecriture_Renversee_To_DB
    Call EJ_Trans_Update_Ecriture_Renversee_Locally
    
    MsgBox _
        Prompt:="L'�criture num�ro '" & gNumeroEcritureARenverser & "' a �t� RENVERS�E avec succ�s", _
        Title:="Confirmation de traitement", _
        Buttons:=vbInformation

    Application.ScreenUpdating = True
    DoEvents
    
    'Reorganise wshGL_EJ
    Application.ScreenUpdating = False
    Dim shp As Shape
    Set shp = wshGL_EJ.Shapes("btnUpdate")
    Call GL_EJ_Forme_Restaurer(shp)
    
    'Renverser les montants (DT --> CT & CT ---> DT)
    For i = 9 To rowEJLast
        If wshGL_EJ.Cells(i, 8).value <> 0 Then
            wshGL_EJ.Cells(i, 9).value = wshGL_EJ.Cells(i, 8).value
            wshGL_EJ.Cells(i, 8).value = ""
        Else
            wshGL_EJ.Cells(i, 8).value = wshGL_EJ.Cells(i, 9).value
            wshGL_EJ.Cells(i, 9).value = ""
        End If
    Next i
    
    wshGL_EJ.Range("F4, K4, F6:k6").Font.Color = vbBlack
    wshGL_EJ.Range("E9:K23").Font.Color = vbBlack

    'Retour � la source
    wshGL_EJ.Range("F4").value = ""
    wshGL_EJ.Range("F6").value = saveDescription
    wshGL_EJ.Range("F4").Select
    
    Application.ScreenUpdating = True
    DoEvents
    
    'Lib�rer la m�moire
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
    Dim rowJEAuto As Long, rowJE As Long
    rowJEAuto = wsdGL_EJ_Recurrente.Cells(wsdGL_EJ_Recurrente.Rows.count, 1).End(xlUp).row  'Last Row used in wshGL_EJRecuurente
    
    Call GL_EJ_Clear_All_Cells
    rowJE = 9
    
    Dim r As Long
    For r = 2 To rowJEAuto
        If wsdGL_EJ_Recurrente.Range("A" & r).value = NoEJAuto And wsdGL_EJ_Recurrente.Range("C" & r).value <> "" Then
            wshGL_EJ.Range("E" & rowJE).value = wsdGL_EJ_Recurrente.Range("D" & r).value
            wshGL_EJ.Range("H" & rowJE).value = wsdGL_EJ_Recurrente.Range("E" & r).value
            wshGL_EJ.Range("I" & rowJE).value = wsdGL_EJ_Recurrente.Range("F" & r).value
            wshGL_EJ.Range("J" & rowJE).value = wsdGL_EJ_Recurrente.Range("G" & r).value
            wshGL_EJ.Range("L" & rowJE).value = wsdGL_EJ_Recurrente.Range("C" & r).value
            rowJE = rowJE + 1
        End If
    Next r
    wshGL_EJ.Range("F6").value = "[Auto]-" & EJAutoDesc
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
        .Range("E6").value = "Description:"
        Application.EnableEvents = True
        wshGL_EJ.Activate
        wshGL_EJ.Range("F4").Select
    End With
    
    'Envl�ve la validation sur la cellule description/client
    Dim cell As Range
    Set cell = wshGL_EJ.Range("F6:K6")
    Call AnnulerValidation(cell)
    
    With ActiveSheet
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    'Lib�rer la m�moire
    Set cell = Nothing
    
    Call Log_Record("modGL_EJ:GL_EJ_Clear_All_Cells", "", startTime)

End Sub

Sub GL_EJ_Construire_Remise_TPS_TVQ(r As Integer)

    Dim dateFin As Date
    dateFin = CDate(wshGL_EJ.Range("K4").value)
    
    'Remplir la description, si elle est vide
    If wshGL_EJ.Range("F6").value = "" Then
        wshGL_EJ.Range("F6").value = "D�claration TPS/TVQ - Du " & _
            Format$(Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arri�re(dateFin), wsdADMIN.Range("B1").value) & " au " & _
            Format$(dateFin, wsdADMIN.Range("B1").value)
    End If
    
    Dim cases() As Currency
    ReDim cases(101 To 213)
    
    'Remplir le formulaire de d�claration
    wshGL_EJ.Range("T5").value = "du " & Format$(Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arri�re(dateFin), wsdADMIN.Range("B1").value)
    wshGL_EJ.Range("V5").value = "du " & Format$(Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arri�re(dateFin), wsdADMIN.Range("B1").value)
    wshGL_EJ.Range("T6").value = "du " & Format$(dateFin, wsdADMIN.Range("B1").value)
    wshGL_EJ.Range("V6").value = "du " & Format$(dateFin, wsdADMIN.Range("B1").value)
    
    Dim rngResultAF As Range
    Call GL_Get_Account_Trans_AF(ObtenirNoGlIndicateur("Revenus de consultation"), Fn_Calcul_Date_Premier_Jour_Trois_Mois_Arri�re(dateFin), dateFin, rngResultAF)
    cases(101) = -Application.WorksheetFunction.Sum(rngResultAF.Columns(7)) _
                    - Application.WorksheetFunction.Sum(rngResultAF.Columns(8))

    With wshGL_EJ.Range("P10")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .value = -cases(101)
    End With
    
    'TPS percues
    cases(105) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TPS Factur�e"), dateFin)
    wshGL_EJ.Range("E" & r).value = "TPS percues"
    If cases(105) <= 0 Then
        wshGL_EJ.Range("H" & r).value = -cases(105)
    Else
        wshGL_EJ.Range("I" & r).value = cases(105)
    End If
    r = r + 1
    With wshGL_EJ.Range("T10")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .value = -cases(105)
    End With
    
    'TVQ percues
    cases(205) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TVQ Factur�e"), dateFin)
    wshGL_EJ.Range("E" & r).value = "TVQ percues"
    If cases(205) <= 0 Then
        wshGL_EJ.Range("H" & r).value = -cases(205)
    Else
        wshGL_EJ.Range("I" & r).value = cases(205)
    End If
    r = r + 1
    With wshGL_EJ.Range("V10")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .value = -cases(205)
    End With
    
    cases(108) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TPS Pay�e"), dateFin)
    wshGL_EJ.Range("E" & r).value = "TPS pay�es"
    If cases(108) <= 0 Then
        wshGL_EJ.Range("H" & r).value = -cases(108)
    Else
        wshGL_EJ.Range("I" & r).value = cases(108)
    End If
    r = r + 1
    With wshGL_EJ.Range("T13")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .value = cases(108)
    End With
    
    cases(208) = Fn_Get_GL_Account_Balance(ObtenirNoGlIndicateur("TVQ Pay�e"), dateFin)
    wshGL_EJ.Range("E" & r).value = "TVQ pay�es"
    If cases(208) <= 0 Then
        wshGL_EJ.Range("H" & r).value = -cases(208)
    Else
        wshGL_EJ.Range("I" & r).value = cases(208)
    End If
    r = r + 1
    With wshGL_EJ.Range("V13")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .value = cases(208)
    End With
    
    cases(113) = -cases(105) - cases(108)
    With wshGL_EJ.Range("T16")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .value = cases(113)
    End With
    
    cases(213) = -cases(205) - cases(208)
    With wshGL_EJ.Range("V16")
        .Font.Bold = True
        .Font.size = 12
        .NumberFormat = "###,##0.00 $"
        .HorizontalAlignment = xlRight
        .value = cases(213)
    End With
    
    Dim net As Double
    If cases(113) + cases(213) > 0 Then
        With wshGL_EJ.Range("X14")
            .Font.Bold = True
            .Font.size = 12
            .NumberFormat = "###,##0.00 $"
            .HorizontalAlignment = xlRight
            .value = cases(113) + cases(213)
        End With
        net = cases(113) + cases(213)
    Else
        With wshGL_EJ.Range("X10")
            .Font.Bold = True
            .Font.size = 12
            .NumberFormat = "###,##0.00 $"
            .HorizontalAlignment = xlRight
            .value = -(cases(113) + cases(213))
        End With
        net = -(cases(113) + cases(213))
    End If
    
    'Encaisse
    wshGL_EJ.Range("E" & r).value = "Encaisse"
    If net <= 0 Then
        wshGL_EJ.Range("H" & r).value = -net
    Else
        wshGL_EJ.Range("I" & r).value = net
    End If
    r = r + 1
    
    With wshGL_EJ
        .Unprotect
        .Range("N:Y").EntireColumn.Hidden = False
    End With

End Sub

Sub GL_EJ_Renverser_Ecriture()

    Dim ws As Worksheet: Set ws = wsdGL_Trans
    
    '1. Demande le num�ro d'�criture � partir d'un ListBox
    Call PreparerAfficherListeEcriture
    Dim no_Ecriture As Long
    If ActiveSheet.Range("B3").value <> -1 Then
        no_Ecriture = ActiveSheet.Range("B3").value
    Else
        MsgBox _
            Prompt:="Vous n'avez s�lectionn� aucune �criture � renverser", _
            Title:="S�lection d'une �criture � renverser", _
            Buttons:=vbInformation
        Application.EnableEvents = False
        wshGL_EJ.Range("F4").value = ""
        wshGL_EJ.Range("F4").Select
        Application.EnableEvents = True
        Exit Sub
    End If
    
    '2. Affiche l'�criture � renverser
    Call GL_Get_JE_Detail_Trans_AF(no_Ecriture)
    Dim lastUsedRowResult As Long
    lastUsedRowResult = ws.Cells(ws.Rows.count, "AC").End(xlUp).row
    If lastUsedRowResult < 2 Then
        MsgBox "Je ne retrouve pas l'�criture '" & no_Ecriture & "'" & vbNewLine & vbNewLine & _
                "Veuillez v�rifier votre num�ro et reessayez", vbInformation, "Num�ro d'�criture invalide"
        Exit Sub
    End If
    Dim rngResult As Range
    Set rngResult = ws.Range("AC1").CurrentRegion.offset(1, 0)
    If InStr(rngResult.Cells(1, 4).value, "ENCAISSEMENT:") <> 0 Or _
        InStr(rngResult.Cells(1, 4).value, "D�BOURS�:") <> 0 Or _
        InStr(rngResult.Cells(1, 4).value, "FACTURE:") <> 0 Or _
        InStr(rngResult.Cells(1, 4).value, "RENVERSEMENT:") <> 0 Then
        MsgBox "Je ne peux renverser ce type d'�criture '" & _
                Left$(rngResult.Cells(1, 4).value, InStr(rngResult.Cells(1, 4).value, ":") - 1) & _
                "'" & vbNewLine & vbNewLine & _
                "Veuillez v�rifier votre num�ro et reessayez", _
                vbInformation, "Type d'�criture impossible � renverser"
        wshGL_EJ.Range("F4").value = ""
        wshGL_EJ.Range("F4").Select
        Exit Sub
    End If
    
    'Cette �criture a-t-elle d�j� �t� RENVERS�E ?
    Dim rng As Range
    Set rng = ws.Columns("D")
    Dim trouve As Range
    Set trouve = rng.Find(What:="RENVERSEMENT:" & no_Ecriture, LookIn:=xlValues, LookAt:=xlWhole)
    If Not trouve Is Nothing Then
        MsgBox "Cette �criture a d�j� �t� RENVERS�E..." & vbNewLine & vbNewLine & _
               "Avec le num�ro d'�criture '" & ws.Cells(trouve.row, 1).value & "'" & vbNewLine & vbNewLine & _
               "En date du " & Format$(ws.Cells(trouve.row, 2).value, wsdADMIN.Range("B1").value) & ".", vbInformation
        Exit Sub
    End If
    
    Application.EnableEvents = False
    wshGL_EJ.Range("K4").value = Format$(rngResult.Cells(1, 2).value, wsdADMIN.Range("B1").value)
    wshGL_EJ.Range("F6").value = rngResult.Cells(1, 3).value
    Dim ligne As Range
    Dim l As Long: l = 9
    For Each ligne In rngResult.Rows
        wshGL_EJ.Range("E" & l).value = ligne.Cells(6).value
        If ligne.Cells(7).value <> 0 Then
            wshGL_EJ.Range("H" & l).value = ligne.Cells(7).value
        End If
        If ligne.Cells(8).value <> 0 Then
            wshGL_EJ.Range("I" & l).value = ligne.Cells(8).value
        End If
        wshGL_EJ.Range("J" & l).value = ligne.Cells(9).value
        wshGL_EJ.Range("L" & l).value = ligne.Cells(5).value
        l = l + 1
    Next ligne
    Application.EnableEvents = True
    
    'On affiche l'�criture � renverser en rouge
    wshGL_EJ.Range("F4, K4, F6:k6").Font.Color = vbRed
    wshGL_EJ.Range("E9:K23").Font.Color = vbRed
    
    'Change le libell� du Bouton & caract�ristiques
    Dim shp As Shape
    Set shp = wshGL_EJ.Shapes("btnUpdate")
    Call GL_EJ_Forme_Modifier(shp)
    
    'Lib�rer la m�moire
    Set ligne = Nothing
    Set rngResult = Nothing
    Set shp = Nothing
    Set ws = Nothing
    
End Sub

Sub GL_EJ_Depot_Client()

    Dim ws As Worksheet: Set ws = wshGL_EJ
    
    'Ajuster le formulaire
    ws.Range("E6").value = "Client:"
            
    'Ajouter la validation des donn�es
    Dim cell As Range: Set cell = wshGL_EJ.Range("F6:K6")
    
    Dim condition As Boolean
    condition = (wshGL_EJ.Range("F4").value = "D�p�t de client")
    
    Call G�rerValidation(cell, "dnrClients_Search_Field_Only", condition)
    
    'Force l'�criture
    wshGL_EJ.Range("E9").value = "Encaisse"
    wshGL_EJ.Range("E10").value = "Produit per�u d'avance"
    
    'Saisie du montant du d�p�t
    wshGL_EJ.Range("K4").Select

    'Lib�rer les objects
    Set cell = Nothing
    Set ws = Nothing
    
End Sub

Sub G�rerValidation(cell As Range, nomPlage As String, condition As Boolean)
    
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
    
    Dim feuilleProt�g�e As Boolean
    feuilleProt�g�e = ws.ProtectContents
    
    If feuilleProt�g�e Then ws.Unprotect
    
    On Error Resume Next
    cell.Validation.Delete 'Supprimer toute validation existante
    On Error GoTo 0
    
    'Ajouter la validation de donn�es
    cell.Validation.Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, _
                        Formula1:="=" & ThisWorkbook.Names(nomPlage).Name

    'Configurer les propri�t�s de la validation de donn�es
    If Not cell.Validation Is Nothing Then
        cell.Validation.IgnoreBlank = True
        cell.Validation.InCellDropdown = True
        cell.Validation.ShowInput = True
        cell.Validation.ShowError = True
    End If
    
    If feuilleProt�g�e Then
        With ws
            .Protect UserInterfaceOnly:=True
            .EnableSelection = xlUnlockedCells
        End With
    End If
    
    'Lib�rer la m�moire
    Set ws = Nothing
    
End Sub

Sub AnnulerValidation(cell As Range)

    cell.Validation.Delete
    
End Sub

Sub GL_EJ_Recurrente_Build_Summary()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Build_Summary", "", 0)
    
    'Build the summary at column K & L
    Dim lastUsedRow1 As Long
    lastUsedRow1 = wsdGL_EJ_Recurrente.Cells(wsdGL_EJ_Recurrente.Rows.count, 1).End(xlUp).row
    
    Dim lastUsedRow2 As Long
    lastUsedRow2 = wsdGL_EJ_Recurrente.Cells(wsdGL_EJ_Recurrente.Rows.count, "J").End(xlUp).row
    If lastUsedRow2 > 1 Then
        wsdGL_EJ_Recurrente.Range("J2:K" & lastUsedRow2).Clear
    End If
    
    With wsdGL_EJ_Recurrente
        Dim i As Long, k As Long, oldEntry As String
        k = 2
        For i = 2 To lastUsedRow1
            If .Range("A" & i).value <> oldEntry Then
                .Range("J" & k).value = .Range("B" & i).value
                .Range("K" & k).value = "'" & Fn_Pad_A_String(.Range("A" & i).value, " ", 5, "L")
                oldEntry = .Range("A" & i).value
                k = k + 1
            End If
        Next i
    End With

    Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Build_Summary", "", startTime)

End Sub

Sub GL_Get_JE_Detail_Trans_AF(noEJ As Long) '2024-11-17 @ 12:08

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_Get_JE_Detail_Trans_AF", "", 0)

    Dim ws As Worksheet: Set ws = wsdGL_Trans
    
    'wsdGL_Trans_AF#2

    'Effacer les donn�es de la derni�re utilisation
    ws.Range("AA6:AA10").ClearContents
    ws.Range("AA6").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    'D�finir le range pour la source des donn�es en utilisant un tableau
    Dim rngData As Range
    Set rngData = ws.Range("l_tbl_GL_Trans[#All]")
    ws.Range("AA7").value = rngData.Address
    
    'D�finir le range des crit�res
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("AA2:AA3")
    ws.Range("AA3").value = noEJ
    ws.Range("AA8").value = rngCriteria.Address
    
    'D�finir le range des r�sultats et effacer avant le traitement
    Dim rngResult As Range
    Set rngResult = ws.Range("AC1").CurrentRegion
    rngResult.offset(1, 0).Clear
    Set rngResult = ws.Range("AC1:AK1")
    ws.Range("AA9").value = rngResult.Address
    
    rngData.AdvancedFilter _
                action:=xlFilterCopy, _
                criteriaRange:=rngCriteria, _
                CopyToRange:=rngResult, _
                Unique:=False
        
    'Quels sont les r�sultats ?
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AC").End(xlUp).row
    ws.Range("AA10").value = lastUsedRow - 1 & " lignes"

    'On tri les r�sultats par noGL / par date?
    If lastUsedRow > 2 Then
        With ws.Sort 'Sort - NoEntr�e, D�bit(D) et Cr�dit (D)
        .SortFields.Clear
            'First sort On NoEntr�e
            .SortFields.Add key:=ws.Range("AC2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            'Second, sort On D�bit(D)
            .SortFields.Add key:=ws.Range("AI2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending, _
                DataOption:=xlSortNormal
            'Third, sort On Cr�dit(D)
            .SortFields.Add key:=ws.Range("AJ2"), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending, _
                DataOption:=xlSortNormal
            .SetRange wsdGL_Trans.Range("AC2:AK" & lastUsedRow)
            .Apply 'Apply Sort
         End With
    End If
    
    'Lib�rer la m�moire
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
    destinationFileName = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'SQL select command to find the next available ID
    Dim strSQL As String
    strSQL = "SELECT MAX(NoEntr�e) AS MaxEJNo FROM [" & destinationTab & "]"

    'Open recordset to find out the MaxID
    rs.Open strSQL, conn
    
    'Get the last used row
    Dim MaxEJNo As Long, lastJE As Long
    If IsNull(rs.Fields("MaxEJNo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastJE = 1
    Else
        lastJE = rs.Fields("MaxEJNo").value
    End If
    
    'Calculate the new JE number
    gNextJENo = lastJE + 1
    wshGL_EJ.Range("B1").value = gNextJENo
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    'Read all line from Journal Entry
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields(fGlTNoEntr�e - 1).value = gNextJENo
            rs.Fields(fGlTDate - 1).value = Format$(CDate(wshGL_EJ.Range("K4").value), "yyyy-mm-dd")
            If wshGL_EJ.Range("F4").value <> "D�p�t de client" Then
                rs.Fields(fGlTDescription - 1).value = wshGL_EJ.Range("F6").value
                rs.Fields(fGlTSource - 1).value = wshGL_EJ.Range("F4").value
            Else
                rs.Fields(fGlTDescription - 1).value = "Client:" & wshGL_EJ.Range("B6").value & " - " & wshGL_EJ.Range("F6").value
                rs.Fields(fGlTSource - 1).value = UCase$(wshGL_EJ.Range("F4").value)
            End If
            rs.Fields(fGlTNoCompte - 1).value = wshGL_EJ.Range("L" & l).value
            rs.Fields(fGlTCompte - 1).value = wshGL_EJ.Range("E" & l).value
            If wshGL_EJ.Range("H" & l).value <> "" <> 0 Then
                rs.Fields(fGlTD�bit - 1).value = CDbl(Replace(wshGL_EJ.Range("H" & l).value, ".", ","))
            End If
            If wshGL_EJ.Range("I" & l).value <> "" Then
                rs.Fields(fGlTCr�dit - 1).value = CDbl(Replace(wshGL_EJ.Range("I" & l).value, ".", ","))
            End If
            rs.Fields(fGlTAutreRemarque - 1).value = wshGL_EJ.Range("J" & l).value
            rs.Fields(fGlTTimeStamp - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rs.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_To_DB", "", startTime)

End Sub

Sub GL_Trans_Add_Record_Locally(r As Long) 'Write records locally
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wshGL_EJ.Range("B1").value
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'What is the last used row in GL_Trans ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wsdGL_Trans.Cells(wsdGL_Trans.Rows.count, "A").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally - r = " & r, -1)
    
    Dim i As Long
    For i = 9 To r
        wsdGL_Trans.Range("A" & rowToBeUsed).value = JENo
        wsdGL_Trans.Range("B" & rowToBeUsed).value = CDate(wshGL_EJ.Range("K4").value)
        If wshGL_EJ.Range("F4").value <> "D�p�t de client" Then
            wsdGL_Trans.Range("C" & rowToBeUsed).value = wshGL_EJ.Range("F6").value
            wsdGL_Trans.Range("D" & rowToBeUsed).value = wshGL_EJ.Range("F4").value
        Else
            wsdGL_Trans.Range("C" & rowToBeUsed) = "Client:" & wshGL_EJ.Range("B6").value & " - " & wshGL_EJ.Range("F6").value
            wsdGL_Trans.Range("D" & rowToBeUsed).value = UCase$(wshGL_EJ.Range("F4").value)
        End If
        wsdGL_Trans.Range("E" & rowToBeUsed).value = wshGL_EJ.Range("L" & i).value
        wsdGL_Trans.Range("F" & rowToBeUsed).value = wshGL_EJ.Range("E" & i).value
        If wshGL_EJ.Range("H" & i).value <> "" Then
            wsdGL_Trans.Range("G" & rowToBeUsed).value = wshGL_EJ.Range("H" & i).value
        End If
        If wshGL_EJ.Range("I" & i).value <> "" Then
            wsdGL_Trans.Range("H" & rowToBeUsed).value = wshGL_EJ.Range("I" & i).value
        End If
        wsdGL_Trans.Range("I" & rowToBeUsed).value = wshGL_EJ.Range("J" & i).value
        wsdGL_Trans.Range("J" & rowToBeUsed).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rowToBeUsed = rowToBeUsed + 1
        
        Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally - i = " & i, -1)

    Next i
    
    Application.ScreenUpdating = True
    
    Call Log_Record("modGL_EJ:GL_Trans_Add_Record_Locally", "", startTime)

End Sub

Sub EJ_Trans_Update_Ecriture_Renversee_To_DB()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:EJ_Trans_Update_Ecriture_Renversee_To_DB", "", 0)
    
    'D�finition des param�tres
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_Trans$"

    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Requ�te SQL pour rechercher la ligne correspondante
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE [NoEntr�e] = " & gNumeroEcritureARenverser

    'Ouvrir le Recordset
    rs.Open strSQL, conn, 1, 3 'adOpenKeyset (1) + adLockOptimistic (3) pour modifier les donn�es

    'V�rifier si des enregistrements existent
    If rs.EOF Then
        MsgBox "Aucun enregistrement trouv�.", vbCritical, "Impossible de mettre � jour les �critures RENVERS�ES"
    Else
        'Boucler � travers les enregistrements
        Do While Not rs.EOF
            rs.Fields(fGlTSource - 1).value = "RENVERS�E par " & wshGL_EJ.Range("B1").value
            rs.Update
        'Passer � l'enregistrement suivant
        rs.MoveNext
        Loop
    End If
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    'Lib�rer la m�moire
    Set conn = Nothing
    Set rs = Nothing

    Call Log_Record("modGL_EJ:EJ_Trans_Update_Ecriture_Renversee_To_DB", "", startTime)
    
End Sub

Sub EJ_Trans_Update_Ecriture_Renversee_Locally()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modEJ_Saisie:EJ_Trans_Update_Ecriture_Renversee_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = wsdGL_Trans
    
    'Derni�re ligne de la table
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    'Boucler sur toutes les lignes pour trouver les correspondances
    Dim cell As Range
    For Each cell In ws.Range("A2:A" & lastUsedRow)
        If cell.value = gNumeroEcritureARenverser Then
            cell.offset(0, fGlTSource - 1).value = "RENVERS�E par " & wshGL_EJ.Range("B1").value
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    'Lib�rer la m�moire
    Set ws = Nothing

    Call Log_Record("modEJ_Saisie:EJ_Trans_Update_Ecriture_Renversee_Locally", "", startTime)
    
End Sub

Sub GL_EJ_Recurrente_Add_Record_To_DB(r As Long) 'Write/Update a record to external .xlsx file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Add_Record_To_DB", "", 0)

    Application.ScreenUpdating = False
    
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "GL_EJ_R�currente$"
    
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
    If IsNull(rs.Fields("MaxEJANo").value) Then
        ' Handle empty table (assign a default value, e.g., 1)
        lastEJA = 1
    Else
        lastEJA = rs.Fields("MaxEJANo").value
    End If
    
    'Calculate the new ID
    nextEJANo = lastEJA + 1
    wsdGL_EJ_Recurrente.Range("B2").value = nextEJANo

    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'Close the previous recordset, no longer needed and open an empty recordset
    rs.Close
    rs.Open "SELECT * FROM [" & destinationTab & "] WHERE 1=0", conn, 2, 3
    
    Dim l As Long
    For l = 9 To r
        rs.AddNew
            'Add fields to the recordset before updating it
            rs.Fields(fGlEjRNoEjR - 1).value = nextEJANo
            rs.Fields(fGlEjRDescription - 1).value = Replace(wshGL_EJ.Range("F6").value, "[Auto]-", "")
            rs.Fields(fGlEjRNoCompte - 1).value = wshGL_EJ.Range("L" & l).value
            rs.Fields(fGlEjRCompte - 1).value = wshGL_EJ.Range("E" & l).value
            If wshGL_EJ.Range("H" & l).value <> "" Then
                rs.Fields(fGlEjRD�bit - 1).value = CDbl(Replace(wshGL_EJ.Range("H" & l).value, ".", ","))
            End If
            If wshGL_EJ.Range("I" & l).value <> "" Then
                rs.Fields(fGlEjRCr�dit - 1).value = CDbl(Replace(wshGL_EJ.Range("I" & l).value, ".", ","))
            End If
            rs.Fields(fGlEjRAutreRemarque - 1).value = wshGL_EJ.Range("J" & l).value
            rs.Fields(fGlEjRTimeStamp - 1).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        rs.Update
    Next l
    
    'Close recordset and connection
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    conn.Close
    
    Application.ScreenUpdating = True

    'Lib�rer la m�moire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Add_Record_To_DB", "", startTime)

End Sub

Sub GL_EJ_Recurrente_Add_Record_Locally(r As Long) 'Write records to local file
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_EJ:GL_EJ_Recurrente_Add_Record_Locally", "", 0)
    
    Application.ScreenUpdating = False
    
    'Get the JE number
    Dim JENo As Long
    JENo = wsdGL_EJ_Recurrente.Range("B2").value
    
    'timeStamp uniforme
    Dim timeStamp As Date
    timeStamp = Now
    
    'What is the last used row in EJ_AUto ?
    Dim lastUsedRow As Long, rowToBeUsed As Long
    lastUsedRow = wsdGL_EJ_Recurrente.Cells(wsdGL_EJ_Recurrente.Rows.count, "C").End(xlUp).row
    rowToBeUsed = lastUsedRow + 1
    
    Dim i As Long
    For i = 9 To r
        wsdGL_EJ_Recurrente.Range("A" & rowToBeUsed).value = JENo
        wsdGL_EJ_Recurrente.Range("B" & rowToBeUsed).value = Replace(wshGL_EJ.Range("F6").value, "[Auto]-", "")
        wsdGL_EJ_Recurrente.Range("C" & rowToBeUsed).value = wshGL_EJ.Range("L" & i).value
        wsdGL_EJ_Recurrente.Range("D" & rowToBeUsed).value = wshGL_EJ.Range("E" & i).value
        If wshGL_EJ.Range("H" & i).value <> "" Then
            wsdGL_EJ_Recurrente.Range("E" & rowToBeUsed).value = wshGL_EJ.Range("H" & i).value
        End If
        If wshGL_EJ.Range("I" & i).value <> "" Then
            wsdGL_EJ_Recurrente.Range("F" & rowToBeUsed).value = wshGL_EJ.Range("I" & i).value
        End If
        wsdGL_EJ_Recurrente.Range("G" & rowToBeUsed).value = wshGL_EJ.Range("J" & i).value
        wsdGL_EJ_Recurrente.Range("H" & rowToBeUsed).value = Format$(timeStamp, "yyyy-mm-dd hh:mm:ss")
        
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
    
    'R�tablir la forme du bouton (Mettre � jour / Renverser)
    Dim shp As Shape
    Set shp = wshGL_EJ.Shapes("btnUpdate")
    Call GL_EJ_Forme_Restaurer(shp)

    'Nouvelle fa�on de faire
    wshGL_EJ.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
    gFromMenu = True
    
    'Lib�rer la m�moire
    Set shp = Nothing
    
End Sub

Sub GL_EJ_Forme_Sauvegarder(forme As Shape)

    'V�rifier si le Dictionary est d�j� instanci�, sinon le cr�er
    If gSauvegardesCaracteristiquesForme Is Nothing Then
        Set gSauvegardesCaracteristiquesForme = CreateObject("Scripting.Dictionary")
    End If

    'Sauvegarder les caract�ristiques originales de la forme
    gSauvegardesCaracteristiquesForme("Left") = forme.Left
    gSauvegardesCaracteristiquesForme("Width") = forme.Width
    gSauvegardesCaracteristiquesForme("Height") = forme.Height
    gSauvegardesCaracteristiquesForme("FillColor") = forme.Fill.ForeColor.RGB
    gSauvegardesCaracteristiquesForme("LineColor") = forme.Line.ForeColor.RGB
    gSauvegardesCaracteristiquesForme("Text") = forme.TextFrame2.TextRange.Text
    gSauvegardesCaracteristiquesForme("TextColor") = forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    
End Sub

Sub GL_EJ_Forme_Modifier(forme As Shape)

    'Appliquer des modifications � la forme
    Application.ScreenUpdating = True
    With forme
        .Left = 470
        .Width = 175
        .Height = 30
        .Fill.ForeColor.RGB = RGB(255, 0, 0)  'Rouge
        .Line.ForeColor.RGB = RGB(255, 255, 255) 'Blanc pur
        .TextFrame2.TextRange.Text = "Renversement"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) 'Blanc pur
    End With
    
    DoEvents
    
    Application.ScreenUpdating = False
    
End Sub

Sub GL_EJ_Forme_Restaurer(forme As Shape)

    'V�rifiez si les caract�ristiques originales sont sauvegard�es
    If gSauvegardesCaracteristiquesForme Is Nothing Then
        Exit Sub
    End If

    'Restaurer les caract�ristiques de la forme
    forme.Left = gSauvegardesCaracteristiquesForme("Left")
    forme.Width = gSauvegardesCaracteristiquesForme("Width")
    forme.Height = gSauvegardesCaracteristiquesForme("Height")
    forme.Fill.ForeColor.RGB = gSauvegardesCaracteristiquesForme("FillColor")
    forme.Line.ForeColor.RGB = gSauvegardesCaracteristiquesForme("LineColor")
    forme.TextFrame2.TextRange.Text = gSauvegardesCaracteristiquesForme("Text")
    forme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = gSauvegardesCaracteristiquesForme("TextColor")

End Sub

Sub PreparerAfficherListeEcriture()

    'Charger la liste des �critures au G/L en m�moire
    Dim ws As Worksheet: Set ws = wsdGL_Trans
    Dim arrData As Variant
    arrData = ws.Range("A1").CurrentRegion.value
    
    'Initialiser le tableau des r�sultats
    Dim resultats() As Variant
    Dim compteur As Long
    ReDim resultats(1 To Round(UBound(arrData, 1) / 2, 0), 1 To 5) 'Maximum = Nombre de lignes / 2
    
    Dim strDejaVu As String, source As String
    Dim i As Long
    compteur = 0
    For i = 2 To UBound(arrData, 1)
        source = CStr(arrData(i, fGlTSource))
        'Seulement les �critures de journal (exclure les autres)
        If source = "" Or Not ExclureTransaction(source) = True Then
            If InStr(strDejaVu, CStr(arrData(i, 1)) & ".|.") = 0 Then
                compteur = compteur + 1
                resultats(compteur, 1) = arrData(i, fGlTNoEntr�e)
                resultats(compteur, 2) = Format$(arrData(i, fGlTDate), wsdADMIN.Range("B1").value)
                resultats(compteur, 3) = arrData(i, fGlTDescription)
                resultats(compteur, 4) = source
                resultats(compteur, 5) = Format$(arrData(i, fGlTTimeStamp), wsdADMIN.Range("B1").value & " hh:mm:ss")
                strDejaVu = strDejaVu & CStr(arrData(i, fGlTNoEntr�e)) & ".|."
            End If
        End If
    Next i
    
    'Est-ce que nous avons des r�sultats
    If compteur = 0 Then
        MsgBox "Aucune �criture � renverser.", vbInformation
        Exit Sub
    End If
   
    'R�duire la taille du tableau resultats
    Call Array_2D_Resizer(resultats, compteur, UBound(resultats, 2))
    
    'Charger les r�sultats dans la ListBox
    With ufListe�critureGL.lsbListe�critureGL
        .ColumnCount = 5
        .ColumnWidths = "35;62;310;125;92"
        .List = resultats
    End With
    
    ufListe�critureGL.lsbListe�critureGL.Clear
    
    'Ajouter chaque ligne de 'resultats' au ListBox
    i = 1
    Do While i <= compteur
        ufListe�critureGL.lsbListe�critureGL.AddItem resultats(i, 1)
        ufListe�critureGL.lsbListe�critureGL.List(ufListe�critureGL.lsbListe�critureGL.ListCount - 1, 1) = resultats(i, 2)
        ufListe�critureGL.lsbListe�critureGL.List(ufListe�critureGL.lsbListe�critureGL.ListCount - 1, 2) = resultats(i, 3)
        ufListe�critureGL.lsbListe�critureGL.List(ufListe�critureGL.lsbListe�critureGL.ListCount - 1, 3) = resultats(i, 4)
        ufListe�critureGL.lsbListe�critureGL.List(ufListe�critureGL.lsbListe�critureGL.ListCount - 1, 4) = resultats(i, 5)
        i = i + 1
    Loop

    'D�placer le focus sur la derni�re ligne
    If ufListe�critureGL.lsbListe�critureGL.ListCount > 0 Then
        ufListe�critureGL.lsbListe�critureGL.ListIndex = ufListe�critureGL.lsbListe�critureGL.ListCount - 1
    End If
    
    'Afficher le UserForm
    ufListe�critureGL.show
    
End Sub

Sub ckbRecurrente_Click()

    If wshGL_EJ.ckbRecurrente.value = True Then
        wshGL_EJ.ckbRecurrente.BackColor = COULEUR_SAISIE
    Else
        wshGL_EJ.ckbRecurrente.BackColor = RGB(217, 217, 217)
    End If

End Sub


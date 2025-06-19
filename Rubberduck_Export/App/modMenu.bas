Attribute VB_Name = "modMenu"
Option Explicit

Sub shpMenuTEC_Click()

    Call menuTEC
    
End Sub

Sub menuTEC()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:menuTEC_Click", "", 0)
    
    wshMenuTEC.Visible = xlSheetVisible
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

    Call Log_Record("modMenu:menuTEC_Click", "", startTime)

End Sub

Sub shpMenuFacturation_Click()

    Call menuFacturation

End Sub

Sub menuFacturation()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:menuFacturation_Click", "", 0)
    
    If GetNomUtilisateur() = "Guillaume" Or _
            GetNomUtilisateur() = "GuillaumeCharron" Or _
            GetNomUtilisateur() = "gchar" Or _
            GetNomUtilisateur() = "RobertMV" Or _
            GetNomUtilisateur() = "robertmv" Or _
            GetNomUtilisateur() = "User" Or _
            GetNomUtilisateur() = "vgervais" Or _
            GetNomUtilisateur() = "Vlad_Portable" Or _
            GetNomUtilisateur() = "Oli_Portable" Then
        wshMenuFAC.Visible = xlSheetVisible
        wshMenuFAC.Activate
        wshMenuFAC.Range("A1").Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If
    
    Call Log_Record("modMenu:menuFacturation_Click", "", startTime)

End Sub

Sub shpMenuComptabilité_Click()

    Call menuComptabilité
    
End Sub

Sub menuComptabilité()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:menuComptabilité", "", 0)
    
    If GetNomUtilisateur() = "Guillaume" Or _
            GetNomUtilisateur() = "GuillaumeCharron" Or _
            GetNomUtilisateur() = "gchar" Or _
            GetNomUtilisateur() = "RobertMV" Or _
            GetNomUtilisateur() = "robertmv" Or _
            GetNomUtilisateur() = "User" Then
        wshMenuGL.Visible = xlSheetVisible
        wshMenuGL.Activate
        wshMenuGL.Range("A1").Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If

    Call Log_Record("modMenu:menuComptabilité", "", startTime)

End Sub

Sub shpParamètres_Click()

    Call Parametres
    
End Sub

Sub Parametres()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:Parametres", "", 0)
    
    If GetNomUtilisateur() = "Guillaume" Or _
            GetNomUtilisateur() = "GuillaumeCharron" Or _
            GetNomUtilisateur() = "gchar" Or _
            GetNomUtilisateur() = "RobertMV" Or _
            GetNomUtilisateur() = "robertmv" Then
        wsdADMIN.Visible = xlSheetVisible
        wsdADMIN.Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If
    
    Call Log_Record("modMenu:Parametres", "", startTime)

End Sub

Sub shpExitApp_Click()

    Call SauvegarderEtSortirApplication

End Sub

Sub SauvegarderEtSortirApplication() '2024-08-30 @ 07:37
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SauvegarderEtSortirApplication", "", 0)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & vbNewLine & _
                        "l'application de gestion (sauvegarde automatique) ?", vbYesNo + vbQuestion, "Confirmation de sortie")
    
    If confirmation = vbYes Then
        Call ApplicationFermetureNormale(GetNomUtilisateur())
    End If
    
'        Application.EnableEvents = False
'        wsdADMIN.Range("B1").value = ""
'        wsdADMIN.Range("B2").value = ""
'        Application.EnableEvents = True
'
'        Call Delete_User_Active_File
'
'        On Error Resume Next
'        Call Log_Record("----- Session terminée NORMALEMENT (modMenu:SauvegarderEtSortirApplication) -----", "", 0)
'        Call Log_Record("", "", -1)
'        On Error GoTo 0
'
'        Application.ScreenUpdating = True
'
'       'On termine ici...
'        Dim wb As Workbook: Set wb = ActiveWorkbook
'        Application.EnableEvents = False
'        ActiveWorkbook.Close SaveChanges:=True
'        Application.EnableEvents = True
'
'        If GetNomUtilisateur() = "RobertMV" Or GetNomUtilisateur() = "robertmv" Then
'            Call StopperSauvegardeAutomatique
'            Call ExporterCodeVBA 'Sauvegarde AUTOMATIQUE du code VBA
'        End If
'
'        DoEvents
'
'        'On tente de quitter l'application EXCEL
'        Application.Application.Quit
'
'    End If
'
'    'Libérer la mémoire
'    Set wb = Nothing
    
End Sub

Sub ApplicationFermetureNormale(ByVal userName As String) 'Nouvelle procédure - 2025-05-30 @ 11:07

    On Error GoTo ExitPoint
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Application.StatusBar = False
    
    Dim ws As Worksheet
    Set ws = wsdADMIN
    With ws
        .Range("B1").Value = ""
        .Range("B2").Value = ""
    End With
    
    'Effacer fichier utilisateur actif + Fermeture de la journalisation
    Call Delete_User_Active_File(GetNomUtilisateur())
    Call Log_Record("----- Session terminée NORMALEMENT (modMenu:SauvegarderEtSortirApplication) -----", "", 0)
    Call Log_Record("", "", -1)

    'Fermer la vérification d'inactivité
    If gProchaineVerification > 0 Then
        On Error Resume Next
        Application.OnTime gProchaineVerification, "VerifierInactivite", , False
        On Error GoTo 0
    End If

    'Fermer la sauvegarde automtique du code VBA (seul le développeur déclenche la sauvegarde automtique)
    If userName = "RobertMV" Or userName = "robertmv" Then
        Call StopperSauvegardeAutomatique
        Call ExporterCodeVBA
    End If
    
    'Fermeture du classeur de l'application uniquement
    ThisWorkbook.Close SaveChanges:=True
    
ExitPoint:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Set ws = Nothing
    On Error GoTo 0
    
End Sub

Sub Hide_All_Worksheets_Except_Menu() '2024-02-20 @ 07:28
    
    DoEvents
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:Hide_All_Worksheets_Except_Menu", "", 0)
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName <> "wshMenu" Then
            If GetNomUtilisateur() <> "RobertMV" Or InStr(ws.CodeName, "wshzDoc") = 0 Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modMenu:Hide_All_Worksheets_Except_Menu", "", startTime)
    
End Sub

Sub HideDevShapesBasedOnUsername(ByVal userName As String) '2025-06-06 @ 11:17
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:HideDevShapesBasedOnUsername", "", 0)
    
    Dim ws As Worksheet: Set ws = wshMenu
    Dim devShapes As Variant
    devShapes = Array( _
        "shpImporterCorrigerMASTER", _
        "shpVérificationIntégrité", _
        "shpTraitementFichiersLog", _
        "shpCompterLignesCode", _
        "shpRechercherCode", _
        "shpCorrigerNomClientTEC", _
        "shpCorrigerNomClientCAR", _
        "shpChercherRéférencesCirculaires", _
        "shpChangerReferenceSystem", _
        "shpListerModulesEtRoutines", _
        "shpVérificationMacrosContrôles" _
    )

    Dim isDevUser As Boolean
    isDevUser = (userName = "RobertMV" Or userName = "robertmv")
    Dim visibleState As MsoTriState
    visibleState = IIf(isDevUser, msoTrue, msoFalse)

    Dim i As Long
    For i = LBound(devShapes) To UBound(devShapes)
        On Error Resume Next 'Ignore erreur si Shape absent
        ws.Shapes(devShapes(i)).Visible = visibleState
        If Err.Number <> 0 Then
            Debug.Print "Forme introuvable: " & devShapes(i)
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    Call Log_Record("modMenu:HideDevShapesBasedOnUsername", "", startTime)

End Sub

Sub Delete_User_Active_File(ByVal userName As String)

'    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:Delete_User_Active_File", "", 0)
    
    Dim traceFilePath As String
    traceFilePath = wsdADMIN.Range("F5").Value & DATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
    If Dir(traceFilePath) <> "" Then
        Kill traceFilePath
    End If

'    Call Log_Record("modMenu:Delete_User_Active_File", userName, startTime)

End Sub

Sub shpImporterCorrigerMASTER_Click()

    If GetNomUtilisateur() <> "RobertMV" And GetNomUtilisateur() <> "robertmv" Then
        Exit Sub
    End If
    
    'Crée un répertoire local et importe les fichiers à analyser
    Call CreerRepertoireEtImporterFichiers
    
    'Ajuste les tableaux (tables) de toutes les feuilles de GCF_BD_MASTER.xlsx
    Call AjusterEpurerTablesDeMaster

End Sub

Sub shpVerificationIntegrite_Click()

    Call modAppli_Utils.VerifierIntegriteTablesLocales

End Sub

Sub shpRechercherCode_Click()

    Call RechercherCodeProjet

End Sub

Sub shpCompterLignesCodeProjet_Click()

    Call CompterLignesCode

End Sub

Sub shpCorrigerNomClientTEC_Click()

    Call modzDataConversion.CorrigeNomClientInTEC
    
End Sub

Sub shpCorrigerNomClientCAR_Click()

    Call modzDataConversion.CorrigeNomClientInCAR
    
End Sub

Sub shpChercherRéférencesCirculaires_Click() '2024-11-22 @ 13:33

    Call Detect_Circular_References_In_Workbook
    
End Sub

Sub shpChangerReferenceSystem_Click() '2024-11-22 @ 13:33

    Call Toggle_A1_R1C1_Reference
    
End Sub

Sub shpListerModulesEtRoutines_Click() '2024-11-22 @ 13:33

    Call List_Subs_And_Functions_All
    
End Sub

Sub shpVérificationMacrosContrôles_Click()

    Call VerifierControlesAssociesToutesFeuilles

End Sub

Sub shpRetourMenuPrincipal_Click()

    Call RetourMenuPrincipal

End Sub

Sub RetourMenuPrincipal()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
        .Activate
        .Range("A1").Select
    End With

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub


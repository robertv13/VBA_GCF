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
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    If userName = "Guillaume" Or _
            userName = "GuillaumeCharron" Or _
            userName = "gchar" Or _
            userName = "Robert M. Vigneault" Or _
            userName = "robertmv" Or _
            userName = "User" Then
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
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    If userName = "Guillaume" Or _
            userName = "GuillaumeCharron" Or _
            userName = "gchar" Or _
            userName = "Robert M. Vigneault" Or _
            userName = "robertmv" Or _
            userName = "User" Then
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
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    If userName = "Guillaume" Or _
            userName = "GuillaumeCharron" Or _
            userName = "gchar" Or _
            userName = "Robert M. Vigneault" Or _
            userName = "robertmv" Or _
            userName = "User" Then
        wshAdmin.Visible = xlSheetVisible
        wshAdmin.Select
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
        Application.EnableEvents = False
        wshAdmin.Range("B1").Value = ""
        wshAdmin.Range("B2").Value = ""
        Application.EnableEvents = True
        
        Call Delete_User_Active_File

        On Error Resume Next
        Call Log_Record("----- Session terminée NORMALEMENT (modMenu:SauvegarderEtSortirApplication) -----", "", 0)
        Call Log_Record("", "", -1)
        On Error GoTo 0
        
        Application.ScreenUpdating = True
        
       'On termine ici...
        Dim wb As Workbook: Set wb = ActiveWorkbook
        Application.EnableEvents = False
        ActiveWorkbook.Close SaveChanges:=True
        Application.EnableEvents = True
        
        'On tente de quitter l'application EXCEL
        Application.Application.Quit
        
    End If
    
    'Libérer la mémoire
    Set wb = Nothing
    
End Sub

Sub Hide_All_Worksheets_Except_Menu() '2024-02-20 @ 07:28
    
    DoEvents
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:Hide_All_Worksheets_Except_Menu", "", 0)
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName <> "wshMenu" Then
            If userName <> "Robert M. Vigneault" Or InStr(ws.CodeName, "wshzDoc") = 0 Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modMenu:Hide_All_Worksheets_Except_Menu", "", startTime)
    
End Sub

Sub HideDevShapesBasedOnUsername()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:HideDevShapesBasedOnUsername", "", 0)
    
    'Set the worksheet where the shapes are located
    Dim ws As Worksheet: Set ws = wshMenu
    
    'Loop through each shape in the worksheet
    Dim shp As Shape
    Dim userName As String
    userName = Fn_Get_Windows_Username
    If userName = "Robert M. Vigneault" Or userName = "robertmv" Then
        ws.Shapes("shpImporterCorrigerMASTER").Visible = msoTrue
        ws.Shapes("shpVérificationIntégrité").Visible = msoTrue
        ws.Shapes("shpRechercherCode").Visible = msoTrue
        ws.Shapes("shpCorrigerNomClientTEC").Visible = msoTrue
        ws.Shapes("shpCorrigerNomClientCAR").Visible = msoTrue
        ws.Shapes("shpChercherRéférencesCirculaires").Visible = msoTrue
        ws.Shapes("shpChangerReferenceSystem").Visible = msoTrue
        ws.Shapes("shpListerModulesEtRoutines").Visible = msoTrue
        ws.Shapes("shpVérificationMacrosContrôles").Visible = msoTrue
        ws.Shapes("shpVérifierDernièresLignes").Visible = msoTrue
        ws.Shapes("shpTraitementFichiersLog").Visible = msoTrue
    Else
        ws.Shapes("shpImporterCorrigerMASTER").Visible = msoFalse
        ws.Shapes("shpVérificationIntégrité").Visible = msoFalse
        ws.Shapes("shpRechercherCode").Visible = msoFalse
        ws.Shapes("shpCorrigerNomClientTEC").Visible = msoFalse
        ws.Shapes("shpCorrigerNomClientCAR").Visible = msoFalse
        ws.Shapes("shpChercherRéférencesCirculaires").Visible = msoFalse
        ws.Shapes("shpChangerReferenceSystem").Visible = msoFalse
        ws.Shapes("shpListerModulesEtRoutines").Visible = msoFalse
        ws.Shapes("shpVérificationMacrosContrôles").Visible = msoFalse
        ws.Shapes("shpVérifierDernièresLignes").Visible = msoFalse
        ws.Shapes("shpTraitementFichiersLog").Visible = msoTrue
    End If
    
    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
    Call Log_Record("modMenu:HideDevShapesBasedOnUsername", "", startTime)

End Sub

Sub Delete_User_Active_File()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:Delete_User_Active_File", "", 0)
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    Dim traceFilePath As String
    traceFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
    If Dir(traceFilePath) <> "" Then
        Kill traceFilePath
    End If

    Call Log_Record("modMenu:Delete_User_Active_File", "", startTime)

End Sub

Sub shpImporterCorrigerMASTER_Click()

    If Not Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        Exit Sub
    End If
    
    'Crée un répertoire local et importe les fichiers à analyser
    Call CreerRepertoireEtImporterFichiers
    
    'Ajuste les tableaux (tables) de toutes les feuilles de GCF_BD_MASTER.xlsx
    Call AjusterTableauxDansMaster

End Sub

Sub shpVérificationIntégrité_Click()

    Call VérifierIntégrité

End Sub

Sub shpRechercherCode_Click()

    Call Code_Search_Everywhere

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




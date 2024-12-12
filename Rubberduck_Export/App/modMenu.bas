Attribute VB_Name = "modMenu"
Option Explicit

'Execute the next menu (next level)
Sub menuTEC_Click()
    
    wshMenuTEC.Visible = xlSheetVisible
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

End Sub

Sub menuFacturation_Click()
    
    If Fn_Get_Windows_Username = "Guillaume" Or _
            Fn_Get_Windows_Username = "GuillaumeCharron" Or _
            Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
            Fn_Get_Windows_Username = "robertmv" Then
        wshMenuFAC.Visible = xlSheetVisible
        wshMenuFAC.Activate
        wshMenuFAC.Range("A1").Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If

End Sub

'CommentOut - 2024-11-25
'Sub MenuDEB_Click()
'
'    If Fn_Get_Windows_Username = "Guillaume" Or _
'            Fn_Get_Windows_Username = "GuillaumeCharron" Or _
'            Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
'            Fn_Get_Windows_Username = "robertmv" Then
'        wshMenuDEB.Visible = xlSheetVisible
'        wshMenuDEB.Activate
'        wshMenuDEB.Range("A1").Select
'    Else
'        Application.EnableEvents = False
'        wshMenu.Activate
'        Application.EnableEvents = True
'    End If
'
'End Sub
'
Sub menuComptabilite_Click()
    
    If Fn_Get_Windows_Username = "Guillaume" Or _
            Fn_Get_Windows_Username = "GuillaumeCharron" Or _
            Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
            Fn_Get_Windows_Username = "robertmv" Then
        wshMenuGL.Visible = xlSheetVisible
        wshMenuGL.Activate
        wshMenuGL.Range("A1").Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If

End Sub

Sub menuParametres_Click()
    
    If Fn_Get_Windows_Username = "Guillaume" Or _
            Fn_Get_Windows_Username = "GuillaumeCharron" Or _
            Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
            Fn_Get_Windows_Username = "robertmv" Then
        wshAdmin.Visible = xlSheetVisible
        wshAdmin.Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If
    
End Sub

Sub Exit_After_Saving() '2024-08-30 @ 07:37
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & vbNewLine & _
                        "l'application de gestion (sauvegarde automatique) ?", vbYesNo + vbQuestion, "Confirmation de sortie")
    
    If confirmation = vbYes Then
        Call Hide_All_Worksheets_Except_Menu
    
        DoEvents
        
        Call Delete_User_Active_File

        Application.ScreenUpdating = True
        
        On Error Resume Next
        Call Log_Record("***** Session terminée NORMALEMENT (modMenu:Exit_After_Saving) *****", 0)
        Call Log_Record("", -1)
        
        On Error GoTo 0
        
        DoEvents
        
        'Really ends here !!!
        Dim wb As Workbook: Set wb = ActiveWorkbook
        ActiveWorkbook.Close SaveChanges:=True
        
        'Never pass here... It's too late
        Application.EnableEvents = True
        Application.Application.Quit
    End If
    
    'Libérer la mémoire
    Set wb = Nothing
    
End Sub

Sub Hide_All_Worksheets_Except_Menu() '2024-02-20 @ 07:28
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:Hide_All_Worksheets_Except_Menu", 0)
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName <> "wshMenu" Then
            If Fn_Get_Windows_Username <> "Robert M. Vigneault" Or InStr(ws.CodeName, "wshzDoc") = 0 Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modAppli:Hide_All_Worksheets_Except_Menu", startTime)
    
End Sub

Sub HideDevShapesBasedOnUsername()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:HideDevShapesBasedOnUsername", 0)
    
    'Set the worksheet where the shapes are located
    Dim ws As Worksheet: Set ws = wshMenu
    
    'Loop through each shape in the worksheet
    Dim shp As Shape
    For Each shp In ws.Shapes
        'Check the username and hide shapes accordingly
        Select Case shp.Name
            Case "Import & Reorganisation de MASTER des Tableaux (MASTER)"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If
                
            Case "VérificationIntégrité"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If
            
            Case "RechercheCode"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If

            Case "Correction nom (TEC)"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If
            
            Case "Correction nom (CAR)"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If
            
            Case "RéférencesCirculaires"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                Else
                    shp.Visible = msoFalse
                End If
           
            Case "ChangeReferenceSystem"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If

            Case "ListeModules&Routines"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Or _
                    Fn_Get_Windows_Username = "robertmv" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If

            Case Else
        End Select
    Next shp

    'Libérer la mémoire
    Set shp = Nothing
    Set ws = Nothing
    
    Call Log_Record("modMenu:HideDevShapesBasedOnUsername", startTime)

End Sub

Sub Delete_User_Active_File()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:Delete_User_Active_File", 0)
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    Dim traceFilePath As String
    traceFilePath = wshAdmin.Range("F5").value & DATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
    If Dir(traceFilePath) <> "" Then
        Kill traceFilePath
    End If

    Call Log_Record("modMenu:Delete_User_Active_File", startTime)

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
        .EnableSelection = xlNoRestrictions
        .Activate
        .Range("A1").Select
    End With

    'Libérer la mémoire
    Set ws = Nothing
    
End Sub




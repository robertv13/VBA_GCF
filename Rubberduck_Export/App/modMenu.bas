Attribute VB_Name = "modMenu"
Option Explicit

Dim width As Long

Sub SlideIn_TEC()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_TEC", 0)

    With wshMenu.Shapes("btnTECMenu")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoTEC").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_TEC", startTime)

End Sub

Sub SlideOut_TEC()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_TEC", 0)

    With wshMenu.Shapes("btnTECMenu")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoTEC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "TEC"
    End With
    
    Call Log_Record("modMenu:SlideOut_TEC", startTime)

End Sub

Sub SlideIn_Facturation()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_Facturation", 0)
    
    With wshMenu.Shapes("btnFacturationMenu")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoFacturation").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_Facturation", startTime)

End Sub

Sub SlideOut_Facturation()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_Facturation", 0)

    With wshMenu.Shapes("btnFacturationMenu")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
          ActiveSheet.Shapes("icoFacturation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Facturation"
    End With

    Call Log_Record("modMenu:SlideOut_Facturation", startTime)

End Sub

Sub SlideIn_Debours()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_Debours", 0)

    With wshMenu.Shapes("btnDeboursMenu")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_Debours", startTime)

End Sub

Sub SlideOut_Debours()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_Debours", 0)

    With wshMenu.Shapes("btnDeboursMenu")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Débours"
    End With

    Call Log_Record("modMenu:SlideOut_Debours", startTime)

End Sub

Sub SlideIn_Comptabilite()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_Comptabilite", 0)

    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_Comptabilite", startTime)

End Sub

Sub SlideOut_Comptabilite()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_Comptabilite", 0)

    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Comptabilité"
    End With

    Call Log_Record("modMenu:SlideOut_Comptabilite", startTime)

End Sub

Sub SlideIn_Parametres()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_Parametres", 0)

    With wshMenu.Shapes("btnParametresOption")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_Parametres", startTime)

End Sub

Sub SlideOut_Parametres()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_Parametres", 0)

    With wshMenu.Shapes("btnParametresOption")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Paramètres"
    End With

    Call Log_Record("modMenu:SlideOut_Parametres", startTime)

End Sub

Sub SlideIn_Exit()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_Exit", 0)

    With wshMenu.Shapes("btnEXIT")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoEXIT").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_Exit", startTime)

End Sub

Sub SlideOut_Exit()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_Exit", 0)

    With ActiveSheet.Shapes("btnEXIT")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoEXIT").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Sortie"
    End With

    Call Log_Record("modMenu:SlideOut_Exit", startTime)

End Sub

'Second level (sub-menu) ---------------------------------------------------------------------------
Sub SlideOut_SaisieHeures()
    
'    With wshMenuTEC.Shapes("btnSaisieHeures")
'        For width = 32 To MAXWIDTH step 8
'            .Height = width
'            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Saisie des Heures"
'    End With

End Sub

Sub SlideIn_SaisieHeures()
    
'    With wshMenuTEC.Shapes("btnSaisieHeures")
'        For width = MAXWIDTH To 32 Step -8
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideOut_TEC_TDB()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = 32 To MAXWIDTH step 8
'            .Height = width
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Tableau de bord"
'    End With

End Sub

Sub SlideIn_TEC_TDB()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = MAXWIDTH To 32 Step -8
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideIn_TEC_Analyse()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = MAXWIDTH To 32 Step -8
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideOut_TEC_Analyse()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = 32 To MAXWIDTH step 8
'            .Height = width
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Tableau de bord"
'    End With

End Sub

Sub SlideIn_PrepFact()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_PrepFact", 0)

    With wshMenuFAC.Shapes("btnPrepFact")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_PrepFact", startTime)

End Sub

Sub SlideOut_PrepFact()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_PrepFact", 0)

    With wshMenuFAC.Shapes("btnPrepFact")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Préparation de facture"
    End With

    Call Log_Record("modMenu:SlideOut_PrepFact", startTime)

End Sub

Sub SlideIn_SuiviCC()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_SuiviCC", 0)

    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_SuiviCC", startTime)

End Sub

Sub SlideOut_SuiviCC()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_SuiviCC", 0)

    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = 32 To MAXWIDTH Step 1
            .Height = width
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Suivi de C/C"
    End With

    Call Log_Record("modMenu:SlideOut_SuiviCC", startTime)

End Sub

Sub SlideIn_Encaissement()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_Encaissement", 0)

    With wshMenuFAC.Shapes("btnEncaissement")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_Encaissement", startTime)

End Sub

Sub SlideOut_Encaissement()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_Encaissement", 0)

    With wshMenuFAC.Shapes("btnEncaissement")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Encaissement"
    End With

    Call Log_Record("modMenu:SlideOut_Encaissement", startTime)

End Sub

Sub SlideIn_FAC_Historique()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_FAC_Historique", 0)

    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_FAC_Historique", startTime)

End Sub

Sub SlideOut_FAC_Historique()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_FAC_Historique", 0)

    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Historique factures"
    End With

    Call Log_Record("modMenu:SlideOut_FAC_Historique", startTime)

End Sub

Sub SlideIn_FAC_Confirmation()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_FAC_Confirmation", 0)

    With wshMenuFAC.Shapes("btnFAC_Confirmation")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFAC_Confirmation").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_FAC_Confirmation", startTime)

End Sub

Sub SlideOut_FAC_Confirmation()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_FAC_Confirmation", 0)

    With wshMenuFAC.Shapes("btnFAC_Confirmation")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoFAC_Confirmation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Confirmation de facture"
    End With

    Call Log_Record("modMenu:SlideOut_FAC_Confirmation", startTime)

End Sub

Sub SlideIn_Paiement()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_Paiement", 0)

    With wshMenuDEB.Shapes("btnPaiement")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_Paiement", startTime)

End Sub

Sub SlideOut_Paiement()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_Paiement", 0)

    With wshMenuDEB.Shapes("btnPaiement")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Déboursé"
    End With

    Call Log_Record("modMenu:SlideOut_Paiement", startTime)

End Sub

Sub SlideIn_EJ()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_EJ", 0)

    With wshMenuGL.Shapes("btnEJ")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_EJ", startTime)

End Sub

Sub SlideOut_EJ()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_EJ", 0)

    With wshMenuGL.Shapes("btnEJ")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Entrée de Journal"
    End With

    Call Log_Record("modMenu:SlideOut_EJ", startTime)

End Sub

Sub SlideIn_BV()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_BV", 0)

    With wshMenuGL.Shapes("btnBV")
        For width = 180 To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoBV").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_BV", startTime)

End Sub

Sub SlideOut_BV()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_BV", 0)

    With wshMenuGL.Shapes("btnBV")
        For width = 32 To 180
            .Height = width
            ActiveSheet.Shapes("icoBV").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Balance de Vérification"
    End With

    Call Log_Record("modMenu:SlideOut_BV", startTime)

End Sub

Sub SlideIn_GL_Report()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_GL_Report", 0)

    With wshMenuGL.Shapes("btnGL")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

    Call Log_Record("modMenu:SlideIn_GL_Report", startTime)

End Sub

Sub SlideOut_GL_Report()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_GL_Report", 0)

    With wshMenuGL.Shapes("btnGL")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Rapport des transactions"
    End With

    Call Log_Record("modMenu:SlideOut_GL_Report", startTime)

End Sub

Sub SlideIn_EF()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideIn_EF", 0)

    With wshMenuGL.Shapes("btnEF")
        For width = MAXWIDTH To 32 Step -8
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    Call Log_Record("modMenu:SlideIn_EF", startTime)

End Sub

Sub SlideOut_EF()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMenu:SlideOut_EF", 0)

    With wshMenuGL.Shapes("btnEF")
        For width = 32 To MAXWIDTH Step 8
            .Height = width
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "États financiers"
    End With

    Call Log_Record("modMenu:SlideOut_EF", startTime)

End Sub

'Execute the next menu (next level)
Sub menuTEC_Click()
    
    Call SlideIn_TEC
    
    wshMenuTEC.Visible = xlSheetVisible
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

End Sub

Sub menuFacturation_Click()
    
    Call SlideIn_Facturation
    
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
            Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        wshMenuFAC.Visible = xlSheetVisible
        wshMenuFAC.Activate
        wshMenuFAC.Range("A1").Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If

End Sub

Sub MenuDEB_Click()
    
    Call SlideIn_Debours
        
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
            Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        wshMenuDEB.Visible = xlSheetVisible
        wshMenuDEB.Activate
        wshMenuDEB.Range("A1").Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If

End Sub

Sub menuComptabilite_Click()
    
    Call SlideIn_Comptabilite
        
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
            Fn_Get_Windows_Username = "Robert M. Vigneault" Then
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
    
    Call SlideIn_Parametres
    
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
            Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        wshAdmin.Visible = xlSheetVisible
        wshAdmin.Select
    Else
        Application.EnableEvents = False
        wshMenu.Activate
        Application.EnableEvents = True
    End If
    
End Sub

Sub Exit_After_Saving() '2024-08-06 @ 14:51
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & vbNewLine & _
                        "l'application de gestion (sauvegarde automatique) ?", vbYesNo + vbQuestion, "Confirmation de sortie")
    
    If confirmation = vbYes Then
        Call SlideIn_Exit
        Call Hide_All_Worksheets_Except_Menu
    
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
        Call Write_Info_On_Main_Menu
        
        Application.EnableEvents = False
        
        Dim wb As Workbook: Set wb = ActiveWorkbook
        ActiveWorkbook.Close saveChanges:=True
        
        Application.EnableEvents = True
        
        DoEvents '2024-08-21 @ 06:11

        Call End_Timer("message:This session has been terminated N O R M A L L Y", 0)
        
        Call Log_Record("***** Session terminée normalement (modMenu:Exit_After_Saving) *****", 0)
        
        Application.Application.Quit
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set wb = Nothing
    
End Sub

'Sub Exit_Click() '2024-07-05 @ 06:37
    
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'
'    Dim confirmation As VbMsgBoxResult
'    confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & vbNewLine & _
'                        "l'application de gestion ?" & vbNewLine, vbYesNo + vbQuestion, "Confirmation de sortie")
'
'    If confirmation = vbYes Then
'        Call SlideIn_Exit
'        Call Hide_All_Worksheets_Except_Menu
'
'        'Delete temporary Worksheet (Feuil*)
'        Dim ws As Worksheet
'        For Each ws In ThisWorkbook.Worksheets
'            If InStr(1, ws.CodeName, "Feuil") > 0 Then
'                Application.DisplayAlerts = False
'                ws.delete
'                Application.DisplayAlerts = True
'            End If
'        Next ws
'
'        Application.ScreenUpdating = True
'        Application.EnableEvents = False
'
'        Call End_Timer("message:This  session  has  been  terminated N O R M A L L Y", 0)
'
'        Application.DisplayAlerts = False
'        ThisWorkbook.Save
'        Application.DisplayAlerts = True
'        ThisWorkbook.Close saveChanges:=False
'
'        'If Personal.xlsb is open, hide it without saving
'        Dim wb As Workbook
'        On Error Resume Next
'        Set wb = Workbooks("PERSONAL.XLSB")
'        If Not wb Is Nothing Then
'            wb.IsAddin = True
'        End If
'        On Error GoTo 0
'
'        'Cleaning - 2024-07-05 @ 06:46
'        Set wb = Nothing
'        Set ws = Nothing
'
'        Application.Quit
'
'    End If
'
'End Sub

Sub Hide_All_Worksheets_Except_Menu() '2024-02-20 @ 07:28
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modAppli:Hide_All_Worksheets_Except_Menu()")
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName <> "wshMenu" Then
            If Fn_Get_Windows_Username <> "Robert M. Vigneault" Or InStr(ws.CodeName, "wshzDoc") = 0 Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
    Call End_Timer("modAppli:Hide_All_Worksheets_Except_Menu()", timerStart)
    
End Sub

Sub Slide_In_All_Menu_Options()

    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modMenu:Slide_In_All_Menu_Options()")
    
    Call SlideIn_TEC
    Call SlideIn_Facturation
    Call SlideIn_Debours
    Call SlideIn_Comptabilite
    Call SlideIn_Parametres
    Call SlideIn_Exit

    Call End_Timer("modMenu:Slide_In_All_Menu_Options()", timerStart)

End Sub

Sub Hide_Dev_Shapes_Based_On_Username()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modMenu:Hide_Dev_Shapes_Based_On_Username()")
    
    'Set the worksheet where the shapes are located
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Menu")
    
    'Loop through each shape in the worksheet
    Dim shp As Shape
    For Each shp In ws.Shapes
        'Check the username and hide shapes accordingly
        Select Case shp.name
            Case "ChangeReferenceSystem"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If

            Case "VérificationIntégritée"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If

            Case "RechercheCode"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If

            Case "RéférencesCirculaires"
                If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If

            Case Else
        End Select
    Next shp

    Call End_Timer("modMenu:Hide_Dev_Shapes_Based_On_Username()", timerStart)

End Sub

Sub BackToMainMenu()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
        .Activate
        .Range("A1").Select
    End With

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub




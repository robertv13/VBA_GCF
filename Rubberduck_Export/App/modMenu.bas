Attribute VB_Name = "modMenu"
Option Explicit

Dim width As Long

Sub SlideIn_TEC()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modMenu:SlideIn_TEC()")
    
    On Error Resume Next
    wshMenu.Unprotect
    On Error GoTo 0
    
    With wshMenu.Shapes("btnTECMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoTEC").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With
    
    wshMenu.Protect userInterfaceOnly:=True

    Call End_Timer("modMenu:SlideIn_TEC()", timerStart)

End Sub

Sub SlideOut_TEC()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnTECMenu")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoTEC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "TEC"
    End With
    
    wshMenu.Protect userInterfaceOnly:=True
    
End Sub

Sub SlideIn_Facturation()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Timer("modMenu:SlideIn_Facturation()")
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnFacturationMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoFacturation").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenu.Protect userInterfaceOnly:=True

    Call End_Timer("modMenu:SlideIn_Facturation()", timerStart)

End Sub

Sub SlideOut_Facturation()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnFacturationMenu")
        For width = 32 To MAXWIDTH
            .Height = width
          ActiveSheet.Shapes("icoFacturation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Facturation"
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_Debours()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnDeboursMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_Debours()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnDeboursMenu")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Débours"
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_Comptabilite()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_Comptabilite()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Comptabilité"
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_Parametres()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnParametresOption")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_Parametres()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnParametresOption")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Paramètres"
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_Exit()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With wshMenu.Shapes("btnEXIT")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoEXIT").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_Exit()
    
    On Error Resume Next: wshMenu.Unprotect: On Error GoTo 0
    
    With ActiveSheet.Shapes("btnEXIT")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEXIT").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Sortie"
    End With

    wshMenu.Protect userInterfaceOnly:=True

End Sub

'Second level (sub-menu) ---------------------------------------------------------------------------
Sub SlideOut_SaisieHeures()
    
'    With wshMenuTEC.Shapes("btnSaisieHeures")
'        For width = 32 To MAXWIDTH
'            .Height = width
'            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Saisie des Heures"
'    End With

End Sub

Sub SlideIn_SaisieHeures()
    
'    With wshMenuTEC.Shapes("btnSaisieHeures")
'        For width = MAXWIDTH To 32 Step -1
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideOut_TEC_TDB()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = 32 To MAXWIDTH
'            .Height = width
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Tableau de bord"
'    End With

End Sub

Sub SlideIn_TEC_TDB()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = MAXWIDTH To 32 Step -1
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideIn_TEC_Analyse()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = MAXWIDTH To 32 Step -1
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideOut_TEC_Analyse()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = 32 To MAXWIDTH
'            .Height = width
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Tableau de bord"
'    End With

End Sub

Sub SlideIn_PrepFact()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnPrepFact")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_PrepFact()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnPrepFact")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Préparation de facture"
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_SuiviCC()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_SuiviCC()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Suivi de C/C"
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_Encaissement()

    wshMenuFAC.Unprotect
    
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

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_Encaissement()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnEncaissement")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Encaissement"
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_FAC_Historique()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_FAC_Historique()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Historique factures"
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_FAC_Confirmation()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnFAC_Confirmation")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFAC_Confirmation").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_FAC_Confirmation()

    wshMenuFAC.Unprotect
    
    With wshMenuFAC.Shapes("btnFAC_Confirmation")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoFAC_Confirmation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Confirmation de facture"
    End With

    wshMenuFAC.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_Paiement()

    wshMenuDEB.Unprotect
    
    With wshMenuDEB.Shapes("btnPaiement")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenuDEB.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_Paiement()

    wshMenuDEB.Unprotect
    
    With wshMenuDEB.Shapes("btnPaiement")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Déboursé"
    End With

    wshMenuDEB.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_EJ()

    wshMenuGL.Unprotect
    
    With wshMenuGL.Shapes("btnEJ")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_EJ()

    wshMenuGL.Unprotect
    
    With wshMenuGL.Shapes("btnEJ")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Entrée de Journal"
    End With

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_BV()

    wshMenuGL.Unprotect
    
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

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_BV()

    wshMenuGL.Unprotect
    
    With wshMenuGL.Shapes("btnBV")
        For width = 32 To 180
            .Height = width
            ActiveSheet.Shapes("icoBV").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Balance de Vérification"
    End With

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_GL_Report()

    wshMenuGL.Unprotect
    
    With wshMenuGL.Shapes("btnGL")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_GL_Report()

    wshMenuGL.Unprotect
    
    With wshMenuGL.Shapes("btnGL")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Rapport des transactions"
    End With

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

Sub SlideIn_EF()

    wshMenuGL.Unprotect
    
    With wshMenuGL.Shapes("btnEF")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        On Error Resume Next
        .TextFrame2.TextRange.Characters.text = ""
        On Error GoTo 0
    End With

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

Sub SlideOut_EF()

    wshMenuGL.Unprotect
    
    With wshMenuGL.Shapes("btnEF")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "États financiers"
    End With

    wshMenuGL.Protect userInterfaceOnly:=True

End Sub

'Execute the next menu (next level)
Sub menuTEC_Click()
    
    Call SlideIn_TEC
    
    wshMenuTEC.Visible = xlSheetVisible
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

End Sub

Sub menuFacturation_Click()
    
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
        Fn_Get_Windows_Username = "Robert M. Vigneault" Then
    
        Call SlideIn_Facturation
        
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
    
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
        Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        
        Call SlideIn_Debours
        
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
    
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
        Fn_Get_Windows_Username = "Robert M. Vigneault" Then
    
        Call SlideIn_Comptabilite
        
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
    
    If Fn_Get_Windows_Username = "GCFiscalite" Or _
        Fn_Get_Windows_Username = "Robert M. Vigneault" Then
    
        Call SlideIn_Parametres
        
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
    wshMenu.Activate
    wshMenu.Range("A1").Select

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub




Attribute VB_Name = "modMenu"
Option Explicit

Dim width As Long

Sub SlideOut_TEC()
    
    With wshMenu.Shapes("btnTECMenu")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoTEC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "TEC"
    End With
    
End Sub

Sub SlideIn_TEC()
        
    With wshMenu.Shapes("btnTECMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoTEC").Left = width - 32
        Next width
        On Error Resume Next
        wshMenu.Unprotect
        On Error GoTo 0
        .TextFrame2.TextRange.Characters.text = ""
    End With
    
End Sub

Sub SlideOut_Facturation()
    
    With wshMenu.Shapes("btnFacturationMenu")
        For width = 32 To MAXWIDTH
            .Height = width
          ActiveSheet.Shapes("icoFacturation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Facturation"
    End With

End Sub

Sub SlideIn_Facturation()
    
    With wshMenu.Shapes("btnFacturationMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoFacturation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Debours()
    
    With wshMenu.Shapes("btnDeboursMenu")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Débours"
    End With

End Sub

Sub SlideIn_Debours()
    
    With wshMenu.Shapes("btnDeboursMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideIn_Comptabilite()
    
    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Comptabilite()
    
    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Comptabilité"
    End With

End Sub

Sub SlideOut_Parametres()
    
    With wshMenu.Shapes("btnParametresOption")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Paramètres"
    End With

End Sub

Sub SlideIn_Parametres()
    
    With wshMenu.Shapes("btnParametresOption")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Exit()
    
    With ActiveSheet.Shapes("btnEXIT")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEXIT").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Sortie"
    End With

End Sub

Sub SlideIn_Exit()
    
    With wshMenu.Shapes("btnEXIT")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoEXIT").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

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

Sub SlideOut_PrepFact()

    With wshMenuFAC.Shapes("btnPrepFact")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Préparation de facture"
    End With

End Sub

Sub SlideIn_PrepFact()

    With wshMenuFAC.Shapes("btnPrepFact")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_SuiviCC()

    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Suivi de C/C"
    End With

End Sub

Sub SlideIn_SuiviCC()

    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Encaissement()

    With wshMenuFAC.Shapes("btnEncaissement")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Encaissement"
    End With

End Sub

Sub SlideIn_Encaissement()

    With wshMenuFAC.Shapes("btnEncaissement")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_FAC_Historique()

    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Historique factures"
    End With

End Sub

Sub SlideIn_FAC_Historique()

    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_FAC_Annulation()

    With wshMenuFAC.Shapes("btnFAC_Annulation")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoFAC_Annulation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Annulation de facture"
    End With

End Sub

Sub SlideIn_FAC_Annulation()

    With wshMenuFAC.Shapes("btnFAC_Annulation")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFAC_Annulation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Regularisation()

    With wshMenuFAC.Shapes("btnRegularisation")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoRegularisation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Régularisation"
    End With

End Sub

Sub SlideIn_Regularisation()

    With wshMenuFAC.Shapes("btnRegularisation")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoRegularisation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Paiement()

    With wshMenuDEB.Shapes("btnPaiement")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Déboursé"
    End With

End Sub

Sub SlideIn_Paiement()

    With wshMenuDEB.Shapes("btnPaiement")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_EJ()

    With wshMenuGL.Shapes("btnEJ")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Entrée de Journal"
    End With

End Sub

Sub SlideIn_EJ()

    With wshMenuGL.Shapes("btnEJ")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_BV()

    With wshMenuGL.Shapes("btnBV")
        For width = 32 To 180
            .Height = width
            ActiveSheet.Shapes("icoBV").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Balance de Vérification"
    End With

End Sub

Sub SlideIn_BV()

    With wshMenuGL.Shapes("btnBV")
        For width = 180 To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoBV").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_GL_Report()

    With wshMenuGL.Shapes("btnGL")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Rapport des transactions"
    End With

End Sub

Sub SlideIn_GL_Report()

    With wshMenuGL.Shapes("btnGL")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_EF()

    With wshMenuGL.Shapes("btnEF")
        For width = 32 To MAXWIDTH
            .Height = width
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "États financiers"
    End With

End Sub

Sub SlideIn_EF()

    With wshMenuGL.Shapes("btnEF")
        For width = MAXWIDTH To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

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
    
    wshMenuFAC.Visible = xlSheetVisible
    wshMenuFAC.Activate
    wshMenuFAC.Range("A1").Select

End Sub

Sub MenuDEB_Click()
    
    Call SlideIn_Debours
    
    wshMenuDEB.Visible = xlSheetVisible
    wshMenuDEB.Activate
    wshMenuDEB.Range("A1").Select

End Sub

Sub menuComptabilite_Click()
    
    Call SlideIn_Comptabilite
    
    wshMenuGL.Visible = xlSheetVisible
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select

End Sub

Sub menuParametres_Click()
    
    Call SlideIn_Parametres
    
    wshAdmin.Visible = xlSheetVisible
    wshAdmin.Select
    
End Sub

Sub Exit_Without_Saving() '2024-06-20 @ 13:48
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & _
                        "l'application de gestion (sans sauvegarde) ?", vbYesNo + vbQuestion, "Confirmation de sortie")
    
    If confirmation = vbYes Then
        Call SlideIn_Exit
        Call Hide_All_Worksheets_Except_Menu
    
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
        Call Output_Timer_Results("message:This  session  has  been  terminated A B N O R M A L L Y", 0)
        
        Dim wb As Workbook: Set wb = ActiveWorkbook
        ActiveWorkbook.Close saveChanges:=False

        Application.Application.Quit
    End If
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set wb = Nothing
    
End Sub

Sub Exit_Click() '2024-07-05 @ 06:37
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & vbNewLine & _
                        "l'application de gestion ?" & vbNewLine, vbYesNo + vbQuestion, "Confirmation de sortie")
    
    If confirmation = vbYes Then
        Call SlideIn_Exit
        Call Hide_All_Worksheets_Except_Menu
        
        'Delete temporary Worksheet (Feuil*)
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If InStr(1, ws.CodeName, "Feuil") > 0 Then
                Application.DisplayAlerts = False
                ws.delete
                Application.DisplayAlerts = True
            End If
        Next ws
    
        Application.ScreenUpdating = True
        Application.EnableEvents = False
        
        Call Output_Timer_Results("message:This  session  has  been  terminated N O R M A L L Y", 0)
        
        Application.DisplayAlerts = False
        ThisWorkbook.Save
        Application.DisplayAlerts = True
        ThisWorkbook.Close saveChanges:=False
    
        'If Personal.xlsb is open, hide it without saving
        Dim wb As Workbook
        On Error Resume Next
        Set wb = Workbooks("PERSONAL.XLSB")
        If Not wb Is Nothing Then
            wb.IsAddin = True
        End If
        On Error GoTo 0
    
        'Cleaning - 2024-07-05 @ 06:46
        Set wb = Nothing
        Set ws = Nothing
        
        Application.Quit
        
    End If
    
End Sub

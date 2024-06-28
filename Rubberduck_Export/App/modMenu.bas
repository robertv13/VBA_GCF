Attribute VB_Name = "modMenu"
Option Explicit

Dim width As Long
Public Const maxWidth As Integer = 150

Sub SlideOut_TEC()
    
    With wshMenu.Shapes("btnTECMenu")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoTEC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "TEC"
    End With
    
End Sub

Sub SlideIn_TEC()
        
    With wshMenu.Shapes("btnTECMenu")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoTEC").Left = width - 32
        Next width
        On Error Resume Next
        wshMenu.Unprotect
        On Error GoTo 0
        .TextFrame2.TextRange.Characters.text = ""
        wshMenu.Protect UserInterfaceOnly:=True
    End With
    
End Sub

Sub SlideOut_Facturation()
    
    With wshMenu.Shapes("btnFacturationMenu")
        For width = 32 To maxWidth
            .Height = width
          ActiveSheet.Shapes("icoFacturation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Facturation"
    End With

End Sub

Sub SlideIn_Facturation()
    
    With wshMenu.Shapes("btnFacturationMenu")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            wshMenu.Shapes("icoFacturation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Debours()
    
    With wshMenu.Shapes("btnDeboursMenu")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Débours"
    End With

End Sub

Sub SlideIn_Debours()
    
    With wshMenu.Shapes("btnDeboursMenu")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub
Sub SlideOut_Comptabilite()
    
    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Comptabilité"
    End With

End Sub

Sub SlideIn_Comptabilite()
    
    With wshMenu.Shapes("btnComptabiliteMenu")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Parametres()
    
    With wshMenu.Shapes("btnParametresOption")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Paramètres"
    End With

End Sub

Sub SlideIn_Parametres()
    
    With wshMenu.Shapes("btnParametresOption")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Exit()
    
    With ActiveSheet.Shapes("btnEXIT")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoEXIT").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Sortie"
    End With

End Sub

Sub SlideIn_Exit()
    
    With wshMenu.Shapes("btnEXIT")
        For width = maxWidth To 32 Step -1
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
'        For width = 32 To maxWidth
'            .Height = width
'            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Saisie des Heures"
'    End With

End Sub

Sub SlideIn_SaisieHeures()
    
'    With wshMenuTEC.Shapes("btnSaisieHeures")
'        For width = maxWidth To 32 Step -1
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideOut_TEC_TDB()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = 32 To maxWidth
'            .Height = width
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Tableau de bord"
'    End With

End Sub

Sub SlideIn_TEC_TDB()
    
'    With wshMenuTEC.Shapes("btnTEC_TDB")
'        For width = maxWidth To 32 Step -1
'            .Height = width
'            .Left = width - 32
'            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = ""
'    End With

End Sub

Sub SlideOut_PrepFact()

    With wshMenuFAC.Shapes("btnPrepFact")
        For width = 32 To 200
            .Height = width
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Préparation de facture"
    End With

End Sub

Sub SlideIn_PrepFact()

    With wshMenuFAC.Shapes("btnPrepFact")
        For width = 200 To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_SuiviCC()

    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Suivi de C/C"
    End With

End Sub

Sub SlideIn_SuiviCC()

    With wshMenuFAC.Shapes("btnSuiviCC")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Encaissement()

    With wshMenuFAC.Shapes("btnEncaissement")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Encaissement"
    End With

End Sub

Sub SlideIn_Encaissement()

    With wshMenuFAC.Shapes("btnEncaissement")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_FAC_Historique()

    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Historique factures"
    End With

End Sub

Sub SlideIn_FAC_Historique()

    With wshMenuFAC.Shapes("btnFAC_Historique")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFAC_Historique").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Regularisation()

    With wshMenuFAC.Shapes("btnRegularisation")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoRegularisation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Régularisation"
    End With

End Sub

Sub SlideIn_Regularisation()

    With wshMenuFAC.Shapes("btnRegularisation")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoRegularisation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_Paiement()

    With wshMenuDEB.Shapes("btnPaiement")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Déboursé"
    End With

End Sub

Sub SlideIn_Paiement()

    With wshMenuDEB.Shapes("btnPaiement")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_EJ()

    With wshMenuGL.Shapes("btnEJ")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Entrée de Journal"
    End With

End Sub

Sub SlideIn_EJ()

    With wshMenuGL.Shapes("btnEJ")
        For width = maxWidth To 32 Step -1
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
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Rapport - GL"
    End With

End Sub

Sub SlideIn_GL_Report()

    With wshMenuGL.Shapes("btnGL")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = ""
    End With

End Sub

Sub SlideOut_EF()

    With wshMenuGL.Shapes("btnEF")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "États financiers"
    End With

End Sub

Sub SlideIn_EF()

    With wshMenuGL.Shapes("btnEF")
        For width = maxWidth To 32 Step -1
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

Sub menuDebours_Click()
    
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

Sub EXIT_NO_SAVE() '2024-06-20 @ 13:48
    
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
        
        Dim wb As Workbook
        Set wb = ActiveWorkbook
        ActiveWorkbook.Close SaveChanges:=False

        Application.Application.Quit
    End If
    
End Sub

Sub EXIT_Click() '2024-06-20 @ 07:11
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox("Êtes-vous certain de vouloir quitter" & vbNewLine & _
                        "l'application de gestion ?", vbYesNo + vbQuestion, "Confirmation de sortie")
    
    If confirmation = vbYes Then
        Call SlideIn_Exit
        Call Hide_All_Worksheets_Except_Menu
    
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
        Call Output_Timer_Results("message:This  session  has  been  terminated N O R M A L L Y", 0)
        
        Dim wb As Workbook
        Set wb = ActiveWorkbook
        ActiveWorkbook.Close SaveChanges:=True

        Application.Application.Quit
    End If
    
End Sub
Sub Test_SaisieHeures()
    
    Dim shp As shape
    Set shp = wshMenuTEC.Shapes("btnSaisieHeures")
    
    Debug.Print "Out - " & shp.name & " H=" & Round(shp.Height, 0) & ", T=" & Round(shp.Top, 0) & ", L=" & Round(shp.Left, 0) & ", W=" & Round(shp.width, 0)
    
    shp.Left = 60
    shp.Top = 10
    shp.Height = 150
    shp.width = 32
    
    MsgBox "On slide_in"
    
    shp.Left = 0
    shp.Top = 10
    shp.Height = 32
    shp.width = 32
    
    Debug.Print "Out - " & shp.name & " H=" & Round(shp.Height, 0) & ", T=" & Round(shp.Top, 0) & ", L=" & Round(shp.Left, 0) & ", W=" & Round(shp.width, 0)
'    With wshMenuTEC.Shapes("btnSaisieHeures")
'        For width = 32 To maxWidth
'            .Height = width
'            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
'        Next width
'        .TextFrame2.TextRange.Characters.text = "Saisie des Heures"
'    End With

End Sub


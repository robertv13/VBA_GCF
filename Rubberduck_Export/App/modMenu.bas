Attribute VB_Name = "modMenu"
Option Explicit

Dim width As Long
Public Const maxWidth As Integer = 150

Sub SlideOut_TEC()
    With wshMenu.Shapes("btnTEC")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoTEC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "TEC"
    End With
End Sub

Sub SlideIn_TEC()
    With wshMenu.Shapes("btnTEC")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoTEC").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnTEC").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Facturation()
    With wshMenu.Shapes("btnFacturation")
        For width = 32 To maxWidth
            .Height = width
          ActiveSheet.Shapes("icoFacturation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Facturation"
    End With
End Sub

Sub SlideIn_Facturation()
    With wshMenu.Shapes("btnFacturation")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoFacturation").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnFacturation").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Debours()
    With wshMenu.Shapes("btnDebours")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Débours"
    End With
End Sub

Sub SlideIn_Debours()
    With wshMenu.Shapes("btnDebours")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoDebours").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnDebours").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub
Sub SlideOut_Comptabilite()
    With wshMenu.Shapes("btnComptabilite")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Comptabilité"
    End With
End Sub

Sub SlideIn_Comptabilite()
    With wshMenu.Shapes("btnComptabilite")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoComptabilite").Left = width - 32
        Next width
            ActiveSheet.Shapes("btnComptabilite").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub
Sub SlideOut_Parametres()
    With wshMenu.Shapes("btnParametres")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Paramètres"
    End With
End Sub

Sub SlideIn_Parametres()
    With wshMenu.Shapes("btnParametres")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoParametres").Left = width - 32
        Next width
            ActiveSheet.Shapes("btnParametres").TextFrame2.TextRange.Characters.text = ""
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
    With ActiveSheet.Shapes("btnEXIT")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEXIT").Left = width - 32
        Next width
            ActiveSheet.Shapes("btnExit").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

'Second level (sub-menu) ---------------------------------------------------------------------------
Sub SlideOut_SaisieHeures()
    With ActiveSheet.Shapes("btnSaisieHeures")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Saisie des Heures"
    End With
End Sub

Sub SlideIn_SaisieHeures()
    With ActiveSheet.Shapes("btnSaisieHeures")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoSaisieHeures").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnSaisieHeures").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_TEC_TDB()
    With wshMenuTEC.Shapes("btnTEC_TDB")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Tableau de bord"
    End With
End Sub

Sub SlideIn_TEC_TDB()
    With wshMenuTEC.Shapes("btnTEC_TDB")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoTEC_TDB").Left = width - 32
        Next width
        wshMenuTEC.Shapes("btnTEC_TDB").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_PrepFact()
    With wshMenuFACT.Shapes("btnPrepFact")
        For width = 32 To 200
            .Height = width
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Préparation de facture"
    End With
End Sub

Sub SlideIn_PrepFact()
    With wshMenuFACT.Shapes("btnPrepFact")
        For width = 200 To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPrepFact").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnPrepFact").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_SuiviCC()
    With wshMenuFACT.Shapes("btnSuiviCC")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Suivi de C/C"
    End With
End Sub

Sub SlideIn_SuiviCC()
    With wshMenuFACT.Shapes("btnSuiviCC")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoSuiviCC").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnSuiviCC").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Encaissement()
    With wshMenuFACT.Shapes("btnEncaissement")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Encaissement"
    End With
End Sub

Sub SlideIn_Encaissement()
    With wshMenuFACT.Shapes("btnEncaissement")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEncaissement").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnEncaissement").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Regularisation()
    With wshMenuFACT.Shapes("btnRegularisation")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoRegularisation").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Régularisation"
    End With
End Sub

Sub SlideIn_Regularisation()
    With wshMenuFACT.Shapes("btnRegularisation")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoRegularisation").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnRegularisation").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Paiement()
    With ActiveSheet.Shapes("btnPaiement")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Paiement"
    End With
End Sub

Sub SlideIn_Paiement()
    With ActiveSheet.Shapes("btnPaiement")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoPaiement").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnPaiement").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_EJ()
    With ActiveSheet.Shapes("btnEJ")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Entrée de Journal"
    End With
End Sub

Sub SlideIn_EJ()
    With ActiveSheet.Shapes("btnEJ")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEJ").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnEJ").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_GL_Report()
    With ActiveSheet.Shapes("btnGL")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Rapport - GL"
    End With
End Sub

Sub SlideIn_GL_Report()
    With ActiveSheet.Shapes("btnGL")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoGL").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnGL").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_BV()
    With ActiveSheet.Shapes("btnBV")
        For width = 32 To 180
            .Height = width
            ActiveSheet.Shapes("icoBV").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "Balance de Vérification"
    End With
End Sub

Sub SlideIn_BV()
    With ActiveSheet.Shapes("btnBV")
        For width = 180 To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoBV").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnBV").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_EF()
    With ActiveSheet.Shapes("btnEF")
        For width = 32 To maxWidth
            .Height = width
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        .TextFrame2.TextRange.Characters.text = "États financiers"
    End With
End Sub

Sub SlideIn_EF()
    With ActiveSheet.Shapes("btnEF")
        For width = maxWidth To 32 Step -1
            .Height = width
            .Left = width - 32
            ActiveSheet.Shapes("icoEF").Left = width - 32
        Next width
        ActiveSheet.Shapes("btnEF").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub menuTEC_Click() '2024-02-13 @ 13:48
    
    Dim timerStart As Double: timerStart = Timer
    
    SlideIn_TEC
    
    wshMenuTEC.Visible = xlSheetVisible
'    wshTEC_Local.Visible = xlSheetVisible
'    wshBD_Clients.Visible = xlSheetVisible
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

    Call Output_Timer_Results("menuTEC_Click()", timerStart)

End Sub

Sub menuFacturation_Click() '2024-02-13 @ 13:48
    
    Dim timerStart As Double: timerStart = Timer

    SlideIn_Facturation
    
    wshMenuFACT.Visible = xlSheetVisible
'    wshFAC_Brouillon.Visible = xlSheetVisible
'    wshBD_Clients.Visible = xlSheetVisible
'    wshFAC_Entête.Visible = xlSheetVisible
'    wshFAC_Détails.Visible = xlSheetVisible
'    wshFAC_Finale.Visible = xlSheetVisible
    
    wshMenuFACT.Activate
    wshMenuFACT.Range("A1").Select

    Call Output_Timer_Results("menuFacturation_Click()", timerStart)

End Sub

Sub menuDebours_Click() '2024-02-13 @ 13:48
    
    Dim timerStart As Double: timerStart = Timer
    
    SlideIn_Debours
    
    wshMenuDEBOURS.Visible = xlSheetVisible
'    wshPaiement.Visible = xlSheetVisible
    
    wshMenuDEBOURS.Activate
    wshMenuDEBOURS.Range("A1").Select

    Call Output_Timer_Results("menuDebours_Click()", timerStart)

End Sub

Sub menuComptabilite_Click() '2024-02-13 @ 13:48
    
    Dim timerStart As Double: timerStart = Timer
    
    SlideIn_Comptabilite
    
    wshMenuCOMPTA.Visible = xlSheetVisible
'    wshGL_EJ.Visible = xlSheetVisible
'    wshGL.Visible = xlSheetVisible
'    wshGL_EJ_Recurrente.Visible = xlSheetVisible
'    wshBV.Visible = xlSheetVisible
        
    wshMenuCOMPTA.Activate
    wshMenuCOMPTA.Range("A1").Select

    Call Output_Timer_Results("menuComptabilite_Click()", timerStart)

End Sub

Sub menuParametres_Click() '2024-02-13 @ 13:48
    
    Dim timerStart As Double: timerStart = Timer

    SlideIn_Parametres
    
    wshAdmin.Visible = xlSheetVisible
    wshAdmin.Select
    
    Call Output_Timer_Results("menuParametres_Click()", timerStart)
    
End Sub

Sub EXIT_Click() '2024-02-13 @ 13:48
    
    Application.EnableEvents = False
    
    Call SlideIn_Exit
    
    Call Hide_All_Worksheets_Except_Menu

    ThisWorkbook.Close SaveChanges:=True
    
    Application.Quit
    
    Application.EnableEvents = True
    
End Sub


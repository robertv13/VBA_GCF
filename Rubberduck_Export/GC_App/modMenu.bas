Attribute VB_Name = "modMenu"
Option Explicit

Dim Wdth As Long
Public Const maxWidth As Integer = 150

Sub SlideOut_TEC()
    With ActiveSheet.Shapes("btnTEC")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoTEC").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "TEC"
    End With
End Sub

Sub SlideIn_TEC()
    With ActiveSheet.Shapes("btnTEC")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoTEC").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnTEC").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Facturation()
    With ActiveSheet.Shapes("btnFacturation")
        For Wdth = 32 To maxWidth
            .Height = Wdth
          ActiveSheet.Shapes("icoFacturation").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Facturation"
    End With
End Sub

Sub SlideIn_Facturation()
    With ActiveSheet.Shapes("btnFacturation")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoFacturation").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnFacturation").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Debours()
    With ActiveSheet.Shapes("btnDebours")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoDebours").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Débours"
    End With
End Sub

Sub SlideIn_Debours()
    With ActiveSheet.Shapes("btnDebours")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoDebours").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnDebours").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub
Sub SlideOut_Comptabilite()
    With ActiveSheet.Shapes("btnComptabilite")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoComptabilite").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Comptabilité"
    End With
End Sub

Sub SlideIn_Comptabilite()
    With ActiveSheet.Shapes("btnComptabilite")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoComptabilite").Left = Wdth - 32
        Next Wdth
            ActiveSheet.Shapes("btnComptabilite").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub
Sub SlideOut_Parametres()
    With ActiveSheet.Shapes("btnParametres")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoParametres").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Paramètres"
    End With
End Sub

Sub SlideIn_Parametres()
    With ActiveSheet.Shapes("btnParametres")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoParametres").Left = Wdth - 32
        Next Wdth
            ActiveSheet.Shapes("btnParametres").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

'Second level (sub-menu) ---------------------------------------------------------------------------
Sub SlideOut_SaisieHeures()
    With ActiveSheet.Shapes("btnSaisieHeures")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoSaisieHeures").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Saisie des Heures"
    End With
End Sub

Sub SlideIn_SaisieHeures()
    With ActiveSheet.Shapes("btnSaisieHeures")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoSaisieHeures").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnSaisieHeures").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_ExportHeures()
    With ActiveSheet.Shapes("btnExportHeures")
        For Wdth = 32 To maxWidth
            .Height = Wdth
          ActiveSheet.Shapes("icoExportHeures").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Export des Heures"
    End With
End Sub

Sub SlideIn_ExportHeures()
    With ActiveSheet.Shapes("btnExportHeures")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoExportHeures").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnExportHeures").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_PrepFact()
    With ActiveSheet.Shapes("btnPrepFact")
        For Wdth = 32 To 182
            .Height = Wdth
            ActiveSheet.Shapes("icoPrepFact").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Préparation de facture"
    End With
End Sub

Sub SlideIn_PrepFact()
    With ActiveSheet.Shapes("btnPrepFact")
        For Wdth = 182 To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoPrepFact").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnPrepFact").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_SuiviCC()
    With ActiveSheet.Shapes("btnSuiviCC")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoSuiviCC").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Suivi de C/C"
    End With
End Sub

Sub SlideIn_SuiviCC()
    With ActiveSheet.Shapes("btnSuiviCC")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoSuiviCC").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnSuiviCC").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Encaissement()
    With ActiveSheet.Shapes("btnEncaissement")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoEncaissement").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Encaissement"
    End With
End Sub

Sub SlideIn_Encaissement()
    With ActiveSheet.Shapes("btnEncaissement")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoEncaissement").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnEncaissement").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Regularisation()
    With ActiveSheet.Shapes("btnRegularisation")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoRegularisation").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Régularisation"
    End With
End Sub

Sub SlideIn_Regularisation()
    With ActiveSheet.Shapes("btnRegularisation")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoRegularisation").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnRegularisation").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_Paiement()
    With ActiveSheet.Shapes("btnPaiement")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoPaiement").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Paiement"
    End With
End Sub

Sub SlideIn_Paiement()
    With ActiveSheet.Shapes("btnPaiement")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoPaiement").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnPaiement").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_EJ()
    With ActiveSheet.Shapes("btnEJ")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoEJ").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Entrée de Journal"
    End With
End Sub

Sub SlideIn_EJ()
    With ActiveSheet.Shapes("btnEJ")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoEJ").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnEJ").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_GL()
    With ActiveSheet.Shapes("btnGL")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoGL").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Grand Livre"
    End With
End Sub

Sub SlideIn_GL()
    With ActiveSheet.Shapes("btnGL")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoGL").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnGL").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_BV()
    With ActiveSheet.Shapes("btnBV")
        For Wdth = 32 To 182
            .Height = Wdth
            ActiveSheet.Shapes("icoBV").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "Balance de Vérification"
    End With
End Sub

Sub SlideIn_BV()
    With ActiveSheet.Shapes("btnBV")
        For Wdth = 182 To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoBV").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnBV").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub SlideOut_EF()
    With ActiveSheet.Shapes("btnEF")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoEF").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.text = "États financiers"
    End With
End Sub

Sub SlideIn_EF()
    With ActiveSheet.Shapes("btnEF")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoEF").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnEF").TextFrame2.TextRange.Characters.text = ""
    End With
End Sub

Sub TEC_Click()
    SlideIn_TEC
    
    wshMenuTEC.Visible = xlSheetVisible
    wshBaseHours.Visible = xlSheetVisible
    wshFilteredHours.Visible = xlSheetVisible
    wshClientDB.Visible = xlSheetVisible
    wshHoursToExport.Visible = xlSheetVisible
    
    wshMenuTEC.Select
End Sub

Sub Facturation_Click()
    SlideIn_Facturation
    
    wshMenuFACT.Visible = xlSheetVisible
    wshFACshInvoice.Visible = xlSheetVisible
    wshFACInvList.Visible = xlSheetVisible
    wshFACInvItems.Visible = xlSheetVisible
    wshFACshBillEntries.Visible = xlSheetVisible
    wshFACServItems.Visible = xlSheetVisible
    wshFACProjects.Visible = xlSheetVisible
    wshFACCustomers.Visible = xlSheetVisible
    wshFACshFactureFinale.Visible = xlSheetVisible
    
    wshMenuFACT.Select
End Sub

Sub Debours_Click()
    SlideIn_Debours
    
    wshMenuDEBOURS.Visible = xlSheetVisible
    wshPaiement.Visible = xlSheetVisible
    
    wshMenuDEBOURS.Select
End Sub

Sub Comptabilite_Click()
    SlideIn_Comptabilite
    
    wshMenuCOMPTA.Visible = xlSheetVisible
    wshJE.Visible = xlSheetVisible
    wshGL.Visible = xlSheetVisible
    wshEJRecurrente.Visible = xlSheetVisible
    wshBV.Visible = xlSheetVisible
        
    wshMenuCOMPTA.Select
End Sub

Sub Parametres_Click()
    SlideIn_Parametres
    With wshAdmin
        .Visible = xlSheetVisible
        .Select
    End With
End Sub

Sub SaisieHeures_Click()
    SlideIn_SaisieHeures
    Load frmSaisieHeures
    frmSaisieHeures.show vbModal
End Sub

Sub ExportHeures_Click()
    SlideIn_ExportHeures
    MsgBox "Activer la fonction 'Export des heures'"
End Sub

Sub PreparationFacture_Click()
    SlideIn_PrepFact
    MsgBox "Activer la fonction 'Préparation de Facture'"
End Sub

Sub SuiviCC_Click()
    SlideIn_SuiviCC
    MsgBox "Activer la fonction 'Suivi des C/C'"
End Sub

Sub Encaissement_Click()
    SlideIn_Encaissement
    MsgBox "Activer la fonction 'Encaissement'"
End Sub

Sub Regularisation_Click()
    SlideIn_Régularisation
    MsgBox "Activer la fonction 'Régularisation'"
End Sub

Sub Paiement_Click()
    SlideIn_Paiement
    MsgBox "Activer la fonction 'Paiement'"
End Sub

Sub EJ_Click()
    SlideIn_EJ
    MsgBox "Activer la fonction 'Entrée de Journal'"
End Sub

Sub GL_Click()
    SlideIn_GL
    MsgBox "Activer la fonction 'Grand Livre'"
End Sub

Sub BV_Click()
    SlideIn_BV
    MsgBox "Activer la fonction 'Balance de vérification'"
End Sub

Sub EF_Click()
    SlideIn_EF
    MsgBox "Activer la fonction 'États Financiers'"
End Sub

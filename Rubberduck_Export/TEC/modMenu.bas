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

Sub TEC_Click()
    SlideIn_TEC
    With wshMenuTEC
        .Visible = xlSheetVisible
        .Select
    End With
'    Load frmSaisieHeures
'    frmSaisieHeures.show vbModal
End Sub

Sub Facturation_Click()
    SlideIn_Facturation
    With wshMenuFACT
        .Visible = xlSheetVisible
        .Select
    End With
End Sub

Sub Debours_Click()
    SlideIn_Debours
    With wshMenuDEBOURS
        .Visible = xlSheetVisible
        .Select
    End With
End Sub

Sub Comptabilite_Click()
    SlideIn_Comptabilite
    With wshMenuCOMPTA
        .Visible = xlSheetVisible
        .Select
    End With
End Sub

Sub Parametres_Click()
    SlideIn_Parametres
    With wshAdmin
        .Visible = xlSheetVisible
        .Select
    End With
End Sub


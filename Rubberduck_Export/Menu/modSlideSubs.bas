Attribute VB_Name = "modSlideSubs"
Option Explicit

Dim Wdth As Long
Public Const maxWidth As Integer = 150

Sub SlideOut_TEC()
    With ActiveSheet.Shapes("btnTEC")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoTEC").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "TEC"
    End With
End Sub

Sub SlideIn_TEC()
    With ActiveSheet.Shapes("btnTEC")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoTEC").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnTEC").TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub

Sub SlideOut_Facturation()
    With ActiveSheet.Shapes("btnFacturation")
        For Wdth = 32 To maxWidth
            .Height = Wdth
          ActiveSheet.Shapes("icoFacturation").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Facturation"
    End With
End Sub

Sub SlideIn_Facturation()
    With ActiveSheet.Shapes("btnFacturation")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoFacturation").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnFacturation").TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub
Sub SlideOut_Debours()
    With ActiveSheet.Shapes("btnDebours")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoDebours").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Débours"
    End With
End Sub

Sub SlideIn_Debours()
    With ActiveSheet.Shapes("btnDebours")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoDebours").Left = Wdth - 32
        Next Wdth
        ActiveSheet.Shapes("btnDebours").TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub
Sub SlideOut_Comptabilite()
    With ActiveSheet.Shapes("btnComptabilite")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoComptabilite").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Comptabilité"
    End With
End Sub

Sub SlideIn_Comptabilite()
    With ActiveSheet.Shapes("btnComptabilite")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoComptabilite").Left = Wdth - 32
        Next Wdth
            ActiveSheet.Shapes("btnComptabilite").TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub
Sub SlideOut_Parametres()
    With ActiveSheet.Shapes("btnParametres")
        For Wdth = 32 To maxWidth
            .Height = Wdth
            ActiveSheet.Shapes("icoParametres").Left = Wdth - 32
        Next Wdth
        .TextFrame2.TextRange.Characters.Text = "Paramètres"
    End With
End Sub

Sub SlideIn_Parametres()
    With ActiveSheet.Shapes("btnParametres")
        For Wdth = maxWidth To 32 Step -1
            .Height = Wdth
            .Left = Wdth - 32
            ActiveSheet.Shapes("icoParametres").Left = Wdth - 32
        Next Wdth
            ActiveSheet.Shapes("btnParametres").TextFrame2.TextRange.Characters.Text = ""
    End With
End Sub

Sub TEC_Click()
    MsgBox "This is the TEC button"
End Sub

Sub Facturation_Click()
    MsgBox "This is the Facturation button"
End Sub

Sub Debours_Click()
    MsgBox "This is the Debours button"
End Sub

Sub Comptabilite_Click()
    MsgBox "This is the Comptabilité button"
End Sub

Sub Parametres_Click()
    MsgBox "This is the Parametres button"
End Sub

Sub GetAllShapes()

    Dim shp As Shape
    'Loop through each shape on ActiveSheet
    For Each shp In ActiveSheet.Shapes
        Debug.Print shp.Name & " est [" & shp.Left & "," & shp.Top & "] - " & shp.Width
'        If shp.Name = "icoParametres" Then
'            shp.Left = 2
'            shp.Top = 310
'            Debug.Print shp.Name & " = " & shp.Top
'        End If
    Next shp

End Sub


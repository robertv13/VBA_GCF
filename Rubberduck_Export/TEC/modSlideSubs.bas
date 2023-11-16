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
    Load frmSaisieHeures
    frmSaisieHeures.show vbModal
End Sub

Sub Facturation_Click()
    SlideIn_Facturation
    MsgBox "This is the Facturation button"
End Sub

Sub Debours_Click()
    SlideIn_Debours
    MsgBox "This is the Debours button"
End Sub

Sub Comptabilite_Click()
    SlideIn_Comptabilite
    MsgBox "This is the Comptabilité button"
End Sub

Sub Parametres_Click()
    SlideIn_Parametres
    wshAdmin.Select
End Sub

Sub GetAllShapes()

    Dim shp As Shape
    'Loop through each shape on ActiveSheet
    For Each shp In ActiveSheet.Shapes
        Debug.Print shp.Name & " est [" & shp.Left & "," & shp.Top & "] - " & shp.Type
'        If shp.Name = "icoParametres" Then
'            shp.Left = 2
'            shp.Top = 310
'            Debug.Print shp.Name & " = " & shp.Top
'        End If
    Next shp

End Sub

Sub GetShapeProperties() 'List Properties of all the shapes

    Dim sShapes As Shape, lLoop As Long
    Dim wshNew As Worksheet
    Set wshNew = Sheets.Add
    'Add headings for our lists. Expand as needed
    wshNew.Range("A1:G1") = Array("Type", "Name", "ZOrder", "Height", "Width", "Left", "Top")

    'Loop through all shapes on active sheet
    For Each sShapes In wshMenu.Shapes
        'Increment Variable lLoop for row numbers
        lLoop = lLoop + 1
        With sShapes
            'Add shape properties
            wshNew.Cells(lLoop + 1, 1) = .Type
            wshNew.Cells(lLoop + 1, 2) = .Name
            wshNew.Cells(lLoop + 1, 3) = .ZOrderPosition
            wshNew.Cells(lLoop + 1, 4) = .Height
            wshNew.Cells(lLoop + 1, 5) = .Width
            wshNew.Cells(lLoop + 1, 6) = .Left
            wshNew.Cells(lLoop + 1, 7) = .Top
            'Follow the same pattern for more
        End With
    Next sShapes
    'AutoFit Columns.
    wshNew.Columns.AutoFit
End Sub


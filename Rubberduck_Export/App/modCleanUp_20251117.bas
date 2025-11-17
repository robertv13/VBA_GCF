Attribute VB_Name = "modCleanUp_20251117"
'Option Explicit
'
'Sub AuditerPlageVersSpecsLanceur()
'
'    Call AuditerPlageVersSpecs(Sheets("Bilan").Range("A1:E50"))
'
'End Sub
'
'Public Sub AuditerPlageVersSpecs(ByVal plage As Range)
'
'    Dim wsSource As Worksheet, wsSpecs As Worksheet
'    Dim nomSpecs As String
'    Dim cell As Range
'
'    If plage Is Nothing Then
'        MsgBox "La plage fournie est vide ou invalide.", vbCritical
'        Exit Sub
'    End If
'
'    Set wsSource = plage.Worksheet
'    nomSpecs = wsSource.Name & "_Specs"
'
'    ' Supprimer ancienne feuille _Specs si elle existe
'    Application.DisplayAlerts = False
'    On Error Resume Next: ThisWorkbook.Worksheets(nomSpecs).Delete: On Error GoTo 0
'    Application.DisplayAlerts = True
'
'    Set wsSpecs = ThisWorkbook.Worksheets.Add(After:=wsSource)
'    wsSpecs.Name = nomSpecs
'
'    ' En-têtes
'    With wsSpecs
'        .Range("A1:P1").Value = Array("Adresse", "Valeur", "Format", "Police", "Taille", "Gras", "Italique", "Couleur police", _
'                              "Couleurfond", "Align Hor", "Align Ver", "WrapText", "Fusion", "Hauteur ligne", "Largeur colonne", "Formule")
'        .Rows(1).Font.Bold = True
'    End With
'
'    Dim auditData() As Variant
'    Dim ligne As Long, col As Long
'    Dim nbLignes As Long: nbLignes = plage.Cells.count
'    Dim nbColonnes As Long: nbColonnes = 16
'    ReDim auditData(1 To nbLignes, 1 To nbColonnes)
'
'    ligne = 1
'    For Each cell In plage.Cells
'            auditData(ligne, 1) = cell.Address
'            auditData(ligne, 2) = cell.Value
'            If Not cell.NumberFormat = "# ##0_);(# ##0)" Then auditData(ligne, 3) = cell.NumberFormat
'            If Not cell.Font.Name = "Verdana" Then auditData(ligne, 4) = cell.Font.Name
'            If Not cell.Font.size = 11 Then auditData(ligne, 5) = cell.Font.size
'            If cell.Font.Bold Then auditData(ligne, 6) = True
'            If cell.Font.Italic Then auditData(ligne, 7) = True
'            If Not CouleurLisible(cell.Font.Color) = "#625850" Then auditData(ligne, 8) = CouleurLisible(cell.Font.Color)
'            If Not CouleurLisible(cell.Interior.Color) = "#FFFFFF" Then auditData(ligne, 9) = CouleurLisible(cell.Interior.Color)
'            If Not TexteAlignementHorizontal(cell.HorizontalAlignment) = "Gauche" Then auditData(ligne, 10) = TexteAlignementHorizontal(cell.HorizontalAlignment)
'            If Not TexteAlignementVertical(cell.VerticalAlignment) = "Bas" Then auditData(ligne, 11) = TexteAlignementVertical(cell.VerticalAlignment)
'            If cell.WrapText = True Then auditData(ligne, 12) = True
'            If cell.MergeCells = True Then auditData(ligne, 13) = True
'            If Not plage.Worksheet.Rows(cell.row).RowHeight = 14.25 Then auditData(ligne, 14) = plage.Worksheet.Rows(cell.row).RowHeight
'            auditData(ligne, 15) = plage.Worksheet.Columns(cell.Column).ColumnWidth
'            auditData(ligne, 16) = IIf(cell.HasFormula, "'" & cell.formula, "")
'            ligne = ligne + 1
'    Next cell
'
'    wsSpecs.Range("A2").Resize(nbLignes, nbColonnes).Value = auditData
'
'    wsSpecs.Columns.AutoFit
'
'    MsgBox "Audit de la plage '" & plage.Address & "' terminé dans '" & nomSpecs & "'.", vbInformation
'
'End Sub
'
'Private Function CouleurLisible(valeurLong As Long) As String
'    If valeurLong = -4105 Then
'        CouleurLisible = "(Automatique)"
'    Else
'        Dim r As Long, g As Long, b As Long
'        r = valeurLong Mod 256
'        g = (valeurLong \ 256) Mod 256
'        b = (valeurLong \ 65536) Mod 256
'        CouleurLisible = "#" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
'    End If
'End Function
'
'Private Function TexteAlignementHorizontal(code As Variant) As String
'
'    Select Case code
'        Case xlGeneral: TexteAlignementHorizontal = "Général"
'        Case xlLeft: TexteAlignementHorizontal = "Gauche"
'        Case xlCenter: TexteAlignementHorizontal = "Centre"
'        Case xlRight: TexteAlignementHorizontal = "Droite"
'        Case xlFill: TexteAlignementHorizontal = "Remplissage"
'        Case xlJustify: TexteAlignementHorizontal = "Justifié"
'        Case xlCenterAcrossSelection: TexteAlignementHorizontal = "Centré sur sélection"
'        Case xlDistributed: TexteAlignementHorizontal = "Distribué"
'        Case Else: TexteAlignementHorizontal = "(Inconnu)"
'    End Select
'
'End Function
'
'Private Function TexteAlignementVertical(code As Variant) As String
'    Select Case code
'        Case xlTop: TexteAlignementVertical = "Haut"
'        Case xlCenter: TexteAlignementVertical = "Centre"
'        Case xlBottom: TexteAlignementVertical = "Bas"
'        Case xlJustify: TexteAlignementVertical = "Justifié"
'        Case xlDistributed: TexteAlignementVertical = "Distribué"
'        Case Else: TexteAlignementVertical = "(Inconnu)"
'    End Select
'End Function
'
'Sub zz_TesterAppMsgBox()
'
'    Dim r As VbMsgBoxResult
'    r = Fn_AppMsgBox("Voulez-vous continuer ?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirmation avant de continuer")
'
'    Debug.Print "#090 - " & r
'
'End Sub
'
'Sub zz_ScannerSuppressionAmbigue_VersFenetreImmediate() '2025-07-01 @ 09:36
'
'    Dim vbComp As Object
'    Dim vbMod As Object
'    Dim numLigne As Long
'    Dim ligneCode As String
'    Dim motsCibles As Variant
'    Dim mot As Variant
'
'    motsCibles = Array("Delete", "xlDialogEditDelete", "Selection.Delete", "SendKeys")
'
'    Debug.Print "?? --- Résultats du scan suppression VBA ---"
'
'    For Each vbComp In ThisWorkbook.VBProject.VBComponents
'        Set vbMod = vbComp.codeModule
'        For numLigne = 1 To vbMod.CountOfLines
'            ligneCode = Trim(vbMod.Lines(numLigne, 1))
'
'            ' Ignorer les lignes vides ou les commentaires purs
'            If ligneCode <> vbNullString And Left(ligneCode, 1) <> "'" Then
'                For Each mot In motsCibles
'                    If InStr(1, ligneCode, mot, vbTextCompare) > 0 Then
'                        ' Écarter les suppressions explicites de lignes ou colonnes
'                        If Not ligneCode Like "*EntireRow.Delete*" And _
'                           Not ligneCode Like "*EntireColumn.Delete*" Then
'                            Debug.Print vbComp.Name & " [Ligne " & numLigne & "] : " & ligneCode
'                        End If
'                        Exit For
'                    End If
'                Next mot
'            End If
'        Next numLigne
'    Next vbComp
'
'    Debug.Print "--- Scan terminé ---"
'
'End Sub
'
'Function Fn_PadDroite(text As String, longueur As Integer) As String '2025-07-03 @ 17:54
'
'    Fn_PadDroite = Left(text & Space(longueur), longueur)
'
'End Function
'
'
'

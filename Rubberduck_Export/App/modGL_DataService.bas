Attribute VB_Name = "modGL_DataService"
Option Explicit

Function ConstruirePlanComptable() As Object '2025-06-01 @ 08:20

    Dim dictComptes As Object
    Set dictComptes = CreateObject("Scripting.Dictionary")
    
    Dim dataPlan As Variant
    dataPlan = Fn_Get_Plan_Comptable(2)
    
    Dim i As Long
    For i = 1 To UBound(dataPlan, 1)
        dictComptes(dataPlan(i, 1)) = dataPlan(i, 2)
    Next i
    
    Set ConstruirePlanComptable = dictComptes
    
End Function

Function CreerCopieTemporaireSolide(onglet As String) As String

    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim wbTmp As Workbook
    Dim sPath As String, sFichier As String
    Dim oldScreenUpdating As Boolean
    Dim lastRow As Long, lastCol As Long

    On Error GoTo ErrHandler

    sPath = ThisWorkbook.path & DATA_PATH & "\"
    If Dir(sPath, vbDirectory) = "" Then
        MsgBox "Le répertoire n'existe pas : " & vbCrLf & sPath, vbCritical
        CreerCopieTemporaireSolide = ""
        Exit Function
    End If

    sFichier = "GL_Temp_" & Environ("Username") & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set wsSrc = ThisWorkbook.Worksheets(onglet)
    Set wbTmp = Workbooks.Add(xlWBATWorksheet)
    Set wsDest = wbTmp.Sheets(1)

    ' Déterminer la zone utilisée
    With wsSrc
        lastRow = .Cells(.Rows.count, 1).End(xlUp).row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
    End With

    ' Copier les valeurs uniquement
    wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(lastRow, lastCol)).Value = _
        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value

    ' Optionnel : nommer la feuille comme l’originale
    On Error Resume Next: wsDest.Name = wsSrc.Name: On Error GoTo 0

    ' Sauvegarde
    Application.DisplayAlerts = False
    wbTmp.SaveAs fileName:=sPath & sFichier, FileFormat:=xlOpenXMLWorkbook
    wbTmp.Close SaveChanges:=False
    Application.DisplayAlerts = True

    Application.ScreenUpdating = oldScreenUpdating
    CreerCopieTemporaireSolide = sPath & sFichier
    Exit Function

ErrHandler:
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "Erreur lors de la création du fichier temporaire : " & Err.description, vbCritical
    CreerCopieTemporaireSolide = ""
    
End Function

'CommentOut - 2025-06-06 @ 11:02
'Function CréerCopieTemporaireSansFlash(onglet As String) As String '2025-06-03 @ 21:07
'
'    Dim wsSrc As Worksheet
'    Dim wbTmp As Workbook
'    Dim sPath As String, sFichier As String
'    Dim oldScreenUpdating As Boolean
'
'    On Error GoTo ErrHandler
'
'    ' Répertoire
'    sPath = ThisWorkbook.path & DATA_PATH & "\"
'    If Dir(sPath, vbDirectory) = "" Then
'        MsgBox "Le répertoire n'existe pas : " & vbCrLf & sPath, vbCritical
'        CréerCopieTemporaireSansFlash = ""
'        Exit Function
'    End If
'
'    ' Fichier temporaire
'    sFichier = "GL_Temp_" & Environ("Username") & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
'    oldScreenUpdating = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'
'    'Feuille source
'    Set wsSrc = ThisWorkbook.Worksheets(onglet)
'    If wsSrc.ProtectContents Then wsSrc.Unprotect
'
'    ' Copier la feuille dans un nouveau classeur
'    wsSrc.Copy ' <-- Crée un nouveau classeur automatiquement !
'
'    'Ce nouveau classeur devient actif après le .Copy
'    Set wbTmp = ActiveWorkbook
'
'    ' Sauvegarde
'    Application.DisplayAlerts = False
'    wbTmp.SaveAs fileName:=sPath & sFichier, FileFormat:=xlOpenXMLWorkbook
'    wbTmp.Close SaveChanges:=False
'    Application.DisplayAlerts = True
'
'    Application.ScreenUpdating = oldScreenUpdating
'    CréerCopieTemporaireSansFlash = sPath & sFichier
'    Exit Function
'
'ErrHandler:
'    Application.ScreenUpdating = oldScreenUpdating
'    MsgBox "Erreur lors de la création du fichier temporaire : " & Err.description, vbCritical
'    CréerCopieTemporaireSansFlash = ""
'
'End Function
'


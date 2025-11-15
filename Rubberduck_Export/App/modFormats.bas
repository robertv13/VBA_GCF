Attribute VB_Name = "modFormats"
Option Explicit

Public Const FMT_DATE As String = "yyyy-mm-dd"
Public Const FMT_DATE_HEURE As String = "yyyy-mm-dd hh:mm:ss"
Public Const FMT_MNT_CURRENCY As String = "#,##0.00"
Public Const FMT_MNT_CURR_DOLLARS As String = "#,##0.00 $"
Public Const FMT_TAUX_PCT_3 As String = "#0.000 %"
Public Const FMT_ENTIER As String = "0"

'Alignements
Public Sub SetAlignLeft(r As Range)

    If Not r Is Nothing Then r.HorizontalAlignment = xlLeft
    
End Sub

Public Sub SetAlignCenter(r As Range)

    If Not r Is Nothing Then r.HorizontalAlignment = xlCenter
    
End Sub

Public Sub SetAlignRight(r As Range)

    If Not r Is Nothing Then r.HorizontalAlignment = xlRight
    
End Sub

'Formats
Public Sub SetNumberFormat(r As Range, ByVal nf As String)

    If Not r Is Nothing Then r.NumberFormat = nf
    
End Sub

'Colonnes
Public Sub SetColWidth(ws As Worksheet, ByVal colIndex As Long, ByVal widthChars As Double)

    ws.Columns(colIndex).ColumnWidth = widthChars
    
End Sub

'Post-traitements communs
Public Sub AppliquerCommonPost(ws As Worksheet, lo As ListObject)

    On Error Resume Next
    lo.Range.EntireColumn.AutoFit
    lo.DataBodyRange.RowHeight = 15
    On Error GoTo 0
    
End Sub

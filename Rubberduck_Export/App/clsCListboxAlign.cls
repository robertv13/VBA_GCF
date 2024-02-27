VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCListboxAlign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PLACE IN A CLASS MODULE
'-------------------------------------------------------------------------------------
'--- Developer: Ryan Wells (wellsr.com)-----------------------------------------------
'--- Date: 05/2017--------------------------------------------------------------------
'--- Class: clsCListboxAlign----------------------------------------------------------
'--- Purpose: Align the text in different columns in a UserForm ListBox differently---
'--- This has been adapted from a few sources I stumbled across many moons ago but ---
'--- I don't recall the sources.------------------------------------------------------
'-------------------------------------------------------------------------------------

Public Sub Center(LBox As MSForms.ListBox, Optional WhichColumn As Integer = 0)
    '
    ' PURPOSE: Center align the text in a listbox column
    ' HOW TO USE:
    ' - First argument is the listbox you want to adjust, the second optional argument is which
    ' column in the listbox you want to align.
    ' - To use this procedure, you would place a statement like the following in your UserForm_Initialize routine:
    ' MyListBoxClass.Center Me.lstEcriture, 1
    '
    Dim labSizer As MSForms.Label
    Dim lngIndex As Long
    Dim intColumn As Integer
    Dim lngTopIndex As Long
    Dim vntColWidths As Variant
    
    'Get label control to help size text
    Set labSizer = m_GetSizer(LBox.Parent)
    If labSizer Is Nothing Then Exit Sub
    
    ReDim sngWidth(LBox.ColumnCount) As Single
    If Len(LBox.ColumnWidths) > 0 Then
        'Decode column widths
        vntColWidths = Split(LBox.ColumnWidths, ";")
        'Fudge for gap between cols
        For intColumn = 1 To LBox.ColumnCount
            sngWidth(intColumn) = Val(vntColWidths(intColumn - 1)) - 5
        Next
    Else
        'Assume default sizes
        For intColumn = 1 To LBox.ColumnCount
            sngWidth(intColumn) = (LBox.width - (15 * LBox.ColumnCount)) / LBox.ColumnCount
        Next intColumn
    End If
    
    'Generic font attributes
    With labSizer
        With .Font
            .name = LBox.Font.name
            .Size = LBox.Font.Size
            .Bold = LBox.Font.Bold
            .Italic = LBox.Font.Italic
        End With
        .WordWrap = False
    End With
    
    'Begin processing column width to center align
    lngTopIndex = LBox.TopIndex
    For intColumn = 1 To LBox.ColumnCount
        If intColumn = WhichColumn Or WhichColumn = -1 Then
            ' if you say to center align this column or center align all columns
            For lngIndex = 0 To LBox.ListCount - 1
                LBox.TopIndex = lngIndex
                labSizer.width = LBox.width
                labSizer.Caption = Trim(LBox.List(lngIndex, intColumn - 1))
                labSizer.AutoSize = True
                Do While labSizer.width < sngWidth(intColumn)
                    labSizer.Caption = " " & labSizer.Caption & " "
                Loop
                LBox.List(lngIndex, intColumn - 1) = labSizer.Caption
            Next lngIndex
        End If
    Next intColumn
    LBox.TopIndex = lngTopIndex
    
    LBox.Parent.Controls.Remove labSizer.name
    Set labSizer = Nothing

End Sub

Public Sub Left(LBox As MSForms.ListBox, Optional WhichColumn As Integer = 0)
    '
    ' PURPOSE: Left align the text in a listbox column
    ' HOW TO USE:
    ' - First argument is the listbox you want to adjust, the second optional argument is which
    ' column in the listbox you want to align.
    ' - To use this procedure, you would place a statement like the following in your UserForm_Initialize routine:
    ' MyListBoxClass.Left Me.lstEcriture, 1
    '
    Dim lngIndex As Long
    Dim intColumn As Integer
    Dim lngTopIndex As Long
    Dim vntColWidths As Variant
    ReDim sngWidth(LBox.ColumnCount) As Single
    
    If Len(LBox.ColumnWidths) > 0 Then
        'Decode column widths
        vntColWidths = Split(LBox.ColumnWidths, ";")
        'Fudge for gap between cols
        For intColumn = 1 To LBox.ColumnCount
            sngWidth(intColumn) = Val(vntColWidths(1)) - 5
        Next
    Else
        'Assume default sizes
        For intColumn = 1 To LBox.ColumnCount
            sngWidth(intColumn) = (LBox.width - (15 * LBox.ColumnCount)) / LBox.ColumnCount
        Next intColumn
    End If
    
    lngTopIndex = LBox.TopIndex
    For intColumn = 1 To LBox.ColumnCount
        If intColumn = WhichColumn Or WhichColumn = -1 Then
            'If you say to left align this column or left align all columns
            For lngIndex = 0 To LBox.ListCount - 1
                LBox.TopIndex = lngIndex
                LBox.List(lngIndex, intColumn - 1) = Trim(LBox.List(lngIndex, intColumn - 1))
            Next lngIndex
        End If
    Next intColumn
    LBox.TopIndex = lngTopIndex
    
End Sub

Public Sub Right(LBox As MSForms.ListBox, Optional WhichColumn As Integer = 1)
    '
    ' PURPOSE: Right align the text in a listbox column
    ' HOW TO USE:
    ' - First argument is the listbox you want to adjust, the second optional argument is which
    ' column in the listbox you want to align.
    ' - To use this procedure, you would place a statement like the following in your UserForm_Initialize routine:
    ' MyListBoxClass.Right Me.lstEcriture, 1
    '
    Dim labSizer As MSForms.Label
    Dim lngIndex As Long
    Dim intColumn As Integer
    Dim lngTopIndex As Long
    Dim vntColWidths As Variant
    'Get label control to help size text
    Set labSizer = m_GetSizer(LBox.Parent)
    If labSizer Is Nothing Then Exit Sub
    
    ReDim sngWidth(LBox.ColumnCount) As Single
    If Len(LBox.ColumnWidths) > 0 Then
        ' decode column widths
        vntColWidths = Split(LBox.ColumnWidths, ";")
        ' fudge for gap between cols
        For intColumn = 1 To LBox.ColumnCount
            sngWidth(intColumn) = Val(vntColWidths(1)) - 10
        Next
    Else
        ' assume default sizes
        For intColumn = 1 To LBox.ColumnCount
            sngWidth(intColumn) = (LBox.width - (25 * LBox.ColumnCount)) / LBox.ColumnCount
        Next intColumn
    End If
    
    With labSizer
        With .Font
            .name = LBox.Font.name
            .Size = LBox.Font.Size
            .Bold = LBox.Font.Bold
            .Italic = LBox.Font.Italic
        End With
        .WordWrap = False
    End With
    
    lngTopIndex = LBox.TopIndex
    For intColumn = 1 To LBox.ColumnCount
        If intColumn = WhichColumn Or WhichColumn = -1 Then
            'If you say to right align this column, or right align all columns
            For lngIndex = 0 To LBox.ListCount - 1
                LBox.TopIndex = lngIndex
                labSizer.width = LBox.width
                labSizer.Caption = Trim(LBox.List(lngIndex, intColumn - 1))
                labSizer.AutoSize = True
                Do While labSizer.width < sngWidth(intColumn)
                    labSizer.Caption = " " & labSizer.Caption
                Loop
                LBox.List(lngIndex, intColumn - 1) = labSizer.Caption
            Next lngIndex
        End If
    Next intColumn
    LBox.TopIndex = lngTopIndex
    LBox.Parent.Controls.Remove labSizer.name
    Set labSizer = Nothing

End Sub

Private Property Get m_GetSizer(Base As MSForms.UserForm) As MSForms.Label
    
    Set m_GetSizer = Base.Controls.Add("Forms.Label.1", "labSizer", True)

End Property


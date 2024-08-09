Attribute VB_Name = "Module1"
Option Explicit

Sub RemoveExternalLinks()
    Dim Links As Variant
    Dim i As Long
    
    ' Get the array of links in the workbook
    Links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    
    ' Check if there are any links
    If Not IsEmpty(Links) Then
        ' Loop through each link
        For i = LBound(Links) To UBound(Links)
            ' Break each link
            ActiveWorkbook.BreakLink name:=Links(i), Type:=xlLinkTypeExcelLinks
        Next i
        MsgBox "All external links have been removed.", vbInformation
    Else
        MsgBox "No external links found.", vbInformation
    End If
End Sub


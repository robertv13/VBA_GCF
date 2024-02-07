Attribute VB_Name = "Utils"
Option Explicit

Sub ExportNamedRanges()
    Dim wsDestination As Worksheet
    Dim namedRange As Name
    Dim destRow As Long
    
    ' Create a new worksheet for exporting named ranges
    Set wsDestination = ThisWorkbook.Sheets.Add
    wsDestination.Name = "NamedRanges"
    
    ' Write headers
    wsDestination.Range("A1").Value = "Name"
    wsDestination.Range("B1").Value = "Refers To"
    wsDestination.Range("C1").Value = "Value"
    
    destRow = 2 ' Start writing from the second row
    
    ' Loop through each named range in the workbook
    For Each namedRange In ThisWorkbook.Names
        ' Write named range name to destination worksheet
        wsDestination.Cells(destRow, 1).Value = namedRange.Name
        ' Write refers to address to destination worksheet
        wsDestination.Cells(destRow, 2).Value = namedRange.RefersTo
        ' Write named range value to destination worksheet
        If Not namedRange.RefersToRange Is Nothing Then
            wsDestination.Cells(destRow, 3).Value = namedRange.RefersToRange.Value
        End If
        destRow = destRow + 1 ' Move to the next row
    Next namedRange
    
    ' Autofit columns for better visibility
    wsDestination.Columns("A:C").AutoFit
    
    MsgBox "Named ranges exported successfully.", vbInformation
End Sub


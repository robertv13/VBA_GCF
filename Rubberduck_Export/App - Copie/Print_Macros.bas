Attribute VB_Name = "Print_Macros"
Option Explicit
Dim LastPage As Long, PageNumb As Long

Sub Invoice_Print()
    With Invoice
        Invoice_SaveUpdate                       'Save Current Invoice
        LastPage = .Range("B12").Value           'Set Last Page
        For PageNumb = 1 To LastPage
            .Range("B11").Value = PageNumb
            Invoice_PageLoad                     'Load Page
            .PrintOut , , , False, True, , , , False
        Next PageNumb
    End With
End Sub

Sub Invoice_SaveAsPDF()
    MsgBox "Please join our Patreon Program to see an updated training and this Save As PDF Feature" & vbCrLf & "https://www.patreon.com/ExcelForFreelancers"
End Sub


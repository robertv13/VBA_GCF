VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCustForm 
   Caption         =   "Add Client"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "AddCustForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCustForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBtn_Click()
    AddCustForm.Hide
End Sub

Private Sub SaveBtn_Click()
    Dim CustomerFld As Control
    Dim CustRow As Long, CustCol As Long
    If Me.Field1.Value = Empty Then
        MsgBox "Please make sure to save a Customer Name before saving"
        Exit Sub
    End If
    With Customers
        If Invoice.Range("B3").Value = Empty Then 'New Customer
            CustRow = .Range("A99999").End(xlUp).Row + 1 'First avail Row
            On Error Resume Next
            .Range("A" & CustRow).Value = Application.WorksheetFunction.Max(.Range("Cust_ID")) + 1
            On Error GoTo 0
            If .Range("A" & CustRow).Value = Empty Then .Range("A" & CustRow).Value = 1 'Set First Cust ID to 1
        Else
            CustRow = Invoice.Range("B3").Value
        End If
        For CustCol = 2 To 8
            Set CustomerFld = Me.Controls("Field" & CustCol - 1)
            .Cells(CustRow, CustCol).Value = CustomerFld.Value
        Next CustCol
        AddCustForm.Hide
        Invoice.Range("E5").Value = Field1.Value 'Set Customer Name
    End With
End Sub


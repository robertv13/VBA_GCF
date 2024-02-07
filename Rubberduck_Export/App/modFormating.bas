Attribute VB_Name = "modFormating"
Option Explicit

Public Sub MyAutoFill()
    
    'Declare range Variables
    Dim selection1 As Range
    Dim selection2 As Range
    
    'Set range variables = their respective ranges
    Set selection1 = Sheet1.Range("A1:A2")
    Set selection2 = Sheet1.Range("A1:A20")
    
    'Autofil
    selection1.AutoFill Destination:=selection2

End Sub


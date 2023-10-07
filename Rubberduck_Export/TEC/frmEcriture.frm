VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEcriture 
   Caption         =   "Écriture"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   OleObjectBlob   =   "frmEcriture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEcriture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'PLACE IN YOUR USERFORM CODE

Private MyListBoxClass As clsCListboxAlign

Private Sub UserForm_Initialize()
    Dim lngRow As Long
    Dim lngIndex As Long
    Set MyListBoxClass = New clsCListboxAlign 'declare the class
    
    'This is just a sample where I add data to a listbox.
    'You'll want to use your own data.
    '-----------------------------------------------------------------------
    
    With Me.lstEcriture
        .ColumnCount = 4
        .ColumnWidths = "190; 65; 65; 160"

        .AddItem
        .AddItem
        .AddItem
        .AddItem
        
        .List(0, 0) = "1020 - Chèques CIBC"
        .List(1, 0) = "3900 - Apport des actionnaires"
        .List(2, 0) = "Third Product"
        .List(3, 0) = "Fourth Product"
        
        .List(0, 1) = "0,99"
        .List(1, 1) = "32,99"
        .List(2, 1) = "332,99"
        .List(3, 1) = "3 332,99"
        
        .List(0, 2) = "33 999,99"
        .List(1, 2) = "333 079,00"
        .List(2, 2) = "3 333 000,00"
        .List(3, 2) = "99 888 777,00"
        
        .List(0, 3) = "Ligne - 1"
        .List(1, 3) = "Ligne - 2"
        .List(2, 3) = "Ligne - 3"
        .List(3, 3) = "Ligne - 4"
        
    End With
    '-----------------------------------------------------------------------

    'This is how you left, center and right align a ListBox.
    MyListBoxClass.Left Me.lstEcriture, 1
    MyListBoxClass.Right Me.lstEcriture, 2
    MyListBoxClass.Right Me.lstEcriture, 3
    MyListBoxClass.Left Me.lstEcriture, 4

End Sub

Private Sub UserForm_Terminate()
    
    'Clear the class declaration
    Set MyListBoxClass = Nothing

End Sub


'Private Sub UserForm_Initialize()
'
'    'Empty all txtBox
'    txtSource = ""
'    txtDate = ""
'    txtDescription = ""
'    txtTotalDebit = ""
'    txtTotalCredit = ""
'
'    'Empty Ecriture List
'    lstEcriture.Clear
'
'    'Fill Ecriture List
'    With lstEcriture
'        .ColumnHeads = False
'        .ColumnCount = 4
'        .ColumnWidths = "180; 70; 70; 150"
'    End With
'
'    lstEcriture.AddItem
'    lstEcriture.List(0, 0) = "1050 - Chèques CIBC"
'    lstEcriture.List(0, 1) = "Débit"
'    lstEcriture.List(0, 2) = "Crédit"
'    lstEcriture.List(0, 3) = "Remarque"
'
'    'Set Focus on txtSource
'    txtSource.SetFocus
'
'End Sub

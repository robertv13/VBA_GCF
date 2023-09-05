VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormBook 
   Caption         =   "Select book"
   ClientHeight    =   5532
   ClientLeft      =   132
   ClientTop       =   648
   ClientWidth     =   11136
   OleObjectBlob   =   "UserFormCompanyBase_20230303_0029.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare the class used to create the searchable dropdown
Private oEventHandler As clsSearchableDropdown
Private m_Cancelled As Boolean

Public Property Get Book() As String
    Book = oEventHandler.SelectedItem
End Property
Public Property Get Cancelled() As String
    Cancelled = m_Cancelled
End Property

Public Property Let ListData(ByVal rg As Range)
    oEventHandler.List = rg.value
End Property

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub UserForm_Initialize()

    Call InitializeSettings

    ' Create the object for the searchable dropdown
    Set oEventHandler = New clsSearchableDropdown
    
    With oEventHandler
    
        ' assign the listbox and textbox to the searchable dropdown class
        Set .SearchListBox = Me.ListBox1
        Set .SearchTextBox = Me.TextBox1
    
        ' Set the text box size to the size selected in the font combobox
        TextBox1.Font.Size = ComboBoxFont.value
        
        ' Set the maximum number of rows to the value selected in the rows combobox
        .MaxRows = ComboBoxRows.value
        
        ' Set show all matches to the values in the show all matches checkbox
        .ShowAllMatches = CheckBoxShowMatches.value
            
    End With

End Sub

Private Sub UserForm_Terminate()
    Set oEventHandler = Nothing
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
' Handle the user clicking on the X to cancel the Userform
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    hide
    m_Cancelled = True
    
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub buttonOk_Click()
    Me.hide
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub CheckBoxShowMatches_Click()
    ' set show all matches on
    If Not oEventHandler Is Nothing Then
        oEventHandler.ShowAllMatches = CheckBoxShowMatches.value
    End If
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub ComboBoxFont_Change()
    ' Set the textbox to the size in the font combobox
    With TextBox1
        .Font.Size = ComboBoxFont.value
        .Height = .Font.Size * 2
    End With
    ' Refilter the listview to set the new font size
    If Not oEventHandler Is Nothing Then
        oEventHandler.FilterListBox
    End If
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub ComboBoxRows_Change()
    If Not oEventHandler Is Nothing Then
        ' Set the max rows
        oEventHandler.MaxRows = ComboBoxRows.value
    End If
End Sub


' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub ComboBoxCase_Change()
    If Not oEventHandler Is Nothing Then
        ' Set the max rows
        oEventHandler.CompareMethod = IIf(ComboBoxCase.value = "Not sensitive", vbTextCompare, vbBinaryCompare)
    End If
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub InitializeSettings()

    m_Cancelled = False

    ' Fill the comboxbox boxes with values
    ComboBoxFont.List = Array(8, 9, 10, 11, 12, 14)
    ComboBoxFont.ListIndex = 4
    ComboBoxRows.List = Array(6, 7, 8, 9, 10, 11, 12)
    ComboBoxRows.ListIndex = 2
    ComboBoxCase.List = Array("Not sensitive", "Sensitive")
    ComboBoxCase.ListIndex = 0

End Sub





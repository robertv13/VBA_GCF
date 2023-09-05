VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormCompany 
   Caption         =   "Recherche de compagnie"
   ClientHeight    =   5316
   ClientLeft      =   132
   ClientTop       =   648
   ClientWidth     =   8880.001
   OleObjectBlob   =   "UserFormCompany_20230303_0018.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: Paul Kelly - //ExcelMacroMastery.com/

'Declare the class used to create the searchable dropdown
Private oEventHandler As clsSearchableDropdown

Private m_Cancelled As Boolean

Public Property Get Company() As String
    
    Company = oEventHandler.SelectedItem

End Property

Public Property Get Cancelled() As String
    
    Cancelled = m_Cancelled

End Property

Public Property Let ListData(ByVal rg As Range)
    
    oEventHandler.List = rg.value

End Property

Private Sub UserForm_Initialize()

    Call InitializeSettings

    'Create the object for the searchable dropdown
    Set oEventHandler = New clsSearchableDropdown
    
    With oEventHandler
    
        'assign the listbox and textbox to the searchable dropdown class
        Set .SearchListBox = Me.ListBox1
        Set .SearchTextBox = Me.TextBox1
    
        'Set the text box size to the size selected in the font combobox
        TextBox1.Font.Size = ComboBoxFont.value
        
        'Set the maximum number of rows to the value selected in the rows combobox
        .MaxRows = ComboBoxRows.value
        
        'Set show all matches to the values in the show all matches checkbox
        .ShowAllMatches = False
            
    End With

End Sub

Private Sub UserForm_Terminate()
    
    Set oEventHandler = Nothing

End Sub

'Handle the user clicking on the X to cancel the Userform
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    'Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    'Hide the Userform and set cancelled to true
    hide
    m_Cancelled = True
    
End Sub

Private Sub buttonOk_Click()
    
    Me.hide

End Sub

Private Sub ComboBoxFont_Change()
    
    'Set the textbox to the size in the font combobox
    With TextBox1
        .Font.Size = ComboBoxFont.value
        .Height = .Font.Size * 2
    End With
    'Refilter the listview to set the new font size
    If Not oEventHandler Is Nothing Then
        oEventHandler.FilterListBox
    End If

End Sub

Private Sub ComboBoxRows_Change()
    
    If Not oEventHandler Is Nothing Then
        'Set the max rows
        oEventHandler.MaxRows = ComboBoxRows.value
    End If

End Sub

Private Sub ComboBoxCase_Change()
    
    If Not oEventHandler Is Nothing Then
        'Set the max rows
        oEventHandler.CompareMethod = IIf(ComboBoxCase.value = "Toutes occurences", vbTextCompare, vbBinaryCompare)
    End If

End Sub

Private Sub InitializeSettings()

    m_Cancelled = False

    'Fill the comboxbox boxes with values
    ComboBoxFont.List = Array(9, 10, 11, 12)
    ComboBoxFont.ListIndex = 3
    
    ComboBoxRows.List = Array(5, 6, 7, 8, 9, 10)
    ComboBoxRows.ListIndex = 5
    
    ComboBoxCase.List = Array("Toutes occurences", "Selon la casse")
    ComboBoxCase.ListIndex = 1

End Sub

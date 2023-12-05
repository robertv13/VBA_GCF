Attribute VB_Name = "modEvents"
Option Explicit

Private Sub Workbook_Open()
    Dim MyMenu As Object
    
    Set MyMenu = Application.ShortcutMenus(xlWorksheetCell) _
                 .MenuItems.AddMenu("This is my Custom Menu", 1)
    
    With MyMenu.MenuItems
        .Add "MyMacro1", "MyMacro1", , 1, , ""
        .Add "MyMacro2", "MyMacro2", , 2, , ""
    End With
    
    Set MyMenu = Nothing
    
End Sub

'Private Sub auto_open()
'    'Code thute when Excel is launched
'    MsgBox "This code ran at Excel start!"
'End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'Catch changes to cell A1
    If Target.Address = "$A$1" Then
        MsgBox "This Code Runs When Cell A1 Changes!"
    End If
End Sub

'Private Sub auto_close()
'    MsgBox "This code ran at Excel close!"
'End Sub


'Private Sub Workbook_Open()
'    'Open a specific tabname
'    Sheets("mytabname").Activate
'End Sub

'Private Sub Workbook_Open()
'    'Shows a UserForm as the workbook is open
'    UserForm1.Show
'End Sub









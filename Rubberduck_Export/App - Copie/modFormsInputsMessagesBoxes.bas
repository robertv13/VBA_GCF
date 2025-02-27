Attribute VB_Name = "modFormsInputsMessagesBoxes"
Option Explicit

Public Sub MyInputBox()

    Dim MyInput As String
    MyInput = InputBox("This is my InputBox", "MyInputTitle", "Enter your input text HERE")

    If MyInput = "Enter your input text HERE" Or MyInput = "" Then
        Exit Sub
    Else
        MsgBox "The text from MyInputBox is " & MyInput
    End If

End Sub

'MsgBox "line 1" & vbCrLf & "line 2"

Private Sub Workbook_Open()
    
    UserForm1.Show

End Sub

Sub YesNoMessageBox()
    
    Dim Answer As String
    Dim MyNote As String

    'Place your text here
    MyNote = "Do you agree?"

    'Display MessageBox
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "???")

    If Answer = vbNo Then
        'Code for No button Press
        MsgBox "You pressed NO!"
    Else
        'Code for Yes button Press
        MsgBox "You pressed Yes!"
    End If

End Sub


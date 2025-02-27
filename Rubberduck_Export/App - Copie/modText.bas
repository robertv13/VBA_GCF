Attribute VB_Name = "modText"
Option Explicit

'MsgBox Right(MyString, 9)

'MsgBox Right(MyString, Len(MyString) - 1)

Sub AddSpaces()
    Dim MyString As String

    MyString = "Hello" & Space(10) & "World"

    MsgBox MyString

End Sub

Function Number_of_Words(Text_String As String) As Integer

    'Function counts the number of words in a string
    'by looking at each character and seeing whether it is a space or not
    Number_of_Words = 0
    Dim String_Length As Integer
    Dim Current_Character As Integer
    String_Length = Len(Text_String)

    For Current_Character = 1 To String_Length

        If (Mid(Text_String, Current_Character, 1)) = " " Then
            Number_of_Words = Number_of_Words + 1
        End If

    Next Current_Character

End Function

Function Acroymn(Original_String As String) As String
    Dim Trimmed_String As String
    Dim Length As Integer
    Dim Pos As Integer

    Trimmed_String = Application.WorksheetFunction.Trim(Original_String)

    'work out the length of the string
    Length = Len(Trimmed_String)

    Acroymn = UCase(Left(Trimmed_String, 1))

    For Pos = 2 To Length - 1
        If (Mid(Trimmed_String, Pos, 1) = " ") Then
            Acroymn = Acroymn & UCase(Mid(Trimmed_String, Pos + 1, 1))
        End If
    Next Pos

End Function

Function Extract_Number_from_Text(Phrase As String) As Double

    Dim Length_of_String As Integer
    Dim Current_Pos As Integer
    Dim Temp As String

    Length_of_String = Len(Phrase)
    Temp = ""

    For Current_Pos = 1 To Length_of_String

        If (Mid(Phrase, Current_Pos, 1) = "-") Then
            Temp = Temp & Mid(Phrase, Current_Pos, 1)
        End If

        If (Mid(Phrase, Current_Pos, 1) = ".") Then
            Temp = Temp & Mid(Phrase, Current_Pos, 1)
        End If

        If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then
            Temp = Temp & Mid(Phrase, Current_Pos, 1)
        End If

    Next Current_Pos

    If Len(Temp) = 0 Then
        Extract_Number_from_Text = 0
    Else
        Extract_Number_from_Text = CDbl(Temp)
    End If

End Function

Public Sub FindSomeText()

    If InStr("Look in this string", "look") = 0 Then
        MsgBox "woops, no match"
    Else
        MsgBox "at least one match"
    End If

End Sub

Private Sub Worksheet_Change(ByVal Target As Excel.Range)

    Application.EnableEvents = False

    If Target.Column = 5 Then
        Target = StrConv(Target, vbProperCase)
    End If

    Application.EnableEvents = True
End Sub

Sub LoopThroughString()

    Dim Counter As Integer
    Dim MyString As String
    MyString = "AutomateExcel"                   'define string

    For Counter = 1 To Len(MyString)
        'do something to each character in string
        'here we'll msgbox each character
        MsgBox Mid(MyString, Counter, 1)
    Next

End Sub

Function Find_nth_word(Phrase As String, n As Integer) As String

    Dim Current_Pos As Long
    Dim Length_of_String As Integer
    Dim Current_Word_No As Integer

    Find_nth_word = ""
    Current_Word_No = 1

    'Remove Leading Spaces
    Phrase = Trim(Phrase)

    Length_of_String = Len(Phrase)

    For Current_Pos = 1 To Length_of_String
        If (Current_Word_No = n) Then
            Find_nth_word = Find_nth_word & Mid(Phrase, Current_Pos, 1)
        End If
        If (Mid(Phrase, Current_Pos, 1) = " ") Then
            Current_Word_No = Current_Word_No + 1
        End If
    Next Current_Pos

    'Remove the rightmost space
    Find_nth_word = Trim(Find_nth_word)

End Function

Option Explicit
Private Sub CommandButton1_Click()

    'Define Variables

    Dim Original_String As String
    Dim Reversed_String As String
    Dim Next_Char As String

    Dim Length As Integer
    Dim Pos As Integer

    'Get the Original String

    Original_String = InputBox("Pls enter the original string: ")

    'Find the revised length of the string

    Length = Len(Original_String)

    'Set up the reversed string
    Reversed_String = ""

    'Progress through the string on a character by character basis
    'Starting at the last character and going towards the first character

    For Pos = Length To 1 Step -1

        Next_Char = Mid(Original_String, Pos, 1)
        Reversed_String = Reversed_String & Next_Char
    Next Pos

    MsgBox "The reversed string is " & Reversed_String

End Sub

Sub SayThisString()

    Dim SayThis As String

    SayThis = "I love Microsoft Excel"
    Application.Speech.Speak (SayThis)

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Sheet1.Cells.CheckSpelling

End Sub


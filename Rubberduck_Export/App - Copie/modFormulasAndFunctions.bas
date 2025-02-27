Attribute VB_Name = "modFormulasAndFunctions"
Option Explicit

'Notes, Subs & Functions - 2023-12-04 from Automate EXCEL

'Application.Calculation = xlAutomatic
'Application.Calculation = xlManual
'MsgBox Right(MyString, 9)
'MsgBox Right(MyString, Len(MyString) - 1)

Sub AddSpaces()
    
    Dim MyString As String
    MyString = "Hello" & Space(10) & "World"
    MsgBox MyString

End Sub

Function Compare_Dates(sDate As Date, eDate As Date, oDate As Date) As Boolean
    'Boolean Function to compare dates
    'Will return TRUE only when Other_Date is between Start_Date and End_Date
    'Otherwise will return FALSE
    'Set outcome to FALSE - default value
    Compare_Dates = False

    'Compare Dates
    If ((oDate >= sDate) And (oDate <= eDate)) Then
        Compare_Dates = True
    End If

End Function

Function Number_of_Words(Text_String As String) As Integer
    'Function counts the number of words in a string by looking at each character
    'and seeing whether it is a space or not
    Dim Number_of_Words As Integer
    Dim String_Length As Integer
    Dim Current_Character As Integer
    
    Number_of_Words = 0
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

    'Work out the length of the string
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

    'Looking for numeric digit, -, . or ,
    For Current_Pos = 1 To Length_of_String
        If (Mid(Phrase, Current_Pos, 1) = "-") Then
            Temp = Temp & Mid(Phrase, Current_Pos, 1)
        End If
        If (Mid(Phrase, Current_Pos, 1) = ".") Or _
           (Mid(Phrase, Current_Pos, 1) = ",") Then
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

Function Max_Each_Column(r As Range) As Variant
    Dim TempArray() As Double, i As Long

    If r Is Nothing Then Exit Function

    With r
        ReDim TempArray(1 To .Columns.Count)
        For i = 1 To .Columns.Count
            TempArray(i) = Application.Max(.Columns(i))
        Next
    End With

    Max_Each_Column = TempArray

End Function

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Application.EnableEvents = False

    If Target.Column = 5 Then
        Target = StrConv(Target, vbProperCase)
    End If

    Application.EnableEvents = True
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

Private Sub ReverseString()
    Dim Original_String As String
    Dim Reversed_String As String
    Dim Next_Char As String
    Dim Length As Integer
    Dim Pos As Integer

    'Get the Original String
    Original_String = InputBox("Pls enter the original string: ")

    'Find the length of the string to be reversed
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

Function Show_Cell_Formulae(c As Range) As String
    
    Show_Cell_Formulae = "Cell " & c.Address & " has the formulae: " & c.Formula & " '"

End Function

Function Color_By_Numbers(Color_Range As Range, Color_Index As Integer) As Double
    Dim Color_By_Numbers As Double
    Dim Cell As String
    
    'Will look at cells that are in the range and if the color interior property
    'matches the cell color required then it will sum

    'Loop Through range
    For Each Cell In Color_Range
        If (Cell.Interior.ColorIndex = Color_Index) Then
            Color_By_Numbers = Color_By_Numbers + Cell.Value
        End If
    Next Cell
End Function

Sub UseFunction()

    MsgBox Application.WorksheetFunction.Combin(42, 6)

End Sub

Function ThreeParameterVlookup(r As Range, Col As Integer, p1 As Variant, p2 As Variant, p3 As Variant) As Variant

    Dim Current_Row As Integer
    Dim No_Of_Rows_in_Range As Integer
    Dim No_Of_Cols_in_Range As Integer
    Dim Matching_Row As Integer

    'Set answer to N/A by default
    ThreeParameterVlookup = CVErr(xlErrNA)
    Matching_Row = 0
    Current_Row = 1

    No_Of_Rows_in_Range = r.Rows.Count
    No_Of_Cols_in_Range = r.Columns.Count

    'Check if Col is greater than number of columns in range
    If (Col > No_Of_Cols_in_Range) Then
        ThreeParameterVlookup = CVErr(xlErrRef)
    End If
    If (Col <= No_Of_Cols_in_Range) Then
        Do
            If ((r.Cells(Current_Row, 1).Value = p1) And _
                (r.Cells(Current_Row, 2).Value = p2) And _
                (r.Cells(Current_Row, 3).Value = p3)) Then
                Matching_Row = Current_Row
            End If
            Current_Row = Current_Row + 1
        Loop Until ((Current_Row = No_Of_Rows_in_Range) Or (Matching_Row <> 0))

        If Matching_Row <> 0 Then
            ThreeParameterVlookup = r.Cells(Matching_Row, Col)
        End If
    End If

End Function



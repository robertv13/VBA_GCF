Attribute VB_Name = "modUtils"
Option Explicit

Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub Log_Record(ByVal procedureName As String, Optional ByVal startTime As Double = 0) '2024-08-12 @ 12:12

    Dim logFile As String, rootPath As String
    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
        rootPath = "P:\Administration\APP\GCF"
    Else
        rootPath = "C:\VBA\GC_FISCALITÉ"
    End If

    logFile = rootPath & "\Log.txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Dim currentTime As String
    currentTime = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    
    Open logFile For Append As #fileNum
    
    If startTime = 0 Then
        startTime = Timer 'Start timing
        Print #fileNum, Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        Replace(currentTime, " ", "_") & "|" & _
                        ThisWorkbook.Name & "|" & _
                        procedureName & " (entrée)"
        Close #fileNum
    Else
        Dim elapsedTime As Double
        elapsedTime = Round(Timer - startTime, 4) 'Calculate elapsed time
        Print #fileNum, Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        Replace(currentTime, " ", "_") & "|" & _
                        ThisWorkbook.Name & "|" & _
                        procedureName & " (sortie)" & "|" & _
                        "Temps écoulé: " & Format(elapsedTime, "0.0000") & " seconds"
        Close #fileNum
    End If
End Sub

Sub Test_Log_Record()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modzDevUtils:Test_Log_Record", 0)

    Call Log_Record("modzDevUtils:Test_Log_Record", startTime)
    
End Sub

Function Fn_Get_Windows_Username() As String 'Function to retrieve the Windows username using the API

    Dim buffer As String * 255
    Dim size As Long: size = 255
    
    If GetUserName(buffer, size) Then
        Fn_Get_Windows_Username = Left$(buffer, size - 1)
    Else
        Fn_Get_Windows_Username = "Unknown"
    End If
    
End Function




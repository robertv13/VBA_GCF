Attribute VB_Name = "modUtils"
Option Explicit

Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub Log_Record(moduleProcName As String, param1 As String, Optional ByVal startTime As Double = 0) '2024-08-22 @ 05:48

    Dim currentTime As String
    currentTime = Format$(Now, "yyyymmdd_hhnnss")
    
    'Determine the location of the Log file
    Dim rootPath As String
    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
        rootPath = "P:\Administration\APP\GCF"
    Else
        rootPath = "C:\VBA\GC_FISCALITÉ"
    End If

    Dim logFile As String
    logFile = rootPath & "\LogClientsApp.txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    
    Dim moduleName As String, procName As String
    If InStr(moduleProcName, ":") Then
        moduleName = Left(moduleProcName, InStr(moduleProcName, ":") - 1)
        procName = Right(moduleProcName, Len(moduleProcName) - InStr(moduleProcName, ":"))
    Else
        moduleName = moduleProcName
        procName = ""
    End If
    
    If startTime = 0 Then
        startTime = Timer 'Start timing
        Print #fileNum, Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        currentTime & "|" & _
                        ThisWorkbook.Name & "|" & _
                        moduleName & "|" & _
                        procName & "|" & _
                        "" & "|" & _
                        param1
    ElseIf startTime <= 0 Then 'Log intermédiaire
        Print #fileNum, Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        currentTime & "|" & _
                        ThisWorkbook.Name & "|" & _
                        moduleName & "|" & _
                        procName & "|" & _
                        "" & "|" & _
                        param1
    Else
        Dim elapsedTime As Double
        elapsedTime = Round(Timer - startTime, 4) 'Calculate elapsed time
        Print #fileNum, Replace(Fn_Get_Windows_Username, " ", "_") & "|" & _
                        currentTime & "|" & _
                        ThisWorkbook.Name & "|" & _
                        moduleName & "|" & _
                        procName & " (sortie)" & "|" & _
                        "Temps écoulé: " & Format(elapsedTime, "#0.0000") & " seconds" & "|" & _
                        param1
    End If
    
    Close #fileNum

End Sub

Sub Test_Log_Record()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modzDevUtils:Test_Log_Record", "Test", 0)

    Call Log_Record("modzDevUtils:Test_Log_Record", "Test sortie", startTime)
    
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




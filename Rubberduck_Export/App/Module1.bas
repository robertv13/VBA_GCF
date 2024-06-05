Attribute VB_Name = "Module1"
Option Explicit

Public Function GetOneDrivePath(ByVal fullWorkbookName As String) As String '2024-05-27 @ 10:10
    
    'Try the 3 key types in the registry to find the file
    Dim oneDrive As Variant
    oneDrive = Array("OneDriveCommercial", "OneDriveConsumer", "OneDrive")
    
    Dim ShellScript As Object
    Set ShellScript = CreateObject("WScript.Shell")
    Dim oneDriveRegLocalPath As String
    
    Dim key As Variant
    For Each key In oneDrive
    
        'Get the Get OneDrive path from the registry - If doesn't exist go to the next key
        On Error Resume Next
        oneDriveRegLocalPath = ShellScript.RegRead("HKEY_CURRENT_USER\Environment\" & key)
        If oneDriveRegLocalPath = vbNullString Then GoTo continue
        On Error GoTo 0
                    
        'Get the end part of the path from the URL name
        Dim fileEndPart As String
        fileEndPart = GetEndPath(fullWorkbookName)
        If Len(fileEndPart) = 0 Then GoTo continue
        
        'Build the final filename by combining registry drive and the end part of url
        GetOneDrivePath = Replace(oneDriveRegLocalPath & fileEndPart, "/", "\")
        
        'Check if the file exists
        If Dir(GetOneDrivePath) = "" Then
            GetOneDrivePath = ""
        Else
            Exit For
        End If
continue:
    Next key
    
    If GetOneDrivePath = "" Then Err.Raise 53, "GetOneDrivePath" _
                , "Could not find the file [" & fullWorkbookName & "] on OneDrive."

End Function

Public Function GetEndPath(ByVal fullWorkbookName As String) As String

    'Remove the url part of the name which is preceded by the text "/Documents"
    If InStr(1, fullWorkbookName, "my.sharepoint.com") <> 0 Then
        'Get the part of the string after "/Documents"
        Dim arr() As String
        arr = Split(fullWorkbookName, "/Documents")
        GetEndPath = arr(UBound(arr))
    ElseIf InStr(1, fullWorkbookName, "d.docs.live.net") <> 0 Then
        'Get the part of the filename without the URL
        Dim firstPart As String
        firstPart = Split(fullWorkbookName, "/")(4)
        GetEndPath = Mid(fullWorkbookName, InStr(fullWorkbookName, firstPart) - 1)
    Else
        GetEndPath = ""
    End If
    
End Function



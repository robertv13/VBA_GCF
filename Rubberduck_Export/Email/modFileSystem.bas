Attribute VB_Name = "modFileSystem"
Option Explicit

Sub GetFileInfo()

    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile("C:\VBA\Git and Github for beginners Tutorial.docx")
    s = "'" & f.name & "' on drive " & UCase(f.Drive) & vbCrLf
    s = s & "Créé le : " & f.DateCreated & vbCrLf
    s = s & "Dernier accès : " & f.DateLastAccessed & vbCrLf
    s = s & "Dernière modification: " & f.DateLastModified
    MsgBox s, 0, "File Access Info"

End Sub



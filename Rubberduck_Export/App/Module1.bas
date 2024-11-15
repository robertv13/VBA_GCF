Attribute VB_Name = "Module1"
Option Explicit

Sub test_Input_Box()

'    Dim Name As String
'    Name = InputBox("What is your name ? ", "Name entry")
    
    Dim output As Variant
    output = Application.InputBox("Saisi d'une chaine de caractères", "Exemple de saisi", "String")
    
'    Dim r As String
'    r = InputBox("Prompt", "Title", "Default")
    
End Sub

Sub Debug_TEC()

    Dim wsTEC As Worksheet: Set wsTEC = wshTEC_Local
    Dim lurTEC As Long
    lurTEC = wsTEC.Cells(wsTEC.rows.count, "A").End(xlUp).row
    
    Dim wsTDB As Worksheet: Set wsTDB = wshTEC_TDB_Data
    Dim lurTDB As Long
    lurTDB = wsTDB.Cells(wsTDB.rows.count, "A").End(xlUp).row
    
    Dim arr() As Variant
    ReDim arr(1 To 2500, 1 To 3)
    
    Dim i As Long
    Dim TECID As Long
    Dim dateCutoff As Date
    dateCutoff = #11/13/2024#
    
    Dim h As Double, hTEC As Double
    'Boucle dans TEC_Local
    Debug.Print "Mise en mémoire TEC_LOCAL"
    For i = 3 To lurTEC
        With wsTEC
            If .Range("D" & i).value > dateCutoff Then Stop
            TECID = CLng(.Range("A" & i).value)
            If arr(TECID, 1) <> "" Then Stop
            arr(TECID, 1) = TECID
            h = .Range("H" & i).value
            If UCase(.Range("N" & i).value) = "VRAI" Then
                h = 0
            End If
            If h <> 0 Then
                If UCase(.Range("J" & i).value) = "VRAI" And Len(.Range("E" & i).value) > 2 Then
                    If UCase(.Range("L" & i).value) = "FAUX" Then
                        If .Range("M" & i).value <= dateCutoff Then
                            arr(TECID, 2) = h
                        Else
                            Stop
                        End If
                    End If
                End If
            End If
        End With
    Next i
    
    'Boucle dans TEC_TDB
    Dim hTDB As Double
    Debug.Print "Mise en mémoire TEC_TDB"
    For i = 2 To lurTDB
        With wsTDB
            If .Range("D" & i).value > dateCutoff Then Stop
            TECID = CLng(.Range("A" & i).value)
            arr(TECID, 1) = TECID
            arr(TECID, 3) = .Range("Q" & i).value
        End With
    Next i
    
    Debug.Print "Analyse des écarts"
    Dim tTEC As Double, tTDB As Double
    For i = 1 To 2500
        tTEC = tTEC + arr(i, 2)
        tTDB = tTDB + arr(i, 3)
        If arr(i, 2) <> arr(i, 3) Then
            Debug.Print arr(i, 1), arr(i, 2), arr(i, 3)
        End If
        wshzTEC_Debug.Range("C" & i).value = arr(i, 2)
        wshzTEC_Debug.Range("D" & i).value = arr(i, 3)
    Next i
    
    Debug.Print "Totaux", Round(tTEC, 2), Round(tTDB, 2)
    
End Sub


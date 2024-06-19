Attribute VB_Name = "modAppliUtils"
Option Explicit

Public Sub ConvertRangeBooleanToText(rng As Range)

    Dim cell As Range
    For Each cell In rng
        Select Case cell.value
            Case 0, "False" 'False
                cell.value = "FAUX"
            Case -1, "True" 'True
                cell.value = "VRAI"
            Case "VRAI", "FAUX"
                
            Case Else
                MsgBox cell.value & " est une valeur INVALIDE pour la cellule " & cell.Address & " de la feuille TEC_Local"
        End Select
    Next cell

End Sub

Public Sub ProtectCells(rng As Range)

    'Lock the checkbox
    rng.Locked = True
    
    'Protect the worksheet
    rng.Parent.Protect UserInterfaceOnly:=True


End Sub

Public Sub UnprotectCells(rng As Range)

    'Lock the checkbox
    rng.Locked = False
    
    'Protect the worksheet
    rng.Parent.Protect UserInterfaceOnly:=True


End Sub

Sub Start_Routine(subName As String) '2024-06-06 @ 10:12

    Dim modeOper As Integer
    modeOper = 2
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Integer: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & "Entering: " & subName & _
            vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A99999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName & " - entering"
        End With
    End If

End Sub

Sub Output_Timer_Results(subName As String, t As Double)

    Dim modeOper As Integer
    modeOper = 2 '2024-03-29 @ 11:37
    
    'Allows message to be used - 2024-06-06 @ 11:05
    If InStr(subName, "message:") = 1 Then
        subName = Right(subName, Len(subName) - 8)
    Else
        subName = subName & " - exiting"
    End If
    
    'modeOper = 1 - Dump to immediate Window
    If modeOper = 1 Then
        Dim l As Integer: l = Len(subName)
        Debug.Print vbNewLine & String(40 + l, "*") & vbNewLine & _
        Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & subName & " = " _
        & Format(Timer - t, "##0.0000") & " secondes" & vbNewLine & String(40 + l, "*")
    End If

    'modeOper = 2 - Dump to worksheet
    If modeOper = 2 Then
        With wshzDocLogAppli
            Dim lastUsedRow As Long
            lastUsedRow = .Range("A9999").End(xlUp).row
            lastUsedRow = lastUsedRow + 1 'Row to write a new record
            .Range("A" & lastUsedRow).value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            .Range("B" & lastUsedRow).value = subName
            If t Then
                .Range("C" & lastUsedRow).value = Format(Round(Timer - t, 4), "##0.0000")
            End If
        End With
    End If

End Sub

Public Sub ArrayToRange(ByRef data As Variant _
                        , ByVal outRange As Range _
                        , Optional ByVal clearExistingData As Boolean = True _
                        , Optional ByVal clearExistingHeaderSize As Long = 1)
                        
    If clearExistingData = True Then
        outRange.CurrentRegion.Offset(clearExistingHeaderSize).Clearcontents
    End If
    
    Dim rows As Long, columns As Long
    rows = UBound(data, 1) - LBound(data, 1) + 1
    columns = UBound(data, 2) - LBound(data, 2) + 1
    outRange.Resize(rows, columns).value = data
    
End Sub

Sub CreateOrReplaceWorksheet(wsName As String)
    
    Dim ws As Worksheet
    Dim wsExists As Boolean
    wsExists = False
    
    'Check if the worksheet exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = wsName Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    ' If the worksheet exists, delete it
    If wsExists Then
        Application.DisplayAlerts = False
        ws.delete
        Application.DisplayAlerts = True
    End If
    
    'Add the new worksheet
    Set ws = ThisWorkbook.Worksheets.add
    ws.name = wsName

End Sub

Private Sub IntegrityVerification()

    'wshBD_Clients
    Call check_Clients

    'wshGL_Trans
    Call check_GL_Trans
    
    'wshTEC_Local
    Call check_TEC
    
End Sub

Private Sub check_Clients()

    'wshBD_Clients
    Dim ws As Worksheet
    Set ws = wshBD_Clients
    Debug.Print ws.name & " (wshBD_Clients)"
    
    Dim arr As Variant
    arr = wshBD_Clients.Range("A1").CurrentRegion.value
    Dim dict_code_client As New Dictionary
    Dim dict_nom_client As New Dictionary
    
    Dim i As Long, code As String, nom As String
    Dim cas_doublon_nom As Long
    Dim cas_doublon_code As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        nom = arr(i, 1)
        code = arr(i, 2)
        If dict_nom_client.Exists(nom) = False Then
            dict_nom_client.add nom, code
        Else
            Debug.Print Tab(5); "Le nom '" & nom & "' est un doublon"; Tab(75); "pour le code '" & code & "'"
            cas_doublon_nom = cas_doublon_nom + 1
        End If
        If dict_code_client.Exists(code) = False Then
            dict_code_client.add code, nom
        Else
            Debug.Print Tab(5); "Le code '" & code & "' est un doublon"; Tab(75); "pour le nom '" & nom & "'"
            cas_doublon_code = cas_doublon_code + 1
        End If
    Next i
    Debug.Print Tab(5); "Un total de "; UBound(arr, 1) - 1; " clients ont été analysés!"
    If cas_doublon_nom = 0 Then
        Debug.Print Tab(10); "Aucun doublon de nom"
    Else
        Debug.Print Tab(10); "Il y a " & cas_doublon_nom & " cas de doublons pour les noms"
    End If
    If cas_doublon_code = 0 Then
        Debug.Print Tab(10); "Aucun doublon de code"
    Else
        Debug.Print Tab(10); "Il y a " & cas_doublon_code & " cas de doublons pour les codes"
    End If
    Debug.Print ""
    
End Sub

Private Sub check_GL_Trans()

    'wshGL_Trans
    Dim ws As Worksheet
    Set ws = wshGL_Trans
    Debug.Print ws.name & " (wshGL_Trans)"
    
    Dim arr As Variant
    arr = wshGL_Trans.Range("A1").CurrentRegion.value
    Dim dict_GL_Entry As New Dictionary
    Dim sum_arr() As Double
    ReDim sum_arr(1 To 5000, 1 To 3)
    
    'Array pointer
    Dim row As Long: row = 1
    Dim currentRow As Long
        
    Dim i As Long, GL_Entry_No As String, dt As Double, ct As Double
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        GL_Entry_No = arr(i, 1)
        dt = arr(i, 7)
        ct = arr(i, 8)
        If dict_GL_Entry.Exists(GL_Entry_No) = False Then
            dict_GL_Entry.add GL_Entry_No, row
            sum_arr(row, 1) = GL_Entry_No
            row = row + 1
        End If
        currentRow = dict_GL_Entry(GL_Entry_No)
        sum_arr(currentRow, 2) = sum_arr(currentRow, 2) + dt
        sum_arr(currentRow, 3) = sum_arr(currentRow, 3) + ct
    Next i
    
    Dim sum_dt As Currency, sum_ct As Currency
    Dim cas_hors_balance As Long
    Dim v As Variant
    For Each v In dict_GL_Entry.items()
        GL_Entry_No = sum_arr(v, 1)
        dt = Round(sum_arr(v, 2), 2)
        ct = Round(sum_arr(v, 3), 2)
        If dt <> ct Then
            Debug.Print "Écriture # " & v & " ne balance pas... " & vbNewLine & "Dt = " & Format(dt, "###,###,##0.00") & vbNewLine & "Ct = " & Format(ct, "###,###,##0.00")
            cas_hors_balance = cas_hors_balance + 1
        End If
        sum_dt = sum_dt + dt
        sum_ct = sum_ct + ct
    Next v
    
    Debug.Print Tab(5); "Un total de"; UBound(arr, 1) - 1; "lignes de transactions ont été analysées"
    Debug.Print Tab(10); "- Un total de"; dict_GL_Entry.count; "écritures ont été analysées"
    If cas_hors_balance = 0 Then
        Debug.Print Tab(10); "- TOUTES les écritures balancent au niveau de l'écriture"
    Else
        Debug.Print Tab(10); "Il y a"; cas_hors_balance; "écriture(s) qui ne balance(nt) pas !!!"
    End If
    Debug.Print Tab(5); "Les totaux des transactions sont:"
    Debug.Print Tab(10); "Dt = " & Format(sum_dt, "###,###,##0.00 $")
    Debug.Print Tab(10); "Ct = " & Format(sum_ct, "###,###,##0.00 $")
    Debug.Print ""
    
End Sub

Private Sub check_TEC()

    'wshTEC_Local
    Dim ws As Worksheet
    Set ws = wshTEC_Local
    Debug.Print ws.name & " (wshTEC_Local)"
    
    Dim arr As Variant
    arr = wshTEC_Local.Range("A1").CurrentRegion.Offset(2)
    Dim dict_TEC_ID As New Dictionary
    Dim dict_prof As New Dictionary
    Dim dict_client As New Dictionary
    
    Dim i As Long, TECID As Long, ProfID As String, Prof As String, dateTEC As Date, testDate As Boolean
    Dim code As String, nom As String, hres As Double, testHres As Boolean, estFacturable As Boolean
    Dim estFacturee As Boolean, estDetruit As Boolean
    Dim cas_doublon_TECID As Long, cas_date_invalide As Long, cas_doublon_prof As Long, cas_doublon_client As Long
    Dim cas_hres_invalide As Long, cas_estFacturable_invalide As Long, cas_estFacturee_invalide As Long
    Dim cas_estDetruit_invalide As Long
    Dim total_hres_inscrites As Double, total_hres_detruites As Double, total_hres_facturees As Double
    Dim total_hres_facturable As Double, total_hres_TEC As Double, total_hres_non_facturable As Double
    
    For i = LBound(arr, 1) To UBound(arr, 1) - 2
        TECID = arr(i, 1)
        ProfID = arr(i, 2)
        Prof = arr(i, 3)
        dateTEC = arr(i, 4)
        testDate = IsDate(dateTEC)
        If testDate = False Then
            Debug.Print "TEC_ID ="; TECID; " a une date invalide '"; dateTEC; " !!!"
            cas_date_invalide = cas_date_invalide + 1
        End If
        code = arr(i, 5)
        nom = arr(i, 6)
        hres = arr(i, 8)
        testHres = IsNumeric(hres)
        If testHres = False Then
            Debug.Print "TEC_ID ="; TECID; " la valeur des heures est invalide '"; hres; " !!!"
            cas_hres_invalide = cas_date_invalide + 1
        End If
        estFacturable = arr(i, 10)
        If InStr("Vrai^Faux^", estFacturable & "^") = 0 Or Len(estFacturable) <> 2 Then
            Debug.Print "TEC_ID ="; TECID; " la valeur de la colonne 'EstFacturable' est invalide '" & estFacturable & "' !!!"
            cas_estFacturable_invalide = cas_estFacturable_invalide + 1
        End If
        estFacturee = arr(i, 12)
        If InStr("Vrai^Faux^", estFacturee & "^") = 0 Or Len(estFacturee) <> 2 Then
            Debug.Print "TEC_ID ="; TECID; " la valeur de la colonne 'EstFacturee' est invalide '" & estFacturee & "' !!!"
            cas_estFacturee_invalide = cas_estFacturee_invalide + 1
        End If
        estDetruit = arr(i, 12)
        If InStr("Vrai^Faux^", estDetruit & "^") = 0 Or Len(estDetruit) <> 2 Then
            Debug.Print "TEC_ID ="; TECID; " la valeur de la colonne 'estDetruit' est invalide '" & estDetruit & "' !!!"
            cas_estDetruit_invalide = cas_estDetruit_invalide + 1
        End If
        
        total_hres_inscrites = total_hres_inscrites + hres
        If estDetruit = "Vrai" Then total_hres_detruites = total_hres_detruites + hres
        
       
        If estDetruit = "Faux" And estFacturable = "Vrai" Then total_hres_facturable = total_hres_facturable + hres
        If estDetruit = "Faux" And estFacturable = "Faux" Then total_hres_non_facturable = total_hres_non_facturable + hres
        If estDetruit = "Faux" And estFacturee = "Vrai" Then total_hres_facturees = total_hres_facturees + hres
        
        'Dictionary
        If dict_TEC_ID.Exists(TECID) = False Then
            dict_TEC_ID.add TECID, 0
        Else
            Debug.Print Tab(5); "Le TEC_ID '" & TECID & "' est un doublon"; Tab(75); "pour la rangée '" & i & "'"
            cas_doublon_TECID = cas_doublon_TECID + 1
        End If
        If dict_prof.Exists(Prof & "-" & ProfID) = False Then
            dict_prof.add Prof & "-" & ProfID, 0
        End If
        If dict_client.Exists(nom & "-" & code) = False Then
            dict_client.add nom & "-" & code, 0
        End If
    Next i
    
    Debug.Print Tab(5); "Un total de"; UBound(arr, 1) - 2; "charges de temps ont été analysées!"
    If cas_doublon_TECID = 0 Then
        Debug.Print Tab(10); "Aucun doublon de TEC_ID"
    Else
        Debug.Print Tab(10); "Il y a " & cas_doublon_TECID & " cas de doublons pour les TEC_ID"
    End If
    If cas_date_invalide = 0 Then
        Debug.Print Tab(10); "Aucune date INVALIDE"
    Else
        Debug.Print Tab(10); "Il y a " & cas_date_invalide & " cas de date INVALIDE"
    End If
    If cas_hres_invalide = 0 Then
        Debug.Print Tab(10); "Aucune heures INVALIDE"
    Else
        Debug.Print Tab(10); "Il y a " & cas_hres_invalide & " cas d'heures INVALIDE"
    End If
    If cas_estFacturable_invalide = 0 Then
        Debug.Print Tab(10); "Aucune valeur 'estFacturable' n'est INVALIDE"
    Else
        Debug.Print Tab(10); "Il y a " & cas_estFacturable_invalide & " cas de valeur 'estFacturable' INVALIDE"
    End If
    If cas_estFacturee_invalide = 0 Then
        Debug.Print Tab(10); "Aucune valeur 'estFacturee' n'est INVALIDE"
    Else
        Debug.Print Tab(10); "Il y a " & cas_estFacturee_invalide & " cas de valeur 'estFacturee' INVALIDE"
    End If
    If cas_estDetruit_invalide = 0 Then
        Debug.Print Tab(10); "Aucune valeur 'estDetruit' n'est INVALIDE"
    Else
        Debug.Print Tab(10); "Il y a " & cas_estDetruit_invalide & " cas de valeur 'estDetruit' INVALIDE"
    End If
    Debug.Print "La somme des heures donne ce resultat:"
    Debug.Print Tab(10); "Heures inscrites       : "; total_hres_inscrites
    Debug.Print Tab(10); "Heures détruites       : "; total_hres_detruites
    Debug.Print Tab(10); "Heures restantes       : "; total_hres_inscrites - total_hres_detruites
    Debug.Print Tab(10); "Heures facturables     : "; total_hres_facturable
    Debug.Print Tab(10); "Heures non_facturables : "; total_hres_non_facturable
    
    'Loop over Keys and get item
    Dim k As Variant
    For Each k In dict_prof.Keys()
        Debug.Print k, dict_prof(k)
    Next k

'    'Loop over Keys and get item
'    Dim d As Variant
'    For Each d In dict_client.Keys()
'        Debug.Print d, dict_client(d)
'    Next d

End Sub


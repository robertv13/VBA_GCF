Attribute VB_Name = "modGL_Rapport"
Option Explicit

Public Sub GL_Report_For_Selected_Accounts()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport:GL_Report_For_Selected_Accounts", 0)
   
    'Reference the worksheet
    Dim ws As Worksheet:  Set ws = wshGL_Rapport

    If ws.Range("F6").value = "" Or ws.Range("H6").value = "" Then
        MsgBox "Vous devez saisir une date de début et une date de fin pour ce rapport!"
        Exit Sub
    End If
    
    If Application.WorksheetFunction.Days(ws.Range("H6").value, ws.Range("F6").value) < 0 Then
        MsgBox "La date de départ doit obligatoirement être antérieure" & vbNewLine & vbNewLine & _
                "ou égale à la date de fin!", vbInformation
        Exit Sub
    End If
    
    'Reference the listBox
    Dim lb As OLEObject: Set lb = ws.OLEObjects("ListBox1")

    'Ensure it is a ListBox
    Dim selectedItems As Collection
    If TypeName(lb.Object) = "ListBox" Then
        Set selectedItems = New Collection

        'Loop through ListBox items and collect selected ones
        Dim i As Long
        With lb.Object
            For i = 0 To .ListCount - 1
                If .Selected(i) And Trim(.List(i)) <> "" Then
                    selectedItems.Add .List(i)
                End If
            Next i
        End With

        'Is there any account selected ?
        If selectedItems.count = 0 Then
            MsgBox "Il n'y a aucune compte de sélectionné!"
            Exit Sub
        End If
        
        'Erase & Create output Worksheet
        Call CreateOrReplaceWorksheet("X_GL_Rapport_Out")
        
        'Setup report header
        Call Set_Up_Report_Headers_And_Columns
        
        'Prepare Variables
        Dim dateDeb As Date, dateFin As Date, sortType As String
        With wshGL_Rapport
            dateDeb = CDate(.Range("F6").value)
            dateFin = CDate(.Range("H6").value)
            If .Range("B3").value = "Vrai" Then
                sortType = "Date"
            Else
                sortType = "Transaction"
            End If
        End With
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        'Process one account at the time...
        Dim item As Variant
        Dim compte As String
        For Each item In selectedItems
            compte = item
            
            Call get_GL_Trans_With_AF(compte, dateDeb, dateFin, sortType)
            
            Call Print_results_From_GL_Trans(compte, dateDeb)
        
        Next item
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        
    End If
    
    Dim h1 As String, h2 As String, h3 As String
    h1 = wshAdmin.Range("NomEntreprise")
    h2 = "Rapport des transactions du Grand Livre"
    h3 = "(Du " & dateDeb & " au " & dateFin & ")"
    Call GL_Rapport_Wrap_Up(h1, h2, h3)
    
    'Clean up
    Set item = Nothing
    Set lb = Nothing
    Set selectedItems = Nothing
    Set ws = Nothing
    
    Call Log_Record("modGL_Rapport:GL_Report_For_Selected_Accounts", startTime)

End Sub

Sub get_GL_Trans_With_AF(compte As String, dateDeb As Date, dateFin As Date, sortType As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport:get_GL_Trans_With_AF", 0)

    Dim glNo As String
    glNo = Left(compte, InStr(compte, " ") - 1)
    
    'Setup ADVANCED FILTER with 3 ranges
    
    'Data source
    With wshGL_Trans
        Dim rgData As Range: Set rgData = .Range("A1").CurrentRegion
        
        'Assign Criteria (3)
        .Range("L3").value = glNo
        .Range("M3").value = ">=" & Format$(dateDeb, "mm-dd-yyyy")
        .Range("N3").value = "<=" & Format$(dateFin, "mm-dd-yyyy")
        Dim rgCriteria As Range: Set rgCriteria = .Range("L2:N3")
        
        'Destination to copy (setup & clear previous results)
        Dim rgCopyToRange As Range: Set rgCopyToRange = .Range("P1").CurrentRegion
        rgCopyToRange.Offset(1).ClearContents
        
        'Do the Advanced Filter
        rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgCopyToRange
        
        Dim lastResultUsedRow
        lastResultUsedRow = .Range("P99999").End(xlUp).Row
        If lastResultUsedRow < 3 Then GoTo NoSort
        With .Sort
            .SortFields.Clear
            If sortType = "Date" Then
                .SortFields.Add key:=wshGL_Trans.Range("Q2:Q" & lastResultUsedRow), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortTextAsNumbers 'Sort Based Date
            Else
                .SortFields.Add key:=wshGL_Trans.Range("P2:P" & lastResultUsedRow), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortTextAsNumbers 'Sort on Transaction #
            End If
            .SetRange wshGL_Trans.Range("P2:Y" & lastResultUsedRow) 'Set Range
            .Apply 'Apply Sort
         End With
    End With

NoSort:
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set rgCriteria = Nothing
    Set rgCopyToRange = Nothing
    Set rgData = Nothing
    
    Call Log_Record("modGL_Rapport:get_GL_Details_For_A_Account", startTime)

End Sub

Sub Print_results_From_GL_Trans(compte As String, dateDeb As Date)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("X_GL_Rapport_Out")
    
    Dim lastRowUsed_AB As Long, lastRowUsed_A As Long, lastRowUsed_B As Long
    Dim saveFirstRow As Long
    Dim solde As Currency
    lastRowUsed_A = ws.Range("A99999").End(xlUp).Row
    lastRowUsed_B = ws.Range("B99999").End(xlUp).Row
    If lastRowUsed_A > lastRowUsed_B Then
        lastRowUsed_AB = lastRowUsed_A
    Else
        lastRowUsed_AB = lastRowUsed_B
    End If
    
    lastRowUsed_AB = lastRowUsed_AB + 2
    ws.Range("A" & lastRowUsed_AB).value = compte
    ws.Range("A" & lastRowUsed_AB).Font.Bold = True
    
    'Solde d'ouverture pour ce compte
    Dim soldeOuverture As Double
    Dim glNo As String
    glNo = Left(compte, InStr(compte, " ") - 1)
    soldeOuverture = Fn_Get_Account_Opening_Balance(glNo, dateDeb)
    solde = soldeOuverture
    ws.Range("D" & lastRowUsed_AB).value = "Solde d'ouverture"
    
    ws.Range("H" & lastRowUsed_AB).value = solde
    ws.Range("H" & lastRowUsed_AB).Font.Bold = True
    lastRowUsed_AB = lastRowUsed_AB + 1
    saveFirstRow = lastRowUsed_AB

    If wshGL_Trans.Range("P2") = "" Then
        Exit Sub
    End If
    
    Dim lastUsedTrans As Long
    lastUsedTrans = wshGL_Trans.Cells(wshGL_Trans.rows.count, "P").End(xlUp).Row '2024-08-15 @ 15:46
    If lastUsedTrans > 1 Then
        Dim i As Long, sumDT As Currency, sumCT As Currency
        'Read thru the rows
        For i = 2 To lastUsedTrans
            ws.Cells(lastRowUsed_AB, 2).value = wshGL_Trans.Range("Q" & i).value
            ws.Cells(lastRowUsed_AB, 3).value = wshGL_Trans.Range("R" & i).value
            ws.Cells(lastRowUsed_AB, 4).value = wshGL_Trans.Range("S" & i).value
            ws.Cells(lastRowUsed_AB, 5).value = wshGL_Trans.Range("P" & i).value
            ws.Cells(lastRowUsed_AB, 6).value = wshGL_Trans.Range("V" & i).value
            ws.Cells(lastRowUsed_AB, 7).value = wshGL_Trans.Range("W" & i).value
            solde = solde + CCur(wshGL_Trans.Range("V" & i).value) - CCur(wshGL_Trans.Range("W" & i).value)
            ws.Cells(lastRowUsed_AB, 8).value = solde
            
            sumDT = sumDT + wshGL_Trans.Range("V" & i).value
            sumCT = sumCT + wshGL_Trans.Range("W" & i).value
            
            lastRowUsed_AB = lastRowUsed_AB + 1
        Next i
    End If
    
    'Ajoute le formatage conditionnel pour les transactions
    With Range("B" & saveFirstRow & ":H" & lastRowUsed_AB - 1)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(LIGNE();2)=1"
        .FormatConditions(Range("B" & saveFirstRow & ":H" & lastRowUsed_AB - 1).FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.14996795556505
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
    
    ws.Range("H" & lastRowUsed_AB - 1).Font.Bold = True
    With ws.Range("F" & lastRowUsed_AB, "G" & lastRowUsed_AB)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    ws.Range("F" & lastRowUsed_AB).value = sumDT
    ws.Range("G" & lastRowUsed_AB).value = sumCT
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Public Sub GL_Rapport_Clear_All_Cells(ws As Worksheet)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport:GL_Rapport_Clear_All_Cells", 0)
    
    With ws
        .Range("B3").value = True 'Sort by Date
        .Range("F4").value = "Dates manuelles"
        .Range("F6").value = ""
        .Range("H6").value = ""
        .Range("F4").Select
    End With
    
    Call Log_Record("modGL_Rapport:GL_Rapport_Clear_All_Cells", startTime)

End Sub

Sub Set_Up_Report_Headers_And_Columns()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("X_GL_Rapport_Out")
    
    With ws
        .Cells(1, 1) = "Compte"
        .Cells(1, 2) = "Date"
        .Cells(1, 3) = "Description"
        .Cells(1, 4) = "Source"
        .Cells(1, 5) = "No.Écr."
        .Cells(1, 6) = "Débit"
        .Cells(1, 7) = "Crédit"
        .Cells(1, 8) = "SOLDE"
        With .Range("A1:H1")
            .Font.Italic = True
            .Font.Bold = True
            .Font.size = 10
            .HorizontalAlignment = xlCenter
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
        End With
    
        With .columns("A")
            .ColumnWidth = 5
        End With
        
        With .columns("B")
            .ColumnWidth = 11
            .HorizontalAlignment = xlCenter
        End With
        
        With .columns("C")
            .ColumnWidth = 50
        End With
        
        With .columns("D")
            .ColumnWidth = 20
        End With
        
        With .columns("E")
            .ColumnWidth = 9
            .HorizontalAlignment = xlCenter
        End With
        
        With .columns("F")
            .ColumnWidth = 15
        End With
        
        With .columns("G")
            .ColumnWidth = 15
        End With
        
        With .columns("H")
            .ColumnWidth = 15
        End With
    End With
    
    ws.Range("A2:H" & 2500).ClearContents '2024-09-28 @ 06:36

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Sub GL_Rapport_Wrap_Up(h1 As String, h2 As String, h3 As String)

    Application.PrintCommunication = False
    
    'Determine the active cells & setup Print Area
    Dim lastUsedRow As Long
    lastUsedRow = ThisWorkbook.Worksheets("X_GL_Rapport_Out").Range("H999999").End(xlUp).Row + 1
    Range("A3:H" & lastUsedRow).Select
    
    With ActiveSheet.PageSetup
        .PrintArea = "$A$3:$H$" & lastUsedRow
        .PrintTitleRows = "$1:$2"
        
        .LeftMargin = Application.InchesToPoints(0.15)
        .RightMargin = Application.InchesToPoints(0.15)
        .TopMargin = Application.InchesToPoints(0.85)
        .BottomMargin = Application.InchesToPoints(0.45)
        .HeaderMargin = Application.InchesToPoints(0.15)
        .FooterMargin = Application.InchesToPoints(0.15)
        
        .LeftHeader = ""
        .CenterHeader = "&""-,Gras""&16" & h1 & _
                        Chr(10) & "&11" & h2 & _
                        Chr(10) & "&11" & h3
        .RightHeader = ""
        
        .LeftFooter = "&9&D - &T"
        .CenterFooter = ""
        .RightFooter = "&9Page &P de &N"

        .FitToPagesWide = 1
        .FitToPagesTall = 99
        
    End With
    
    'Keep header rows always displayed
    ActiveWindow.SplitRow = 2

    Range("A" & lastUsedRow).Select
    
    MsgBox "Le rapport a été généré avec succès", vbInformation, "Rapport des transactions du Grand Livre"
    
    Application.PrintCommunication = True

End Sub

Sub GL_Rapport_Back_To_Menu()
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport:GL_Rapport_Back_To_Menu", 0)
   
    wshGL_Rapport.Visible = xlSheetHidden
    On Error Resume Next
    ThisWorkbook.Worksheets("X_GL_Rapport_Out").Visible = xlSheetHidden
    On Error GoTo 0

    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
    Call Log_Record("modGL_Rapport:GL_Rapport_Back_To_Menu", startTime)
    
End Sub



Attribute VB_Name = "modAppli"
Option Explicit

Public Const APP_VERSION_NO As String = "v3.9.1" '2024-07-01 @ 09:37
Public Const NB_MAX_LIGNE_FAC As Integer = 35 '2024-06-18 @ 12:18
Public Const HIGHLIGHT_COLOR As String = &HCCFFCC 'Light green (Pastel Green)

Public interior_color_current_cell As Long
Public userName As String

'Using Enum to specify the column number of worksheets (data)
Public Enum DEB_Trans_data_Columns
    fdebtFirst = 1
    fdebtNo_Entr�e = fdebtFirst
    fdebtDate
    fdebtType
    fdebtBeneficiaire
    fdebtReference
    fdebtNo_Compte
    fdebtCodeTaxe
    fdebtTOTAL
    fdebtTPS
    fdebtTVQ
    fdebtCr�dit_TPS
    fdebtCr�dit_TVQ
    fdebtAutreRemarque
    fdebtTimeStamp
    fdebtLast = fdebtTimeStamp
End Enum

Public Enum FAC_Ent�te_Data_Columns
    ffacEntFirst = 1
    ffacEntInv_No = ffacEntFirst
    ffacEntDate_Facture
    ffacEntCust_ID
    ffacEntContact
    ffacEntNom_Client
    ffacEntAdresse1
    ffacEntAdresse2
    ffacEntAdresse3
    ffacEntHonoraires
    ffacEntAF1_Desc
    ffacEntAutresFrais_1
    ffacEntAF2_Desc
    ffacEntAutresFrais_2
    ffacEntAF3_Desc
    ffacEntAutresFrais_3
    ffacEntTaux_TPS
    ffacEntMnt_TPS
    ffacEntTaux_TVQ
    ffacEntMntTVQ
    ffacEntAR_Total
    ffacEntD�p�t
    ffacEntLast = ffacEntD�p�t
End Enum

Public Enum GL_EJ_Auto_Data_Columns
    fglejaFirst = 1
    fglejaNo_EJA = fglejaFirst
    fglejaDescription
    fglejaNo_Compte
    fglejaCompte
    fglejaD�bit
    fglejaCr�dit
    fglejaAutreRemarque
    fglejaLast = fglejaAutreRemarque
End Enum

Public Enum GL_Trans_Data_Columns
    fgltFirst = 1
    fgltEntryNo = fgltFirst
    fgltDate
    fgltDescr
    fgltSource
    fgltGLNo
    fgltCompte
    fgltdt
    fgltct
    fgltRem
    fgltTStamp
    fgltLast = fgltTStamp
End Enum

Public Enum TEC_Data_Columns
    ftecFirst = 1
    ftec_ID = ftecFirst
    ftecProf_ID
    ftecProf
    ftecDate
    ftecClient_ID
    ftecClientNom
    ftecDescription
    ftecHeures
    ftecCommentaireNote
    ftecEstFacturable
    ftecDateSaisie
    ftecEstFacturee
    ftecDateFacturee
    ftecEstDetruit
    ftecVersionApp
    ftecNoFacture
    ftecLast = ftecNoFacture
End Enum

Sub BackToMainMenu()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.name <> "Menu" Then ws.Visible = xlSheetHidden
    Next ws
    wshMenu.Activate
    wshMenu.Range("A1").Select

    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Private Sub auto_open() '2024-03-06 @ 14:36

    userName = Environ("Username") '2024-03-27 @ 06:54

End Sub

Private Sub auto_close() '2024-03-06 @ 14:36

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modAppli:auto_close()")

    MsgBox "Auto_Close..."
    
    Call Output_Timer_Results("modAppli:auto_close()", timerStart)
    
End Sub

Sub Dynamic_Range_Redefine_Plan_Comptable() '2024-06-06 @ 07:41
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modAppli:Dynamic_Range_Redefine_Plan_Comptable()")

    'Delete existing dynamic named range (assuming it could exists)
    On Error Resume Next
    ThisWorkbook.Names("dnrPlanComptableDescription").delete
    On Error GoTo 0
    
    'Define a new dynamic named range for 'dnrPlanComptableDescription'
    Dim newRangeFormula As String
    newRangeFormula = "=OFFSET(Admin!$T$11,,,COUNTA(Admin!$T:$T)-2,1)"
    
    'Create a new dynamic named range
    ThisWorkbook.Names.add name:="dnrPlanComptableDescription", RefersTo:=newRangeFormula
    
    Call Output_Timer_Results("modAppli:Dynamic_Range_Redefine_Plan_Comptable()", timerStart)
    
End Sub

Sub Hide_All_Worksheets_Except_Menu() '2024-02-20 @ 07:28
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modAppli:Hide_All_Worksheets_Except_Menu()")
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.codeName <> "wshMenu" And _
            InStr(ws.codeName, "wshzDoc") = 0 Then
                ws.Visible = xlSheetHidden
        End If
    Next ws
    
    Call Output_Timer_Results("modAppli:Hide_All_Worksheets_Except_Menu()", timerStart)
    
    'Cleaning memory - 2024-07-01 @ 09:34
    Set ws = Nothing
    
End Sub

Sub Buttons_Enabled_True_Or_False(clear As Boolean, add As Boolean, _
                                  update As Boolean, delete As Boolean)
    With ufSaisieHeures
        .cmdClear.Enabled = clear
        .cmdAdd.Enabled = add
        .cmdUpdate.Enabled = update
        .cmdDelete.Enabled = delete
    End With

End Sub

Sub Invalid_Date_Message() '2024-03-03 @ 07:45 - TBD ??

''    MsgBox Prompt:="La valeur saisie ne peut �tre utilis�e comme une date valide", _
''        Title:="Validation de la date", _
''        Buttons:=vbCritical

End Sub

Sub Erreur_Totaux_DT_CT()

    MsgBox Prompt:="Les totaux (D�bit vs. Cr�dit) sont diff�rents !!!", _
        Title:="Validation des totaux du G/L", _
        Buttons:=vbCritical

End Sub

Sub Pause_Application(s As Double)
    
    If s > 5 Then Stop
    
    Dim endTime As Double
    endTime = Timer + s 'Set end time to 's' seconds from now
    
    Do While Timer < endTime
        'Sleep
    Loop
    
End Sub

Sub Slide_In_All_Menu_Options()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modAppli:Slide_In_All_Menu_Options()")
    
    Call SlideIn_TEC
    Call SlideIn_Facturation
    Call SlideIn_Debours
    Call SlideIn_Comptabilite
    Call SlideIn_Parametres
    Call SlideIn_Exit

    Call Output_Timer_Results("modAppli:Slide_In_All_Menu_Options()", timerStart)

End Sub

Sub MsgBoxInvalidDate() '2024-06-13 @ 12:40

    MsgBox "La date saisie ne peut �tre accept�e tel qu'elle est entr�e." & vbNewLine & vbNewLine & _
           "Elle doit �tre obligatoirement de format:" & vbNewLine & _
           "     'jj', " & vbNewLine & _
           "     'jj-mm' ou " & vbNewLine & _
           "     'jj-mm-aaaa'" & vbNewLine & vbNewLine & _
           "Veuillez saisir la date de nouveau SVP", _
           vbCritical, _
           "La date saisie est INVALIDE"

End Sub

Sub SetTabOrder(ws As Worksheet) '2024-06-15 @ 13:58

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modAppli:SetTabOrder()")
    
    'Clear previous settings AND protect the worksheet
    ws.EnableSelection = xlNoRestrictions
    ws.Protect UserInterfaceOnly:=True

    'Collect all unprotected cells
    Dim cell As Range, unprotectedCells As Range
    For Each cell In ws.usedRange
        If Not cell.Locked Then
            If unprotectedCells Is Nothing Then
                Set unprotectedCells = cell
            Else
                Set unprotectedCells = Union(unprotectedCells, cell)
            End If
        End If
    Next cell

    'Sort to ensure cells are sorted left-to-right, top-to-bottom
    Dim sortedCells As Range: Set sortedCells = unprotectedCells.SpecialCells(xlCellTypeVisible)
    Debug.Print ws.name & " - Unprotected cells are '" & sortedCells.Address & "' - " & sortedCells.count & " - " & Format(Now(), "dd/mm/yyyy hh:mm:ss")

    'Enable TAB through unprotected cells
    Application.EnableEvents = False
    Dim i As Long
    For i = 1 To sortedCells.count
        If i = sortedCells.count Then
            sortedCells.Cells(i).Next.Select
        Else
            sortedCells.Cells(i).Next.Select
            sortedCells.Cells(i + 1).Activate
        End If
    Next i
    
    Application.EnableEvents = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    Set unprotectedCells = Nothing
    Set sortedCells = Nothing
    
    Call Output_Timer_Results("modAppli:SetTabOrder()", timerStart)

    'Cleaning memory - 2024-07-01 @ 09:34
    Set cell = Nothing
    
End Sub

Sub BackupMasterFile()

    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modAppli:BackupMasterFile()")
    
    Application.ScreenUpdating = False
    
    'Open the master file
    Dim masterWorkbook As Workbook
    Set masterWorkbook = Workbooks.Open("C:\VBA\GC_FISCALIT�\DataFiles\GCF_BD_Sortie.xlsx")
    
    'Get the current date and time in the format YYYYMMDD_HHMMSS
    Dim currentDateAndTime As String
    currentDateAndTime = Format(Now, "YYYYMMDD_HHMMSS")

    'Create the backup file name
    Dim backupFileName As String
    backupFileName = Left(masterWorkbook.name, InStrRev(masterWorkbook.name, ".") - 1) & "_" & currentDateAndTime & ".xlsx"

    'Define the backup file path (same directory as the master file)
    Dim backupFilePath As String
    backupFilePath = masterWorkbook.path & "\" & backupFileName

    'Save a copy of the master workbook with the new name
    masterWorkbook.SaveCopyAs backupFilePath

    'Close the master workbook
    masterWorkbook.Close SaveChanges:=False

'    'Optional: Notify the user
'    MsgBox "Backup created: " & vbNewLine & vbNewLine & "'" & backupFilePath & "'"

    Application.ScreenUpdating = True

    'Cleaning memory - 2024-07-01 @ 09:34
    Set masterWorkbook = Nothing
    
    Call Output_Timer_Results("modAppli:BackupMasterFile()", timerStart)

End Sub

Sub TEST_GetOneDrivePath()

    On Error GoTo eh
    Debug.Print "Original Path is: " & ThisWorkbook.path & "/" & ThisWorkbook.name
    Debug.Print "The Path is     : " & GetOneDrivePath(ThisWorkbook.FullName)
    Exit Sub
eh:
    MsgBox Err.Description
    
End Sub


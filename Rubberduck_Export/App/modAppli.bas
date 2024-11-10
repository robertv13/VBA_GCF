Attribute VB_Name = "modAppli"
'@Folder(GC_FISCALIT�.Main)

Option Explicit

#If VBA7 Then
    'D�claration pour les environnements 64 bits
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'    Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
'    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#Else
    '32-bit Excel (anciennes versions)
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'    Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If

'Public Const VK_NUMLOCK As Long = &H90

Public Const NB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18
Public Const HIGHLIGHT_COLOR As String = &HCCFFCC 'Light green (Pastel Green)
Public Const BASIC_COLOR As Long = 16777215 '2024-07-23 @ 08:15

Public Const DATA_PATH As String = "\DataFiles"
Public Const FACT_PDF_PATH As String = "\Factures_PDF"
Public Const FACT_EXCEL_PATH As String = "\Factures_Excel"

'Variable utilis�e pour �viter l'�v�nement Activate � chaque fois que l'on revient dans une feuille
Public fromMenu As Boolean '2024-09-03 @ 06:14

'Niveau de d�tail pour le log de SaisieHeures
Public logSaisieHeuresVeryDetailed As Boolean

'Pour assurer un contr�le dans Facture Finale
Public flagEtapeFacture As Integer

'Using Enum to specify the column number of worksheets (data)
Public Enum DB_Clients '2024-10-26 @ 17:41
    fClntMFFirst = 1
    fClntMFClientNom = fClntMFFirst
    fClntMFClient_ID
    fClntMFNomClientSyst�me
    fClntMFContactFacturation
    fClntMFTitreContactFacturation
    fClntMFCourrielFacturation
    fClntMFAdresse_1
    fClntMFAdresse_2
    fClntMFVille
    fClntMFProvince
    fClntMFCodePostal
    fClntMFPays
    fClntMFR�f�r�Par
    fClntMFFinAnn�e
    fClntMFComptable
    fClntMFNotaire_Avocat
    fClntMFTimeStamp
    fClntMFLast = fClntMFTimeStamp
End Enum

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
    ffacEntFouP
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
    ftecTEC_ID = ftecFirst
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

Sub Set_Root_Path(ByRef rootPath As String)
   
    If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        rootPath = "C:\VBA\GC_FISCALIT�"
    Else
        rootPath = "P:\Administration\APP\GCF"
    End If

End Sub

Sub Write_Info_On_Main_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:Write_Info_On_Main_Menu", 0)
    
    Application.EnableEvents = False
    wshMenu.Unprotect
    Application.ScreenUpdating = True
    
    With wshMenu.Range("$A$30")
        .Font.size = 8
        .Font.Color = vbBlue
        .value = "'" & CStr("Heure - " & Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    End With
    
    With wshMenu.Range("$A$31")
        .Font.size = 8
        .Font.Color = vbBlack
        .value = "'" & CStr("Version - " & ThisWorkbook.Name)
    End With
    
    With wshMenu.Range("$A$32")
        .Font.size = 8
        .Font.Color = vbBlack
        .value = "'" & CStr("Utilisateur - " & Fn_Get_Windows_Username)
    End With
    
    Dim env As String: env = wshAdmin.Range("F5").value
    With wshMenu.Range("$A$33")
        .Font.size = 8
        .Font.Color = vbRed
        .value = "'" & CStr("Environnement - " & wshAdmin.Range("F5").value)
    End With

    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    DoEvents '2024-08-23 @ 06:21

    Call Log_Record("modAppli:Write_Info_On_Main_Menu", startTime)

End Sub

Sub Handle_Rubberduck_Reference()

    Dim ref As Object

    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
        On Error Resume Next 'In case the reference doesn't exist
        For Each ref In ThisWorkbook.VBProject.References
            If ref.Name = "Rubberduck Addin" Then 'Rubberduck is the name of the reference to remove
                ThisWorkbook.VBProject.References.Remove ref
                Exit For
            End If
        Next ref
        On Error GoTo 0
    End If

    'Lib�rer la m�moire
    Set ref = Nothing
    
End Sub

Sub UpdatePivotTables()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:UpdatePivotTables", 0)
    
    Dim ws As Worksheet: Set ws = wshStatsHeuresPivotTables
    Dim pt As pivotTable
    
    'Parcourt tous les PivotTables dans chaque feuille
    For Each pt In ws.PivotTables
        On Error Resume Next
        Application.EnableEvents = False
        pt.PivotCache.Refresh 'Actualise le cache Pivot
        Application.EnableEvents = True
        On Error GoTo 0
    Next pt

    'Lib�rer la m�moire
    Set pt = Nothing
    Set ws = Nothing
    
    Call Log_Record("modAppli:UpdatePivotTables", startTime)
    
End Sub

Public Sub Get_GL_Trans_With_AF(glCode As String, dateDeb As Date, dateFin As Date) '2024-11-08 @ 09:34

    Dim ws As Worksheet: Set ws = wshGL_Trans
    
    'O� allons-nous mettre les r�sultats ?
    Dim rngResult As Range
    Set rngResult = ws.Range("P1").CurrentRegion.Offset(1, 0)
    rngResult.ClearContents
    Set rngResult = ws.Range("P1").CurrentRegion
    
    'O� sont les donn�es � traiter ?
    Dim rngSource As Range
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    'Rien � traiter
    If lastUsedRow < 2 Then
        Exit Sub
    End If
    Set rngSource = ws.Range("A1:J" & lastUsedRow)
    
    'Quels sont les crit�res ?
    Dim rngCriteria As Range
    Set rngCriteria = ws.Range("L2:N3")
    With ws
        .Range("L3").value = glCode
        .Range("M3").value = ">=" & CLng(dateDeb)
        .Range("N3").value = "<=" & CLng(dateFin)
    End With
    
    'On documente le processus
    ws.Range("M8").value = "Derni�re utilisation: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    ws.Range("M9").value = rngSource.Address
    ws.Range("M10").value = rngCriteria.Address
    ws.Range("M11").value = rngResult.Address
    
    'Go, on execute le AdvancedFilter
    rngSource.AdvancedFilter xlFilterCopy, _
                             rngCriteria, _
                             rngResult, _
                             False
    
    'Combien y a-t-il de transactions dans le r�sultat ?
    lastUsedRow = ws.Cells(ws.rows.count, "P").End(xlUp).row
    ws.Range("M12").value = lastUsedRow
    Set rngResult = ws.Range("P1:Y" & lastUsedRow)

    If lastUsedRow > 2 Then
        With ws.Sort
            .SortFields.Clear
                .SortFields.Add _
                    key:=ws.Range("Q2"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par date de transaction
                .SortFields.Add _
                    key:=ws.Range("P2"), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal 'Trier par num�ro d'�criture
            .SetRange rngResult
            .Header = xlYes
            .Apply
        End With
    End If

End Sub

Attribute VB_Name = "modAppli"
'@Folder(GC_FISCALITÉ.Main)

Option Explicit

#If VBA7 Then
    '64-bit Excel (VBA7 and later)
'    Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'    Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#Else
    '32-bit Excel
    Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If

Public Const NB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18
Public Const HIGHLIGHT_COLOR As String = &HCCFFCC 'Light green (Pastel Green)
Public Const BASIC_COLOR As Long = 16777215 '2024-07-23 @ 08:15
Public Const MAXWIDTH As Long = 192

Public Const DATA_PATH As String = "\DataFiles"
Public Const FACT_PDF_PATH As String = "\Factures_PDF"
Public Const FACT_EXCEL_PATH As String = "\Factures_Excel"

Public fromMenu As Boolean '2024-09-03 @ 06:14

'Constantes pour les touches du clavier
Private Const VK_NUMLOCK As Byte = &H90
Private Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Private Const KEYEVENTF_KEYUP As Long = &H2

'Using Enum to specify the column number of worksheets (data)
Public Enum DEB_Trans_data_Columns
    fdebtFirst = 1
    fdebtNo_Entrée = fdebtFirst
    fdebtDate
    fdebtType
    fdebtBeneficiaire
    fdebtReference
    fdebtNo_Compte
    fdebtCodeTaxe
    fdebtTOTAL
    fdebtTPS
    fdebtTVQ
    fdebtCrédit_TPS
    fdebtCrédit_TVQ
    fdebtAutreRemarque
    fdebtTimeStamp
    fdebtLast = fdebtTimeStamp
End Enum

Public Enum FAC_Entête_Data_Columns
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
    ffacEntDépôt
    ffacEntLast = ffacEntDépôt
End Enum

Public Enum GL_EJ_Auto_Data_Columns
    fglejaFirst = 1
    fglejaNo_EJA = fglejaFirst
    fglejaDescription
    fglejaNo_Compte
    fglejaCompte
    fglejaDébit
    fglejaCrédit
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
        rootPath = "C:\VBA\GC_FISCALITÉ"
    Else
        rootPath = "P:\Administration\APP\GCF"
    End If

End Sub

Sub Write_Info_On_Main_Menu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:Write_Info_On_Main_Menu", 0)
    
    Application.EnableEvents = False
    wshMenu.Unprotect
    Application.ScreenUpdating = True
    
    With wshMenu.Range("$A$32")
        .Font.size = 8
        .Font.Color = vbBlue
        .value = "'" & CStr("Heure - " & Format$(Now(), "dd-mm-yyyy hh:mm:ss"))
    End With
    
    With wshMenu.Range("$A$33")
        .Font.size = 8
        .Font.Color = vbBlack
        .value = "'" & CStr("Version - " & ThisWorkbook.name)
    End With
    
    With wshMenu.Range("$A$34")
        .Font.size = 8
        .Font.Color = vbBlack
        .value = "'" & CStr("Utilisateur - " & Fn_Get_Windows_Username)
    End With
    
    Dim env As String: env = wshAdmin.Range("F5").value
    With wshMenu.Range("$A$35")
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

    Call Log_Record("modAppli:Write_Info_On_Main_Menu()", startTime)

End Sub

Sub Ensure_NumLock_Is_On()

    'Vérifie l'état actuel de NumLock
    Dim NumLockState As Boolean
    NumLockState = GetKeyState(VK_NUMLOCK) And 1
    Debug.Print "DB.#190 - NumLockState (avant) = '" & NumLockState & "' - "; Now

    If NumLockState = False Then
            Application.SendKeys "{NUMLOCK}", True
    End If

'    'Si NumLock est désactivé, on l'active
'    If Not NumLockState Then
'        'Simule l'appui sur la touche NumLock
'        Debug.Print "     DB.#195 - Simulate PRESS NUM LOCK"
'        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY, 0
'        'Simule le relâchement de la touche NumLock
'        Debug.Print "     DB.#198 - Simulate RELEASE NUM LOCK"
'        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
'
'        'Petite pause pour laisser le temps au système de traiter la modification
'        Application.Wait Now + TimeValue("00:00:02")
'
'        'Check the status AFTER modifications...
'        NumLockState = GetKeyState(VK_NUMLOCK) And 1
'        Debug.Print "DB.#203 - NumLockState (after) = "; NumLockState; " "; Now
'    End If

    NumLockState = GetKeyState(VK_NUMLOCK) And 1
    Debug.Print "DB.#210 - NumLockState (après) = '" & NumLockState & "' - "; Now


End Sub

Sub AssurerNumLockActive()
    Const VK_NUMLOCK As Long = &H90 ' Code de la touche NUM LOCK
    
    ' Vérifie si NUM LOCK est désactivé
    If Not (GetKeyState(VK_NUMLOCK) And 1) Then
        'Si désactivé, on l'active en envoyant {NUMLOCK}
        Application.SendKeys "{NUMLOCK}", True
        MsgBox "NUM LOCK a été activé."
    Else
        MsgBox "NUM LOCK est déjà activé."
    End If
End Sub

Sub VerifierNumLockAvecGetKeyboardState()

    Dim keyState(255) As Byte
    
    'Appel de l'API pour récupérer l'état des touches du clavier
    GetKeyboardState keyState(0)
    
    ' La touche NUM LOCK est représentée par l'index 144 (0x90)
    If keyState(&H90) And 1 = 0 Then
        MsgBox "Le NUM LOCK est désactivé.", vbInformation
        'Reactive le NUM LOCK
        Application.SendKeys "{NUMLOCK}", True
    End If
    
End Sub

Sub Handle_Rubberduck_Reference()

    Dim ref As Object

    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
        On Error Resume Next 'In case the reference doesn't exist
        For Each ref In ThisWorkbook.VBProject.References
            If ref.name = "Rubberduck Addin" Then 'Rubberduck is the name of the reference to remove
                ThisWorkbook.VBProject.References.Remove ref
                Exit For
            End If
        Next ref
        On Error GoTo 0
    End If

End Sub

Attribute VB_Name = "modAppli"
Option Explicit

#If VBA7 Then
    'D�claration pour les environnements 64 bits
'    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
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
    fDebTFirst = 1
    fDebTNo_Entr�e = fDebTFirst
    fDebTDate
    fDebTType
    fDebTBeneficiaire
    fDebTReference
    fDebTNo_Compte
    fDebTCodeTaxe
    fDebTTOTAL
    fDebTTPS
    fDebTTVQ
    fDebTCr�dit_TPS
    fDebTCr�dit_TVQ
    fDebTAutreRemarque
    fDebTTimeStamp
    fDebTLast = fDebTTimeStamp
End Enum

Public Enum FAC_Ent�te_Data_Columns
    fFacEntFirst = 1
    fFacEntInv_No = fFacEntFirst
    fFacEntDate_Facture
    fFacEntFouP
    fFacEntCust_ID
    fFacEntContact
    fFacEntNom_Client
    fFacEntAdresse1
    fFacEntAdresse2
    fFacEntAdresse3
    fFacEntHonoraires
    fFacEntAF1_Desc
    fFacEntAutresFrais_1
    fFacEntAF2_Desc
    fFacEntAutresFrais_2
    fFacEntAF3_Desc
    fFacEntAutresFrais_3
    fFacEntTaux_TPS
    fFacEntMnt_TPS
    fFacEntTaux_TVQ
    fFacEntMntTVQ
    fFacEntAR_Total
    fFacEntD�p�t
    fFacEntLast = fFacEntD�p�t
End Enum

Public Enum GL_EJ_Recurrente_Data_Columns
    fGLEJrFirst = 1
    fGLEJrNo_EJA = fGLEJrFirst
    fGLEJrDescription
    fGLEJrNo_Compte
    fGLEJrCompte
    fGLEJrD�bit
    fGLEJrCr�dit
    fGLEJrAutreRemarque
    fGLEJrLast = fGLEJrAutreRemarque
End Enum

Public Enum GL_Trans_Data_Columns
    fGLtFirst = 1
    fGLtEntryNo = fGLtFirst
    fGLtDate
    fGLtDescr
    fGLtSource
    fGLtGLNo
    fGLtCompte
    fGLtdt
    fGLtct
    fGLtRem
    fGLtTStamp
    fGLtLast = fGLtTStamp
End Enum

Public Enum TEC_Data_Columns
    fTEClFirst = 1
    fTEClTEC_ID = fTEClFirst
    fTEClProf_ID
    fTEClProf
    fTEClDate
    fTEClClient_ID
    fTEClClientNom
    fTEClDescription
    fTEClHeures
    fTEClCommentaireNote
    fTEClEstFacturable
    fTEClDateSaisie
    fTEClEstFacturee
    fTEClDateFacturee
    fTEClEstDetruit
    fTEClVersionApp
    fTEClNoFacture
    fTEClLast = fTEClNoFacture
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
        .Font.color = vbBlue
        .value = "'" & CStr("Heure - " & Format$(Now(), wshAdmin.Range("B1").value & " hh:mm:ss"))
    End With
    
    With wshMenu.Range("$A$31")
        .Font.size = 8
        .Font.color = vbBlack
        .value = "'" & CStr("Version - " & ThisWorkbook.Name)
    End With
    
    With wshMenu.Range("$A$32")
        .Font.size = 8
        .Font.color = vbBlack
        .value = "'" & CStr("Utilisateur - " & Fn_Get_Windows_Username)
    End With
    
'    Dim env As String: env = wshAdmin.Range("F5").value
    With wshMenu.Range("$A$33")
        .Font.size = 8
        .Font.color = vbRed
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

'CommentOut - 2024-11-14
'Sub Handle_Rubberduck_Reference()
'
'    Dim ref As Object
'
'    If Fn_Get_Windows_Username <> "Robert M. Vigneault" Then
'        On Error Resume Next 'In case the reference doesn't exist
'        For Each ref In ThisWorkbook.VBProject.References
'            If ref.Name = "Rubberduck Addin" Then 'Rubberduck is the name of the reference to remove
'                ThisWorkbook.VBProject.References.Remove ref
'                Exit For
'            End If
'        Next ref
'        On Error GoTo 0
'    End If
'
'    'Lib�rer la m�moire
'    Set ref = Nothing
'
'End Sub
'

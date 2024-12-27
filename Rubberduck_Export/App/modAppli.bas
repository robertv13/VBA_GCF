Attribute VB_Name = "modAppli"
Option Explicit

#If VBA7 Then
    'Déclaration pour les environnements 64 bits
'    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'    Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
'    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#Else
    '32-bit Excel (anciennes versions)
'    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'    Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If

'Public Const VK_NUMLOCK As Long = &H90

Public Const NB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18
Public Const HIGHLIGHT_COLOR As String = &HCCFFCC 'Light green (Pastel Green)
Public Const BASIC_COLOR As Long = 16777215 '2024-07-23 @ 08:15

Public Const DATA_PATH As String = "\DataFiles"
Public Const FACT_PDF_PATH As String = "\Factures_PDF"
Public Const FACT_EXCEL_PATH As String = "\Factures_Excel"

'Variable utilisée pour éviter l'évènement Activate à chaque fois que l'on revient dans une feuille
Public fromMenu As Boolean '2024-09-03 @ 06:14

'Niveau de détail pour le log de SaisieHeures
Public logSaisieHeuresVeryDetailed As Boolean

'Pour assurer un contrôle dans Facture Finale
Public flagEtapeFacture As Integer

'Using Enum to specify the column number of worksheets (data)
Public Enum BD_Clients '2024-10-26 @ 17:41
    [_First] = 1
    fClntFMClientNom = [_First]
    fClntFMClientID
    fClntFMNomClientSystème
    fClntFMContactFacturation
    fClntFMTitreContactFacturation
    fClntFMCourrielFacturation
    fClntFMAdresse1
    fClntFMAdresse2
    fClntFMVille
    fClntFMProvince
    fClntFMCodePostal
    fClntFMPays
    fClntFMRéféréPar
    fClntFMFinAnnée
    fClntFMComptable
    fClntFMNotaireAvocat
    fClntFMTimeStamp
    [_Last]
End Enum

Public Enum BD_Fournisseurs '2024-12-24 @ 07:34
    [_First] = 1
    fFourFMNomFournisseur = [_First]
    fFourFMFournID
    fFourFMContact
    fFourFMCourrielContact
    fFourFMAdresse1
    fFourFMAdresse2
    fFourFMVille
    fFourFMProvince
    fFourFMCodePostal
    fFourFMPays
    fFourFMNoTPS
    fFourFMNoTVQ
    [_Last]
End Enum

Public Enum DEB_Récurrent
    [_First] = 1
    fDebRNoDebRec = [_First]
    fDebRDate
    fDebRType
    fDebRBeneficiaire
    fDebRReference
    fDebRNoCompte
    fDebRCompte
    fDebRCodeTaxe
    fDebRTotal
    fDebRTPS
    fDebRTVQ
    fDebRCréditTPS
    fDebRCréditTVQ
    [_Last]
End Enum

Public Enum DEB_Trans
    [_First] = 1
    fDebTNoEntrée = [_First]
    fDebTDate
    fDebTType
    fDebTBeneficiaire
    fDebTFournID
    fDebTDescription
    fDebTReference
    fDebTNoCompte
    fDebTCompte
    fDebTCodeTaxe
    fDebTTotal
    fDebTTPS
    fDebTTVQ
    fDebTCréditTPS
    fDebTCréditTVQ
    fDebTDépense
    fDebTAutreRemarque
    fDebTTimeStamp
    [_Last]
End Enum

Public Enum ENC_Détails
    [_First] = 1
    fEncDPayID = [_First]
    fEncDInvNo
    fEncDCustomer
    fEncDPayDate
    fEncDPayAmount
    [_Last]
End Enum

Public Enum ENC_Entête
    [_First] = 1
    fEncEPayID = [_First]
    fEncEPayDate
    fEncECustomer
    fEncECodeClient
    fEncEPayType
    fEncEAmount
    fEncENotes
    [_Last]
End Enum

Public Enum FAC_Comptes_Clients
    [_First] = 1
    fFacCCInvNo = [_First]
    fFacCCInvoiceDate
    fFacCCCustomer
    fFacCCCodeClient
    fFacCCStatus
    fFacCCTerms
    fFacCCDueDate
    fFacCCTotal
    fFacCCTotalPaid
    fFacCCBalance
    fFacCCDaysOverdue
    [_Last]
End Enum

Public Enum FAC_Détails
    [_First] = 1
    fFacDInvNo = [_First]
    fFacDDescription
    fFacDHeures
    fFacDTaux
    fFacDHonoraires
    fFacDInvRow
    [_Last]
End Enum

Public Enum FAC_Entête
    [_First] = 1
    fFacEInvNo = [_First]
    fFacEDateFacture
    fFacEACouC
    fFacECustID
    fFacEContact
    fFacENomClient
    fFacEAdresse1
    fFacEAdresse2
    fFacEAdresse3
    fFacEHonoraires
    fFacEAF1Desc
    fFacEAutresFrais1
    fFacEAF2Desc
    fFacEAutresFrais2
    fFacEAF3Desc
    fFacEAutresFrais3
    fFacETauxTPS
    fFacEMntTPS
    fFacETauxTVQ
    fFacEMntTVQ
    fFacEARTotal
    fFacEDépôt
    [_Last]
End Enum

Public Enum FAC_Projets_Détails
    [_First] = 1
    fFacPDProjetID = [_First]
    fFacPDNomClient
    fFacPDClientID
    fFacPDTECID
    fFacPDProfID
    fFacPDDate
    fFacPDProf
    fFacPDHeures
    fFacPDestDetruite
    fFacPDTimeStamp
    [_Last]
End Enum

Public Enum FAC_Projets_Entête
    [_First] = 1
    fFacPEProjetID = [_First]
    fFacPENomClient
    fFacPEClientID
    fFacPEDate
    fFacPEHonoTotal
    fFacPEProf1
    fFacPEHres1
    fFacPETauxH1
    fFacPEHono1
    fFacPEProf2
    fFacPEHres2
    fFacPETauxH2
    fFacPEHono2
    fFacPEProf3
    fFacPEHres3
    fFacPETauxH3
    fFacPEHono3
    fFacPEProf4
    fFacPEHres4
    fFacPETauxH4
    fFacPEHono4
    fFacPEProf5
    fFacPEHres5
    fFacPETauxH5
    fFacPEHono5
    fFacPEestDetruite
    fFacPETimeStamp
    [_Last]
End Enum

Public Enum FAC_Sommaire_Taux
    [_First] = 1
    fFacSTInvNo = [_First]
    fFacSTSéquence
    fFacSTProf
    fFacSTHeures
    fFacSTTaux
    [_Last]
End Enum

Public Enum GL_EJ_Récurrente
    [_First] = 1
    fGlEjRNoEjR = [_First]
    fGlEjRDescription
    fGlEjRNoCompte
    fGlEjRCompte
    fGlEjRDébit
    fGlEjRCrédit
    fGlEjRAutreRemarque
    [_Last]
End Enum

Public Enum GL_Trans
    [_First] = 1
    fGlTNoEntrée = [_First]
    fGlTDate
    fGlTDescription
    fGlTSource
    fGlTNoCompte
    fGlTCompte
    fGlTDébit
    fGlTCrédit
    fGlTAutreRemarque
    fGlTTimeStamp
    [_Last]
End Enum

Public Enum TEC_Local
    [_First] = 1
    fTECTECID = [_First]
    fTECProfID
    fTECProf
    fTECDate
    fTECClientID
    fTECClientNom
    fTECDescription
    fTECHeures
    fTECCommentaireNote
    fTECEstFacturable
    fTECDateSaisie
    fTECEstFacturee
    fTECDateFacturee
    fTECEstDetruit
    fTECVersionApp
    fTECNoFacture
    [_Last]
End Enum

Public Enum TEC_TDB_Data
    [_First] = 1
    fTECTDBTECID = [_First]
    fTECTDBProfID
    fTECTDBProf
    fTECTDBDate
    fTECTDBClientID
    fTECTDBClientNom
    fTECTDBEstClntFact
    fTECTDBH_Saisies
    fTECTDBEstFacturable
    fTECTDBEstFacturee
    fTECTDBEstDetruite
    fTECTDBH_Détruites
    fTECTDBH_ND
    fTECTDBH_Facturables
    fTECTDBH_NonFact
    fTECTDBH_Facturées
    fTECTDBH_TEC
    [_Last]
End Enum

Sub Set_Root_Path(ByRef rootPath As String)
   
    DoEvents
    
    If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        rootPath = "C:\VBA\GC_FISCALITÉ"
    Else
        rootPath = "P:\Administration\APP\GCF"
    End If

End Sub

Sub WriteInfoOnMainMenu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:WriteInfoOnMainMenu", 0)
    
    wshMenu.Unprotect
    
    Application.EnableEvents = False
    
    With wshMenu
        .Range("A30").Value = "Heure - " & Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss")
        .Range("A31").Value = "Version - " & ThisWorkbook.Name
        .Range("A32").Value = "Utilisateur - " & Fn_Get_Windows_Username
        .Range("A33").Value = "Environnement - " & wshAdmin.Range("F5").Value
    End With
    
    Application.EnableEvents = True
    
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Call Log_Record("modAppli:WriteInfoOnMainMenu", startTime)

End Sub

Sub WriteInfoOnMainMenu_OK()

    DoEvents
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:WriteInfoOnMainMenu", 0)
    
    Application.EnableEvents = False
    wshMenu.Unprotect
    Application.ScreenUpdating = True
    
    With wshMenu.Range("$A$30")
        .Font.size = 8
        .Font.color = vbBlue
        .Value = "'" & CStr("Heure - " & Format$(Now(), wshAdmin.Range("B1").Value & " hh:mm:ss"))
    End With
    
    With wshMenu.Range("$A$31")
        .Font.size = 8
        .Font.color = vbBlack
        .Value = "'" & CStr("Version - " & ThisWorkbook.Name)
    End With
    
    With wshMenu.Range("$A$32")
        .Font.size = 8
        .Font.color = vbBlack
        .Value = "'" & CStr("Utilisateur - " & Fn_Get_Windows_Username)
    End With
    
    With wshMenu.Range("$A$33")
        .Font.size = 8
        .Font.color = vbRed
        .Value = "'" & CStr("Environnement - " & wshAdmin.Range("F5").Value)
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    DoEvents '2024-08-23 @ 06:21

    Call Log_Record("modAppli:WriteInfoOnMainMenu", startTime)

End Sub


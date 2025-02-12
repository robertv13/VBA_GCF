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

Public Const DATA_PATH As String = "\DataFiles"
Public Const FACT_PDF_PATH As String = "\Factures_PDF"
Public Const FACT_EXCEL_PATH As String = "\Factures_Excel"

Public Const NB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18

Public Const COULEUR_SAISIE As String = &HCCFFCC 'Light green (Pastel Green)
Public Const COULEUR_BASE_TEC As Long = 6740479
Public Const COULEUR_BASE_FACTURATION As Long = 11854022
Public Const COULEUR_BASE_COMPTABILITÉ As Long = 14277081

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
    fClntFMNomClientPlusNomClientSystème
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
    fFourFMTimeStamp
    [_Last]
End Enum

Public Enum CC_Régularisations
    [_First] = 1
    fREGULRegulID = [_First]
    fREGULInvNo
    fREGULDate
    fREGULClientID
    fREGULClientNom
    fREGULHono
    fREGULFrais
    fREGULTPS
    fREGULTVQ
    fREGULDescription
    fREGULTimeStamp
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
    fDebRTimeStamp
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
    fEncDTimeStamp
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
    fEncETimeStamp
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
    fFacCCTotalRegul
    fFacCCBalance
    fFacCCDaysOverdue
    fFacCCTimeStamp
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
    fFacDTimeStamp
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
    fFacETimeStamp
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
    fFacSTTimeStamp
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
    fGlEjRTimeStamp
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

Sub Auto_Open() '2024-12-28 @ 11:09

    Call CodeEssentielDepart
    
End Sub

Sub CodeEssentielDepart()

    If Application.EnableEvents = False Then Application.EnableEvents = True
    
    On Error GoTo ErrorHandler
    
    ' Réinitialiser les paramètres globaux d'Excel
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'Le serveur est-il disponible ?
    If Fn_Is_Server_Available() = False Then
        MsgBox "Le répertoire (P:\) ne semble pas accessible", vbCritical, "Le serveur n'est pas disponible"
        Application.Quit
    End If
    
    Dim rootPath As String
    Call Set_Root_Path(rootPath)

    Application.EnableEvents = False
    wshAdmin.Range("F5").Value = rootPath
    Application.EnableEvents = True
   
    'Vérification si le chemin est accessible
    If Fn_Check_Server_Access(rootPath) = False Then
        MsgBox "Le répertoire principal (P:\) n'est pas accessible." & vbNewLine & vbNewLine & _
               "Veuillez vérifier votre connexion au serveur SVP", vbCritical, rootPath
        Exit Sub
    End If

    'Log initial activity
    Dim startTime As Double: startTime = Timer: Call Log_Record("----- Début d'une nouvelle session (modAppli:CodeEssentielDepart) -----", "", 0)
    Application.EnableEvents = True
    
    Call Log_Record("Validation d'accès serveur terminée", "", Timer)
    
    'Création d'un fichier qui indique de l'utilisateur utilise l'application
    Call CreateUserActiveFile
    
    Call SetupUserDateFormat
    
    'Call the BackupMasterFile (GCF_BD_MASTER.xlsx) macro at each application startup
    Call BackupMasterFile
    
    Call WriteInfoOnMainMenu
    wshMenu.Range("A1").Value = wshAdmin.Range("NomEntreprise").Value
    
    Call HideDevShapesBasedOnUsername
    
    'Protection de la feuille wshMenu
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells '2024-10-14 @ 11:28
    End With
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
   
    'Efface les feuilles dont le codename n'est pas wsh* -ET- dont le nom commence par 'Feuil'
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If InStr(ws.CodeName, "wsh") <> 1 And InStr(ws.CodeName, "Feuil") = 1 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    wshMenu.Activate

    'Libérer la mémoire
    Set wb = Nothing
    Set ws = Nothing
    
    Call Log_Record("modAppli:CodeEssentielDepart", "", startTime)
    
    Exit Sub
    
ErrorHandler:
    Call Log_Record("Erreur dans modAppli:CodeEssentielDepart : " & Err.Description, Timer)

End Sub

Sub Set_Root_Path(ByRef rootPath As String)
   
    DoEvents
    
    If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        rootPath = "C:\VBA\GC_FISCALITÉ"
    Else
        rootPath = "P:\Administration\APP\GCF"
    End If

End Sub

Sub CreateUserActiveFile()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:CreateUserActiveFile", "", 0)
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    Dim traceFilePath As String
    traceFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
    Dim FileNumber As Integer
    FileNumber = FreeFile
    
    On Error GoTo Error_Handling
    Open traceFilePath For Output As FileNumber
    On Error GoTo 0
    
    Print #FileNumber, "Utilisateur " & userName & " a ouvert l'application à " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    Close FileNumber
    
    Call Log_Record("modAppli:CreateUserActiveFile", "", startTime)
    
    Exit Sub

Error_Handling:
    MsgBox "Erreur en tentant d'accéder le répertoire" & vbNewLine & vbNewLine & _
            "'" & traceFilePath & "'" & vbNewLine & vbNewLine & _
            "Erreur # " & Err.Number & " - " & Err.Description, vbCritical, "Accès à " & traceFilePath

End Sub

Sub SetupUserDateFormat()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:SetupUserDateFormat", "", 0)

    Dim userDateFormat As String
    
    Select Case Fn_Get_Windows_Username
        Case "GuillaumeCharron", "Guillaume"
            userDateFormat = "dd/mm/yy"
        Case "vgervais", "user"
            userDateFormat = "dd/mm/yyyy"
        Case "Annie"
            userDateFormat = "yyyy/mm/dd"
        Case "Robert M. Vigneault", "robertmv"
            userDateFormat = "dd/mm/yyyy"
        Case Else
            userDateFormat = "dd/mm/yyyy"
    End Select

    wshAdmin.Range("B1").Value = userDateFormat
    
    Call Log_Record("modAppli:SetupUserDateFormat", "", startTime)
    
End Sub

Sub BackupMasterFile()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:BackupMasterFile", "", 0)
    
    On Error GoTo MASTER_NOT_AVAILABLE
    
'    Application.ScreenUpdating = False
    
    'Chemin source (fichier principal) et destination (sauvegarde)
    Dim masterFilePath As String
    masterFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
    
    Dim backupFilePath As String
    backupFilePath = wshAdmin.Range("F5").Value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    
    'Créer directement une copie du fichier sans ouvrir Excel
    FileCopy masterFilePath, backupFilePath

    Call Log_Record("modAppli:BackupMasterFile", "", startTime)
    
    Exit Sub
    
MASTER_NOT_AVAILABLE:
    MsgBox "Le fichier GCF_MASTER.xlsx ne peut être accédé..." & vbNewLine & vbNewLine & _
            "Le fichier nécessite une réparation manuelle", _
            vbCritical, _
            "Situation anormale (" & Err.Number & " " & Err.Description & ")"
    Application.Quit

End Sub

Sub WriteInfoOnMainMenu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:WriteInfoOnMainMenu", "", 0)
    
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
    
    Call Log_Record("modAppli:WriteInfoOnMainMenu", "", startTime)

End Sub


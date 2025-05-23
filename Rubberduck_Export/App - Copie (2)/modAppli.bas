Attribute VB_Name = "modAppli"
Option Explicit

'#If VBA7 Then
'    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'#Else
'    Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'#End If
'
Public Const DATA_PATH As String = "\DataFiles"
Public Const FACT_PDF_PATH As String = "\Factures_PDF"
Public Const FACT_EXCEL_PATH As String = "\Factures_Excel"

Public Const NB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18

Public Const COULEUR_SAISIE As String = &HCCFFCC 'Light green (Pastel Green)
Public Const COULEUR_BASE_TEC As Long = 6740479
Public Const COULEUR_BASE_FACTURATION As Long = 11854022
Public Const COULEUR_BASE_COMPTABILIT� As Long = 14277081

'Variable qui contient l'addresse de la derni�re cellule s�lectionn�e
Public gPreviousCellAddress As String

'Variable utilis�e pour �viter l'�v�nement Activate � chaque fois que l'on revient dans une feuille
Public gFromMenu As Boolean '2024-09-03 @ 06:14

'Niveau de d�tail pour le log de SaisieHeures
Public gLogSaisieHeuresVeryDetailed As Boolean

'Pour assurer un contr�le dans Facture Finale
Public gFlagEtapeFacture As Long

'Sauvegarde AUTOMATIQUE du code VBA - 2025-03-03 @ 07:18
Public gNextBackupTime As Date
Public Const INTERVALLE_MINUTES As Double = 10

'Using Enum to specify the column number of worksheets (data)
Public Enum BD_Clients '2024-10-26 @ 17:41
    [_First] = 1
    fClntFMClientNom = [_First]
    fClntFMClientID
    fClntFMNomClientSyst�me
    fClntFMContactFacturation
    fClntFMTitreContactFacturation
    fClntFMCourrielFacturation
    fClntFMAdresse1
    fClntFMAdresse2
    fClntFMVille
    fClntFMProvince
    fClntFMCodePostal
    fClntFMPays
    fClntFMR�f�r�Par
    fClntFMFinAnn�e
    fClntFMComptable
    fClntFMNotaireAvocat
    fClntFMNomClientPlusNomClientSyst�me
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

Public Enum CC_R�gularisations
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

Public Enum DEB_R�current
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
    fDebRCr�ditTPS
    fDebRCr�ditTVQ
    fDebRTimeStamp
    [_Last]
End Enum

Public Enum DEB_Trans
    [_First] = 1
    fDebTNoEntr�e = [_First]
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
    fDebTCr�ditTPS
    fDebTCr�ditTVQ
    fDebTD�pense
    fDebTAutreRemarque
    fDebTTimeStamp
    [_Last]
End Enum

Public Enum ENC_D�tails
    [_First] = 1
    fEncDPayID = [_First]
    fEncDInvNo
    fEncDCustomer
    fEncDPayDate
    fEncDPayAmount
    fEncDTimeStamp
    [_Last]
End Enum

Public Enum ENC_Ent�te
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

Public Enum FAC_D�tails
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

Public Enum FAC_Ent�te
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
    fFacED�p�t
    fFacETimeStamp
    [_Last]
End Enum

Public Enum FAC_Projets_D�tails
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

Public Enum FAC_Projets_Ent�te
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
    fFacSTS�quence
    fFacSTProf
    fFacSTHeures
    fFacSTTaux
    fFacSTTimeStamp
    [_Last]
End Enum

Public Enum GL_EJ_R�currente
    [_First] = 1
    fGlEjRNoEjR = [_First]
    fGlEjRDescription
    fGlEjRNoCompte
    fGlEjRCompte
    fGlEjRD�bit
    fGlEjRCr�dit
    fGlEjRAutreRemarque
    fGlEjRTimeStamp
    [_Last]
End Enum

Public Enum GL_Trans
    [_First] = 1
    fGlTNoEntr�e = [_First]
    fGlTDate
    fGlTDescription
    fGlTSource
    fGlTNoCompte
    fGlTCompte
    fGlTD�bit
    fGlTCr�dit
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
    fTECTDBH_D�truites
    fTECTDBH_ND
    fTECTDBH_Facturables
    fTECTDBH_NonFact
    fTECTDBH_Factur�es
    fTECTDBH_TEC
    [_Last]
End Enum

Private Sub Auto_Open() '2024-12-28 @ 11:09

    Call CodeEssentielDepart
    
End Sub

Sub CodeEssentielDepart()

    If Application.EnableEvents = False Then Application.EnableEvents = True
    
    On Error GoTo ErrorHandler
    
    ' R�initialiser les param�tres globaux d'Excel
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'Le serveur est-il disponible ?
    If Fn_Is_Server_Available() = False Then
        MsgBox "Le r�pertoire (P:\) ne semble pas accessible", vbCritical, "Le serveur n'est pas disponible"
        Application.Quit
    End If
    
    Dim rootPath As String
    rootPath = FN_Get_Root_Path

    Application.EnableEvents = False
    wsdADMIN.Range("F5").value = rootPath
    Application.EnableEvents = True
   
    'V�rification si le chemin est accessible
    If Fn_Check_Server_Access(rootPath) = False Then
        MsgBox "Le r�pertoire principal (P:\) n'est pas accessible." & vbNewLine & vbNewLine & _
               "Veuillez v�rifier votre connexion au serveur SVP", vbCritical, rootPath
        Exit Sub
    End If

    'Log initial activity
    Dim startTime As Double: startTime = Timer: Call Log_Record("----- D�but d'une nouvelle session (modAppli:CodeEssentielDepart) -----", "", 0)
    Application.EnableEvents = True
    
    Call Log_Record("Validation d'acc�s serveur termin�e", "", Timer)
    
    'Cr�ation d'un fichier qui indique de l'utilisateur utilise l'application
    Call CreateUserActiveFile
    
    Call SetupUserDateFormat
    
    'Call the BackupMasterFile (GCF_BD_MASTER.xlsx) macro at each application startup
    Call BackupMasterFile
    
    Call EcrireInformationsConfigAuMenu
    wshMenu.Range("A1").value = wsdADMIN.Range("NomEntreprise").value
    
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

    'Lib�rer la m�moire
    Set wb = Nothing
    Set ws = Nothing
    
    If Fn_Get_Windows_Username = "Robert M. Vigneault" Or Fn_Get_Windows_Username = "robertmv" Then
'        Call ExporterCodeVBA 'Sauvegarde AUTOMATIQUE du code VBA
        Call DemarrerSauvegardeAutomatique
    End If
    
    Call Log_Record("modAppli:CodeEssentielDepart", "", startTime)
    
    Exit Sub
    
ErrorHandler:
    Call Log_Record("Erreur dans modAppli:CodeEssentielDepart : " & Err.Description, Timer)

End Sub

Function FN_Get_Root_Path() As String '2025-03-03 @ 20:28
   
    DoEvents
    
    If Fn_Get_Windows_Username = "Robert M. Vigneault" Then
        FN_Get_Root_Path = "C:\VBA\GC_FISCALIT�"
    Else
        FN_Get_Root_Path = "P:\Administration\APP\GCF"
    End If

End Function

Sub CreateUserActiveFile()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:CreateUserActiveFile", "", 0)
    
    Dim userName As String
    userName = Fn_Get_Windows_Username
    
    Dim traceFilePath As String
    traceFilePath = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & "Actif_" & userName & ".txt"
    
    Dim FileNumber As Long
    FileNumber = FreeFile
    
    On Error GoTo Error_Handling
    Open traceFilePath For Output As FileNumber
    On Error GoTo 0
    
    Print #FileNumber, "Utilisateur " & userName & " a ouvert l'application � " & Format$(Now(), "yyyy-mm-dd hh:mm:ss") & " - Version " & ThisWorkbook.Name
    Close FileNumber
    
    Call Log_Record("modAppli:CreateUserActiveFile", "", startTime)
    
    Exit Sub

Error_Handling:
    MsgBox _
        Prompt:="Erreur en tentant d'acc�der le r�pertoire" & vbNewLine & vbNewLine & _
                    "'" & traceFilePath & "'" & vbNewLine & vbNewLine & _
                    "Erreur # " & Err.Number & " - " & Err.Description, _
        Title:="Acc�s � " & traceFilePath, _
        Buttons:=vbCritical

End Sub

Sub SetupUserDateFormat()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:SetupUserDateFormat", "", 0)

    Dim userDateFormat As String
    
    Select Case Fn_Get_Windows_Username
        Case "GuillaumeCharron", "Guillaume", "gchar"
            userDateFormat = "dd/mm/yy"
        Case "vgervais", "Vlad_Portable", "User", "Oli_Portable"
            userDateFormat = "dd/mm/yyyy"
        Case "Annie"
            userDateFormat = "yyyy/mm/dd"
        Case "Robert M. Vigneault", "robertmv"
            userDateFormat = "dd/mm/yyyy"
        Case Else
            userDateFormat = "dd/mm/yyyy"
    End Select

    wsdADMIN.Range("B1").value = userDateFormat
    
    Call Log_Record("modAppli:SetupUserDateFormat", "", startTime)
    
End Sub

Sub BackupMasterFile()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:BackupMasterFile", "", 0)
    
    On Error GoTo MASTER_NOT_AVAILABLE
    
    'Chemin source (fichier principal) et destination (sauvegarde)
    Dim masterFilePath As String
    masterFilePath = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & "GCF_BD_MASTER.xlsx"
    
    Dim backupFilePath As String
    backupFilePath = wsdADMIN.Range("F5").value & DATA_PATH & Application.PathSeparator & _
                     "GCF_BD_MASTER_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    
    'Cr�er directement une copie du fichier sans ouvrir Excel
    FileCopy masterFilePath, backupFilePath

    Call Log_Record("modAppli:BackupMasterFile", "", startTime)
    
    Exit Sub
    
MASTER_NOT_AVAILABLE:
    MsgBox _
        Prompt:="Le fichier GCF_MASTER.xlsx ne peut �tre acc�d�..." & vbNewLine & vbNewLine & _
                    "Le fichier n�cessite une r�paration manuelle", _
        Title:="Situation anormale (" & Err.Number & " " & Err.Description & ")", _
        Buttons:=vbCritical
'    msgBox "Le fichier GCF_MASTER.xlsx ne peut �tre acc�d�..." & vbNewLine & vbNewLine & _
'            "Le fichier n�cessite une r�paration manuelle", _
'            vbCritical, _
'            "Situation anormale (" & Err.Number & " " & Err.Description & ")"
    Application.Quit

End Sub

Sub EcrireInformationsConfigAuMenu()

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:EcrireInformationsConfigAuMenu", "", 0)
    
    wshMenu.Unprotect
    
    Application.EnableEvents = False
    
    With wshMenu
        .Range("A30").value = "Heure - " & Format$(Now(), wsdADMIN.Range("B1").value & " hh:mm:ss")
        .Range("A31").value = "Version - " & ThisWorkbook.Name
        .Range("A32").value = "Utilisateur - " & Fn_Get_Windows_Username
        .Range("A33").value = "Environnement - " & wsdADMIN.Range("F5").value
        .Range("A34").value = "Format de la date - " & wsdADMIN.Range("B1").value
    End With
    
    Application.EnableEvents = True
    
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
    
    Call Log_Record("modAppli:EcrireInformationsConfigAuMenu", "", startTime)

End Sub



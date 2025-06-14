Attribute VB_Name = "modAppli"
Option Explicit

Public Const DATA_PATH As String = "\DataFiles"
Public Const FACT_PDF_PATH As String = "\Factures_PDF"
Public Const FACT_EXCEL_PATH As String = "\Factures_Excel"

Public Const NB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18

Public Const COULEUR_SAISIE As String = &HCCFFCC 'Light green (Pastel Green)
Public Const COULEUR_BASE_TEC As Long = 6740479
Public Const COULEUR_BASE_FACTURATION As Long = 11854022
Public Const COULEUR_BASE_COMPTABILIT� As Long = 14277081

'Variable qui contient le code d'utilisateur Windows
Public gUtilisateurWindows As String

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
Public Const INTERVALLE_MINUTES_SAUVEGARDE As Double = 10

'Fermeture AUTOMATIQUE de l'application - 2025-05-30 @ 11:36
Public gDerniereActivite As Date
Public gProchaineVerification As Date
Public Const FREQUENCE_VERIFICATION_INACTIVITE As Long = 5
Public Const INTERVALLE_MAXIMUM_INACTIVITE As Long = 15

'Pour capturer �v`nements sur tous les controls des userForm - 2025-05-30 @ 13:11
Public colWrappers As Collection

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

    gDerniereActivite = Now
'    Debug.Print "L'application a d�marr� � " & gDerniereActivite
    gProchaineVerification = Now + TimeSerial(0, FREQUENCE_VERIFICATION_INACTIVITE, 0)
'    Debug.Print "La prochaine v�rification est pr�vue � " & gProchaineVerification
    Application.OnTime gProchaineVerification, "VerifierInactivite"

    Call DemarrageApplication
    
End Sub

Sub DemarrageApplication() '2025-06-06 @ 11:40

    Dim startTime As Double: startTime = Timer: Call Log_Record("----- D�BUT D'UNE NOUVELLE SESSION (modAppli:DemarrageApplication) -----", "", 0)
    
    'Quel est l'utilisateur Windows ?
    gUtilisateurWindows = GetNomUtilisateur()
    Debug.Print "DemarrageApplication - GetNomUtilisateur() = " & gUtilisateurWindows
    
    On Error GoTo ErrorHandler
    
    If Application.EnableEvents = False Then Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Dim rootPath As String
    rootPath = FN_Get_Root_Path

    Application.EnableEvents = False
    wsdADMIN.Range("F5").value = rootPath
    Application.EnableEvents = True
   
    Application.StatusBar = "V�rification de l'acc�s au r�pertoire principal"
    If Fn_Check_Server_Access(rootPath) = False Then
        MsgBox "Le r�pertoire principal '" & rootPath & "' n'est pas accessible." & vbNewLine & vbNewLine & _
               "Veuillez v�rifier votre connexion au serveur SVP", vbCritical, rootPath
        Exit Sub
    End If
    Application.StatusBar = False

    Call CreateUserActiveFile(gUtilisateurWindows)
    Call SetupUserDateFormat(gUtilisateurWindows)
    Call BackupMasterFile
    Call EcrireInformationsConfigAuMenu(gUtilisateurWindows)
    wshMenu.Range("A1").value = wsdADMIN.Range("NomEntreprise").value
    Call HideDevShapesBasedOnUsername(gUtilisateurWindows)
    
    'Protection de la feuille wshMenu
    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
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

    If gUtilisateurWindows = "RobertMV" Or gUtilisateurWindows = "robertmv" Then
        Call DemarrerSauvegardeAutomatique
    End If
    
    'Lib�rer la m�moire
    Set wb = Nothing
    Set ws = Nothing
    
    Call Log_Record("modAppli:DemarrageApplication", "", startTime)
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True 'On s'assure de toujours restaurer l'�tat
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Call Log_Record("modAppli:DemarrageApplication (ERREUR) : " & Err.Description, Timer)
    
End Sub

Function FN_Get_Root_Path() As String '2025-03-03 @ 20:28
   
    DoEvents
    
    If GetNomUtilisateur() = "RobertMV" Or GetNomUtilisateur() = "robertmv" Then
        FN_Get_Root_Path = "C:\VBA\GC_FISCALIT�"
    Else
        FN_Get_Root_Path = "P:\Administration\APP\GCF"
    End If

End Function

Sub CreateUserActiveFile(ByVal userName As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:CreateUserActiveFile", "", 0)
    
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

Sub SetupUserDateFormat(ByVal user As String)

    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:SetupUserDateFormat", "", 0)

    Dim userDateFormat As String
    
    Select Case user
        Case "GuillaumeCharron", "Guillaume", "gchar"
            userDateFormat = "dd/mm/yy"
        Case "vgervais", "Vlad_Portable", "User", "Oli_Portable"
            userDateFormat = "dd/mm/yyyy"
        Case "Annie"
            userDateFormat = "yyyy/mm/dd"
        Case "RobertMV", "robertmv"
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
    Application.Quit

End Sub

Sub EcrireInformationsConfigAuMenu(ByVal user As String)
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:EcrireInformationsConfigAuMenu", "", 0)
    
    Dim oldEnableEvents As Boolean
    Dim heure As String, version As String, utilisateur As String
    Dim environnement As String, formatDate As String
    Dim valeurs As Variant
    
    oldEnableEvents = Application.EnableEvents
    On Error GoTo CleanUp

    wshMenu.Unprotect

    Application.EnableEvents = False

    ' R�cup�ration des valeurs
    formatDate = wsdADMIN.Range("B1").value
    environnement = wsdADMIN.Range("F5").value

    valeurs = Array( _
        "Heure - " & Format$(Now(), formatDate & " hh:mm:ss"), _
        "Version - " & ThisWorkbook.Name, _
        "Utilisateur - " & user, _
        "Environnement - " & environnement, _
        "Format de la date - " & formatDate)

    ' Ecriture en une seule op�ration
    wshMenu.Range("A30:A34").value = Application.WorksheetFunction.Transpose(valeurs)

    With wshMenu
        .Protect UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With

CleanUp:
    Application.EnableEvents = oldEnableEvents
    
    Call Log_Record("modAppli:EcrireInformationsConfigAuMenu", "", startTime)
    
End Sub

'CommentOut - 2025-06-06 @ 11:09
'Sub EcrireInformationsConfigAuMenu(ByVal user As String)
'
''    Dim startTime As Double: startTime = Timer: Call Log_Record("modAppli:EcrireInformationsConfigAuMenu", "", 0)
'
'    wshMenu.Unprotect
'
'    Application.EnableEvents = False
'
'    With wshMenu
'        .Range("A30").value = "Heure - " & Format$(Now(), wsdADMIN.Range("B1").value & " hh:mm:ss")
'        .Range("A31").value = "Version - " & ThisWorkbook.Name
'        .Range("A32").value = "Utilisateur - " & user
'        .Range("A33").value = "Environnement - " & wsdADMIN.Range("F5").value
'        .Range("A34").value = "Format de la date - " & wsdADMIN.Range("B1").value
'    End With
'
'    Application.EnableEvents = True
'
'    With wshMenu
'        .Protect UserInterfaceOnly:=True
'        .EnableSelection = xlUnlockedCells
'    End With
'
''    Call Log_Record("modAppli:EcrireInformationsConfigAuMenu", "", startTime)
'
'End Sub
'
'Mettre � jour � chaque activit�
Public Sub RafraichirActivite(Optional ByVal msg As String = "") '2025-05-30 @ 12:22

'    Debug.Print "Activit� d�tect�e � " & Format(gDerniereActivite, "hh:nn:ss") & " - " & msg
    gDerniereActivite = Now
    Application.StatusBar = False
    
End Sub

'V�rifie l'inactivit� et ferme si plus de 60 minutes
Public Sub VerifierInactivite() '2025-05-30 @ 12:22

    On Error GoTo GestionErreur
    
    Dim heureActuelle As Double
    heureActuelle = Time
    
    'V�rifier si on est dans la plage 18:00 � 23:59
    If heureActuelle < TimeValue("06:00:00") Then
        'Replanifier tout de m�me la prochaine v�rification
        gProchaineVerification = Now + TimeSerial(0, FREQUENCE_VERIFICATION_INACTIVITE, 0)
        Application.OnTime gProchaineVerification, "VerifierInactivite"
        Exit Sub
    End If
    
    If gDerniereActivite = 0 Then
        Debug.Print "gDerniereActivite n'est pas initialis�e..."
        Exit Sub
    End If
    
    'D�terminer le moment pr�cis la derni�re activit� en minutes
    Dim minutesInactive As Double
    minutesInactive = (Now - gDerniereActivite) * 24 * 60 'Convertir en minutes
    Application.StatusBar = "Aucune activit� depuis " & Format(minutesInactive, "#0") & " minute(s)"

    If minutesInactive >= INTERVALLE_MAXIMUM_INACTIVITE Then
        If Not ApplicationIsActive Then
            Application.DisplayAlerts = False
            Call ApplicationFermetureNormale(GetNomUtilisateur())
        End If
    End If

'    If GetNomUtilisateur() <> "RobertMV" And GetNomUtilisateur() <> "Robertmv" Then
'        If minutesInactive >= INTERVALLE_MAXIMUM_INACTIVITE Then
'            Application.DisplayAlerts = False
'            Call ApplicationFermetureNormale
'        End If
'    End If
    
    'Reprogrammer la v�rification
    gProchaineVerification = Now + TimeSerial(0, FREQUENCE_VERIFICATION_INACTIVITE, 0) 'V�rifie toutes les 5 minutes
    Application.OnTime gProchaineVerification, "VerifierInactivite"

    Exit Sub
    
GestionErreur:
    Debug.Print "Erreur dans proc�dure 'VerifierInactivite' : " & Err.Number & " - " & Err.Description
    
End Sub

Private Sub ConnectControlsRecursive(ctrls As MSForms.Controls) '2025-05-30 @ 13:12

    Dim ctrl As MSForms.Control
    For Each ctrl In ctrls
        Debug.Print "Contr�le : " & ctrl.Name & " - Type : " & TypeName(ctrl)

        Select Case TypeName(ctrl)
            Case "Frame", "TabStrip"
                ConnectControlsRecursive ctrl.Controls 'R�cursif pour atteindre tous les niveaux
            Case "MultiPage"
                Dim i As Integer
                For i = 0 To ctrl.Pages.count - 1
                    ConnectControlsRecursive ctrl.Pages(i).Controls
                Next i
            Case "Label"
                'Ignorer les labels (contr�les passifs)
            Case Else
                On Error Resume Next
                Dim wrapper As New clsControlWrapper
                Set wrapper.ctrl = ctrl
                colWrappers.Add wrapper, ctrl.Name
                On Error GoTo 0
        End Select
    Next ctrl
    
End Sub



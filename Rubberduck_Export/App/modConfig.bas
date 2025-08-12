Attribute VB_Name = "modConfig"
'=======================================================
' modConfig.bas - 2025-07-02 @ 09:49
' Déclaration des constantes et des variables globales
'             ainsi que les ENum pour colonnes de tables
'=======================================================

Option Explicit

Public Const gVERSION_APPLICATION As String = "7.B.07"

Public Const gDATA_PATH As String = "\DataFiles"
Public Const gFACT_PDF_PATH As String = "\Factures_PDF"
Public Const gFACT_EXCEL_PATH As String = "\Factures_Excel"

Public Const gNB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18

Public Const gCOULEUR_SAISIE As String = &HCCFFCC 'Light green (Pastel Green)
Public Const gCOULEUR_BASE_TEC As Long = 6740479
Public Const gCOULEUR_BASE_FACTURATION As Long = 11854022
Public Const gCOULEUR_BASE_COMPTABILITE As Long = 14277081

'Variable qui contient le code d'utilisateur Windows
Public gUtilisateurWindows As String

'Variable qui contient l'addresse de la dernière cellule sélectionnée
Public gPreviousCellAddress As String

'Variable utilisée pour éviter l'évènement Activate à chaque fois que l'on revient dans une feuille
Public gFromMenu As Boolean '2024-09-03 @ 06:14

'Niveau de détail pour le log de SaisieHeures
Public gLogSaisieHeuresVeryDetailed As Boolean

'Pour assurer un contrôle dans Facture Finale
Public gFlagEtapeFacture As Long

'Sauvegarde AUTOMATIQUE du code VBA - 2025-03-03 @ 07:18
Public gNextBackupTime As Date
Public Const INTERVALLE_MINUTES_SAUVEGARDE As Double = 15

'Fermeture AUTOMATIQUE de l'application - 2025-07-02 @ 09:47
Public gDerniereActivite As Date
Public gProchaineVerification As Date
Public gFermeturePlanifiee As Date
Public Const gFREQUENCE_VERIFICATION_INACTIVITE As Long = 15
Public Const gMAXIMUM_MINUTES_INACTIVITE As Long = 60
Public Const gDELAI_GRACE_SECONDES As Long = 300
Public Const gHEURE_DEBUT_SURVEILLANCE As Long = 20

'On affiche ou pas certains Debug.print (mécanisme de fermeture automatique de l'application
Public Const gMODE_DEBUG As Boolean = True

'Pour capturer évènements sur tous les controls des userForm - 2025-05-30 @ 13:11
Public colWrappers As Collection

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
    [_Last]
End Enum

Public Enum CC_Regularisations
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

Public Enum DEB_Recurrent
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

Public Enum ENC_Details
    [_First] = 1
    fEncDPayID = [_First]
    fEncDInvNo
    fEncDCustomer
    fEncDPayDate
    fEncDPayAmount
    fEncDTimeStamp
    [_Last]
End Enum

Public Enum ENC_Entete
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

Public Enum FAC_Details
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

Public Enum FAC_Entete
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

Public Enum FAC_Projets_Details
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

Public Enum FAC_Projets_Entete
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

Public Enum GL_EJ_Recurrente
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

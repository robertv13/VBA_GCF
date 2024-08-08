Attribute VB_Name = "modAppli"
'@Folder(GC_FISCALITÉ.Main)

Option Explicit

Public Const APP_VERSION_NO As String = "v4.C.1.xlsb" '2024-08-07 @ 17:03
Public Const NB_MAX_LIGNE_FAC As Long = 35 '2024-06-18 @ 12:18
Public Const HIGHLIGHT_COLOR As String = &HCCFFCC 'Light green (Pastel Green)
Public Const BASIC_COLOR As Long = 16777215 '2024-07-23 @ 08:15
Public Const MAXWIDTH As Long = 192

Public Const DATA_PATH As String = "\DataFiles"
Public Const FACT_PDF_PATH As String = "\Factures_PDF"
Public Const FACT_EXCEL_PATH As String = "\Factures_Excel"

Public userName As String

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

Sub Set_Root_Path()

    Dim rootPath As String
    
    If Not Environ("username") = "Robert M. Vigneault" Then
        rootPath = "P:\Administration\APP\GCF"
    Else
        rootPath = "C:\VBA\GC_FISCALITÉ"
    End If

    wshAdmin.Range("F5").value = rootPath 'Évite de perdre la valeur de la variable wshAdmin.Range("F5").value

End Sub


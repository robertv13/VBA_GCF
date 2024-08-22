Attribute VB_Name = "modFunctions"
Option Explicit

Function Fn_Does_Client_Code_Exist() As Boolean

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Fn_Does_Client_Code_Exist", "", 0)
    
    Fn_Does_Client_Code_Exist = False
    
    Dim ws As Worksheet: Set ws = wshClients
    Dim iCodeClient As String
    iCodeClient = frmForm.txtCodeClient.Value
    If iCodeClient = "" Then
        GoTo Clean_Exit
    End If
    
    'Validating Duplicate Entries
    If Not ws.Range("B:B").Find(what:=iCodeClient, lookat:=xlWhole) Is Nothing Then
        Call Log_Record("modMain:Fn_Does_Client_Code_Exist", "VRAI", -1)
        Fn_Does_Client_Code_Exist = True
    Else
        Call Log_Record("modMain:Fn_Does_Client_Code_Exist", "FAUX", -1)
    End If

Clean_Exit:

    Call Log_Record("modMain:Fn_Does_Client_Code_Exist", "", startTime)
    
    Exit Function

End Function

Function Fn_Fix_Txt_Fin_Annee(fyem As String) As String

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Fn_Fix_Txt_Fin_Annee", "", 0)
    
    'Add the last day of the month to Fiscal Year end month
    Dim fiscalYearEndString As String
    Select Case fyem
        Case "Janvier"
            fiscalYearEndString = "31/01"
        Case "Février"
            fiscalYearEndString = "28/02"
        Case "Mars"
            fiscalYearEndString = "31/03"
        Case "Avril"
            fiscalYearEndString = "30/04"
        Case "Mai"
            fiscalYearEndString = "31/05"
        Case "Juin"
            fiscalYearEndString = "30/06"
        Case "Juillet"
            fiscalYearEndString = "31/07"
        Case "Août"
            fiscalYearEndString = "31/08"
        Case "Septembre"
            fiscalYearEndString = "30/09"
        Case "Octobre"
            fiscalYearEndString = "31/10"
        Case "Novembre"
            fiscalYearEndString = "30/11"
        Case "Décembre"
            fiscalYearEndString = "31/12"
        Case Else
            fiscalYearEndString = frmForm.cmbFinAnnee.Value
        End Select
        
    Fn_Fix_Txt_Fin_Annee = fiscalYearEndString

    Call Log_Record("modMain:Fn_Fix_Txt_Fin_Annee", "", startTime)

End Function

Function Fn_Selected_List() As Long

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Fn_Selected_List", "", 0)
    
    Fn_Selected_List = 0
    
    Dim i As Long
    For i = 0 To frmForm.lstDonnées.ListCount - 1
        If frmForm.lstDonnées.Selected(i) = True Then
            Fn_Selected_List = i + 1
            frmForm.cmdEdit.Enabled = True
            Exit For
        End If
        frmForm.cmdEdit.Enabled = False
    Next i

    Call Log_Record("modMain:Fn_Selected_List", "", startTime)

End Function

Function Fn_ValidateEntries() As Boolean

    Dim startTime As Double: startTime = Timer: Call Log_Record("modMain:Fn_ValidateEntries", "", 0)
    
    Fn_ValidateEntries = True
    
    Dim sh As Worksheet: Set sh = ThisWorkbook.Sheets("Données")

    Dim iCodeClient As Variant
    iCodeClient = frmForm.txtCodeClient.Value
    
    With frmForm
        'Default Color
        .txtCodeClient.BackColor = vbWhite
        .txtNomClient.BackColor = vbWhite
        .txtContactFact.BackColor = vbWhite
        .txtTitreContact.BackColor = vbWhite
        .txtCourrielFact.BackColor = vbWhite
        .txtAdresse1.BackColor = vbWhite
        .txtAdresse2.BackColor = vbWhite
        .txtVille.BackColor = vbWhite
        .txtProvince.BackColor = vbWhite
        .txtCodePostal.BackColor = vbWhite
        .txtPays.BackColor = vbWhite
        .txtReferePar.BackColor = vbWhite
        .txtFinAnnee.BackColor = vbWhite
        .txtComptable.BackColor = vbWhite
        .txtNotaireAvocat.BackColor = vbWhite
        
        'Valeur OBLIGATOIRE
        If Trim(.txtCodeClient.Value) = "" Then
            MsgBox "SVP, saisir un code de client.", vbOKOnly + vbInformation, "Code de client"
            Fn_ValidateEntries = False
            .txtCodeClient.BackColor = vbRed
            .txtCodeClient.Enabled = True
            .txtCodeClient.SetFocus
            GoTo Clean_Exit
        End If
    
        'Valeur OBLIGATOIRE
        If Trim(.txtNomClient.Value) = "" Then
            MsgBox "SVP, saisir le nom du client.", vbOKOnly + vbInformation, "Nom de client"
            Fn_ValidateEntries = False
            .txtNomClient.BackColor = vbRed
            .txtNomClient.SetFocus
            GoTo Clean_Exit
        End If
        
        'Validation de la structure de l'adresse courriel
        If .txtCourrielFact.Value <> "" Then
            If Fn_ValiderCourriel(.txtCourrielFact.Value) = False Then
                MsgBox "SVP, saisir une adresse courriel valide.", vbOKOnly + vbInformation, "Structure d'adresse courriel non-respecté"
                Fn_ValidateEntries = False
                .txtCourrielFact.BackColor = vbRed
                .txtCourrielFact.SetFocus
                GoTo Clean_Exit
            End If
        End If
        
    End With

Clean_Exit:

    Call Log_Record("modMain:Fn_ValidateEntries", "", startTime)

    Exit Function
    
End Function

Function Fn_ValiderCourriel(ByVal courriel As String) As Boolean
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    'Définir le pattern pour l'expression régulière
    regex.Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    regex.IgnoreCase = True
    regex.Global = False
    
    'Vérifier si l'adresse courriel correspond au pattern
    Fn_ValiderCourriel = regex.Test(courriel)
    
End Function


Attribute VB_Name = "modFunctions"
'@IgnoreModule UnassignedVariableUsage
'@Folder("Gestion_Clients")
Option Explicit

Function Fn_Is_Client_Code_Already_Used() As Boolean

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:Fn_Is_Client_Code_Already_Used", "", 0)
    
    Fn_Is_Client_Code_Already_Used = False
    
    Dim ws As Worksheet: Set ws = wshClients
    Dim iCodeClient As String
    iCodeClient = ufClientMF.txtCodeClient.Value
    If iCodeClient = "" Then
        GoTo Clean_Exit
    End If
    
    'Validating Duplicate Entries
    If Not ws.Range("B:B").Find(What:=iCodeClient, LookAt:=xlWhole) Is Nothing Then
        Call CM_Log_Activities("modMain:Fn_Is_Client_Code_Already_Used", "VRAI", -1)
        Fn_Is_Client_Code_Already_Used = True
    Else
        Call CM_Log_Activities("modMain:Fn_Is_Client_Code_Already_Used", "FAUX", -1)
    End If

Clean_Exit:

    Call CM_Log_Activities("modMain:Fn_Is_Client_Code_Already_Used", "", startTime)
    
    Exit Function

End Function

Function Fn_Fix_Txt_Fin_Annee(fyem As String) As String

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:Fn_Fix_Txt_Fin_Annee", "", 0)
    
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
            fiscalYearEndString = ufClientMF.cmbFinAnnee.Value
        End Select
        
    Fn_Fix_Txt_Fin_Annee = fiscalYearEndString

    Call CM_Log_Activities("modMain:Fn_Fix_Txt_Fin_Annee", "", startTime)

End Function

Function Fn_Get_Windows_Username() As String 'Function to retrieve the Windows username using the API

    Dim buffer As String * 255
    Dim size As Long: size = 255
    
    If GetUserName(buffer, size) Then
        Fn_Get_Windows_Username = Left$(buffer, size - 1)
    Else
        Fn_Get_Windows_Username = "Unknown"
    End If
    
End Function

Function Fn_Incremente_Code(c As String) As String

    'Parcourir le code pour extraire la partie numérique
    Dim i As Integer
    Dim numericPart As String
    Dim suffix As String
    For i = 1 To Len(c)
        If IsNumeric(Mid(c, i, 1)) Then
            numericPart = numericPart & Mid(c, i, 1)
        Else
            'Dès qu'on trouve un caractère non numérique, on considère que c'est le suffixe
            suffix = Mid(c, i)
            Exit For
        End If
    Next i
    
    'Si la partie numérique est valide, on ajoute 1
    Dim newCode As String
    If Len(numericPart) > 0 Then
        newCode = CStr(CLng(numericPart) + 1) ' Convertir la partie numérique en nombre et ajouter 1
    Else
        newCode = "" 'Si aucun chiffre n'a été trouvé, laisser la nouvelle valeur vide
    End If
    
    'Retourne le nouveau code
    Fn_Incremente_Code = newCode

End Function

Function Fn_Selected_List() As Long

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:Fn_Selected_List", "", 0)
    
    Fn_Selected_List = 0
    
    Dim i As Long
    For i = 0 To ufClientMF.lstDonnées.ListCount - 1
        If ufClientMF.lstDonnées.Selected(i) = True Then
            Fn_Selected_List = i + 1
            ufClientMF.cmdEdit.Enabled = True
            Exit For
        End If
        ufClientMF.cmdEdit.Enabled = False
    Next i

    Call CM_Log_Activities("modMain:Fn_Selected_List", "", startTime)

End Function

Function Fn_ValidateEntries() As Boolean

    Dim startTime As Double: startTime = Timer: Call CM_Log_Activities("modMain:Fn_ValidateEntries", "", 0)
    
    Fn_ValidateEntries = True
    
    Dim sh As Worksheet: Set sh = ThisWorkbook.Worksheets("Données")
    
    Dim iCodeClient As Variant
    iCodeClient = ufClientMF.txtCodeClient.Value
    
    With ufClientMF
        'Default Color
        .txtCodeClient.BackColor = vbWhite
        .txtNomClient.BackColor = vbWhite
        .txtNomClientSysteme.BackColor = vbWhite
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
        .txtNomClientPlusNomClientSystème = vbWhite
        
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
        
        'Validation de la structure de l'adresse courriel, si ce n'est pas inconnu
        If .txtCourrielFact.Value <> "" And .txtCourrielFact.Value <> "inconnu" Then
            If Fn_ValiderCourriel(.txtCourrielFact.Value) = False Then
                MsgBox _
                    Prompt:="SVP, saisir une ou des adresses courriel valide(s).", _
                    Title:="Structure d'adresse courriel non-respecté", _
                    Buttons:=vbOKOnly + vbInformation
                Fn_ValidateEntries = False
                .txtCourrielFact.BackColor = vbRed
                .txtCourrielFact.SetFocus
                GoTo Clean_Exit
            Else
                .txtCourrielFact = Fn_NormaliserAdressesCourriel(.txtCourrielFact)
            End If
        End If
    End With

Clean_Exit:

    Call CM_Log_Activities("modMain:Fn_ValidateEntries", "", startTime)

    Exit Function
    
End Function

Function Fn_ValiderCourriel(courriel As String) As Boolean
    
    'Supporte de 0 à 2 courriels (séparés par '; ')
    
    Fn_ValiderCourriel = False
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    'Initialisation de l'expression régulière pour valider une adresse courriel
    With regex
        .Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With
    
'    'Normaliser le champs s'il contient plus d'une adresse courriel
'    courriel = Fn_NormaliserAdressesCourriel(courriel)

    'Diviser le paremètre (courriel) en adresses individuelles
    Dim arrAdresse() As String
    arrAdresse = Split(courriel, "; ")
    
    'Vérifier chaque adresse
    Dim adresse As Variant
    For Each adresse In arrAdresse
        adresse = Trim(adresse)
        'Passer si l'adresse est vide (Aucune adresse est aussi permis)
        If adresse <> "" Then
            'Si l'adresse ne correspond pas au pattern, renvoyer Faux
            If Not regex.Test(adresse) Then
                Dim msgValue As VbMsgBoxResult
                msgValue = MsgBox("'" & courriel & "'" & vbNewLine & vbNewLine & _
                                    "N'est pas structurée selon les standards..." & vbNewLine & vbNewLine & _
                                    "Désirez-vous quand même conserver cette adresse ?", _
                                    vbYesNo + vbInformation, "Structure de courriel non standard")
                If msgValue = vbYes Then
                    Fn_ValiderCourriel = True
                Else
                    Fn_ValiderCourriel = False
                    Exit Function
                End If
            End If
        End If
    Next adresse
    
    'Toutes les adresses sont valides
    Fn_ValiderCourriel = True

    'Nettoyer la mémoire
    Set adresse = Nothing
    Set regex = Nothing
    
End Function

Function Fn_NormaliserAdressesCourriel(ByVal strEmails As String) As String

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    'Expression régulière pour capturer les adresses email
    regex.Pattern = "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}" ' Format d'une adresse courriel
    regex.IgnoreCase = True
    regex.Global = True
    
    Dim matches As Object
    Set matches = regex.Execute(strEmails)
    
    'Construit la chaîne normalisée avec "; " comme séparateur
    Dim resultat As String
    Dim i As Integer
    For i = 0 To matches.Count - 1
        resultat = resultat & matches(i).Value & "; "
    Next i
    
    'Supprime le dernier "; " s'il y a au moins une adresse
    If Len(resultat) > 0 Then resultat = Left(resultat, Len(resultat) - 2)
    
    Fn_NormaliserAdressesCourriel = resultat
    
End Function



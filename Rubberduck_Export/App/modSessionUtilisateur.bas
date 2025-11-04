Attribute VB_Name = "modSessionUtilisateur"
Option Explicit

'Module : modSessionUtilisateur
'Date : 2025-10-19
'Auteur : Robert + Copilot
'Rôle : Initialisation de session utilisateur, chargement des données métier, journalisation

Public UtilisateurActif As Scripting.Dictionary

'--- Étape 1 : Récupération UtilisateurID à partir de l'utilisateur Windows ---
Public Function Fn_InfosWindows() As Object

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication( _
        "modSessionUtilisateur:Fn_InfosWindows", vbNullString, 0)

    Dim cn As Object, rs As Object
    Dim infos As Object
    Set infos = CreateObject("Scripting.Dictionary")

    Dim cheminMASTER As String
    cheminMASTER = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                   "GCF_BD_MASTER.xlsx"
    If Dir(cheminMASTER) = "" Then
        Call modAppli.AfficherErreurCritique("Fichier GCF_BD_MASTER est introuvable" & vbNewLine & vbNewLine & cheminMASTER)
        Exit Function
    End If

    Dim nomWindows As String
    nomWindows = Replace(Fn_UtilisateurWindows(), "'", "''")
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminMASTER & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"

    Set rs = cn.Execute("SELECT * FROM [UtilisateursWindows$] WHERE UtilisateurWindows = '" & Fn_UtilisateurWindows() & "'")

    If Not rs.EOF Then
        infos("UtilisateurID") = rs("UtilisateurID")
    Else
        MsgBox "Utilisateur Windows non reconnu dans GCF_BD_MASTER.", _
            vbCritical
        Call modAppli.AfficherErreurCritique("Utilisateur Windows non reconnu dans GCF_BD_MASTER" & _
            vbNewLine & vbNewLine & Fn_UtilisateurWindows)
    End If

    rs.Close
    cn.Close
    
    Set Fn_InfosWindows = infos
    
    Call modDev_Utils.EnregistrerLogApplication("modSessionUtilisateur:Fn_InfosWindows", vbNullString, startTime)

End Function

' --- Étape 2 : Chargement des données d'utilisateur ---
Public Function Fn_ChargerUtilisateur(utilisateurID As Long) As Scripting.Dictionary

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication( _
        "modSessionUtilisateur:Fn_ChargerUtilisateur", vbNullString, 0)
    
    Dim cn As Object, rs As Object
    Dim user As Scripting.Dictionary
    Set user = New Scripting.Dictionary

    Dim cheminMASTER As String
    cheminMASTER = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                   "GCF_BD_MASTER.xlsx"
    If Dir(cheminMASTER) = "" Then
        Call modAppli.AfficherErreurCritique("Fichier GCF_BD_MASTER est introuvable" & vbNewLine & vbNewLine & cheminMASTER)
        Exit Function
    End If

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminMASTER & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"

    Set rs = cn.Execute("SELECT * FROM [Utilisateurs$] WHERE UtilisateurID = " & utilisateurID)

    If Not rs.EOF Then
        Dim champ As Variant
        For Each champ In rs.Fields
            user(champ.Name) = champ.Value
        Next champ
    Else
        MsgBox "UtilisateurID introuvable dans GCF_BD_MASTER.", _
            vbCritical
        Call modAppli.AfficherErreurCritique("UtilisateurID est introuvable dans GCF_BD_MASTER" & _
            vbNewLine & vbNewLine & utilisateurID)
    End If

    rs.Close
    cn.Close
    
    Set Fn_ChargerUtilisateur = user
    
    Call modDev_Utils.EnregistrerLogApplication("modSessionUtilisateur:Fn_ChargerUtilisateur", vbNullString, startTime)

End Function

'--- Étape 3 : Initialisation complète ---
Public Sub InitialiserSessionUtilisateur()

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication( _
        "modSessionUtilisateur:InitialiserSessionUtilisateur", vbNullString, 0)

    Dim infos As Object
    Set infos = Fn_InfosWindows()

    Set UtilisateurActif = Fn_ChargerUtilisateur(infos("UtilisateurID"))

    If UCase(UtilisateurActif("Actif")) <> "VRAI" Then
        MsgBox "Votre utilisateur n'est pas actif dans l'application.", _
            vbCritical
        Call modAppli.AfficherErreurCritique("Votre utilisateur n'est pas actif dans l'application" & _
            vbNewLine & vbNewLine & UtilisateurActif("Prenom"))
    End If

    Call MettreAJourDerniereConnexion(infos("UtilisateurID"))
    
    Call modDev_Utils.EnregistrerLogApplication("modSessionUtilisateur:InitialiserSessionUtilisateur", vbNullString, startTime)

End Sub

' --- Étape 4 : Journalisation ---
Private Sub MettreAJourDerniereConnexion(utilisateurID As Long)

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication( _
        "modSessionUtilisateur:MettreAJourDerniereConnexion", vbNullString, 0)

    Dim cheminMASTER As String
    cheminMASTER = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                   "GCF_BD_MASTER.xlsx"
    If Dir(cheminMASTER) = "" Then
        Call modAppli.AfficherErreurCritique("Fichier GCF_BD_MASTER est introuvable" & vbNewLine & vbNewLine & cheminMASTER)
        Exit Sub
    End If

    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminMASTER & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"

    Dim sql As String
    sql = "UPDATE [Utilisateurs$] SET DateDernLogin = #" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "# WHERE UtilisateurID = " & utilisateurID
    cn.Execute sql

    cn.Close
    
    Call modDev_Utils.EnregistrerLogApplication("modSessionUtilisateur:MettreAJourDerniereConnexion", vbNullString, startTime)
    
End Sub


Attribute VB_Name = "modSurveillance"
Option Explicit

Public gNomFeuilleDerniereActivite As String
Public gAdresseDerniereSelection As String
Public gDerniereInteraction As Double

Sub LancerSurveillance() '2025-11-06 @ 15:57

    Application.OnTime Now + TimeSerial(0, 5, 0), "modSurveillance.VerifierActivite" 'Aux 5 minutes
    
End Sub

Sub VerifierActivite() '2025-11-06 @ 15:57

    If Hour(Now) < 16 Then Exit Sub 'Pas avant 18:00 heures

    Debug.Print Now() & " VerifierActivite"
    If Not ActiviteDetectee() Then
        Call ProposerFermeture
    Else
        Call LancerSurveillance 'Relance la surveillance
    End If
    
End Sub

Function ActiviteDetectee() As Boolean '2025-11-06 @ 15:57

    ActiviteDetectee = False
    
    Debug.Print Now() & " Test sur activité avec [ActiviteDetectee] ?"
    If ActiveSheet.Name <> gNomFeuilleDerniereActivite Then
        Debug.Print Now() & " La feuille a changé"
        ActiviteDetectee = True
    End If
    
    If Selection.Address <> gAdresseDerniereSelection Then
        Debug.Print Now() & " La sélection a changé"
        ActiviteDetectee = True
    End If
    
    If Timer - gDerniereInteraction < 300 Then 'Moins de 5 min
        Debug.Print Now() & " Activité dans les 5 dernières minutes"
        ActiviteDetectee = True
    End If
    
    If ActiviteDetectee = False Then
        Debug.Print Now() & " Aucune activité selon 'ActiviteDetectee'"
    End If
    
End Function

Sub ProposerFermeture() '2025-11-06 @ 15:57

    Dim choix As VbMsgBoxResult
    choix = MsgBox("Aucune activité détectée depuis 5 minutes." & vbCrLf & _
                   "Souhaitez-vous fermer l'application ?", _
                   vbYesNoCancel + vbCritical, "Fermeture automatique")

    Select Case choix
        Case vbYes: Call FermerApplication
        Case vbNo: Call LancerSurveillance 'Relance après délai
        Case vbCancel: gDerniereInteraction = Timer 'Reset
    End Select
    
End Sub

Sub FermerApplication()  '2025-11-06 @ 15:57

'    Call PurgerObjets
'    Call EnregistrerLogApplication("modSurveillance", "Fermeture automatique", "CRITICAL")
'    Application.Quit
    
End Sub

Public Sub InitialiserSurveillanceForm(frm As Object, ByRef wrappers As Collection) '2025-11-06 @ 16:21

    Set wrappers = New Collection
    Dim ctrl As MSForms.Control
    Dim wrapper As Object

    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                Set wrapper = New clsWrapperTextBox
                Set wrapper.tb = ctrl
            Case "ComboBox"
                Set wrapper = New clsWrapperComboBox
                Set wrapper.cb = ctrl
            Case "CommandButton"
                Set wrapper = New clsWrapperCommandButton
                Set wrapper.btn = ctrl
            Case "ListBox"
                Set wrapper = New clsWrapperListBox
                Set wrapper.lb = ctrl
            Case "CheckBox"
                Set wrapper = New clsWrapperCheckBox
                Set wrapper.chk = ctrl
            Case "OptionButton"
                Set wrapper = New clsWrapperOptionButton
                Set wrapper.opt = ctrl
            Case Else
                Debug.Print "Contrôle ignoré : " & ctrl.Name & " (" & TypeName(ctrl) & ")"
                Set wrapper = Nothing
        End Select
        If Not wrapper Is Nothing Then wrappers.Add wrapper
    Next ctrl
    
End Sub

Public Sub EnregistrerActivite(source As String) '2025-11-06 @ 17:36

    gDerniereInteraction = Timer
    Dim ligne As String
    ligne = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & source & " | Feuille: " & Fn_NomFeuilleActive()
    
    Debug.Print ligne
    Call EcrireDansLog(ligne)
    
End Sub

Public Sub EcrireDansLog(texte As String) '2025-11-07 @ 04:36

    Dim chemin As String
    chemin = ThisWorkbook.path & "\journal_activite.txt"
    
    Dim fso As Object, fichier As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.fileExists(chemin) Then
        Set fichier = fso.OpenTextFile(chemin, 8) 'Append
    Else
        Set fichier = fso.CreateTextFile(chemin, True)
    End If
    
    fichier.WriteLine texte
    fichier.Close
    
End Sub

Public Function Fn_NomFeuilleActive() As String '2025-11-07 @ 04:38

    On Error Resume Next
    Fn_NomFeuilleActive = ActiveSheet.Name
    
End Function


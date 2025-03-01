VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufGL_Rapport 
   Caption         =   "Rapport des transactions du G/L"
   ClientHeight    =   9090.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11190
   OleObjectBlob   =   "ufGL_Rapport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufGL_Rapport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Call GL_Trans_Import_All
    
    'Efface le contenu de la listBox
    Me.lsbComptes.Clear
    
    'Obtenir le plan comptable
    Dim arr As Variant
    arr = Fn_Get_Plan_Comptable(2) 'Returns an array with 2 columns
        
    'Ajoute les comptes, un � un dans la listBox
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Me.lsbComptes.AddItem arr(i, 1) & " " & arr(i, 2)
    Next i

    'Options du comboBox pour le type de rapport
    With Me.cmbTypeRapport
        .Clear
        .AddItem "Par compte / par date"
        .AddItem "Par num�ro d'�criture"
        .ListIndex = 0 ' S�lection par d�faut
    End With
    
    'Options du comboBox pour les p�riodes
    cmbPeriode.Clear
    
    'V�rifier si la plage nomm�e existe
    On Error Resume Next
    Dim plage As Range
    Set plage = Range("dnrDateRange")
    On Error GoTo 0
    
    If Not plage Is Nothing Then
        'Ajouter chaque �l�ment de la plage � la ComboBox
        Dim cellule As Range
        For Each cellule In plage
            cmbPeriode.AddItem cellule.value
        Next cellule
    Else
        MsgBox "La plage nomm�e 'dnrDateRange' est introuvable.", vbExclamation, "Contacter le d�veloppeur"
        Exit Sub
    End If
    
    'Valeur s�lectionn�e par d�faut
    cmbPeriode.ListIndex = 1

    'Lib�rer la m�moire
    Set cellule = Nothing
    Set plage = Nothing
    
End Sub

Private Sub cmbTypeRapport_Change()

    'Masquer/Afficher les contr�les selon le type de rapport
    If cmbTypeRapport.ListIndex = 0 Then
        'Mode "Par compte / par date"
        fraParDate.Visible = True
        fraParEcriture.Visible = False
        'Activer la s�lection de comptes
        lsbComptes.Enabled = True
        cmdSelectAll.Enabled = True
        cmdDeselectAll.Enabled = True
    Else
        'Mode "Par num�ro d'�criture"
        fraParDate.Visible = False
        fraParEcriture.Visible = True
        'D�sactiver la s�lection de comptes (car non applicable)
        lsbComptes.Enabled = False
        cmdSelectAll.Enabled = False
        cmdDeselectAll.Enabled = False
    End If
    
End Sub

Private Sub cmbPeriode_Change()

        Select Case cmbPeriode.value
            Case "Aujourd'hui"
                txtDateDebut.value = Format$(Date, wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(Date, wshAdmin.Range("B1").value)
            Case "Mois Courant"
                txtDateDebut.value = Format$(wshAdmin.Range("MoisDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("MoisA"), wshAdmin.Range("B1").value)
            Case "Mois Dernier"
                txtDateDebut.value = Format$(wshAdmin.Range("MoisPrecDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("MoisPrecA"), wshAdmin.Range("B1").value)
            Case "Trimestre courant"
                txtDateDebut.value = Format$(wshAdmin.Range("TrimDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("TrimA"), wshAdmin.Range("B1").value)
            Case "Trimestre pr�c�dent"
                txtDateDebut.value = Format$(wshAdmin.Range("TrimPrecDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("TrimPrecA"), wshAdmin.Range("B1").value)
            Case "Ann�e courante"
                txtDateDebut.value = Format$(wshAdmin.Range("AnneeDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("AnneeA"), wshAdmin.Range("B1").value)
            Case "Ann�e pr�c�dente"
                txtDateDebut.value = Format$(wshAdmin.Range("AnneePrecDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("AnneePrecA"), wshAdmin.Range("B1").value)
            Case "7 derniers jours"
                txtDateDebut.value = Format$(wshAdmin.Range("SeptJoursDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("SeptJoursA"), wshAdmin.Range("B1").value)
            Case "15 derniers jours"
                txtDateDebut.value = Format$(wshAdmin.Range("QuinzeJoursDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("QuinzeJoursA"), wshAdmin.Range("B1").value)
            Case "Semaine"
                txtDateDebut.value = Format$(wshAdmin.Range("DateDebutSemaine"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("DateFinSemaine"), wshAdmin.Range("B1").value)
            Case "Toutes les dates"
                txtDateDebut.value = Format$(#1/1/2024#, wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("AnneeA"), wshAdmin.Range("B1").value)
            Case Else
                txtDateDebut.value = ""
                txtDateFin.value = ""
        End Select

End Sub

Private Sub cmdGenerer_Click()

    'V�rification que le type de rapport est s�lectionn�
    Dim TypeRapport As String
    If Me.cmbTypeRapport.ListIndex = -1 Then
        MsgBox "Veuillez s�lectionner un type de rapport.", vbExclamation, "Erreur"
        Exit Sub
    Else
        TypeRapport = Me.cmbTypeRapport.value
    End If

    'On efface/cree une feuille pour le rapport
    Dim strWsRapport$
    strWsRapport$ = "X_GL_Rapport"
    Call CreateOrReplaceWorksheet(strWsRapport$)
    Dim wsRapport As Worksheet
    Set wsRapport = ThisWorkbook.Sheets(strWsRapport$)
    
    'V�rification des crit�res selon le type de rapport
    If TypeRapport = "Par compte / par date" Then
        'Validation des dates
        Dim dateDebut As Date, dateFin As Date
        If IsDate(Me.txtDateDebut.value) And IsDate(Me.txtDateFin.value) Then
            dateDebut = CDate(Me.txtDateDebut.value)
            dateFin = CDate(Me.txtDateFin.value)
            If dateDebut > dateFin Then
                MsgBox "La date de d�but doit �tre ant�rieure ou �gale � la date de fin.", vbExclamation, "Erreur dans les crit�res de date"
                Exit Sub
            End If
        Else
            MsgBox "Veuillez entrer des dates valides.", vbExclamation, "Erreur dans les crit�res de date"
            Exit Sub
        End If
        
        'V�rification qu'au moins un compte est s�lectionn�
        Dim ligneS�lectionn�e As Boolean
        ligneS�lectionn�e = EstLigneSelectionnee(Me.lsbComptes)
        If ligneS�lectionn�e = False Then
            MsgBox "Veuillez s�lectionner au moins un num�ro de compte.", vbExclamation, "Erreur"
            Exit Sub
        End If
        
        'Les validations sont termin�es, on appelle la proc�dure pour le rapport par compte
        Call GenererRapportGL_Compte(wsRapport, dateDebut, dateFin)

    Else
        'Validation des num�ros d'�criture
        Dim noEcritureDebut As Long, noEcritureFin As Long
        If IsNumeric(Me.txtNoEcritureDebut.value) And IsNumeric(Me.txtNoEcritureFin.value) Then
            noEcritureDebut = CLng(Me.txtNoEcritureDebut.value)
            noEcritureFin = CLng(Me.txtNoEcritureFin.value)
            
            'V�rification de l'ordre des num�ros d'�criture
            If noEcritureDebut > noEcritureFin Then
                MsgBox "Le num�ro d'�criture de d�but doit �tre inf�rieur ou �gal au num�ro de fin.", _
                            vbExclamation, "Erreur dans les crit�res de num�ro d'�criture"
                Exit Sub
            End If
        Else
            MsgBox "Veuillez entrer des num�ros d'�criture valides.", vbExclamation, _
                        "Erreur dans les crit�res de num�ro d'�criture"
            Exit Sub
        End If
        
        'Les validations sont termin�es, on appelle la proc�dure pour le rapport par compte
        Call GenererRapportGL_Ecriture(wsRapport, noEcritureDebut, noEcritureFin)
    End If

End Sub

Private Sub cmdSelectAll_Click()

    'Boucle � travers tous les �l�ments du ListBox et les s�lectionne
    Dim i As Integer
    For i = 0 To lsbComptes.ListCount - 1
        lsbComptes.Selected(i) = True
    Next i
    
End Sub

Private Sub cmdDeselectAll_Click()

    'Boucle � travers tous les �l�ments du ListBox et les d�s�lectionne
    Dim i As Integer
    For i = 0 To lsbComptes.ListCount - 1
        lsbComptes.Selected(i) = False
    Next i

End Sub



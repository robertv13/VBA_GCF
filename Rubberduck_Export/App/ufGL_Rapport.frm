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
        
    'Ajoute les comptes, un à un dans la listBox
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Me.lsbComptes.AddItem arr(i, 1) & " " & arr(i, 2)
    Next i

    'Options du comboBox pour le type de rapport
    With Me.cmbTypeRapport
        .Clear
        .AddItem "Par compte / par date"
        .AddItem "Par numéro d'écriture"
        .ListIndex = 0 ' Sélection par défaut
    End With
    
    'Options du comboBox pour les périodes
    cmbPeriode.Clear
    
    'Vérifier si la plage nommée existe
    On Error Resume Next
    Dim plage As Range
    Set plage = Range("dnrDateRange")
    On Error GoTo 0
    
    If Not plage Is Nothing Then
        'Ajouter chaque élément de la plage à la ComboBox
        Dim cellule As Range
        For Each cellule In plage
            cmbPeriode.AddItem cellule.value
        Next cellule
    Else
        msgBox "La plage nommée 'dnrDateRange' est introuvable.", vbExclamation, "Contacter le développeur"
        Exit Sub
    End If
    
    'Valeur sélectionnée par défaut
    cmbPeriode.ListIndex = 1

    'Libérer la mémoire
    Set cellule = Nothing
    Set plage = Nothing
    
End Sub

Private Sub cmbTypeRapport_Change()

    'Masquer/Afficher les contrôles selon le type de rapport
    If cmbTypeRapport.ListIndex = 0 Then
        'Mode "Par compte / par date"
        fraParDate.Visible = True
        fraParEcriture.Visible = False
        'Activer la sélection de comptes
        lsbComptes.Enabled = True
        cmdSelectAll.Enabled = True
        cmdDeselectAll.Enabled = True
    Else
        'Mode "Par numéro d'écriture"
        fraParDate.Visible = False
        fraParEcriture.Visible = True
        'Désactiver la sélection de comptes (car non applicable)
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
            Case "Trimestre précédent"
                txtDateDebut.value = Format$(wshAdmin.Range("TrimPrecDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("TrimPrecA"), wshAdmin.Range("B1").value)
            Case "Année courante"
                txtDateDebut.value = Format$(wshAdmin.Range("AnneeDe"), wshAdmin.Range("B1").value)
                txtDateFin.value = Format$(wshAdmin.Range("AnneeA"), wshAdmin.Range("B1").value)
            Case "Année précédente"
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

Private Sub txtDateDebut_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim dateCorrigee As String
    
    If Trim(txtDateDebut.Text) <> "" Then
        dateCorrigee = CorrigerDate(txtDateDebut.Text)
        If dateCorrigee = "" Then
            msgBox "La date saisie est invalide, veuillez saisir une date sous un" & vbNewLine & vbNewLine & _
                    "format valide (jj ou jj/mm ou jj/mm/aaaa ou aaaa/mm/jj)" & vbNewLine & vbNewLine & _
                    "Notez que le séparateur peut être '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpréter la date saisie"
            Cancel = True
            txtDateDebut.SelStart = 0
            txtDateDebut.SelLength = Len(txtDateDebut.Text)
        Else
            txtDateDebut.Text = dateCorrigee
        End If
    End If
End Sub

Private Sub txtDateFin_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim dateCorrigee As String
    
    If Trim(txtDateFin.Text) <> "" Then
        dateCorrigee = CorrigerDate(txtDateFin.Text)
        If dateCorrigee = "" Then
            msgBox "La date saisie est invalide, veuillez saisir une date sous un" & vbNewLine & vbNewLine & _
                    "format valide (jj ou jj/mm ou jj/mm/aaaa ou aaaa/mm/jj)" & vbNewLine & vbNewLine & _
                    "Notez que le séparateur peut être '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpréter la date saisie"
            Cancel = True
            txtDateFin.SelStart = 0
            txtDateFin.SelLength = Len(txtDateFin.Text)
        Else
            txtDateFin.Text = dateCorrigee
        End If
    End If
End Sub

Private Sub cmdGenerer_Click()

    'Vérification que le type de rapport est sélectionné
    Dim TypeRapport As String
    If Me.cmbTypeRapport.ListIndex = -1 Then
        msgBox "Veuillez sélectionner un type de rapport.", vbExclamation, "Erreur"
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
    
    'Vérification des critères selon le type de rapport
    If TypeRapport = "Par compte / par date" Then
        'Validation des dates
        Dim dateDebut As Date, dateFin As Date
        If IsDate(Me.txtDateDebut.value) And IsDate(Me.txtDateFin.value) Then
            dateDebut = CDate(Me.txtDateDebut.value)
            dateFin = CDate(Me.txtDateFin.value)
            If dateDebut > dateFin Then
                msgBox "La date de début doit être antérieure ou égale à la date de fin.", vbExclamation, "Erreur dans les critères de date"
                Exit Sub
            End If
        Else
            msgBox "Veuillez entrer des dates valides.", vbExclamation, "Erreur dans les critères de date"
            Exit Sub
        End If
        
        'Vérification qu'au moins un compte est sélectionné
        Dim ligneSélectionnée As Boolean
        ligneSélectionnée = EstLigneSelectionnee(Me.lsbComptes)
        If ligneSélectionnée = False Then
            msgBox "Veuillez sélectionner au moins un numéro de compte.", vbExclamation, "Erreur"
            Exit Sub
        End If
        
        'Les validations sont terminées, on appelle la procédure pour le rapport par compte
        Call GenererRapportGL_Compte(wsRapport, dateDebut, dateFin)

    Else
        'Validation des numéros d'écriture
        Dim noEcritureDebut As Long, noEcritureFin As Long
        If IsNumeric(Me.txtNoEcritureDebut.value) And IsNumeric(Me.txtNoEcritureFin.value) Then
            noEcritureDebut = CLng(Me.txtNoEcritureDebut.value)
            noEcritureFin = CLng(Me.txtNoEcritureFin.value)
            
            'Vérification de l'ordre des numéros d'écriture
            If noEcritureDebut > noEcritureFin Then
                msgBox "Le numéro d'écriture de début doit être inférieur ou égal au numéro de fin.", _
                            vbExclamation, "Erreur dans les critères de numéro d'écriture"
                Exit Sub
            End If
        Else
            msgBox "Veuillez entrer des numéros d'écriture valides.", vbExclamation, _
                        "Erreur dans les critères de numéro d'écriture"
            Exit Sub
        End If
        
        'Les validations sont terminées, on appelle la procédure pour le rapport par compte
        Call GenererRapportGL_Ecriture(wsRapport, noEcritureDebut, noEcritureFin)
    End If

End Sub

Private Sub cmdSelectAll_Click()

    'Boucle à travers tous les éléments du ListBox et les sélectionne
    Dim i As Integer
    For i = 0 To lsbComptes.ListCount - 1
        lsbComptes.Selected(i) = True
    Next i
    
End Sub

Private Sub cmdDeselectAll_Click()

    'Boucle à travers tous les éléments du ListBox et les désélectionne
    Dim i As Integer
    For i = 0 To lsbComptes.ListCount - 1
        lsbComptes.Selected(i) = False
    Next i

End Sub

'Public Sub testValiderDateDernierJourDuMois()
'
'    Dim y As Integer, m As Integer, j As Integer
'    y = 2025
'    m = 6
'    j = 31
'
'    Debug.Print ValiderDateDernierJourDuMois(y, m, j)
'
'End Sub
'
Function CorrigerDate(txtDate As String) As String

    Dim d As Integer, m As Integer, y As Integer
    Dim arr() As String
    Dim dt As Date
    Dim currentDate As Date
    Dim maxDayInMonth As Integer
    On Error GoTo ErrorHandler

    ' Récupérer la date actuelle
    currentDate = Date

    ' Supprimer les espaces et remplacer les séparateurs
    txtDate = Trim(txtDate)
    txtDate = Replace(txtDate, "-", "/")
    txtDate = Replace(txtDate, ".", "/")
    txtDate = Replace(txtDate, " ", "/")

    'Si la saisie est uniquement un jour (par exemple '5' ou '12'), on prend le mois et l'année actuels
    If IsNumeric(txtDate) And Len(txtDate) <= 2 Then
        'Si l'utilisateur entre un chiffre seul, c'est le jour du mois courant avec le mois et l'année courants
        d = CInt(txtDate)
        m = month(currentDate)
        y = year(currentDate)
        GoTo DerniereValidation
    End If

    'S'il y a un séparateur, on décompose la chaîne dans arr()
    arr = Split(txtDate, "/")
    
    'Y a-t-il 3 parties dans la date (Unound(arr) = 2) ?
    If UBound(arr) = 2 Then
        'jj/mm/aaaa ou aaaa/mm/jj
        If Len(arr(0)) = 4 Then
            'Format aaaa/mm/jj
            y = CInt(arr(0))
            m = CInt(arr(1))
            d = CInt(arr(2))
        Else
            'Format jj/mm/aaaa
            d = CInt(arr(0))
            m = CInt(arr(1))
            y = CInt(arr(2))
        End If
        
        GoTo DerniereValidation
        
    ElseIf UBound(arr) = 1 Then 'Deux parties dans la date saisie
        'L'une des 2 parties est vide
        If arr(0) = "" Or arr(1) = "" Then
            CorrigerDate = ""
            Exit Function
        End If
        
        If IsNumeric(arr(0)) And IsNumeric(arr(1)) Then
            'Format jj/mm
            d = CInt(arr(0))
            m = CInt(arr(1))
            y = year(currentDate) ' L'année courante par défaut
            GoTo DerniereValidation
        End If
    End If
    
DerniereValidation:
    If ValiderDateDernierJourDuMois(y, m, d) <> "" Then
        'Conversion en date (Toutes les validations sont terminées)
        dt = DateSerial(y, m, d)
        CorrigerDate = Format(dt, "dd/mm/yyyy")
    Else
        CorrigerDate = ""
    End If
    Exit Function

ErrorHandler:
    CorrigerDate = "" ' Retourne une chaîne vide si une erreur se produit
End Function


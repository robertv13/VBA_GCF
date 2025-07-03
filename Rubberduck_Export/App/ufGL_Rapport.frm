VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufGL_Rapport 
   Caption         =   "Rapport des transactions du G/L"
   ClientLeft      =   -45
   ClientTop       =   -270
   OleObjectBlob   =   "ufGL_Rapport.frx":0000
End
Attribute VB_Name = "ufGL_Rapport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()

    Call modImport.ImporterGLTransactions
    
    'Noter l'activité
    Call ConnectFormControls(Me)
    
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
        .AddItem "Par date de saisie"
        .ListIndex = 0 ' Sélection par défaut
    End With
    
    'Options du comboBox pour les périodes
    cmbPeriode.Clear
    
    'Vérifier si la plage nommée existe
    On Error Resume Next
    Dim plage As Range
    Set plage = wsdADMIN.Range("dnrDateRange")
    On Error GoTo 0
    
    If Not plage Is Nothing Then
        'Ajouter chaque élément de la plage à la ComboBox
        Dim cellule As Range
        For Each cellule In plage
            cmbPeriode.AddItem cellule.Value
        Next cellule
    Else
        MsgBox "La plage nommée 'dnrDateRange' est introuvable.", vbExclamation, "Contacter le développeur"
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
        fraParDateSaisie.Visible = False
        fraTypeEcriture.Visible = False
        fraParDate.Top = 50
        'Activer la sélection de comptes
        lsbComptes.Enabled = True
        cmdSelectAll.Enabled = True
        cmdDeselectAll.Enabled = True
    ElseIf cmbTypeRapport.ListIndex = 1 Then
        'Mode "Par numéro d'écriture"
        fraParDate.Visible = False
        fraParEcriture.Visible = True
        fraParEcriture.Top = 50
        fraParDateSaisie.Visible = False
        fraTypeEcriture.Visible = True
        fraTypeEcriture.Top = 145
        txtNoEcritureDebut = 1
        Dim maxNumeroEcriture As Long
        maxNumeroEcriture = Application.WorksheetFunction.Max(wsdGL_Trans.Range("A:A"))
        txtNoEcritureFin = maxNumeroEcriture
        'Désactiver la sélection de comptes (car non applicable)
        lsbComptes.Enabled = False
        cmdSelectAll.Enabled = False
        cmdDeselectAll.Enabled = False
    Else
        'Mode "Par date de saisie"
        fraParDate.Visible = False
        fraParEcriture.Visible = False
        fraParDateSaisie.Visible = True
        fraParDateSaisie.Top = 50
        fraTypeEcriture.Visible = True
        fraTypeEcriture.Top = 145
        'Désactiver la sélection de comptes (car non applicable)
        lsbComptes.Enabled = False
        cmdSelectAll.Enabled = False
        cmdDeselectAll.Enabled = False
    End If
    
End Sub

Private Sub cmbPeriode_Change()

        Select Case cmbPeriode.Value
            Case "Aujourd'hui"
                txtDateDebut.Value = Format$(Date, wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(Date, wsdADMIN.Range("B1").Value)
            Case "Mois Courant"
                txtDateDebut.Value = Format$(wsdADMIN.Range("MoisDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("MoisA"), wsdADMIN.Range("B1").Value)
            Case "Mois Dernier"
                txtDateDebut.Value = Format$(wsdADMIN.Range("MoisPrecDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("MoisPrecA"), wsdADMIN.Range("B1").Value)
            Case "Trimestre courant"
                txtDateDebut.Value = Format$(wsdADMIN.Range("TrimDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("TrimA"), wsdADMIN.Range("B1").Value)
            Case "Trimestre précédent"
                txtDateDebut.Value = Format$(wsdADMIN.Range("TrimPrecDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("TrimPrecA"), wsdADMIN.Range("B1").Value)
            Case "Année courante"
                txtDateDebut.Value = Format$(wsdADMIN.Range("AnneeDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("AnneeA"), wsdADMIN.Range("B1").Value)
            Case "Année précédente"
                txtDateDebut.Value = Format$(wsdADMIN.Range("AnneePrecDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("AnneePrecA"), wsdADMIN.Range("B1").Value)
            Case "7 derniers jours"
                txtDateDebut.Value = Format$(wsdADMIN.Range("SeptJoursDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("SeptJoursA"), wsdADMIN.Range("B1").Value)
            Case "15 derniers jours"
                txtDateDebut.Value = Format$(wsdADMIN.Range("QuinzeJoursDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("QuinzeJoursA"), wsdADMIN.Range("B1").Value)
            Case "Semaine"
                txtDateDebut.Value = Format$(wsdADMIN.Range("DateDebutSemaine"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("DateFinSemaine"), wsdADMIN.Range("B1").Value)
            Case "Toutes les dates"
                txtDateDebut.Value = Format$(#1/1/2024#, wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("AnneeA"), wsdADMIN.Range("B1").Value)
            Case Else
                txtDateDebut.Value = ""
                txtDateFin.Value = ""
        End Select

End Sub

Private Sub txtDateDebut_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim dateCorrigee As String
    
    If Trim$(txtDateDebut.Text) <> "" Then
        dateCorrigee = CorrigerDate(txtDateDebut.Text)
        If dateCorrigee = "" Then
            MsgBox "La date saisie est invalide, veuillez saisir une date sous un" & vbNewLine & vbNewLine & _
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
    
    If Trim$(txtDateFin.Text) <> "" Then
        dateCorrigee = CorrigerDate(txtDateFin.Text)
        If dateCorrigee = "" Then
            MsgBox "La date saisie est invalide, veuillez saisir une date sous un" & vbNewLine & vbNewLine & _
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

Private Sub txtDateSaisieDebut_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim dateCorrigee As String
    
    If Trim$(txtDateSaisieDebut.Text) <> "" Then
        dateCorrigee = CorrigerDate(txtDateSaisieDebut.Text)
        If dateCorrigee = "" Then
            MsgBox "La date saisie est invalide, veuillez saisir une date sous un" & vbNewLine & vbNewLine & _
                    "format valide (jj ou jj/mm ou jj/mm/aaaa ou aaaa/mm/jj)" & vbNewLine & vbNewLine & _
                    "Notez que le séparateur peut être '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpréter la date saisie"
            Cancel = True
            txtDateSaisieDebut.SelStart = 0
            txtDateSaisieDebut.SelLength = Len(txtDateSaisieDebut.Text)
        Else
            txtDateSaisieDebut.Text = dateCorrigee
        End If
    End If
End Sub

Private Sub txtDateSaisieFin_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Debug.Print "DateSaisieFin_Exit déclenché"
    
    Dim dateCorrigee As String
    
    If Trim$(txtDateSaisieFin.Text) <> "" Then
        dateCorrigee = CorrigerDate(txtDateSaisieFin.Text)
        If dateCorrigee = "" Then
            MsgBox "La date saisie est invalide, veuillez saisir une date sous un" & vbNewLine & vbNewLine & _
                    "format valide (jj ou jj/mm ou jj/mm/aaaa ou aaaa/mm/jj)" & vbNewLine & vbNewLine & _
                    "Notez que le séparateur peut être '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpréter la date saisie"
            Cancel = True
            txtDateSaisieFin.SelStart = 0
            txtDateSaisieFin.SelLength = Len(txtDateSaisieFin.Text)
        Else
            txtDateSaisieFin.Text = dateCorrigee
        End If
    End If
End Sub

Private Sub txtBoxSuivant_GotFocus()

    Debug.Print "Le focus est maintenant sur txtBoxSuivant"
    
End Sub

Private Sub cmdGenerer_Click()

    'Vérification que le type de rapport est sélectionné
    Dim TypeRapport As String
    If Me.cmbTypeRapport.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un type de rapport.", vbExclamation, "Erreur"
        Exit Sub
    Else
        TypeRapport = Me.cmbTypeRapport.Value
    End If

    'On efface/cree une feuille pour le rapport
    Dim strWsRapport As String
    strWsRapport = "X_GL_Rapport"
    Call CreateOrReplaceWorksheet(strWsRapport)
    Dim wsRapport As Worksheet
    Set wsRapport = ThisWorkbook.Sheets(strWsRapport)
    
    With wsRapport '2025-06-30 @ 20:11
        .Activate
        .Range("A3").Select
        .Application.ActiveWindow.FreezePanes = False
        .Application.ActiveWindow.SplitColumn = 0
        .Application.ActiveWindow.SplitRow = 2 'ligne au-dessus de la 3e
        .Application.ActiveWindow.FreezePanes = True
    End With

    'Vérification des critères selon le type de rapport
    If TypeRapport = "Par compte / par date" Then
        'Validation des dates
        Dim dateDebut As Date, dateFin As Date
        If IsDate(Me.txtDateDebut.Value) And IsDate(Me.txtDateFin.Value) Then
            dateDebut = CDate(Me.txtDateDebut.Value)
            dateFin = CDate(Me.txtDateFin.Value)
            If dateDebut > dateFin Then
                MsgBox "La date de début doit être antérieure ou égale à la date de fin.", vbExclamation, "Erreur dans les critères de date"
                Exit Sub
            End If
        Else
            MsgBox "Veuillez entrer des dates valides.", vbExclamation, "Erreur dans les critères de date"
            Exit Sub
        End If
        
        'Vérification qu'au moins un compte est sélectionné
        Dim ligneSélectionnée As Boolean
        ligneSélectionnée = EstLigneSelectionnee(Me.lsbComptes)
        If ligneSélectionnée = False Then
            MsgBox "Veuillez sélectionner au moins un numéro de compte.", vbExclamation, "Erreur"
            Exit Sub
        End If
        
        'Les validations sont terminées, on appelle la procédure pour le rapport par compte
        Call GenererRapportGL_Compte(wsRapport, dateDebut, dateFin)

    ElseIf TypeRapport = "Par numéro d'écriture" Then
        'Validation des numéros d'écriture
        Dim noEcritureDebut As Long, noEcritureFin As Long
        If IsNumeric(Me.txtNoEcritureDebut.Value) And IsNumeric(Me.txtNoEcritureFin.Value) Then
            noEcritureDebut = CLng(Me.txtNoEcritureDebut.Value)
            noEcritureFin = CLng(Me.txtNoEcritureFin.Value)
            
            'Vérification logique des numéros d'écriture
            If noEcritureDebut > noEcritureFin Then
                MsgBox "Le numéro d'écriture de début doit être inférieur ou égal au numéro de fin.", _
                            vbExclamation, "Erreur dans les critères de numéro d'écriture"
                Exit Sub
            End If
        Else
            MsgBox "Veuillez entrer des numéros d'écriture valides.", vbExclamation, _
                        "Erreur dans les critères de numéro d'écriture"
            Exit Sub
        End If
        
        'As-t-on MINIMALEMENT un type de transaction à imprimer ?
        If Not chkDebourse And _
            Not chkEncaissement And _
            Not chkDepotClient And _
            Not chkEJ And _
            Not chkFacture And _
            Not chkRegularisation Then
            MsgBox _
                Prompt:="Vous devez MINIMALEMENT choisir un type de transaction", _
                Title:="Selon les critères choisis, rien ne sera imprimé", _
                Buttons:=vbInformation
            Exit Sub
        End If
        
        'Les validations sont terminées, on appelle la procédure pour le rapport par compte
        Call GenererRapportGL_Ecriture(wsRapport, noEcritureDebut, noEcritureFin)
    Else
        'Validation des dates de saisie
        Dim dateSaisieDebut As Date, dateSaisieFin As Date
        If IsDate(Me.txtDateSaisieDebut.Value) And IsDate(Me.txtDateSaisieFin.Value) Then
            dateSaisieDebut = CDate(Me.txtDateSaisieDebut.Value)
            dateSaisieFin = CDate(Me.txtDateSaisieFin.Value)
            If dateSaisieDebut > dateSaisieFin Then
                MsgBox _
                    Prompt:="La date de début doit être antérieure ou égale à la date de fin.", _
                    Title:="Erreur dans les critères de date", _
                    Buttons:=vbExclamation
                Exit Sub
            End If
        Else
            MsgBox _
                Prompt:="Veuillez entrer des dates valides.", _
                Title:="rreur dans les critères de date", _
                Buttons:=vbExclamation
            Exit Sub
        End If
        
        'As-t-on MINIMALEMENT un type de transaction à imprimer ?
        If Not chkDebourse And _
            Not chkEncaissement And _
            Not chkDepotClient And _
            Not chkEJ And _
            Not chkFacture And _
            Not chkRegularisation Then
            MsgBox _
                Prompt:="Vous devez MINIMALEMENT choisir un type de transaction", _
                Title:="Selon les critères choisis, rien ne sera imprimé", _
                Buttons:=vbInformation
            Exit Sub
        End If
        
        Debug.Print CLng(dateSaisieDebut), CLng(dateSaisieFin)
        dateSaisieDebut = DateSerial(year(dateSaisieDebut), month(dateSaisieDebut), day(dateSaisieDebut))
        dateSaisieFin = DateSerial(year(dateSaisieFin), month(dateSaisieFin), day(dateSaisieFin))
        
        'Les validations sont terminées, on appelle la procédure pour le rapport par date de saisie
        Call GenererRapportGL_DateSaisie(wsRapport, dateSaisieDebut, dateSaisieFin)
       
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

Function CorrigerDate(txtDate As String) As String

    Dim d As Integer, m As Integer, Y As Integer
    Dim arr() As String
    Dim dt As Date
    Dim currentDate As Date
    Dim maxDayInMonth As Integer
    On Error GoTo ErrorHandler

    'Récupérer la date actuelle
    currentDate = Date

    'Supprimer les espaces & n'accepter que les caractères valides
    txtDate = Trim$(txtDate)
    If Not EstDateCaractereValide(txtDate) Then
        CorrigerDate = ""
        Exit Function
    End If
    'Uniformiser les séparateurs
    txtDate = Replace(txtDate, "-", "/")
    txtDate = Replace(txtDate, ".", "/")
    txtDate = Replace(txtDate, " ", "/")

    'Si la saisie est uniquement un jour (par exemple '5' ou '12'), on prend le mois et l'année actuels
    If IsNumeric(txtDate) And Len(txtDate) <= 2 Then
        'Si l'utilisateur entre un chiffre seul, c'est le jour du mois courant avec le mois et l'année courants
        d = CInt(txtDate)
        m = month(currentDate)
        Y = year(currentDate)
        GoTo DerniereValidation
    End If

    'Cas particulier 4, 6 ou 8 caractères sans séparateur
    If EstSeulementChiffres(txtDate) Then
        If Len(txtDate) = 4 Then
            txtDate = Left$(txtDate, 2) & "/" & Right$(txtDate, 2)
        ElseIf Len(txtDate) = 6 Then
            txtDate = Left$(txtDate, 2) & "/" & Mid$(txtDate, 3, 2) & "/" & Right$(txtDate, 2)
        Else
            txtDate = Left$(txtDate, 2) & "/" & Mid$(txtDate, 3, 2) & "/" & Right$(txtDate, 4)
        End If
    End If
    
    'S'il y a un séparateur, on décompose la chaîne dans arr()
    arr = Split(txtDate, "/")
    
    'S'il n'y a qu'une partie dans la date et que le nombre de caractères est de 4, on insère un séparateur
    If UBound(arr) = 0 And Len(arr(0)) = 4 Then
        arr(0) = Left$(arr(0), 2) & "/" & Right$(arr(0), 2)
    End If
    
    'Y a-t-il 3 parties dans la date (Unound(arr) = 2) ?
    If UBound(arr) = 2 Then
        'jj/mm/aaaa ou aaaa/mm/jj
        If Len(arr(0)) = 4 Then
            'Format aaaa/mm/jj
            Y = CInt(arr(0))
            m = CInt(arr(1))
            d = CInt(arr(2))
        Else
            'Format jj/mm/aaaa
            d = CInt(arr(0))
            m = CInt(arr(1))
            Y = CInt(arr(2))
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
            Y = year(currentDate) 'L'année courante par défaut
            GoTo DerniereValidation
        End If
    End If
    
DerniereValidation:
    If ValiderDateDernierJourDuMois(Y, m, d) <> "" Then
        'Conversion en date (Toutes les validations sont terminées)
        dt = DateSerial(Y, m, d)
        CorrigerDate = Format$(dt, "dd/mm/yyyy")
    Else
        CorrigerDate = ""
    End If
    Exit Function

ErrorHandler:
    CorrigerDate = "" ' Retourne une chaîne vide si une erreur se produit
End Function

Function EstDateCaractereValide(ByVal txt As String) As Boolean '2025-03-03 @ 09:49

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    'Expression régulière : accepte uniquement chiffres (0-9) et les séparateurs . / - et espace
    regex.Pattern = "^[0-9./\-\s]+$"
    regex.IgnoreCase = True
    regex.Global = False

    'Teste si la chaîne correspond au modèle
    EstDateCaractereValide = regex.test(txt)

    'Libérer la mémoire
    Set regex = Nothing
    
End Function

Function EstSeulementChiffres(ByVal txt As String) As Boolean '2025-03-03 @ 09:49

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    'Expression régulière en fonction du nombre de caractères
    regex.Pattern = "^\d{4}$|^\d{6}$|^\d{8}$"
    regex.IgnoreCase = True
    regex.Global = False

    'Teste si la chaîne correspond au modèle
    EstSeulementChiffres = regex.test(txt)

    'Libérer la mémoire
    Set regex = Nothing
    
End Function

Private Sub chkToutesEcritures_Click() '2025-03-03 @ 08:36

    Dim Activer As Boolean
    Activer = Me.chkToutesEcritures.Value 'True si coché, False sinon

    'Parcourir toutes les cases du Frame sauf "chkToutesEcritures"
    Dim ctrl As Control
    For Each ctrl In Me.fraTypeEcriture.Controls
        If TypeName(ctrl) = "CheckBox" And ctrl.Name <> "chkToutesEcritures" Then
            ctrl.Enabled = Not Activer 'Désactive si "Toutes les écritures" est cochée
            ctrl.Value = Activer       'Coche si "Toutes les écritures" est cochée
        End If
    Next ctrl
    
End Sub

Private Sub CheckBox_Click() '2025-03-03 @ 08:37

    'Empêcher de décocher une case si "Toutes les écritures" est cochée
    If Me.chkToutesEcritures.Value = True Then
        Application.EnableEvents = False
        Me.chkToutesEcritures.Value = False
        Application.EnableEvents = True
    End If
    
End Sub



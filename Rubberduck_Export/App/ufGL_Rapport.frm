VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufGL_Rapport 
   Caption         =   "Rapport des transactions du G/L"
   ClientHeight    =   9105.001
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

    Call modImport.ImporterGLTransactions
    
    'Noter l'activit�
    Call ConnectFormControls(Me)
    Call RafraichirActivite("Activit� dans userForm '" & Me.Name & "'")
    
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
        .AddItem "Par date de saisie"
        .ListIndex = 0 ' S�lection par d�faut
    End With
    
    'Options du comboBox pour les p�riodes
    cmbPeriode.Clear
    
    'V�rifier si la plage nomm�e existe
    On Error Resume Next
    Dim plage As Range
    Set plage = wsdADMIN.Range("dnrDateRange")
    On Error GoTo 0
    
    If Not plage Is Nothing Then
        'Ajouter chaque �l�ment de la plage � la ComboBox
        Dim cellule As Range
        For Each cellule In plage
            cmbPeriode.AddItem cellule.Value
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
        fraParDateSaisie.Visible = False
        fraTypeEcriture.Visible = False
        fraParDate.Top = 50
        'Activer la s�lection de comptes
        lsbComptes.Enabled = True
        cmdSelectAll.Enabled = True
        cmdDeselectAll.Enabled = True
    ElseIf cmbTypeRapport.ListIndex = 1 Then
        'Mode "Par num�ro d'�criture"
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
        'D�sactiver la s�lection de comptes (car non applicable)
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
        'D�sactiver la s�lection de comptes (car non applicable)
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
            Case "Trimestre pr�c�dent"
                txtDateDebut.Value = Format$(wsdADMIN.Range("TrimPrecDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("TrimPrecA"), wsdADMIN.Range("B1").Value)
            Case "Ann�e courante"
                txtDateDebut.Value = Format$(wsdADMIN.Range("AnneeDe"), wsdADMIN.Range("B1").Value)
                txtDateFin.Value = Format$(wsdADMIN.Range("AnneeA"), wsdADMIN.Range("B1").Value)
            Case "Ann�e pr�c�dente"
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
                    "Notez que le s�parateur peut �tre '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpr�ter la date saisie"
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
                    "Notez que le s�parateur peut �tre '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpr�ter la date saisie"
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
                    "Notez que le s�parateur peut �tre '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpr�ter la date saisie"
            Cancel = True
            txtDateSaisieDebut.SelStart = 0
            txtDateSaisieDebut.SelLength = Len(txtDateSaisieDebut.Text)
        Else
            txtDateSaisieDebut.Text = dateCorrigee
        End If
    End If
End Sub

Private Sub txtDateSaisieFin_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Debug.Print "DateSaisieFin_Exit d�clench�"
    
    Dim dateCorrigee As String
    
    If Trim$(txtDateSaisieFin.Text) <> "" Then
        dateCorrigee = CorrigerDate(txtDateSaisieFin.Text)
        If dateCorrigee = "" Then
            MsgBox "La date saisie est invalide, veuillez saisir une date sous un" & vbNewLine & vbNewLine & _
                    "format valide (jj ou jj/mm ou jj/mm/aaaa ou aaaa/mm/jj)" & vbNewLine & vbNewLine & _
                    "Notez que le s�parateur peut �tre '-' ou '/' ou ' '", vbExclamation, _
                    "Impossible d'interpr�ter la date saisie"
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

    'V�rification que le type de rapport est s�lectionn�
    Dim TypeRapport As String
    If Me.cmbTypeRapport.ListIndex = -1 Then
        MsgBox "Veuillez s�lectionner un type de rapport.", vbExclamation, "Erreur"
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
    
    'V�rification des crit�res selon le type de rapport
    If TypeRapport = "Par compte / par date" Then
        'Validation des dates
        Dim dateDebut As Date, dateFin As Date
        If IsDate(Me.txtDateDebut.Value) And IsDate(Me.txtDateFin.Value) Then
            dateDebut = CDate(Me.txtDateDebut.Value)
            dateFin = CDate(Me.txtDateFin.Value)
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

    ElseIf TypeRapport = "Par num�ro d'�criture" Then
        'Validation des num�ros d'�criture
        Dim noEcritureDebut As Long, noEcritureFin As Long
        If IsNumeric(Me.txtNoEcritureDebut.Value) And IsNumeric(Me.txtNoEcritureFin.Value) Then
            noEcritureDebut = CLng(Me.txtNoEcritureDebut.Value)
            noEcritureFin = CLng(Me.txtNoEcritureFin.Value)
            
            'V�rification logique des num�ros d'�criture
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
        
        'As-t-on MINIMALEMENT un type de transaction � imprimer ?
        If Not chkDebourse And _
            Not chkEncaissement And _
            Not chkDepotClient And _
            Not chkEJ And _
            Not chkFacture And _
            Not chkRegularisation Then
            MsgBox _
                Prompt:="Vous devez MINIMALEMENT choisir un type de transaction", _
                Title:="Selon les crit�res choisis, rien ne sera imprim�", _
                Buttons:=vbInformation
            Exit Sub
        End If
        
        'Les validations sont termin�es, on appelle la proc�dure pour le rapport par compte
        Call GenererRapportGL_Ecriture(wsRapport, noEcritureDebut, noEcritureFin)
    Else
        'Validation des dates de saisie
        Dim dateSaisieDebut As Date, dateSaisieFin As Date
        If IsDate(Me.txtDateSaisieDebut.Value) And IsDate(Me.txtDateSaisieFin.Value) Then
            dateSaisieDebut = CDate(Me.txtDateSaisieDebut.Value)
            dateSaisieFin = CDate(Me.txtDateSaisieFin.Value)
            If dateSaisieDebut > dateSaisieFin Then
                MsgBox _
                    Prompt:="La date de d�but doit �tre ant�rieure ou �gale � la date de fin.", _
                    Title:="Erreur dans les crit�res de date", _
                    Buttons:=vbExclamation
                Exit Sub
            End If
        Else
            MsgBox _
                Prompt:="Veuillez entrer des dates valides.", _
                Title:="rreur dans les crit�res de date", _
                Buttons:=vbExclamation
            Exit Sub
        End If
        
        'As-t-on MINIMALEMENT un type de transaction � imprimer ?
        If Not chkDebourse And _
            Not chkEncaissement And _
            Not chkDepotClient And _
            Not chkEJ And _
            Not chkFacture And _
            Not chkRegularisation Then
            MsgBox _
                Prompt:="Vous devez MINIMALEMENT choisir un type de transaction", _
                Title:="Selon les crit�res choisis, rien ne sera imprim�", _
                Buttons:=vbInformation
            Exit Sub
        End If
        
        Debug.Print CLng(dateSaisieDebut), CLng(dateSaisieFin)
        dateSaisieDebut = DateSerial(year(dateSaisieDebut), month(dateSaisieDebut), day(dateSaisieDebut))
        dateSaisieFin = DateSerial(year(dateSaisieFin), month(dateSaisieFin), day(dateSaisieFin))
        
        'Les validations sont termin�es, on appelle la proc�dure pour le rapport par date de saisie
        Call GenererRapportGL_DateSaisie(wsRapport, dateSaisieDebut, dateSaisieFin)
       
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

Function CorrigerDate(txtDate As String) As String

    Dim d As Integer, m As Integer, Y As Integer
    Dim arr() As String
    Dim dt As Date
    Dim currentDate As Date
    Dim maxDayInMonth As Integer
    On Error GoTo ErrorHandler

    'R�cup�rer la date actuelle
    currentDate = Date

    'Supprimer les espaces & n'accepter que les caract�res valides
    txtDate = Trim$(txtDate)
    If Not EstDateCaractereValide(txtDate) Then
        CorrigerDate = ""
        Exit Function
    End If
    'Uniformiser les s�parateurs
    txtDate = Replace(txtDate, "-", "/")
    txtDate = Replace(txtDate, ".", "/")
    txtDate = Replace(txtDate, " ", "/")

    'Si la saisie est uniquement un jour (par exemple '5' ou '12'), on prend le mois et l'ann�e actuels
    If IsNumeric(txtDate) And Len(txtDate) <= 2 Then
        'Si l'utilisateur entre un chiffre seul, c'est le jour du mois courant avec le mois et l'ann�e courants
        d = CInt(txtDate)
        m = month(currentDate)
        Y = year(currentDate)
        GoTo DerniereValidation
    End If

    'Cas particulier 4, 6 ou 8 caract�res sans s�parateur
    If EstSeulementChiffres(txtDate) Then
        If Len(txtDate) = 4 Then
            txtDate = Left$(txtDate, 2) & "/" & Right$(txtDate, 2)
        ElseIf Len(txtDate) = 6 Then
            txtDate = Left$(txtDate, 2) & "/" & Mid$(txtDate, 3, 2) & "/" & Right$(txtDate, 2)
        Else
            txtDate = Left$(txtDate, 2) & "/" & Mid$(txtDate, 3, 2) & "/" & Right$(txtDate, 4)
        End If
    End If
    
    'S'il y a un s�parateur, on d�compose la cha�ne dans arr()
    arr = Split(txtDate, "/")
    
    'S'il n'y a qu'une partie dans la date et que le nombre de caract�res est de 4, on ins�re un s�parateur
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
            Y = year(currentDate) 'L'ann�e courante par d�faut
            GoTo DerniereValidation
        End If
    End If
    
DerniereValidation:
    If ValiderDateDernierJourDuMois(Y, m, d) <> "" Then
        'Conversion en date (Toutes les validations sont termin�es)
        dt = DateSerial(Y, m, d)
        CorrigerDate = Format$(dt, "dd/mm/yyyy")
    Else
        CorrigerDate = ""
    End If
    Exit Function

ErrorHandler:
    CorrigerDate = "" ' Retourne une cha�ne vide si une erreur se produit
End Function

Function EstDateCaractereValide(ByVal txt As String) As Boolean '2025-03-03 @ 09:49

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    'Expression r�guli�re : accepte uniquement chiffres (0-9) et les s�parateurs . / - et espace
    regex.pattern = "^[0-9./\-\s]+$"
    regex.IgnoreCase = True
    regex.Global = False

    'Teste si la cha�ne correspond au mod�le
    EstDateCaractereValide = regex.test(txt)

    'Lib�rer la m�moire
    Set regex = Nothing
    
End Function

Function EstSeulementChiffres(ByVal txt As String) As Boolean '2025-03-03 @ 09:49

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    'Expression r�guli�re en fonction du nombre de caract�res
    regex.pattern = "^\d{4}$|^\d{6}$|^\d{8}$"
    regex.IgnoreCase = True
    regex.Global = False

    'Teste si la cha�ne correspond au mod�le
    EstSeulementChiffres = regex.test(txt)

    'Lib�rer la m�moire
    Set regex = Nothing
    
End Function

Private Sub chkToutesEcritures_Click() '2025-03-03 @ 08:36

    Dim Activer As Boolean
    Activer = Me.chkToutesEcritures.Value 'True si coch�, False sinon

    'Parcourir toutes les cases du Frame sauf "chkToutesEcritures"
    Dim ctrl As Control
    For Each ctrl In Me.fraTypeEcriture.Controls
        If TypeName(ctrl) = "CheckBox" And ctrl.Name <> "chkToutesEcritures" Then
            ctrl.Enabled = Not Activer 'D�sactive si "Toutes les �critures" est coch�e
            ctrl.Value = Activer       'Coche si "Toutes les �critures" est coch�e
        End If
    Next ctrl
    
End Sub

Private Sub CheckBox_Click() '2025-03-03 @ 08:37

    'Emp�cher de d�cocher une case si "Toutes les �critures" est coch�e
    If Me.chkToutesEcritures.Value = True Then
        Application.EnableEvents = False
        Me.chkToutesEcritures.Value = False
        Application.EnableEvents = True
    End If
    
End Sub


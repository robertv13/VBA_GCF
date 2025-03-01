Attribute VB_Name = "modGL_Rapport_TBD"
Option Explicit

'Sub shp_GL_Rapport_GO_Click()
'
'    Call GL_Report_For_Selected_Accounts
'
'End Sub
'
'Public Sub GL_Report_For_Selected_Accounts()
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport:GL_Report_For_Selected_Accounts", "", 0)
'
'    'Reference the worksheet
'    Dim ws As Worksheet: Set ws = wshGL_Rapport
'
'    If ws.Range("F8").value = "" Or ws.Range("H8").value = "" Then
'        MsgBox "Vous devez saisir une date de début et une date de fin pour ce rapport!"
'        Exit Sub
'    End If
'
'    If ws.Range("H8").value < ws.Range("F8").value Then
'        MsgBox "La date de départ doit obligatoirement être antérieure" & vbNewLine & vbNewLine & _
'                "ou égale à la date de fin!", vbInformation
'        Exit Sub
'    End If
'
'    'Reference the listBox
'    Dim lb As OLEObject: Set lb = ws.OLEObjects("ListBox1")
'
'    'Ensure it is a ListBox
'    Dim selectedItems As Collection
'    If TypeName(lb.Object) = "ListBox" Then
'        Set selectedItems = New Collection
'
'        'Loop through ListBox items and collect selected ones
'        Dim i As Long
'        With lb.Object
'            For i = 0 To .ListCount - 1
'                If .Selected(i) And Trim(.List(i)) <> "" Then
'                    selectedItems.Add .List(i)
'                End If
'            Next i
'        End With
'
'        'Is there any account selected ?
'        If selectedItems.count = 0 Then
'            MsgBox "Il n'y a aucune compte de sélectionné!"
'            Exit Sub
'        End If
'
'        'Erase & Create output Worksheet
'        Call CreateOrReplaceWorksheet("X_GL_Rapport_Out")
'
'        'Setup report header
'        Call SetUpGLReportHeadersAndColumns_Compte
'
'        'Prepare Variables
'        Dim dateDeb As Date, dateFin As Date, sortType As String
'        With wshGL_Rapport
'            dateDeb = CDate(.Range("F8").value)
'            dateFin = CDate(.Range("H8").value)
'            If .Range("B3").value = "Vrai" Then
'                sortType = "Date"
'            Else
'                sortType = "Transaction"
'            End If
'        End With
'
'        Application.ScreenUpdating = False
'        Application.DisplayAlerts = False
'
'        'Process one account at the time...
'        Dim item As Variant
'        Dim compte As String
'        Dim descGL As String
'        Dim GL As String
'        For Each item In selectedItems
'            compte = item
'            GL = Left(compte, InStr(compte, " ") - 1)
'            descGL = Right(compte, Len(compte) - InStr(compte, " "))
'            'Obtenir le solde d'ouverture & les transactions
'            Dim soldeOuverture As Currency
'            soldeOuverture = Fn_Get_GL_Account_Balance(GL, dateDeb - 1)
'
'            'Impression des résultats
'            Call Print_Results_From_GL_Trans(GL, descGL, soldeOuverture, dateDeb, dateFin)
'
'        Next item
'
'        Application.DisplayAlerts = True
'        Application.ScreenUpdating = True
'
'    End If
'
'    Dim h1 As String, h2 As String, h3 As String
'    h1 = wshAdmin.Range("NomEntreprise")
'    h2 = "Rapport des transactions du Grand Livre"
'    h3 = "(Du " & dateDeb & " au " & dateFin & ")"
'    Call GL_Rapport_Wrap_Up(h1, h2, h3)
'
'    'Libérer la mémoire
'    Set item = Nothing
'    Set lb = Nothing
'    Set selectedItems = Nothing
'    Set ws = Nothing
'
'    Call Log_Record("modGL_Rapport:GL_Report_For_Selected_Accounts", "", startTime)
'
'End Sub
'
'Sub shp_GL_Rapport_Exit_Click()
'
'    Call GL_Rapport_Back_To_Menu
'
'End Sub
'
'Sub GL_Rapport_Back_To_Menu()
'
'    Dim startTime As Double: startTime = Timer: Call Log_Record("modGL_Rapport:GL_Rapport_Back_To_Menu", "", 0)
'
'    wshGL_Rapport.Visible = xlSheetHidden
'    On Error Resume Next
'    ThisWorkbook.Worksheets("X_GL_Rapport_Out").Visible = xlSheetHidden
'    On Error GoTo 0
'
'    wshMenuGL.Activate
'    wshMenuGL.Range("A1").Select
'
'    Call Log_Record("modGL_Rapport:GL_Rapport_Back_To_Menu", "", startTime)
'
'End Sub
'
'

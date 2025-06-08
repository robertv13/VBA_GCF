Attribute VB_Name = "modDEV_Debug"
Option Explicit

Sub IdentifierEcartsComptesClientsEtGL() '2025-04-02 @ 07:46

    Dim wsCC As Worksheet
    Set wsCC = wsdFAC_Comptes_Clients
    Dim usedRowCC As Long
    usedRowCC = wsCC.Cells(wsCC.Rows.count, 1).End(xlUp).row
    
    Dim wsGL As Worksheet
    Set wsGL = wsdGL_Trans
    Dim usedRowGL As Long
    usedRowGL = wsGL.Cells(wsGL.Rows.count, 1).End(xlUp).row
    
    Dim wsENC As Worksheet
    Set wsENC = wsdENC_D�tails
    Dim usedRowENC As Long
    usedRowENC = wsENC.Cells(wsENC.Rows.count, 1).End(xlUp).row
    
    'Matrice pour comparer les d�p�ts
    Dim matENC() As Currency
    ReDim matENC(1 To 500, 1 To 2)
    Dim matFAC() As Currency
    ReDim matFAC(24475 To 26000, 1 To 2)
    
    'Additionne TOUS les encaissements � partir de ENC_D�tails et accumule dans dictionary
    Dim totalENC_D�tails As Currency
    Dim i As Long
    For i = 2 To usedRowENC
        totalENC_D�tails = totalENC_D�tails + wsENC.Cells(i, 5).value
        matENC(wsENC.Cells(i, 1).value, 1) = matENC(wsENC.Cells(i, 1).value, 1) + wsENC.Cells(i, 5).value
    Next i
    Debug.Print "ENC_D�tails        ", "Total des encaissements = " & Format$(totalENC_D�tails, "#,##0.00 $") & " pour " & usedRowENC & " lignes"

    'Additionne TOUS les encaissements � partir de FAC_Comptes_Clients
    Dim totalCC_D�tails As Currency
    Dim noFact As String
    For i = 3 To usedRowCC
        totalCC_D�tails = totalCC_D�tails + wsCC.Cells(i, 9).value
        noFact = wsCC.Cells(i, 1).value
        If InStr(noFact, "-") Then
            noFact = Right(noFact, 5)
        End If
        If noFact > 24474 Then 'Premi�re facture cr��e par le logiciel
            matFAC(noFact, 1) = matFAC(noFact, 1) + wsCC.Cells(i, 8).value
        End If
    Next i
    Debug.Print "FAC_Comptes_Clients", "Total des encaissements = " & Format$(totalCC_D�tails, "#,##0.00 $") & " pour " & usedRowCC & " lignes"

    'Analyse TOUS les �critures au G/L
    Dim totalGL_D�tails As Currency
    Dim Source As String, noEnc As Long
    For i = 2 To usedRowGL
        Source = wsGL.Cells(i, 4).value
        If wsGL.Cells(i, 5).value = "1100" Then
            If InStr(Source, "ENCAISSEMENT:") = 1 Or InStr(Source, "D�P�T DE CLIENT:") = 1 Then
                totalGL_D�tails = totalGL_D�tails - wsGL.Cells(i, 7).value + wsGL.Cells(i, 8).value
                noEnc = Mid$(Source, InStr(Source, ":") + 1, Len(Source) - InStr(Source, ":"))
                
                matENC(noEnc, 2) = matENC(noEnc, 2) - wsGL.Cells(i, 7).value + wsGL.Cells(i, 8).value
            Else
                If InStr(Source, "FACTURE:") = 1 Then
                    noFact = Right(Source, 5)
                    matFAC(noFact, 2) = matFAC(noFact, 2) + wsGL.Cells(i, 7).value - wsGL.Cells(i, 8).value
                End If
            End If
            
        End If
    Next i
    Debug.Print "GL_Trans          ", "Total des encaissements = " & Format$(totalGL_D�tails, "#,##0.00 $") & " pour " & usedRowGL & " lignes"

    'Compare les deux valeurs de matENC
    For i = 1 To UBound(matENC, 1)
        If matENC(i, 1) <> matENC(i, 2) Then
            Debug.Print "Encaissement # " & i & " - Encaissement = " & Format$(matENC(i, 1), "#,##0.00 $") & " vs. GL = " & Format$(matENC(i, 2), "#,##0.00 $")
        End If
    Next i
    
    
    'Compare les deux valeurs de matFAC
    For i = LBound(matFAC, 1) To UBound(matFAC, 1)
        If matFAC(i, 1) <> matFAC(i, 2) Then
            Debug.Print "Facture # " & i & " - FAC_Comptes_Clients = " & Format$(matFAC(i, 1), "#,##0.00 $") & " vs. GL = " & Format$(matFAC(i, 2), "#,##0.00 $") & " n'est pas comptabilis�e !!!"
        End If
    Next i
    
    MsgBox "Fin de la v�rification"
    
End Sub

Sub V�rifierTousLesContr�lesFeuillesEtUserForms()

    Dim ws As Worksheet
    Dim ctrl As OLEObject
    Dim rapport As String
    Dim testValue As Variant
    Dim erreurTrouv�e As Boolean
    Dim vbComp As Object
    Dim uf As Object
    Dim ctrlUF As MSForms.Control

    rapport = "?? Contr�les ActiveX corrompus (feuilles + UserForms)" & vbCrLf & String(60, "-") & vbCrLf

    ' V�rification des feuilles
    For Each ws In ThisWorkbook.Worksheets
        For Each ctrl In ws.OLEObjects
            On Error Resume Next
            testValue = ctrl.Object.value
            If Err.Number <> 0 Then
                rapport = rapport & "?? Feuille: " & ws.Name & _
                          " - Contr�le: " & ctrl.Name & _
                          " ? Erreur : " & Err.Description & vbCrLf
            End If
            Err.Clear
            On Error GoTo 0
        Next ctrl
    Next ws

    ' V�rification des UserForms
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 3 Then ' 3 = vbext_ct_MSForm (UserForm)
            On Error Resume Next
            Set uf = VBA.UserForms.Add(vbComp.Name)
            If Err.Number <> 0 Then
                rapport = rapport & "? UserForm: " & vbComp.Name & " - Erreur d'ouverture : " & Err.Description & vbCrLf
                Err.Clear
                GoTo NextForm
            End If

            For Each ctrlUF In uf.Controls
                On Error Resume Next
                testValue = ctrlUF.value
                If Err.Number <> 0 Then
                    rapport = rapport & "?? UserForm: " & vbComp.Name & _
                              " - Contr�le: " & ctrlUF.Name & _
                              " ? Erreur : " & Err.Description & vbCrLf
                End If
                Err.Clear
                On Error GoTo 0
            Next ctrlUF

NextForm:
            Unload uf
        End If
    Next vbComp

    If InStr(rapport, "? Erreur") > 0 Or InStr(rapport, "?") > 0 Then
        MsgBox rapport, vbExclamation, "?? Probl�mes d�tect�s"
    Else
        MsgBox "? Aucun contr�le probl�matique trouv� sur les feuilles ou les UserForms.", vbInformation, "Tout est OK"
    End If
End Sub



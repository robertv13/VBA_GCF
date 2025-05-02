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
    Set wsENC = wsdENC_Détails
    Dim usedRowENC As Long
    usedRowENC = wsENC.Cells(wsENC.Rows.count, 1).End(xlUp).row
    
    'Matrice pour comparer les dépôts
    Dim matENC() As Currency
    ReDim matENC(1 To 500, 1 To 2)
    Dim matFAC() As Currency
    ReDim matFAC(24475 To 24888, 1 To 2)
    
    'Additionne TOUS les encaissements à partir de ENC_Détails et accumule dans dictionary
    Dim totalENC_Détails As Currency
    Dim i As Long
    For i = 2 To usedRowENC
        totalENC_Détails = totalENC_Détails + wsENC.Cells(i, 5).value
        matENC(wsENC.Cells(i, 1).value, 1) = matENC(wsENC.Cells(i, 1).value, 1) + wsENC.Cells(i, 5).value
    Next i
    Debug.Print "ENC_Détails        ", "Total des encaissements = " & Format$(totalENC_Détails, "#,##0.00 $") & " pour " & usedRowENC & " lignes"

    'Additionne TOUS les encaissements à partir de FAC_Comptes_Clients
    Dim totalCC_Détails As Currency
    Dim noFact As String
    For i = 3 To usedRowCC
        totalCC_Détails = totalCC_Détails + wsCC.Cells(i, 9).value
        noFact = wsCC.Cells(i, 1).value
        If InStr(noFact, "-") Then
            noFact = Right(noFact, 5)
        End If
        If noFact > 24474 Then 'Première facture créée par le logiciel
            matFAC(noFact, 1) = matFAC(noFact, 1) + wsCC.Cells(i, 8).value
        End If
    Next i
    Debug.Print "FAC_Comptes_Clients", "Total des encaissements = " & Format$(totalCC_Détails, "#,##0.00 $") & " pour " & usedRowCC & " lignes"

    'Analyse TOUS les écritures au G/L
    Dim totalGL_Détails As Currency
    Dim source As String, noEnc As Long
    For i = 2 To usedRowGL
        source = wsGL.Cells(i, 4).value
        If wsGL.Cells(i, 5).value = "1100" Then
            If InStr(source, "ENCAISSEMENT:") = 1 Or InStr(source, "DÉPÔT DE CLIENT:") = 1 Then
                totalGL_Détails = totalGL_Détails - wsGL.Cells(i, 7).value + wsGL.Cells(i, 8).value
                noEnc = Mid$(source, InStr(source, ":") + 1, Len(source) - InStr(source, ":"))
                
                matENC(noEnc, 2) = matENC(noEnc, 2) - wsGL.Cells(i, 7).value + wsGL.Cells(i, 8).value
            Else
                If InStr(source, "FACTURE:") = 1 Then
                    noFact = Right(source, 5)
                    matFAC(noFact, 2) = matFAC(noFact, 2) + wsGL.Cells(i, 7).value - wsGL.Cells(i, 8).value
                End If
            End If
            
        End If
    Next i
    Debug.Print "GL_Trans          ", "Total des encaissements = " & Format$(totalGL_Détails, "#,##0.00 $") & " pour " & usedRowGL & " lignes"

    'Compare les deux valeurs de matENC
    For i = 1 To UBound(matENC, 1)
        If matENC(i, 1) <> matENC(i, 2) Then
            Debug.Print "Encaissement # " & i & " - Encaissement = " & Format$(matENC(i, 1), "#,##0.00 $") & " vs. GL = " & Format$(matENC(i, 2), "#,##0.00 $")
        End If
    Next i
    
    
    'Compare les deux valeurs de matFAC
    For i = LBound(matFAC, 1) To UBound(matFAC, 1)
        If matFAC(i, 1) <> matFAC(i, 2) Then
            Debug.Print "Facture # " & i & " - FAC_Comptes_Clients = " & Format$(matFAC(i, 1), "#,##0.00 $") & " vs. GL = " & Format$(matFAC(i, 2), "#,##0.00 $") & " n'est pas comptabilisée !!!"
        End If
    Next i
    
End Sub


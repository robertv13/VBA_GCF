Attribute VB_Name = "modDEV_Test"
Option Explicit

Sub test()

    Dim ws As Worksheet
    Set ws = wsdFAC_Projets_Détails
    
    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.count, "AA").End(xlUp).row
    Debug.Print lastUsedRow
    
End Sub

Sub TestGetSummary()

    Dim rs As ADODB.Recordset
    Dim dateMin As Date, dateMax As Date
    
    dateMin = DateSerial(2025, 1, 1)
    dateMax = DateSerial(2025, 6, 30)
    
    Set rs = Get_Summary_By_GL_Account(dateMin, dateMax)
    
    ' Par exemple, lire le recordset et afficher les comptes + totaux
    Do While Not rs.EOF
        Debug.Print rs.Fields("NoCompte").value, rs.Fields("TotalDebit").value, rs.Fields("TotalCredit").value
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Sub Test_Balance_Verification() '2025-06-01 @ 10:53

    Dim arr As Variant
    Dim i As Long
    
    Call Get_Summary_By_GL_Account(DateSerial(2024, 7, 31), DateSerial(2025, 6, 1))
    
    If Not IsArray(arr) Then
        MsgBox "Aucune donnée dans cette période."
    Else
        For i = 0 To UBound(arr, 1)
            Debug.Print "Compte: " & arr(i, 0) & _
                        " - Description: " & arr(i, 1) & _
                        " - Débit: " & Format$(arr(i, 2), "#,##0.00 $") & _
                        " Crédit: " & Format$(arr(i, 3), "#,##0.00 $")
        Next i
    End If
End Sub

Sub Test_Ouverture_ADO()

    Dim cn As Object
    Dim f As String
    
    f = "C:\VBA\GC_FISCALITÉ\DataFiles\GL_Temp_RobertMV_20250601_120734.xlsx" ' <-- remplace par ton chemin exact

    Set cn = CreateObject("ADODB.Connection")
    
    On Error GoTo ErreurADO
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & f & ";" & _
            "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    MsgBox "Connexion réussie"
    cn.Close
    Exit Sub

ErreurADO:
    MsgBox "Erreur de connexion ADO : " & Err.Description, vbCritical
    
End Sub

Sub TestConnexionADO() '2025-06-08 @ 09:17
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;'"
    MsgBox "Connexion OK"
    cn.Close
End Sub

Sub TestADO()
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    On Error GoTo ErrADO
    Dim destinationFileName As String
    destinationFileName = "C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx"
    
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"

'    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\VBA\GC_FISCALITÉ\DataFiles\GCF_BD_MASTER.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;'"
    MsgBox "Connexion OK"
    cn.Close
    Exit Sub

ErrADO:
    MsgBox "Erreur : " & Err.Description
End Sub

Sub TestAccesVBProject()
    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Debug.Print vbComp.Name
    Next vbComp
End Sub


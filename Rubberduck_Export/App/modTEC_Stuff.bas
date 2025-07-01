Attribute VB_Name = "modTEC_Stuff"
Option Explicit

Public Sub Convertir_NF_en_Facturable_Dans_BD(tecID As Long) '2025-01-15 @ 09:44

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Stuff:Convertir_NF_en_Facturable_Dans_BD", CStr(tecID), 0)

    Application.ScreenUpdating = False
    
    'Classeur & feuille à mettre à jour
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("F5").Value & gDATA_PATH & Application.PathSeparator & _
                          "GCF_BD_MASTER.xlsx"
    destinationTab = "TEC_Local$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & _
        ";Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    'Open the recordset for the specified tecID
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID
    rs.Open strSQL, conn, 2, 3
    If Not rs.EOF Then
        'Update EstFacturee, DateFacturee & NoFacture
        rs.Fields(fTECEstFacturable - 1).Value = "VRAI"
    Else
        'On ne trouve pas le tecID - ANORMAL !!!
        MsgBox "L'enregistrement avec le TECID '" & tecID & "' ne peut être trouvé!", vbOK + vbCritical, "Problème avec la convertion (N/FACT ---> FACT)"
        rs.Close
        conn.Close
        Exit Sub
    End If
    
    'Update the recordset
    rs.Update
    
    'Close recordset and connection
    rs.Close
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set rs = Nothing
    
    Call Log_Record("modTEC_Stuff:Convertir_NF_en_Facturable_Dans_BD", "", startTime)

End Sub

Public Sub Convertir_NF_en_Facturable_Locally(tecID As Long) '2025-01-15 @ 09:44

    Dim startTime As Double: startTime = Timer: Call Log_Record("modTEC_Stuff:Convertir_NF_en_Facturable_Locally", CStr(tecID), 0)
    
    Dim ws As Worksheet
    Set ws = wsdTEC_Local
    
    'Déterminer la plage à rechercher dans TEC_Local
    Dim lastTECRow As Long
    lastTECRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Dim lookupRange As Range
    Set lookupRange = ws.Range("A3:A" & lastTECRow)
    
    Dim rowToBeUpdated As Long
    rowToBeUpdated = Fn_Find_Row_Number_TECID(tecID, lookupRange)
    
    'Convertir à Facturable
    ws.Cells(rowToBeUpdated, fTECEstFacturable).Value = "VRAI"

    Call Log_Record("modTEC_Stuff:Convertir_NF_en_Facturable_Locally", "", startTime)

End Sub


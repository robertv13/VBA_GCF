Attribute VB_Name = "modTEC_Stuff"
Option Explicit

Public Sub ConvertirNFenFacturableBDMaster(tecID As Long) '2025-01-15 @ 09:44

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Stuff:ConvertirNFenFacturableBDMaster", CStr(tecID), 0)

    Application.ScreenUpdating = False
    
    'Classeur & feuille à mettre à jour
    Dim destinationFileName As String, destinationTab As String
    destinationFileName = wsdADMIN.Range("PATH_DATA_FILES").Value & gDATA_PATH & Application.PathSeparator & _
                          wsdADMIN.Range("MASTER_FILE").Value
    destinationTab = "TEC_Local$"
    
    'Initialize connection, connection string & open the connection
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destinationFileName & ";" & _
              "Extended Properties=""Excel 12.0 XML;HDR=YES"";"
    Dim recSet As Object: Set recSet = CreateObject("ADODB.Recordset")

    'Open the recordset for the specified tecID
    Dim strSQL As String
    strSQL = "SELECT * FROM [" & destinationTab & "] WHERE TECID=" & tecID
    recSet.Open strSQL, conn, 2, 3
    If Not recSet.EOF Then
        'Update EstFacturee, DateFacturee & NoFacture
        recSet.Fields(fTECEstFacturable - 1).Value = "VRAI"
    Else
        'On ne trouve pas le tecID - ANORMAL !!!
        MsgBox "L'enregistrement avec le TECID '" & tecID & "' ne peut être trouvé!", vbOK + vbCritical, "Problème avec la convertion (N/FACT ---> FACT)"
        recSet.Close
        conn.Close
        Exit Sub
    End If
    
    'Update the recordset
    recSet.Update
    
    'Close recordset and connection
    recSet.Close
    conn.Close
    
    Application.ScreenUpdating = True

    'Libérer la mémoire
    Set conn = Nothing
    Set recSet = Nothing
    
    Call modDev_Utils.EnregistrerLogApplication("modTEC_Stuff:ConvertirNFenFacturableBDMaster", vbNullString, startTime)

End Sub

Public Sub ConvertirNFenFacturableBDLocale(tecID As Long) '2025-01-15 @ 09:44

    Dim startTime As Double: startTime = Timer: Call modDev_Utils.EnregistrerLogApplication("modTEC_Stuff:ConvertirNFenFacturableBDLocale", CStr(tecID), 0)
    
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

    Call modDev_Utils.EnregistrerLogApplication("modTEC_Stuff:ConvertirNFenFacturableBDLocale", vbNullString, startTime)

End Sub


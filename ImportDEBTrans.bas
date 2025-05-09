Sub ImporterDebTrans() '2025-05-07 @ 14:25
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterDebTrans", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdDEB_Trans
    Dim onglet As String, table As String
    onglet = "DEB_Trans"
    table = "l_tbl_DEB_Trans"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterDebTrans", "", startTime)

End Sub
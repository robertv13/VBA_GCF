Sub ImporterCCRegularisations() '2025-05-07 @ 13:58
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modImport:ImporterCCRegularisations", "", 0)
    
    'Mettre en place les variables (paramètres)
    Dim sourceWb As String
    sourceWb = "GCF_BD_MASTER.xlsx"
    Dim ws As Worksheet
    Set ws = wsdCC_Régularisations
    Dim onglet As String, table As String
    onglet = "CC_Régularisations"
    table = "l_tbl_CC_Régularisations"

    Call ImportGeneriqueDuMaster(sourceWb, ws, onglet, table)
    
    Application.ScreenUpdating = True
    
    'Libérer la mémoire
    Set ws = Nothing
    
    Call Log_Record("modImport:ImporterCCRegularisations", "", startTime)

End Sub
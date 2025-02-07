Attribute VB_Name = "modTEC_TDB"
Option Explicit

Sub shpTEC_TDB_BackToMenu_Click()

    Call TEC_TDB_BackToMenu

End Sub

Sub TEC_TDB_BackToMenu()

    wshTEC_TDB.Visible = xlSheetHidden
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

End Sub

Sub shpActualiserTEC_TDB_Click()

    Call ActualiserTEC_TDB

End Sub

Sub ActualiserTEC_TDB()

    startTime = Timer: Call Log_Record("modTEC_TDB:ActualiserTEC_TDB", "", 0)
    
    Call TEC_Update_TDB_From_TEC_Local
    Call TEC_TdB_Refresh_All_Pivot_Tables
    
    Call Log_Record("modTEC_TDB:ActualiserTEC_TDB", "", startTime)

End Sub


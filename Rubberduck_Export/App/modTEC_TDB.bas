Attribute VB_Name = "modTEC_TDB"
Option Explicit

Sub shp_TEC_TDB_Back_To_TEC_Menu_Click()

    Call TEC_TDB_Back_To_TEC_Menu

End Sub

Sub TEC_TDB_Back_To_TEC_Menu()

    wshTEC_TDB.Visible = xlSheetHidden
    
    wshMenuTEC.Activate
    wshMenuTEC.Range("A1").Select

End Sub


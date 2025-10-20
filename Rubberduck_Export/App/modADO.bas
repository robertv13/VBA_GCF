Attribute VB_Name = "modADO"
Option Explicit

Function OuvrirRecordsetADO(fichier As String, feuille As String) As Object '2025-10-19 @ 09:04

    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fichier & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
    rs.Open "SELECT * FROM [" & feuille & "]", cn, 1, 1

    Set OuvrirRecordsetADO = rs
    
End Function


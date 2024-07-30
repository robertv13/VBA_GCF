Attribute VB_Name = "modDataConversion"
Option Explicit

Sub Conversion_CC_Dans_Sortie()

    Dim ws As Worksheet: Set ws = wshFAC_Comptes_Clients
    
End Sub

Sub Conversion_GL_Dans_Sortie()

    Dim ws As Worksheet: Set ws = wshGL_Trans

    'Utilisation des ENum - fglt
    
End Sub

Sub Conversion_TEC_Dans_Sortie()

    Dim ws As Worksheet: Set ws = wshTEC_Local

    'Utilisation des ENum - ftec

End Sub

Sub Import_Rows_Clients_With_Column_Mapping()
    Dim rs As Object
    Dim query As String
    
    'D�finir le chemin du workbook (source)
    Dim sourcePath As String
    sourcePath = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                "Clients.xlsx" '2024-07-29 @ 17:43
    Dim sourceTab As String
    sourceTab = "Clients"
    'D�finir le chemin du workbook (destination)
    Dim destPath As String
    destPath = wshAdmin.Range("FolderSharedData").value & Application.PathSeparator & _
                "GCF_BD_Entr�e.xlsx" '2024-07-29 @ 17:43
    
    'Cr�er des objets Connection (source & destination)
    Dim sourceConn As Object
    Set sourceConn = CreateObject("ADODB.Connection")
    Dim destConn As Object
    Set destConn = CreateObject("ADODB.Connection")
    
    'Ouvrir la connexion vers le workbook source
    sourceConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'Ouvrir la connexion vers le workbook destination
    destConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & destPath & ";Extended Properties=""Excel 12.0;HDR=Yes"";"
    
    'D�finir la requ�te SQL pour s�lectionner les rang�es � copier
    query = "SELECT * FROM [Feuil1$]"
    
    'Ex�cuter la requ�te et obtenir les donn�es du workbook source
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open query, sourceConn, 1, 3
    
    'Mappage des colonnes du workbook source vers le workbook destination
    Dim columnMapping As Object: Set columnMapping = CreateObject("Scripting.Dictionary")
    columnMapping.add "SourceCol1", "DestColA"
    columnMapping.add "SourceCol2", "DestColB"
    columnMapping.add "SourceCol3", "DestColC"

    'Copier les donn�es dans le workbook destination
    Do While Not rs.EOF
        'Construire la requ�te d'insertion pour chaque ligne
        Dim insertQuery As String
        insertQuery = "INSERT INTO [Feuil1$] ("
        
        'Ajouter les noms des colonnes de destination
        Dim destColumns As String: destColumns = ""
        Dim values As String: values = ""
        
        Dim key As Variant
        For Each key In columnMapping.keys
            If destColumns <> "" Then
                destColumns = destColumns & ", "
                values = values & ", "
            End If
            destColumns = destColumns & columnMapping(key)
            values = values & "'" & rs.Fields(key).value & "'"
        Next key
        
        'Pr�parer et ex�cuter la requ�te d'insertion
        insertQuery = insertQuery & destColumns & ") VALUES (" & values & ")"
        destConn.Execute insertQuery
        
        'Passer � la ligne suivante
        rs.MoveNext
    Loop
    
    'Fermer les objets Recordset et Connection
    rs.Close
    sourceConn.Close
    destConn.Close
    
    Set destConn = Nothing
    Set columnMapping = Nothing
    Set rs = Nothing
    Set sourceConn = Nothing
    
    MsgBox "Les donn�es ont �t� copi�es avec succ�s !"
    
End Sub


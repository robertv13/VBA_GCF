Attribute VB_Name = "modGL_Stats_CA"
Option Explicit

Sub shp_GL_PrepEF_Actualiser_Click()

    Dim ws As Worksheet
    Set ws = wshGL_PrepEF
    
    Call Actualiser_Stats_CA
    
End Sub
Sub Actualiser_Stats_CA()

    Dim ws As Worksheet
    Set ws = wshGL_Stats_CA
    
    'Postes de revenus à considérer dans les REVENUS
    Dim glREV(1 To 2) As String
    Dim GL_Revenus_Consultation As String
    glREV(1) = ObtenirNoGlIndicateur("Revenus de consultation")
    Dim GL_Revenus_TEC As String
    glREV(2) = ObtenirNoGlIndicateur("Revenus - Travaux en cours")
    
    'Déterminer le dernier mois complété
    Dim moisPrécédent As Integer
    moisPrécédent = month(DateSerial(year(Date), month(Date), 0))
    Dim dateFinMoisPrécédent As Date
    dateFinMoisPrécédent = DateSerial(year(Date), month(Date), 0)
    
    Dim moisFinAnnéeFinancière As Integer
    moisFinAnnéeFinancière = wshAdmin.Range("MoisFinAnnéeFinancière").Value
    
    'Mémoriser les colonnes pour chacun des 12 mois de l'année financière
    Dim colMois(1 To 12, 1 To 2) As String
    Dim m As Integer, aaf As Integer, maf As Integer, col As Integer
    aaf = ws.Range("C9").Value - 1
    col = 4
    For m = 1 To 12
        maf = m + moisFinAnnéeFinancière
        If maf > 12 Then
            maf = maf - 12
            aaf = aaf + 1
        End If
        colMois(m, 1) = col
        colMois(m, 2) = Format$(aaf, "0000") & "-" & Format$(maf, "00")
        Debug.Print m, col, Format$(aaf, "0000") & "-" & Format$(maf, "00"), colMois(m, 2)
        col = col + 1
    Next m
    
    Dim dateFinMois As Date
    Dim revenus As Double
    Dim r As Integer
    For m = 1 To 12
        col = colMois(m, 1)
        dateFinMois = DateSerial(year(colMois(m, 2)), month(colMois(m, 2)) + 1, 0)
        revenus = 0
        For r = 1 To UBound(glREV, 1)
            revenus = revenus + Fn_Get_GL_Trans_Total(glREV(r), dateFinMois)
        Next r
        Debug.Print m, col, dateFinMois, revenus
        ws.Cells(9, col).Value = revenus
    Next m

End Sub

Sub shp_GL_Stats_CA_Exit_Click()

    Call GL_Stats_CA_Back_To_Menu

End Sub

Sub GL_Stats_CA_Back_To_Menu()
    
    wshGL_Stats_CA.Visible = xlSheetHidden
    
    wshMenuGL.Activate
    wshMenuGL.Range("A1").Select
    
End Sub


Sub DEB_Saisie_Update()

    If wshDEB_Saisie.Shapes("btnUPdate").TextFrame2.TextRange.Text = "Renversement" Then
        Call DEB_Renversement_Update
        Exit Sub
    End If
    
    Dim startTime As Double: startTime = Timer: Call Log_Record("modDEB_Saisie:DEB_Saisie_Update", "", 0)
    
    'Remove highlight from last cell
    If wshDEB_Saisie.Range("B4").Value <> "" Then
        wshDEB_Saisie.Range(wshDEB_Saisie.Range("B4").Value).Interior.Color = xlNone
    End If
    
    'Date is not valid OR the transaction does not balance
    If Fn_Is_Date_Valide(wshDEB_Saisie.Range("O4").Value) = False Or _
        Fn_Is_Debours_Balance = False Then
            Exit Sub
    End If
    
    'Is every line of the transaction well entered ?
    Dim rowDebSaisie As Long
    rowDebSaisie = wshDEB_Saisie.Range("E23").End(xlUp).row  'Last Used Row in wshDEB_Saisie
    If Fn_Is_Deb_Saisie_Valid(rowDebSaisie) = False Then Exit Sub
    
    'Get the FournID
    wshDEB_Saisie.Range("B5").Value = Fn_GetID_From_Fourn_Name(wshDEB_Saisie.Range("J4").Value)

    'Transfert des données vers DEB_Trans, entête d'abord puis une ligne à la fois
    Call DEB_Trans_Add_Record_To_DB(rowDebSaisie)
    Call DEB_Trans_Add_Record_Locally(rowDebSaisie)
    
    'GL posting
    Call DEB_Saisie_GL_Posting_Preparation
    
    If wshDEB_Saisie.ckbRecurrente = True Then
        Call Save_DEB_Recurrent(rowDebSaisie)
    End If
    
    'Retrieve the CurrentDebours number
    Dim CurrentDeboursNo As String
    CurrentDeboursNo = wshDEB_Saisie.Range("B1").Value
    
    MsgBox "Le déboursé, numéro '" & CurrentDeboursNo & "' a été reporté avec succès"
    
    'Get ready for a new one
    Call DEB_Saisie_Clear_All_Cells
    
    Application.EnableEvents = True
    
    wshDEB_Saisie.Activate
    wshDEB_Saisie.Range("F4").Select
        
    Call Log_Record("modDEB_Saisie:DEB_Saisie_Update", "", startTime)
        
End Sub


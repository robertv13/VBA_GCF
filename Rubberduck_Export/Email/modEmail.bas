Attribute VB_Name = "modEmail"
Option Explicit

Sub EmailSignatureTXT()

    'Create an Email with a TXT version of GCF signature - 2023-10-09
    Dim var As String
    Dim O As Outlook.Application
    Set O = New Outlook.Application

    Dim oMail As Outlook.MailItem
    Set oMail = O.CreateItem(olMailItem)
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim myFile As Scripting.TextStream
    Set myFile = fso.GetFile("C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Signatures\" & _
                 "GCF (robertv13@me.com).txt").OpenAsTextStream(ForReading, TristateTrue)

    var = myFile.ReadAll
    
    With oMail

        .To = "robertv13@me.com"
        .subject = "Test de la routine 'EmailSignature' du module 'modEmail'"
        .Body = "Bonjour," & vbNewLine & _
            vbNewLine & _
            "Ci-joint, le rapport quotidien." & _
            vbNewLine & _
            vbNewLine & _
            var
        .Display

    End With
    
    Set O = Nothing
    Set oMail = Nothing

End Sub

Sub EmailSignatureHTML()

    'Create an Email with a HTML version of GCF signature - 2023-10-09 - YouTube LearnAccess
    'https://www.youtube.com/watch?v=RB54YTVTmRk
    
    'Déclaration des variables
    Dim MaMessagerie As Object
    Dim MonMessage As Object
    Dim MaSignature As String
    Dim MonFichier As String
    
    'On affecte les variables de type Object - Initialisations (SET)
    Set MaMessagerie = CreateObject("Outlook.Application")
    Set MonMessage = MaMessagerie.CreateItem(0)
    
    'On affiche le courriel
    MonMessage.Display
    'On récupère la signature
    MaSignature = MonMessage.HTMLBody
        
    'On construit le message
    With MonMessage
        'Adresse du destinataire
        .To = "robertv13@me.com"
        .CC = "robertv13@gmail.com"
        .BCC = "robertv13@hotmail.com; marie.guay@outlook.com"
        .subject = "Exemple de courriel (format RTF)"
        'Corps du courriel
        .HTMLBody = "Bonjour <b>" & "Robert" & ",</b><br><br>" & _
                    "Veuillez trouver ci-joint le tableau de bord" & "<br>" & _
                    "<ul>" & _
                    "<li><b style='color:red;'>" & "Point n° 1" & "</b></li>" & _
                    "<li><b style='color:green;'>" & "Point n° 2" & "</b></li>" & _
                    "<li><b style='color:blue;'>" & "Point n° 3" & "</b></li>" & _
                    "</ul>" & _
                    MaSignature
        'On récupère le fichier à joindre au courriel
        MonFichier = "C:\VBA\GC_FISCALITÉ\Factures_PDF\" & "00031.pdf"
        'MonFichier = ActiveWorkbook.Path & "\" & "Factures_PDF\" & "00031.pdf"
        'Insertion d'une pièce jointe
        .Attachments.Add (MonFichier)
        'Affiche le courriel
'        .Display
        'Envoi direct du courriel
        .Send
    
    End With
    
    'Libérer de la mémoire des objets
    Set MaMessagerie = Nothing
    Set MonMessage = Nothing
  
End Sub

Sub EmailSignatureImage()

    'YouTube: https://www.youtube.com/watch?v=VBlmJxgJp8s - Ajay Kumar
    'Learn Excel - Video 501 - VBA - How to add signature logo in outlook emails

    Dim O As Object
    Set O = New Outlook.Application
    
    Dim oMail As Outlook.MailItem
    Set oMail = O.CreateItem(olMailItem)
    
'    Dim p As String
'    p = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Signatures\GCF (robertv13@me.com)_fichiers\image003.png"
    
    With oMail
        .To = "robertv13@me.com"
        .CC = "robertv13@gmail.com"
        .BCC = "robertv13@hotmail.com"
        .subject = "Test d'envoi de courriel avec image en signature"
        .HTMLBody = "Bonjour," & "<br><br>" & "Vous trouverez ci-joint une copie du rapport." & _
                    "<br><br><br><br><br><br>" & _
                    "<img src = 'C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Signatures\GCF (robertv13@me.com)_fichiers\image003.png' />"
        'Utilisation de la variable 'p' pour le l'emplacement du fichier - NE FONCTIONNE PAS !!!
'        .HTMLBody = "Bonjour," & "<br><br>" & "Vous trouverez ci-joint une copie du rapport." & _
'                    "<br><br><br><br><br></br>" & _
'                    "<img src = " & "" & p & "" & "/>"
        .Display
    End With
    
    Set O = Nothing
    Set oMail = Nothing
    
End Sub

Sub EmailSpecificSignatureHTML() 'Good Routine - 2023-10-09 @ 17:32
    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    Dim strbody As String
    Dim sigString As String
    Dim signature As String

    'Application.ScreenUpdating = False
    Set OutApp = CreateObject("Outlook.Application")
    sigString = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Signatures\GCF (robertv13@me.com).htm"
    If Dir(sigString) <> "" Then
        signature = GetBoiler(sigString)
    Else
        signature = ""
    End If
    On Error GoTo cleanup
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        .To = "robertv13@me.com"
        .subject = "Test d'envoi de courriel avec le format HTML et signature spécifique"
        .HTMLBody = "Bonjour, " & "<br><br>" & _
                    "Vous trouverez ci-joint notre note d'honoraires" & "<br><br><br><br>" & _
                    signature
        .Display  'Or use Send
    End With
    On Error GoTo 0

    Set OutMail = Nothing

cleanup:
    Set OutApp = Nothing
    Set OutMail = Nothing
    'Application.ScreenUpdating = True

End Sub

Sub EmailSpecificSignatureFullHTML()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    Dim strbody As String
    Dim sigString As String
    Dim signature As String

    'Application.ScreenUpdating = False
    Set OutApp = CreateObject("Outlook.Application")
    sigString = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Signatures\GCF (robertv13@me.com).htm"
    If Dir(sigString) <> "" Then
        signature = GetBoiler(sigString)
    Else
        signature = ""
    End If
    On Error GoTo cleanup
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        .To = "robertv13@me.com"
        .subject = "Test d'envoi de courriel avec le format HTML et signature spécifique"
        .HTMLBody = "<p style='margin-right:0cm;margin-left:0cm;font-size:15px;font-family:""Calibri"",sans-serif;margin:0cm;margin-bottom:12.0pt;'>" & _
                    "<span style=""color:black;"">Bonsoir Christian,</span></p>" & _
                    "<p style='margin-right:0cm;margin-left:0cm;font-size:15px;font-family:""Calibri"",sans-serif;margin:0cm;margin-top:6.0pt;margin-bottom:6.0pt;'><span style=""color:black;"">J&rsquo;en suis rendu &agrave; mettre &agrave; jour ma facturation. Tu trouveras donc ci-jointe ma facturation &agrave; jour en date d&apos;aujourd&apos;hui.</span></p>" & _
                    "<p style='margin-right:0cm;margin-left:0cm;font-size:15px;font-family:""Calibri"",sans-serif;margin:0cm;margin-top:6.0pt;margin-bottom:6.0pt;'><span style=""color:black;"">N&apos;h&eacute;site pas s&apos;il y a quoique ce soit!</span></p>" & _
                    signature
        .Display  'Or use Send
    End With
    On Error GoTo 0

    Set OutMail = Nothing

cleanup:
    Set OutApp = Nothing
    Set OutMail = Nothing
    'Application.ScreenUpdating = True

End Sub

Function GetBoiler(ByVal sFile As String) As String
'**** Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function

Sub Email_FromTemplate() '2023-10-10 @ 11h39
    'Create/Send an Email from a Outlook Template, with specific signature
    Dim adrBCC As String
    Dim adrCC As String
    Dim adrTo As String
    Dim pieceJointe As String
    Dim path As String
    Dim signature As String
    Dim sigString As String
    Dim sujet, prenom As String
    Dim pathEtPieceJointe As String
    
    Dim myAttachments As Object
    Dim outlookInsp As Object
    Dim oRang As Object
    Dim outlookApp As Object
    Dim outlookMailItem As Object
    Dim wordDoc As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    'Call the template
    Set outlookMailItem = outlookApp.CreateItemFromTemplate("C:\VBA\GC_FISCALITÉ\Template_Courriel_Envoi_De_Facture.oft")
    Set myAttachments = outlookMailItem.Attachments
    sigString = "C:\Users\Robert M. Vigneault\AppData\Roaming\Microsoft\Signatures\GCF (robertv13@me.com).htm"
    If Dir(sigString) <> "" Then
        signature = GetBoiler(sigString)
    Else
        signature = ""
    End If
    On Error GoTo cleanup
    
    adrTo = "robertv13@me.com"
    adrCC = ""
    adrBCC = ""
    prenom = "Christian"
    sujet = "TEST - Facturation - Fiscalité - TEST"
    'Define path for the attachments
    path = "C:\VBA\GC_FISCALITÉ\Factures_PDF\"
    pieceJointe = "00025.pdf"
    pathEtPieceJointe = path & pieceJointe
    
    outlookMailItem.Display
    outlookMailItem.HTMLBody = Replace(outlookMailItem.HTMLBody, "Robert M. Vigneault", "")

    With outlookMailItem
        .To = adrTo
        .CC = adrCC
        .BCC = adrBCC
        .subject = sujet
        .Attachments.Add pathEtPieceJointe, 1
        Set outlookInsp = outlookMailItem.GetInspector
        Set wordDoc = outlookInsp.WordEditor
        Set oRang = wordDoc.Range
        With oRang.Find
            Do While .Execute(FindText:="{{Prenom}}")
                oRang.Text = prenom
                Exit Do
            Loop
        End With
        .HTMLBody = .HTMLBody & signature
        .Display
    End With

cleanup:
    Set myAttachments = Nothing
    Set outlookInsp = Nothing
    Set oRang = Nothing
    Set outlookApp = Nothing
    Set outlookMailItem = Nothing
    Set wordDoc = Nothing

End Sub

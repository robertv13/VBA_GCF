Attribute VB_Name = "Module1"
Option Explicit

Sub NouvelleRechercheClick()

    With wshCode
        .Range("F2:H2").ClearContents
        .Range("F4:F6").ClearContents
        .Range("F8").ClearContents
        .Range("J2, J4, J6, J8").ClearContents
        .Range("F11:L999").ClearContents
        .Range("F2").Activate
    End With

End Sub

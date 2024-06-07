Attribute VB_Name = "modGL_Rapport"
Option Explicit

Public Sub GL_Report_For_Selected_Accounts()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_Rapport:GL_Report_For_Selected_Accounts()")
   
    Dim ws As Worksheet
    Dim lb As OLEObject
    Dim selectedItems As Collection
    Dim i As Integer

    ' Reference the worksheet and ListBox
    Set ws = ThisWorkbook.Sheets("GL_Rapport") ' Adjust to your sheet name
    Set lb = ws.OLEObjects("ListBox1")

    ' Ensure it is a ListBox
    If TypeName(lb.Object) = "ListBox" Then
        Set selectedItems = New Collection

        ' Loop through ListBox items and collect selected ones
        With lb.Object
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    selectedItems.add .List(i)
                End If
            Next i
        End With

        ' Output the selected items
        Dim item As Variant
        For Each item In selectedItems
            Debug.Print item
        Next item
    End If
    
    Call Output_Timer_Results("modGL_Rapport:GL_Report_For_Selected_Accounts()", timerStart)

End Sub

Sub GL_Rapport_Back_To_Menu()
    
    Dim timerStart As Double: timerStart = Timer: Call Start_Routine("modGL_Rapport:GL_Rapport_Back_To_Menu()")
   
    wshMenuCOMPTA.Activate
    wshMenuCOMPTA.Range("A1").Select
    
    Call Output_Timer_Results("modGL_Rapport:GL_Rapport_Back_To_Menu()", timerStart)
    
End Sub


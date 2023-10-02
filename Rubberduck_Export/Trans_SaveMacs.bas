Attribute VB_Name = "Trans_SaveMacs"
Option Explicit

Sub Trans_Save()
    Dim LastTransRow As Long
    Dim SelRow As Long
    Dim TransRow As Long
    Dim TransNumb As Long
    
    With Sheet1
        SelRow = .Range("B2").Value 'Selected Row
        'Check for balancing Transaction ?
        If .Range("G" & SelRow).Value <> .Range("H" & SelRow + 1).Value Or .Range("G" & SelRow + 1).Value <> .Range("H" & SelRow).Value Then
            MsgBox "Veuillez vous assurez que l'écriture balance (Débits = Crédits)"
            Exit Sub
        End If
        'Check for MINIMUM 2 accounts
        If .Range("E" & SelRow).Value = Empty Or .Range("E" & SelRow + 1).Value = Empty Then
            MsgBox "Au minimum, deux comptes sont nécessaires (DE/À)"
            Exit Sub
        End If
        'Check for Date
        If .Range("D" & SelRow).Value = Empty Then
            MsgBox "Veuillez entrer une date valide"
            Exit Sub
        End If
        'New or existing Transaction?
        If .Range("B4").Value = Empty Then 'New Transaction
                TransRow = Sheet3.Range("D999999").End(xlUp).Row + 1
                TransNumb = .Range("B7").Value 'New Transaction #
        Else: 'Existing Transaction
            TransRow = .Range("B4").Value 'Existing Transaction Row
            TransNumb = .Range("B3").Value 'Existing Transaction #
        End If
        
        'Transactions Worksheet
        Sheet3.Range("D" & TransRow, "D" & TransRow + 1).Value = TransNumb 'Transaction Number
        Sheet3.Range("E" & TransRow, "E" & TransRow + 1).Value = .Range("D" & SelRow).Value 'Transaction Date
        Sheet3.Range("F" & TransRow, "F" & TransRow + 1).Value = .Range("D" & SelRow + 1).Value 'Transaction Type
        Sheet3.Range("G" & TransRow, "G" & TransRow + 1).Value = .Range("F" & SelRow).Value 'Name/Vendor
    
        If .Range("G" & SelRow).Value <> "" Then
            Sheet3.Range("H" & TransRow, "H" & TransRow + 1).Value = .Range("E" & SelRow).Value 'Debit Account
            Sheet3.Range("I" & TransRow, "I" & TransRow + 1).Value = .Range("E" & SelRow + 1).Value 'Credit Account
        Else:
            Sheet3.Range("H" & TransRow, "H" & TransRow + 1).Value = .Range("E" & SelRow + 1).Value 'Credit Account
            Sheet3.Range("I" & TransRow, "I" & TransRow + 1).Value = .Range("E" & SelRow).Value 'Debit Account
        End If
    
        'Debit and Credit Amounts
        If .Range("G" & SelRow).Value <> "" Then
            Sheet3.Range("J" & TransRow).Value = .Range("G" & SelRow).Value         'Debit Amount
            Sheet3.Range("K" & TransRow + 1).Value = .Range("H" & SelRow + 1).Value 'Credit Amount
            Sheet3.Range("J" & TransRow + 1).ClearContents                          'Clear Other Two Fields
            Sheet3.Range("K" & TransRow).ClearContents                              'Clear Other Two Fields
        Else:
            Sheet3.Range("J" & TransRow + 1).Value = .Range("G" & SelRow + 1).Value 'Debit Amount
            Sheet3.Range("K" & TransRow).Value = .Range("H" & SelRow).Value         'CreditAmount
            Sheet3.Range("J" & TransRow).ClearContents                              'Clear Other Two Fields
            Sheet3.Range("K" & TransRow + 1).ClearContents                          'Clear Other Two Fields
        End If
    
        'Memo
        Sheet3.Range("M" & TransRow, "M" & TransRow + 1).Value = .Range("F" & SelRow + 1).Value
        .Range("C" & SelRow).Value = TransNumb
        .Shapes("SaveBtn").Visible = msoFalse 'Hide save button
        .Range("B5").Value = SelRow
        Application.Wait Now + TimeValue("00:00:02")
        .Range("B5").ClearContents
    End With
    
    'Transactions Worksheet
    With Sheet3
        'Resort tramsactions list
        LastTransRow = .Range("D999999").End(xlUp).Row
        'Sort List Based on Date
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.Range("E5"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range("D5:M" & LastTransRow)
            .Apply
        End With
    
        ResetCalc
    
    End With

End Sub

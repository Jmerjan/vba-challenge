Attribute VB_Name = "Module2"
Sub Bonus()

    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        'Bonus Column Headers
        WS.Cells(1, "P").Value = "Ticker"
        WS.Cells(1, "Q").Value = "Value"
        WS.Cells(2, "O").Value = "Greatest % Increase"
        WS.Cells(3, "O").Value = "Greatest % Decrease"
        WS.Cells(4, "O").Value = "Greatest Total Volume"
        

        Dim x As Long
        
        'Bonus
        ' Determine Last Row for the Summarized Data
        
        ILastRow = WS.Cells(Rows.Count, "I").End(xlUp).Row
    
        For x = 2 To ILastRow
            If Cells(x, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & ILastRow)) Then
                Cells(2, "P").Value = Cells(x, "I").Value
                Cells(2, "Q").Value = Cells(x, "K").Value
                Cells(2, "Q").NumberFormat = "0.00%"
            ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & ILastRow)) Then
                Cells(3, "P").Value = Cells(x, "I").Value
                Cells(3, "Q").Value = Cells(x, "K").Value
                Cells(3, "Q").NumberFormat = "0.00%"
            ElseIf Cells(x, "L").Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & ILastRow)) Then
                Cells(4, "P").Value = Cells(x, "I").Value
                Cells(4, "Q").Value = Cells(x, "L").Value
            End If
        Next x
        
    Next WS
        
    
End Sub

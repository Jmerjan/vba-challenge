Attribute VB_Name = "Module1"
Sub Stocks()

   'Make Loop to go through all sheets
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    Debug.Print WS.Name
    

        ' Column Heading
        WS.Cells(1, "I").Value = "Ticker"
        WS.Cells(1, "J").Value = "Yearly Change"
        WS.Cells(1, "K").Value = "Percent Change"
        WS.Cells(1, "L").Value = "Total Stock Volume"
        WS.Cells(2, "O").Value = "Greatest % Increase"
        WS.Cells(3, "O").Value = "Greatest % Decrease"
        WS.Cells(4, "O").Value = "Greatest Total Volume"
        WS.Cells(1, "P").Value = "Ticker"
        WS.Cells(1, "Q").Value = "Value"
        
        
        'Create Variables
        Dim Ticker As String
        Dim Year_Open As Double
        Dim Year_Close As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Sum_Row As Double
        Sum_Row = 2
        Dim i As Long
        Dim Start As Long
        Start = 2
        Dim x As Long
        
        
        
        ' Determine the Last Sum_Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
 
        For i = 2 To LastRow
         ' Check if we are still within the same ticker symbol, if it is not...
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                
                ' Set Ticker name
                Ticker = WS.Cells(i, 1).Value
                WS.Cells(Sum_Row, 9).Value = Ticker
                
                If WS.Cells(Start, 3).Value = 0 Then
                    For FindValue = Start To i
                        If WS.Cells(FindValue, 3).Value <> 0 Then
                            Start = FindValue
                            Exit For
                        End If
                    Next FindValue
                End If
                
                ' Set Close Price
                Year_Close = WS.Cells(i, 6).Value
                
                ' Add Yearly Change
                Yearly_Change = WS.Cells(i, 6).Value - WS.Cells(Start, 3).Value
                WS.Cells(Sum_Row, "J").Value = Yearly_Change
                WS.Cells(Sum_Row, "J").NumberFormat = "$ 0.00"
                
                'Find Percent Change'
                Percent_Change = Yearly_Change / WS.Cells(Start, 3).Value
                WS.Cells(Sum_Row, 11).Value = Percent_Change
                WS.Cells(Sum_Row, 11).NumberFormat = "0.00%"
                
               'Create Color Index'
                If WS.Cells(Sum_Row, "J").Value >= 0 Then
                    WS.Cells(Sum_Row, "J").Interior.ColorIndex = 10
                Else
                    WS.Cells(Sum_Row, "J").Interior.ColorIndex = 3
                End If
                
                Start = i + 1
                ' Add Total Volume
                Volume = Volume + WS.Cells(i, 7).Value
                WS.Cells(Sum_Row, "L").Value = Volume
                
                ' Add one to the summary table Sum_Row
                Sum_Row = Sum_Row + 1
                
                'Reset Yearly Change
                Yearly_Change = 0
                
                'Reset Volume Total
                Volume = 0

            Else
                Volume = Volume + WS.Cells(i, 7).Value
            End If
        Next i
           
        'Whenever I run the bonus in this formula I get an overflow error
        'I made a second module for for Bonus
        
        
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



Sub StockCrunch():
    'Pull ticker symbol, Yearly change, Percent change, and stock volume
    'This VBA script assumes that each sheet has headers and each sheet is sorted by <ticker> and then <date>
    
    
    ' For each Loop - run scipt on all worksheets
    For Each ws In Worksheets
    
        ' Add Headers and other static cells for the Summary tables
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
     
        'Declare
        Dim Lastrow As Double
        Dim Stock_Name As String
        Dim Volume_Total As Double
        Dim Summary_Row As Integer
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Summary_Lastrow As Long
        Dim MaxIncrease As Double
        Dim MaxDecrease As Double
        Dim MaxVolume As Double
        
                
        'Assign
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Volume_Total = 0
        Summary_Row = 2
        Open_Price = ws.Cells(2, 3)
        
        
        'Loop through ticker(column A) to find each ticker symbol, yearly change, percent change, and total stock volume
        
        For i = 2 To Lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Assign Stock name
                Stock_Name = ws.Cells(i, 1).Value
                
                'Sum of Stock volume
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
                'record the stock name
                ws.Range("I" & Summary_Row).Value = Stock_Name
                
                'record the stock volume total
                ws.Range("L" & Summary_Row).Value = Volume_Total
                
                'reset volume total
                Volume_Total = 0
                
                'Set Close Price
                Close_Price = ws.Cells(i, 6).Value
                
                'check for no Open price within the year
                If ws.Cells(i, 3).Value = 0 Then
                    Open_Price = 0
                
                End If
                
                
                'calculate yearly change
                Yearly_Change = Close_Price - Open_Price
                ws.Range("J" & Summary_Row).Value = Yearly_Change
                
                'calculate Percent change
                If Open_Price > 0 Then
                    Percent_Change = Yearly_Change / Open_Price
                    
                ElseIf Open_Price = 0 Then
                    Percent_Change = 0
                    
                End If
                
                'record percent change and format cell
                ws.Range("K" & Summary_Row).Value = Percent_Change
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                
                'reset Open_price if Open price is not zero
                If ws.Cells(i + 1, 3) > 0 Then
                    Open_Price = ws.Cells(i + 1, 3)
                    
                End If
                
                'advance to the next row in summary
                Summary_Row = Summary_Row + 1
                                                      
                                                      
            Else
                'Sum of stock volume
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
                'reset Open_price if first occurance of Open price is not zero
                If ws.Cells(i, 3) > 0 And ws.Cells(i - 1, 3) = 0 Then
                    Open_Price = ws.Cells(i, 3)
                
                End If
                            
            End If
            
        Next i
        
        
        
        'Format Yearly Change column (column J):  positive change in green and negative change in red
        
        ' Find the last row of the ticker summary column (column I)
        Summary_Lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To Summary_Lastrow
        
            'Set positve Yearly Change to green
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            'set negative Yearly Change to red
            ElseIf ws.Cells(j, 10) < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
            
            End If
                            
        Next j
        
        
        'Find Vaule of Greatest % increase, Greatest % decrease and Greatest total volume
        ws.Cells(2, 17).Formula = "=Max(K:K)"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        MaxIncrease = ws.Cells(2, 17).Value
        
        
        ws.Cells(3, 17).Formula = "=Min(K:K)"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        MaxDecrease = ws.Cells(3, 17).Value
        
        
        ws.Cells(4, 17).Formula = "=Max(L:L)"
        MaxVolume = ws.Cells(4, 17).Value
        
        
        
        'Find Ticker symbol for Greatest % increase
        For k = 2 To Summary_Lastrow
        
            If ws.Cells(k, 11).Value = MaxIncrease Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
            End If
            
        Next k
        
        
        'Find Ticker symbol for Greatest % decrease
        For m = 2 To Summary_Lastrow
            
            If ws.Cells(m, 11).Value = MaxDecrease Then
            ws.Cells(3, 16).Value = ws.Cells(m, 9).Value
            
            End If
            
        Next m
        
        
        'Find Ticker symbol for Greatest total volume
        For n = 2 To Summary_Lastrow
            
            If ws.Cells(n, 12).Value = MaxVolume Then
            ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
            
            End If
            
        Next n
        
        'Resize columns for easy viewing
        ws.Columns("A:Q").AutoFit

        
    Next ws
    
End Sub

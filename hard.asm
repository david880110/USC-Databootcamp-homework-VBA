Sub hard()

    For Each ws In Worksheets
    
             
        'Get last row
            
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'add Headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest % Total Volume"
                
        'declare variables
                   
        Dim ticker As String
        Dim ttl_volume As Double
        Dim current_row As Integer
        
        Dim first_close, first_open As Double
        Dim yr_percentage As Double

        Dim ticker_increase As String
        Dim ticker_decrease As String
        Dim ticker_volume As String
        
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double

        'set initial values
        current_row = 2
        ttl_volume = 0
        
        first_close = ws.Range("F2").Value
        first_open = ws.Range("C2").Value
        
        greatest_volume = ws.Cells(2, 10).Value
        greatest_increase = 0
        greatest_decrease = 0
                        
        For Row = 2 To lrow
        
            'If the next row contents a same ticker as the current one, then...
            
            If (ws.Cells(Row + 1, 1).Value = ws.Cells(Row, 1).Value) Then
            
            'Add to the ttl_volume
            
            ttl_volume = ttl_volume + ws.Cells(Row, 7).Value
                                           
            'If the next row contents a different ticker then the current one, then...
            
            Else
                
            'print the Ticker and the Total Volume.
            
            ws.Range("I" & current_row).Value = ws.Cells(Row, 1).Value
            ws.Range("L" & current_row).Value = ttl_volume
            ws.Range("J" & current_row).Value = yr_change
            ws.Range("K" & current_row).Value = FormatPercent(yr_percentage)
            
            'calculate the yearly change and yearly percentage change and print it out
                
            yr_percentage = (ws.Cells(Row, 6).Value - first_open) / first_open
            yr_change = (ws.Cells(Row, 6).Value - first_open)
                                
                'If yearly change is negative set the cell color to red if it is positive, to green.
              
                If (ws.Cells(current_row, 10).Value < 0) Then
                    ws.Range("J" & current_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & current_row).Interior.ColorIndex = 4
                End If
            
            'calculate greatest percentage increase, greatest percentage decrease, greatest_total volume
            
            If ttl_volume > greatest_volume Then
            
            greatest_volume = ttl_volume
            ticker_volume = ws.Cells(Row, 1).Value

            End If

            If yr_percentage > greatest_increase Then
                
            greatest_increase = yr_percentage
            ticker_increase = ws.Cells(Row, 1).Value
            
            ElseIf yr_percentage < greatest_decrease Then

            greatest_decrease = yr_percentage
            ticker_decrease = ws.Cells(Row, 1).Value

            End If
            
            'reset the ttl_volume and increase current_row and get  the first_close of next ticker
               
            ttl_volume = 0
            current_row = current_row + 1
            first_close = ws.Cells(Row + 1, 6)
            
            End If
 
        Next Row

            'Print Out greatest percentage increase, greatest percentage decrease, greatest_total volume
            
            ws.Range("P2").Value = ticker_increase
            ws.Range("P3").Value = ticker_decrease
            ws.Range("P4").Value = ticker_volume
            ws.Range("Q2").Value = FormatPercent(greatest_increase)
            ws.Range("Q3").Value = FormatPercent(greatest_decrease)
            ws.Range("Q4").Value = greatest_volume
 
    Next ws

End Sub


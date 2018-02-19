Sub moderate()

    For Each ws In Worksheets
    
             
        'Get last row
            
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'add Headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
                
        'declare variables
                   
        Dim ticker As String
        Dim ttl_volume As Double
        
        Dim first_close, first_open As Double
        Dim yr_percentage As Double
        Dim current_row As Integer

        current_row = 2
        first_close = ws.Range("F2").Value
        first_open = ws.Range("C2").Value

        ttl_volume = 0
                
        For Row = 2 To lrow
        
			'If the next row contents a same ticker as the current one, then...
			
            If (ws.Cells(Row + 1, 1).Value = ws.Cells(Row, 1).Value) Then
			
			'Add to the ttl_volume
			
            ttl_volume = ttl_volume + ws.Cells(Row, 7).Value
                           

                
            'If the next row contents a different ticker then the current one, then...
			
            Else
                
			'Print the Ticker and the Total Volume.          
			
            ws.Range("I" & current_row).Value = ws.Cells(Row, 1).Value
            ws.Range("L" & current_row).Value = ttl_volume
            ws.Range("J" & current_row).Value = yr_change
			ws.Range("K" & current_row).Value = FormatPercent(yr_percentage)
				
            'Calculate the yearly change and yearly percentage change and print it out
                
            yr_percentage = (ws.Cells(Row, 6).Value - first_open) / first_open
            yr_change = (ws.Cells(Row, 6).Value - first_open)
                                
				'If yearly change is negative set the cell color to red if it is positive, to green.
              
                If (ws.Cells(current_row, 10).Value < 0) Then
                    ws.Range("J" & current_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & current_row).Interior.ColorIndex = 4
                End If
                
                                
            'Reset the ttl_volume and increase current_row and get  the first_close of next ticker
               
            ttl_volume = 0
            current_row = current_row + 1
            first_close = ws.Cells(Row + 1, 6)    

            
            End If
       
        Next Row
        

 
    Next ws

End Sub

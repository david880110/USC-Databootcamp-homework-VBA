Sub easy()

    For Each ws In Worksheets
    
             
        '- Get the Last Row
            
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).row

        ' Add Headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
                
        '- Declare variables
              
        Dim ticker As String
        Dim ttl_volume As Double

        Dim current_row As Integer

        'set initial value to current_row and ttl_volume         
        current_row = 2
        ttl_volume = 0
                
        For row = 2 To lrow
        
			' if the next row contents a same ticker as the current one, then...
            If ws.Cells(row + 1, 1).Value = ws.Cells(row, 1).Value Then
            
                'Add the Total Volume
                
                ttl_volume = ttl_volume + ws.Cells(row, 7).Value
                
                
            ' If the next row contents a different ticker, then...
            Else
                
                '- Retrive the Ticker and the Total Volume.
                
                ws.Range("I" & current_row).Value = ws.Cells(row, 1).Value
                ws.Range("J" & current_row).Value = ttl_volume                
                                              
                ' Reset the ttl_volume and increase current_row 
               
                ttl_volume = 0
                current_row = current_row + 1
            
            End If
       
        Next row
        
    Next ws

End Sub

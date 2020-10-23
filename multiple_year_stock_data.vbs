Sub multiyear_stockdata()

  For Each ws In Worksheets
    
    Dim stock_name As String
    
    Dim total_volume As Double
    
    Dim display_row As Integer
    
    Dim lrow As Long
    
    Dim yearly_change As Double
    
    Dim open_price As Double
    
    Dim close_price As Double
    
    Dim percent_change As Double
        
        
    total_volume = 0
    
    display_row = 2
    
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    open_price = ws.Cells(2, 3).Value
    
    
    ws.Cells(1, 9).Value = "Ticker"

    ws.Cells(1, 10).Value = "Total Stock Volume"

    ws.Cells(1, 11).Value = "Yearly Change"
    
    ws.Cells(1, 12).Value = "Percent Change"


    ws.Columns("J").NumberFormat = "#,##0"    
                
    ws.Columns("K").NumberFormat = "$#,##0.00_);($#,##0.00)"    
            
    ws.Columns("L").NumberFormat = "0.00%"

       
    
    
    For i = 2 To lrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            stock_name = ws.Cells(i, 1).Value
            
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            
            ws.Cells(display_row, 9).Value = stock_name
            
            ws.Cells(display_row, 10).Value = total_volume
            
            
            
            close_price = ws.Cells(i, 6).Value            
            
            yearly_change = close_price - open_price
           
            ws.Cells(display_row, 11).Value = yearly_change
            
            
           
           
                If open_price = 0 Then

                    ws.Cells(display_row, 12).Value = 0

                Else

                    percent_change = yearly_change / open_price
              
                    ws.Cells(display_row, 12).Value = percent_change
                    
                
                End If
                
            
            
             open_price = ws.Cells(i + 1, 3).Value
                             
             display_row = display_row + 1
            
             total_volume = 0
            
            
        Else
        
            total_volume = total_volume + ws.Cells(i, 7).Value
            
        
        End If
        
    Next i
    
                ws.Cells(2, 14).Value = "Greatest % Increase"

                ws.Cells(3, 14).Value = "Greatest % Decrease"
                
                ws.Cells(4, 14).Value = "Greatest Total Volume"
                
                
                
                ws.Cells(1, 15).Value = "Ticker"
                
                ws.Cells(1, 16).Value = "Value"
                
                
                ws.Cells(2, 16).NumberFormat = "0.00%"

                ws.Cells(3, 16).NumberFormat = "0.00%"
                
                ws.Cells(4, 16).NumberFormat = "#,##0"

                
                ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
                
                ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Range("L:L"))
                
                ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("J:J"))
                
                
                
    
            Dim lastrow As Long
            
            lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            
            For i = 2 To lastrow
            
            
                If ws.Cells(i, 12).Value = ws.Cells(2, 16).Value Then
                
                    ws.Cells(2, 15).Value = ws.Cells(i, 9)
                 
                 
                ElseIf ws.Cells(i, 12).Value = ws.Cells(3, 16).Value Then
                 
                    ws.Cells(3, 15).Value = ws.Cells(i, 9)
                 
                 
                ElseIf ws.Cells(i, 10).Value = ws.Cells(4, 16).Value Then
                 
                    ws.Cells(4, 15).Value = ws.Cells(i, 9)
                 
                 
                End If
                 



                If ws.Cells(i, 11).Value >= 0 Then

                    ws.Cells(i, 11).Interior.ColorIndex = 4

                Else

                    ws.Cells(i, 11).Interior.ColorIndex = 3

                End If
                
            Next i
                
                
    Next ws
                              
 End Sub
Attribute VB_Name = "Module1"
Sub tester()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

    'Define all variables'
    
        Dim ticker As String
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim price_change As Double
        price_change = 0
        Dim percent_change As Double
        percent_change = 0
        Dim stock_volume As Double
        stock_volume = 0
        
     'Define summary table'
        
        Dim summary_table_row As Integer
        summary_table_row = 1
        
    ' Define end of loop'
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      'loop'
      
        For i = 2 To lastrow
                'Ticker'
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                   summary_table_row = summary_table_row + 1
                   
                   ticker = ws.Cells(i, 1).Value
                   
                   ws.Cells(summary_table_row, "I").Value = ticker
                   
                 'define data for variables'
                 
                   open_price = ws.Cells(i, 3).Value
                   close_price = ws.Cells(i, 6).Value
                   stock_volume = ws.Cells(i, 7).Value
                   
                   stock_volume = stock_volume + ws.Cells(i, 7).Value
                   
                  'Set range for variable outputs'
                  
                   ws.Range("I" & summary_table_row).Value = ticker
                   ws.Range("J" & summary_table_row).Value = price_change
                   ws.Range("K" & summary_table_row).Value = percent_change
                   ws.Range("L" & summary_table_row).Value = stock_volume
                   
                   'calculate price change'
                   
                   price_change = (close_price - open_price)
                   
                     
                   ElseIf open_price <> 0 Then
                   
                   percent_change = (price_change / open_price) * 100
            
                    
            End If
       
        Next i
        
   Next ws
   
End Sub


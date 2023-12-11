VBA Script


'declare a function'

Sub alph_testing()

'start loop to go through every ws'
Dim ws As Worksheet
For Each ws In Worksheets



                ' Declaring variables and new columns'
               
                        ws.Range("P1").Value = "Ticker"
                        ws.Range("Q1").Value = "Value"
                        ws.Range("O2").Value = "Greatest % Increase"
                        ws.Range("O3").Value = "Greatest % Decrease"
                        ws.Range("O4").Value = "Greatest Total Volume"
                        ws.Range("I1").Value = "Ticker"
                        ws.Range("J1").Value = "Yearly Change"
                        ws.Range("K1").Value = "Percent Change"
                        ws.Range("L1").Value = "Total Stock Volume"
                        
                        
                    'variables for the price change'
                            Dim open_price As Double
                            open_price = 0
                            Dim close_price As Double
                            close_price = 0
                            Dim price_change As Double
                            price_change = 0
                            Dim price_change_percent As Double
                            price_change_percent = 0
                            
                
                    'variables for the ticker'
                              Dim TickerRow As Long: TickerRow = 1
                              Dim Ticker As String
                              Ticker = " "
                              Dim Ticker_volume As Double
                              Ticker_volume = 0
                                
                                                       
                                                        'Set initial and last row for worksheet
                                                        Dim Lastrow As Long
                                                        Dim i As Long
                                                      
                                
                                                        'Define Lastrow of worksheet and create loop
                                                        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                                                        
                                                        For i = 2 To Lastrow
                                                        
                                                        'Ticker symbol output
                                                        
                                                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                                        TickerRow = TickerRow + 1
                                                        Ticker = ws.Cells(i, 1).Value
                                                        ws.Cells(TickerRow, "I").Value = Ticker
                                                        
                                                        'Calculate change in Price
                                                        
                                                        close_price = ws.Cells(i, 6).Value
                                                        price_change_percent = close_price - open_price
                                                        
                                                        'Fixing the open price equal zero problem
                                                        
                                                        ElseIf open_price <> 0 Then
                                                        price_change_percent = (price_change_percent / open_price) * 100
                                                        
                                                        End If
                                                        
                                                        Next i
        
        
       

Next ws

End Sub






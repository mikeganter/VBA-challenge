Attribute VB_Name = "Module1"
Sub VBA_hw():

Dim ws As Worksheet
For Each ws In Worksheets

' define variable for last row to apply code to future data sets
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' define variables for calculations used for summary table
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim stock_volume As Double
Dim stock_ticker As String
Dim j As Integer

' set initial values for variables outside for loop
open_price = ws.Cells(2, 3).Value
stock_volume = 0
stock_ticker = ws.Cells(2, 1).Value
j = 1

' each loop, (i) add to stock volume counter; (ii) check to see if next row contains same ticker; (iii) if new ticker in following row, complete calculations and send values to the summary table
For i = 2 To lastrow
'For i = 2 To 264

        
        
        ' add stock volume in current row into counter
        stock_volume = stock_volume + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                close_price = ws.Cells(i, 6).Value
                j = j + 1
                
                ' complete calculations
                yearly_change = close_price - open_price
                
                If open_price > 0 Then
                    percent_change = (close_price / open_price) - 1
                Else
                    percent_change = 1
                End If
                
                ' send values to the summary table: Ticker, Yearly Change, Percent Change, Stock Volume
                
                    ws.Cells(j, 9).Value = stock_ticker
                    ws.Cells(j, 10).Value = yearly_change
                    ws.Cells(j, 11).Value = percent_change
                    ws.Cells(j, 12).Value = stock_volume
                    
                    ' color formatting based on positive/negative value of yearly change
                    If yearly_change >= 0 Then
                        ws.Cells(j, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(j, 10).Interior.ColorIndex = 3
                    End If
                
                ' set new values for next loop
                open_price = ws.Cells(i + 1, 3).Value
                stock_volume = 0
                stock_ticker = ws.Cells(i + 1, 1).Value
            
            End If

Next i

Next ws

End Sub

Sub The_VBA_of_Wall_Street_Hard_RRZ()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        ' MODERATE ASSIGNMENT
        ' --------------------------------------------

        ' Set an initial variable to hold the ticker symbol
        Dim Ticker_Symbol As String

        ' Set an initial variable to hold the total volume for each stock
        Dim Volume_Total As Double
        Volume_Total = 0

        ' Set an initial variable to hold the opening price for each stock
        Dim Price_Open As Double
        Price_Open = ws.Cells(2, 3).Value
    
        ' Set an initial variable to hold the closing price for each stock
        Dim Price_Close As Double
    
        ' Set an initial variable to hold the yearly change for each stock
        Dim Yearly_Change As Double
    
        ' Set an initial variable to hold the percent change for each stock
        Dim Percent_Change As Double
    
        ' Keep track of the location for each stock in the summary table
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2

        ' Determine the last row of the stocks data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' --------------------------------------------
        ' SET UP THE SUMMARY TABLE
        ' --------------------------------------------

        ' Add the Ticker column header in the summary table
        ws.Range("I1").Value = "Ticker"
    
        ' Add the Yearly Change column header in the summary table
        ws.Range("J1").Value = "Yearly Change"
    
        ' Add the Percent Change column header in the summary table
        ws.Range("K1").Value = "Percent Change"
    
        ' Add the Total Stock Volume column header in the summary table
        ws.Range("L1").Value = "Total Stock Volume"

        ' --------------------------------------------
        ' POPULATE THE SUMMARY TABLE
        ' --------------------------------------------

        ' Loop through all of the stocks
        For i = 2 To LastRow

            ' Check if we are still within the same ticker symbol, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the ticker symbol
                Ticker_Symbol = ws.Cells(i, 1).Value

                ' Set the closing price
                Price_Close = ws.Cells(i, 6).Value
                
                ' Set the yearly change
                Yearly_Change = Price_Close - Price_Open
                
                ' To avoid a #DIV/0 error, check if the opening price is equal to zero, if not...
                If Price_Open <> 0 Then
                
                    ' Set the percent change
                    Percent_Change = Yearly_Change / Price_Open
                    
                ' If the opening price is equal to zero...
                Else
                    
                    ' Set the percent change to zero
                    Percent_Change = 0
                    
                    MsgBox ("Note: " + Ticker_Symbol + " has no value.")
                    
                End If

                ' Add to the total volume
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value

                ' Print the ticker symbol in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            
                ' Print the yearly change, from the opening price to the closing price, to the summary table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00000000"
            
                ' Check if the yearly change is positive or negative, if positive...
                If (Price_Close - Price_Open) > 0 Then
                
                    ' Set the background color of the yearly change to green
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                ' If the yearly change is negative...
                ElseIf (Price_Close - Price_Open) < 0 Then
            
                    ' Set the background color of the yearly change to red
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
                ' Print the percent change, from the opening price to the closing price, to the summary table
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
                ' Print the total volume to the summary table
                ws.Range("L" & Summary_Table_Row).Value = Volume_Total

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the total volume for the next stock
                Volume_Total = 0
            
                ' Reset the opening price for the next stock
                Price_Open = ws.Cells(i + 1, 3).Value

            ' If the cell immediately following a row is within the same ticker symbol...
            Else

                ' Add to the total volume
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
            End If

        Next i

        ' --------------------------------------------
        ' HARD ASSIGNMENT
        ' --------------------------------------------

        ' Set an initial variable to hold the ticker symbol of the stock with the greatest % increase
        Dim Ticker_Percent_Max As String
        Ticker_Percent_Max = ws.Cells(2, 9).Value
    
        ' Set an initial variable to hold the ticker symbol of the stock with the greatest % decrease
        Dim Ticker_Percent_Min As String
        Ticker_Percent_Min = ws.Cells(2, 9).Value
    
        ' Set an initial variable to hold the ticker symbol of the stock with the greatest total volume
        Dim Ticker_Volume_Max As String
        Ticker_Volume_Max = ws.Cells(2, 9).Value
    
        ' Set an initial variable to hold the greatest % increase
        Dim Percent_Max As Double
        Percent_Max = ws.Cells(2, 11).Value
    
        ' Set an initial variable to hold the greatest % decrease
        Dim Percent_Min As Double
        Percent_Min = ws.Cells(2, 11).Value
    
        ' Set an initial variable to hold the greatest total volume
        Dim Volume_Max As Double
        Volume_Max = ws.Cells(2, 12).Value

        ' Determine the last row of the summary table
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' --------------------------------------------
        ' SET UP THE OVERVIEW TABLE
        ' --------------------------------------------

        ' Add the Greatest % Increase row header in the overview table
        ws.Range("O2").Value = "Greatest % Increase"
    
        ' Add the Greatest % Decrease row header in the overview table
        ws.Range("O3").Value = "Greatest % Decrease"
    
        ' Add the Greatest Total Volume row header in the overview table
        ws.Range("O4").Value = "Greatest Total Volume"
    
        ' Add the Ticker column header in the overview table
        ws.Range("P1").Value = "Ticker"
    
        ' Add the Value column header in the overview table
        ws.Range("Q1").Value = "Value"
    
        ' --------------------------------------------
        ' POPULATE THE OVERVIEW TABLE
        ' --------------------------------------------
    
        ' Loop through all of the stocks in the summary table
        For i = 2 To LastRow
    
            ' Check if the percent change is greater than the greatest % increase, if so...
            If ws.Cells(i, 11).Value > Percent_Max Then
        
                ' Set the ticker symbol of the stock with the greatest % increase to the ticker symbol
                Ticker_Percent_Max = ws.Cells(i, 9).Value
            
                ' Set the greatest % increase to the percent change
                Percent_Max = ws.Cells(i, 11).Value
        
            ' Check if the percent change is less than the greatest % decrease, if so...
            ElseIf ws.Cells(i, 11).Value < Percent_Min Then
        
                ' Set the ticker symbol of the stock with the greatest % decrease to the ticker symbol
                Ticker_Percent_Min = ws.Cells(i, 9).Value
            
                ' Set the greatest % decrease to the percent change
                Percent_Min = ws.Cells(i, 11).Value
        
            End If
    
            ' Check if the total stock volume is greater than the greatest total volume, if so...
            If ws.Cells(i, 12).Value > Volume_Max Then
        
                ' Set the ticker symbol of the stock with the greatest total volume to the ticker symbol
                Ticker_Volume_Max = ws.Cells(i, 9).Value
            
                ' Set the greatest total volume to the total stock volume
                Volume_Max = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
        ' Print the ticker symbol of the stock with the greatest % increase to the overview table
        ws.Range("P2").Value = Ticker_Percent_Max
    
        ' Print the ticker symbol of the stock with the greatest % decrease to the overview table
        ws.Range("P3").Value = Ticker_Percent_Min
    
        ' Print the ticker symbol of the stock with the greatest total volume to the overview table
        ws.Range("P4").Value = Ticker_Volume_Max
    
        ' Print the greatest % increase to the overview table
        ws.Range("Q2").Value = Percent_Max
        ws.Range("Q2").NumberFormat = "0.00%"
    
        ' Print the greatest % decrease to the overview table
        ws.Range("Q3").Value = Percent_Min
        ws.Range("Q3").NumberFormat = "0.00%"
    
        ' Print the greatest total volume to the overview table
        ws.Range("Q4").Value = Volume_Max

    Next ws

    MsgBox ("Stock market data analysis complete.")

End Sub
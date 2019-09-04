Sub Mulitple_year_stock_data()

' Loop through all sheets

    For Each ws In Worksheets
    ws.Activate
        
        'Set label column header for summary table
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("K:M").EntireColumn.AutoFit
        Range("O:O").EntireColumn.AutoFit


        ' Set an initial variable for holding the ticker symbol
        Dim Ticker_Symbol As String

        ' Set an initial variable for holding the total stock volume per ticker symbol
        Dim Tot_Stock_Volume As Double
        Tot_Stock_Volume = 0

        ' Set an initial variable for holding the last row for each sheet
        Dim Lastrow As Long

        ' Set the open year date
        Dim Open_Year_Date As Boolean
        Open_Year_Date = True

        ' Set initial variable for Open year price, end of year price and yearly change
        Dim Open_Year_Price, End_Year_Price, Yearly_Change As Double

        ' Set Percent_Change
        Dim Percent_Change As Double

        ' Keep track of the location for each ticker symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2


        ' Count the number of rows in each sheet
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        ' Loop through all the ticker
        For i = 2 To Lastrow

            ' Check if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ' Set the ticker symbol
                Ticker_Symbol = Cells(i, 1).Value
                
                ' Add the total stock volume
                Tot_Stock_Volume = Tot_Stock_Volume + Cells(i, 7).Value
                
                ' Print the ticker symbol in the Summary Table
                Range("J" & Summary_Table_Row).Value = Ticker_Symbol

                ' Print the total stock volume in the Summary Table
                Range("M" & Summary_Table_Row).Value = Tot_Stock_Volume
                
                ' Yearly Change from what the stock opened the year at to what the closing price was.
                End_Year_Price = Cells(i, 6).Value
                Yearly_Change = End_Year_Price - Open_Year_Price
                Range("K" & Summary_Table_Row).Value = Yearly_Change
                        
                ' The percent change from what the stock price opened the year at to what it closed
                Percent_Change = (Yearly_Change / Open_Year_Price) * 100
                Range("L" & Summary_Table_Row).Value = Percent_Change
                Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ' Conditional formatting highlighting positive change in green and negative change in red
                
                If Range("K" & Summary_Table_Row).Value > 0 Then
                    Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the total stock volume
                Tot_Stock_Volume = 0
                
                ' Set Open_Year-Date to true
                Open_Year_Date = True
                
            ' If the cell immediately following a row is the same ticker symbol...
            Else
                
                ' Add to the total stock volume
                Tot_Stock_Volume = Tot_Stock_Volume + Cells(i, 7).Value
                
                ' Get the price the stock opened the year
                If Open_Year_Date And Cells(i, 3).Value <> 0 Then
                    Open_Year_Price = Cells(i, 3).Value
                    Open_Year_Date = False
                
                End If
                
                    
            End If
            
        Next i

        ' Locate stock with the "Greatest % increase", "Greatest % Decrease"
        ' and "Greatest total volume"

        ' Count the number of row in the Summary Table
        Dim n As Integer
        Dim Max_Inc, Min_Dec As Long
        Dim Max_Vol As Double
        Max_Inc = 0
        Min_Dec = 0
        Max_Vol = 0

        n = ws.Cells(Rows.Count, 11).End(xlUp).Row

        For j = 2 To n
            
            If Range("L" & j).Value > Max_Inc Then
                Max_Inc = Range("L" & j).Value
                Range("Q2").Value = Max_Inc
                Range("P2").Value = Range("J" & j).Value
                        
            ElseIf Range("L" & j).Value < Min_Dec Then
                Min_Dec = Range("L" & j).Value
                Range("Q3").Value = Range("L" & j).Value
                Range("P3").Value = Range("J" & j).Value

            End If
            
            If Range("M" & j).Value > Max_Vol Then
                Max_Vol = Range("M" & j).Value
                Range("Q4").Value = Max_Vol
                Range("P4").Value = Range("J" & j).Value
            
            End If
            
            Range("Q2:Q3").NumberFormat = "0.00%"
            Range("Q:Q").EntireColumn.AutoFit

            
        Next j

    Next ws
End Sub

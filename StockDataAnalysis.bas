Attribute VB_Name = "Module1"
Sub TickerTotalVolume()

    ' Set an initial variable for holding the ticker
    Dim ticker As String

    ' Set an initial variable for holding the total stock volume per ticker
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    ' Set header for the summary table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"

    ' Keep track of the location for each stock ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    ' Set total of rows of the whole table
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set variables for yearly change beginning on row 2
    Dim firstRow As Long
    Dim openPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Variant
    firstRow = 2
    
    ' Set variable for percentage change
    Dim percentageChange As Double

    ' Loop through all ticker stocks
    For i = 2 To LastRow

        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the ticker name.
            ticker = Cells(i, 1).Value
            
            ' Get the value for open and closing price.
            openPrice = Cells(firstRow, 3).Value
            closingPrice = Cells(i, 6).Value
            
            ' Get yearlyChange value.
            yearlyChange = closingPrice - openPrice
                        
            ' Get the percentage change at the end of the year.
            If yearlyChange > 0 Then
                If closingPrice = 0 Then
                    percentageChange = 0
                Else
                    percentageChange = (yearlyChange / closingPrice)
                End If
            Else
                If openPrice = 0 Then
                    percentageChange = 0
                Else
                    percentageChange = (yearlyChange / openPrice)
                End If
            End If
            
            ' Get the next start row for the next ticket.
            firstRow = i + 1
    
            ' Add to the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
            ' Print the ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker
            
            ' Print the yearly change to the summary Table
            Range("J" & Summary_Table_Row).NumberFormat = "0.000000000"
            Range("J" & Summary_Table_Row).Value = yearlyChange
            
                ' Set conditional format to highlight positive and negative change
                If yearlyChange > 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
            
            ' Print the percentage change to the summary Table
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("K" & Summary_Table_Row).Value = percentageChange
    
            ' Print the Volume Amount to the Summary Table
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
          
            ' Reset the Total Stock Volume
            Total_Stock_Volume = 0

        ' If the cell immediately following a row is the same ticker...
        Else
        
            ' Add to the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        End If

    Next i

End Sub

Sub MaxMinValues()

    ' Set Headers for challenges.
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    ' Set variable for challenges
    Dim tickerMax, tickerMin, tickerMaxVol As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim maxTotalVolume As Double
    
    ' Set total of rows of the whole table
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 11).End(xlUp).Row
    
    ' Set maximun value
    greatestIncrease = 0
    greatestDecrease = 0
    maxTotalVolume = 0
    
    ' Loop through all ticker stocks
    For i = 2 To LastRow

        ' Check if we are still within the same ticker, if it is not...
        If Cells(i, 11).Value > greatestIncrease Then
        
            tickerMax = Cells(i, 9).Value
            greatestIncrease = Cells(i, 11).Value
            
        ElseIf Cells(i, 11).Value < greatestDecrease Then
            
            tickerMin = Cells(i, 9).Value
            greatestDecrease = Cells(i, 11).Value
        End If
        
        If Cells(i, 12).Value > maxTotalVolume Then
            
            tickerMaxVol = Cells(i, 9).Value
            maxTotalVolume = Cells(i, 12).Value
            
        End If
        
    Next i
    
    ' Print the values in the respective columns
    Range("P2").Value = tickerMax
    Range("P3").Value = tickerMin
    Range("P4").Value = tickerMaxVol
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q2").Value = greatestIncrease
    Range("Q3").Value = greatestDecrease
    
    Range("Q4").Value = maxTotalVolume
    
End Sub

Sub Main()

    ' Run subroutines in all worksheets
    For Each ws In Worksheets
    
        ws.Select
        Call TickerTotalVolume
        Call MaxMinValues
        
    Next ws

End Sub

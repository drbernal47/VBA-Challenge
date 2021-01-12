Attribute VB_Name = "Module1"
Sub StockLooper()

    ' Loop through this program for each worksheet
    
    For Each ws In Worksheets
    
        ' Create analysis headers in spreadsheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N1").Value = "Stock Feature"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
    
    
        ' Create variables to read and store data from each row
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim Volume As Double
        Dim NextTicker As String
    
        ' Create a variable to help us add data in new rows to the data analysis section (starting with row 2).
        Dim TickerNumber As Double
        TickerNumber = 2
    
        ' Create variables to help us calculate Yearly & Percent Change
        Dim YearOpenPrice As Double
        Dim YearClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
    
        ' Create variables to help us consistently format cells later
        PercentFormat = "#.00%"
        GreenColor = 4
        RedColor = 3
    
        ' Create a variable to help us keep track of which row begins each ticker section
        Dim OpeningRow As Double
        OpeningRow = 2
    
        ' Create variables to help us count total stock volume
        Dim VolumeCount As Double
        VolumeCount = 0

        ' Count the number of rows from the data to determine the loop length
        Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
        ' Loop through and start adding volume, keeping track of when the next ticker changes
        For r = 2 To LastRow
            
            ' Read and store this row's data
            Ticker = ws.Cells(r, 1).Value
            OpenPrice = ws.Cells(r, 3).Value
            ClosePrice = ws.Cells(r, 6).Value
            Volume = ws.Cells(r, 7).Value
            NextTicker = ws.Cells(r + 1, 1).Value
            
            ' Add current row's value to the total stock volume
            VolumeCount = VolumeCount + Volume
        
            ' Check to see if this is the last row of the ticker by peeking ahead
            If Ticker <> NextTicker Then
        
                ' Add data to the analysis section, starting with Ticker
                ws.Cells(TickerNumber, 9).Value = Ticker
            
                ' Calculate the Yearly Change by finding the Opening & Closing Prices
                YearOpenPrice = ws.Cells(OpeningRow, 3).Value
                YearClosePrice = ClosePrice
                YearlyChange = YearClosePrice - YearOpenPrice
                ws.Cells(TickerNumber, 10).Value = YearlyChange
            
                ' Format Yearly Change based on positive or negative value
                If YearlyChange > 0 Then
            
                    ws.Cells(TickerNumber, 10).Interior.ColorIndex = GreenColor
                
                ElseIf YearlyChange < 0 Then
            
                    ws.Cells(TickerNumber, 10).Interior.ColorIndex = RedColor
                
                End If
            
            
                ' Format PercentChange for readiability
                ws.Cells(TickerNumber, 11).NumberFormat = PercentFormat
            
                ' Calculate the Percent Change, addressing the case where YearOpenPrice is zero
                If YearOpenPrice = 0 Then
                    ' Percent Change is undefined, so we leave it empty
                
                ElseIf YearOpenPrice <> 0 Then
                    PercentChange = YearlyChange / YearOpenPrice
                    ws.Cells(TickerNumber, 11).Value = PercentChange
                
                End If

                ' Report the total volume of the stock (we have been counting all along)
                ws.Cells(TickerNumber, 12).Value = VolumeCount
            
                'Reset the Volume Count
                VolumeCount = 0
            
                ' Calculate a new Opening Row value for the next ticker section
                OpeningRow = r + 1
            
                ' Increase the Ticker Number count to ensure we use a fresh row for the next ticker
                TickerNumber = TickerNumber + 1
            
            End If
        
        Next r
    
    
        ' Count the number of rows in the data analysis section
        LastTickerRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'Create variables to help us track the desired features
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
    
        ' Create variables to help us track which rows have "winning" data and set to default
        Dim IncreaseLeadTicker As String
        Dim DecreaseLeadTicker As String
        Dim VolumeLeadTicker As String
        
        ' Create variables to read and store data from our analysis rows
        Dim CandidateTicker As String
        Dim CandidateChange As Double
        Dim CandidateVolume As Double
    
    
        ' Loop through our data analysis to find the desired featured tickers
        For r = 2 To LastTickerRow
    
            ' Read and store data from the analysis row
            CandidateTicker = ws.Cells(r, 9).Value
            CandidateChange = ws.Cells(r, 11).Value
            CandidateVolume = ws.Cells(r, 12).Value
    
            ' Compare this row's values to the stored feature values and adjust as needed
            If CandidateChange > GreatestIncrease Then
        
                GreatestIncrease = CandidateChange
                IncreaseLeadTicker = CandidateTicker
        
            End If
        
            If CandidateChange < GreatestDecrease Then
        
                GreatestDecrease = CandidateChange
                DecreaseLeadTicker = CandidateTicker
            
            End If
        
            If CandidateVolume > GreatestVolume Then
        
                GreatestVolume = CandidateVolume
                VolumeLeadTicker = CandidateTicker
            
            End If
    
        Next r
    
        'Print values that were found
    
        ws.Range("O2").Value = IncreaseLeadTicker
        ws.Range("P2").Value = GreatestIncrease
        ws.Range("P2").NumberFormat = PercentFormat
        ws.Range("O3").Value = DecreaseLeadTicker
        ws.Range("P3").Value = GreatestDecrease
        ws.Range("P3").NumberFormat = PercentFormat
        ws.Range("O4").Value = VolumeLeadTicker
        ws.Range("P4").Value = GreatestVolume

    Next ws

End Sub

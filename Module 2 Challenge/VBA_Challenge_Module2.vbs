Sub VBA_CHALLENGE_MODULE_2()
Dim ws As Worksheet
    Dim TickerSymbol As String
    Dim LastRow As Long
    Dim StockVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
    Dim Summary_Table_Row As Integer

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Reset variables for each worksheet
        StockVolume = 0
        Summary_Table_Row = 2

        ' Find the last row with data in the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Summary Table Headers and adjust Column
        ws.Range("I1").Value = "Ticker"
        ws.Columns("I:I").AutoFit
        ws.Range("J1").Value = "Quarterly Change"
        ws.Columns("J:J").AutoFit
        ws.Range("K1").Value = "Percent Change"
        ws.Columns("K:K").AutoFit
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("L:L").AutoFit
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Columns("O:O").AutoFit
        ws.Range("P1").Value = "Ticker"
        ws.Columns("P:P").AutoFit
        ws.Range("Q1").Value = "Value"
        ws.Columns("Q:Q").AutoFit

        ' Initialize variables for greatest calculations
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double
        Dim Ticker_Greatest_Increase As String
        Dim Ticker_Greatest_Decrease As String
        Dim Ticker_Greatest_Volume As String


        ' Loop through rows of data
        For i = 2 To LastRow
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

                OpenPrice = ws.Cells(i, 3).Value
                StockVolume = StockVolume + ws.Cells(i, 7).Value

            ElseIf ws.Cells(i+1,1).Value <> ws.Cells(i, 1).Value Then  

                TickerSymbol = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                Quarterly_Change = ClosePrice - OpenPrice
                Percent_Change = ((ClosePrice - OpenPrice) / OpenPrice) 
                StockVolume = StockVolume + ws.Cells(i, 7).Value

                ws.Range("I" & Summary_Table_Row).Value = TickerSymbol
                ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Table_Row).Value = StockVolume


                ' Check for greatest values
                If Percent_Change > Greatest_Increase Then
                    Greatest_Increase = Percent_Change
                    Ticker_Greatest_Increase = TickerSymbol
                End If
                
                If Percent_Change < Greatest_Decrease Or Greatest_Decrease = 0 Then
                    Greatest_Decrease = Percent_Change
                    Ticker_Greatest_Decrease = TickerSymbol
                End If
                
                If StockVolume > Greatest_Volume Then
                    Greatest_Volume = StockVolume
                    Ticker_Greatest_Volume = TickerSymbol
                End If

                Summary_Table_Row = Summary_Table_Row + 1
                StockVolume = 0
            
            Else
                StockVolume = StockVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Output of the other calculations table
        ws.Range("P2").Value = Ticker_Greatest_Increase
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P3").Value = Ticker_Greatest_Decrease
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P4").Value = Ticker_Greatest_Volume
        ws.Range("Q4").Value = Greatest_Volume
        ws.Columns("Q:Q").AutoFit
        'Output of the cell interior color
        Dim cell As Range
    ' Loop through each cell in the specified range
    For Each cell In ws.Range("J2:J123")
        If cell.Value > 0 Then
            cell.Interior.ColorIndex = 4 ' Set the cell color to green for positive numbers
        ElseIf cell.Value < 0 Then
            cell.Interior.ColorIndex = 3 ' Set the cell color to red for negative numbers
        Else
            cell.Interior.ColorIndex = 2 ' Set the cell color to white for zero
        End If
    Next cell

    Next ws


End Sub
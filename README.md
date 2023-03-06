# VBA-challenge
Module 2 Challenge - VBA

Instructions:

Create a script that loops through all the stocks for one year and outputs the following information:

-The ticker symbol

-Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

-The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

-The total stock volume of the stock.

-Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

-Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.


Solution:

-Screenshots of results for each year are attached.

-.vbs  file with code attached.

-VBA code additionally pasted below.




    ' Initiate subroutine for the assignment.
    Sub ticker_challenge()


        ' Executes program across all sheets in book
        For Each ws In Worksheets

            ' Set initial varialbles for handling ticker name
            Dim WorksheetName As String
            Dim ticker_name As String
            Dim next_ticket_name As String
            Dim yearly_price_change As Double
            Dim percent_change As Double
            Dim start As Long

            ' Set an initial variable for holding the total volume per ticker
            Dim volume_total As LongLong
            volume_total = 0

            ' Define amount of rows in sheet
            Dim rowCount As Long
            rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
            WorksheetName = ws.Name

            ' Keep track of the location for each ticker in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2

            ' Label  summary table headers
            ws.Range("J1").Value = "Ticker"
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("L1").Value = "Percent Change"
            ws.Range("M1").Value = "Total Stock Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"

            ' Loop through all tickers
            For i = 2 To rowCount

                ' Check if we are still within the same ticker, if it is not...
                ticker_name = ws.Cells(i, 1).Value
                next_ticker_name = ws.Cells(i + 1, 1).Value

                If ws.Cells(i - 1, 1).Value <> ticker_name Then
                    open_price = ws.Cells(i, 3).Value

                End If

                If next_ticker_name <> ticker_name Then

                    ' track closing price
                    close_price = ws.Cells(i, 6).Value

                    ' Print the ticker in the Summary table
                    ws.Cells(Summary_Table_Row, 10).Value = ticker_name

                    ' Add to the volume total
                    volume_total = volume_total + ws.Cells(i, 7).Value

                    ' Print the ticker total to the Summary Table
                    ws.Cells(Summary_Table_Row, 13).Value = volume_total

                    ' Calculate Yearly Change
                    ' yearly change =Cells(End,6).Value - Cells(Start,3).Value
                    ' end = row
                    yearly_price_change = close_price - open_price

                    ' Calculate Yearly Percentage change
                    percent_change = (yearly_price_change / open_price)

                    ' Print yearly price change to the Summary Table
                    ws.Cells(Summary_Table_Row, 11).Value = yearly_price_change

                        ' Color code price change by green/red
                        If yearly_price_change > 0 Then
                            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4

                        Else
                            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3

                        End If

                    ' Print yearly percent change to the Summary Table
                    ws.Cells(Summary_Table_Row, 12).Value = percent_change
                    ws.Cells(Summary_Table_Row, 12).NumberFormat = "0.00%"

                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1

                    'Reset the volume total
                    volume_total = 0

                    'Reset the start
                    start = i + 1

                ' If the cell immediately following a row is the same ticker
                Else

                    ' Add to the ticker volume total
                    volume_total = volume_total + ws.Cells(i, 7).Value

                End If

            Next i

                ' Calculate maximum and minnimum values from sumarry table and add to new table
                last_sum_row = Summary_Table_Row - 1

                ws.Cells(2, 17) = ws.Application.WorksheetFunction.Max(Range(ws.Cells(2, 12), ws.Cells(last_sum_row, 12)))
                ws.Cells(3, 17) = ws.Application.WorksheetFunction.Min(Range(ws.Cells(2, 12), ws.Cells(last_sum_row, 12)))
                ws.Cells(4, 17) = ws.Application.WorksheetFunction.Max(Range(ws.Cells(2, 13), ws.Cells(last_sum_row, 13)))

                Greatest_Increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
                ws.Cells(2, 16) = ws.Cells(Greatest_Increase + 1, 10)

                Greatest_Decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
                ws.Cells(3, 16) = ws.Cells(Greatest_Decrease + 1, 10)

                Greatest_Volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & rowCount)), ws.Range("M2:M" & rowCount), 0)
                ws.Cells(4, 16) = ws.Cells(Greatest_Volume + 1, 10)

                ws.Cells(2, 17).NumberFormat = "0.00%"
                ws.Cells(3, 17).NumberFormat = "0.00%"

        Next ws

    End Sub

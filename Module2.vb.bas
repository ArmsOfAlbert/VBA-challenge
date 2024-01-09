Attribute VB_Name = "Module1"
Sub part_one_multiple_sheets()
    Dim ws As Worksheet
    Dim Ticker_name As String
    Dim Yearly_change As Double
    Dim Percent_change As Double
    Dim Stock_Value As Long
    Dim first_day As Double
    Dim Last_day As Double
    Dim Total_stock_value As LongLong
    Dim table_row As Integer
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volume_ticker As String
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        table_row = 2
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0

        ' Loop through all tickers in the current sheet
        For i = 2 To 22771
            ' Check if it's a new Ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                Ticker_name = ws.Cells(i, 1).Value
                first_day = ws.Cells(i, 3).Value ' Set first_day for the new Ticker
                Total_stock_value = 0
            End If

            Total_stock_value = Total_stock_value + ws.Cells(i, 7).Value

            ' Check if it's the last day of the Ticker
            If ws.Cells(i + 1, 1).Value <> Ticker_name Then
                Last_day = ws.Cells(i, 6).Value

                ' Calculate Yearly_change and Percent_change
                Yearly_change = Last_day - first_day
                If first_day <> 0 Then
                    Percent_change = Round((Yearly_change / first_day) * 100, 2)
                Else
                    Percent_change = 0
                End If

                ' Write data to the table
                ws.Cells(table_row, 12).Value = Ticker_name
                ws.Cells(table_row, 13).Value = Yearly_change
                ws.Cells(table_row, 14).Value = Percent_change
                ws.Cells(table_row, 15).Value = Total_stock_value

                ' Apply conditional formatting based on Yearly_change
                If Yearly_change > 0 Then
                    ws.Cells(table_row, 13).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf Yearly_change < 0 Then
                    ws.Cells(table_row, 13).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                ' Update greatest increase, greatest decrease, and greatest volume
                If Percent_change > greatest_increase Then
                    greatest_increase = Percent_change
                    greatest_increase_ticker = Ticker_name
                End If

                If Percent_change < greatest_decrease Then
                    greatest_decrease = Percent_change
                    greatest_decrease_ticker = Ticker_name
                End If

                If Total_stock_value > greatest_volume Then
                    greatest_volume = Total_stock_value
                    greatest_volume_ticker = Ticker_name
                End If

                ' Reset values for the next Ticker
                table_row = table_row + 1
                first_day = 0 ' Reset first_day for the next Ticker
                Total_stock_value = 0
            End If
        Next i

        ' Output the stocks with the greatest increase, decrease, and volume on each sheet
        ws.Cells(2, 18).Value = "Greatest % Increase"
        ws.Cells(3, 18).Value = "Greatest % Decrease"
        ws.Cells(4, 18).Value = "Greatest Volume"

        ws.Cells(2, 19).Value = greatest_increase_ticker
        ws.Cells(3, 19).Value = greatest_decrease_ticker
        ws.Cells(4, 19).Value = greatest_volume_ticker

        ws.Cells(2, 20).Value = greatest_increase
        ws.Cells(3, 20).Value = greatest_decrease
        ws.Cells(4, 20).Value = greatest_volume
    Next ws
End Sub


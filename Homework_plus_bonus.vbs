Sub GetTickerData()
    Dim data_row As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume As LongLong
    Dim cur_tkr As String
    Dim next_tkr As String
    '---------------------------
    'Bonus
    Dim max_change As Double
    Dim min_change As Double
    Dim max_volume As LongLong
    Dim max_tkr As String
    Dim min_tkr As String
    Dim max_vol_tkr As String
    '---------------------------

    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Setup Intial values
        data_row = 2
        volume = 0
        max_change = 0
        min_change = 0
        max_volume = 0
        cur_tkr = ws.Cells(2, 1).Value
        open_price = ws.Cells(2, 3).Value
        close_price = ws.Cells(2, 6).Value
        For i = 2 To LastRow
            next_tkr = ws.Cells(i, 1).Value
            If cur_tkr = next_tkr Then
                close_price = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
            Else
                yearly_change = close_price - open_price

                'Check for 0 to avoid overflow error
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If

                ws.Range("I" & data_row).Value = cur_tkr
                ws.Range("J" & data_row).Value = yearly_change

                If yearly_change < 0 Then
                    ws.Range("J" & data_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & data_row).Interior.ColorIndex = 4
                End If

                ws.Range("K" & data_row).Value = percent_change
                ws.Range("K" & data_row).NumberFormat = "0.00%"
                ws.Range("L" & data_row).Value = volume

                '---------------------------
                'Bonus
                If percent_change > max_change Then
                    max_change = percent_change
                    max_tkr = cur_tkr
                End If

                If percent_change < min_change Then
                    min_change = percent_change
                    min_tkr = cur_tkr
                End If

                If volume > max_volume Then
                    max_volume = volume
                    max_vol_tkr = cur_tkr
                End If
                '---------------------------

                'Reset Values
                data_row = data_row + 1
                cur_tkr = next_tkr
                volume = 0
                open_price = ws.Cells(i, 3).Value
                close_price = ws.Cells(2, 6).Value
            End If
        Next i

        '---------------------------
        'Bonus
        ws.Range("P2").Value = max_tkr
        ws.Range("P3").Value = min_tkr
        ws.Range("P4").Value = max_vol_tkr

        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q2").Value = max_change
        ws.Range("Q3").Value = min_change
        ws.Range("Q4").Value = max_volume
        '---------------------------

        ws.Columns("A:Q").AutoFit
    Next ws

End Sub

Sub stock_summary()

    ' Loop through all worksheets
    For Each ws In Worksheets

        ' Set variables
        Dim stock_ticker As String

        Dim stock_open As Double
        stock_open = 0

        dim stock_close As Double
        stock_close = 0

        Dim stock_volume As Double
        stock_volume = 0

        Dim stock_summary As Integer
        stock_summary = 2

        ' Defining table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Define the last row of the sheet
        last_row_ticker = ws.Range("A1").End(xlDown).Row

        ' Start the loop beginning at the second row, looking at column A
        For i = 2 To last_row_ticker

            ' If this is the first time the ticker appears in column A, then give the stock_open
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                stock_open = ws.Cells(i, 3).Value
            End If

            ' If this is the last time the ticker appears in column A, then do the following
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Define more variables
                stock_close = ws.Cells(i, 6).Value
                stock_ticker = ws.Cells(i, 1).Value
                stock_volume = stock_volume + ws.Cells(i, 7).Value

                ' Add ticker, volume per ticker, yearly change to the summary table
                ws.Range("I" & stock_summary).Value = stock_ticker
                ws.Range("L" & stock_summary).Value = stock_volume
                ws.Range("J" & stock_summary).Value = stock_close - stock_open
                
                ' Conditional formatting
                If ws.Range("J" & stock_summary).Value > 0 Then
                    ws.Range("J" & stock_summary).Interior.Color = RGB(0, 255, 0)
                ElseIf ws.Range("J" & stock_summary).Value < 0 Then
                    ws.Range("J" & stock_summary).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Calculate the percent change and print that to the summary table
                If stock_open <> 0 Then
                    ws.Range("K" & stock_summary).Value = ws.Range("J" & stock_summary).Value / stock_open
                    ws.Range("K" & stock_summary).NumberFormat = "0.00%"
                Else
                    ws.Range("K" & stock_summary).Value = "N/A"
                End If
    
                stock_summary = stock_summary + 1
                stock_volume = 0
            Else

                ' Add to the Total Stock Volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                    
            End If
        Next i



    ' Check if cell is numeric to account for "N/A" when Opening Stock Price is zero
    For i = 2 To last_row_ticker
        If IsNumeric(ws.Cells(i, 11).Value) = True Then
            If ws.Cells(i, 11).Value > max Then
                max = ws.Cells(i, 11).Value
                stock_max = ws.Cells(i, 9).Value
            End If
        End If
    Next i
        
    ' Find Greatest % Decrease and Ticker Name
    Dim min As Variant
    Dim stock_min As String
    min = 0

    For i = 2 To last_row_ticker
        If IsNumeric(ws.Cells(i, 11).Value) = True Then
            If ws.Cells(i, 11).Value < min Then
                min = ws.Cells(i, 11).Value
                stock_min = ws.Cells(i, 9).Value
            End If
        End If
    Next i
        
    ' Find Greatest Total Volume and Ticker Name
    Dim stock_max_volume As Variant
    Dim stock_max_volumeTicker As String
    stock_max_volume = 0

    For i = 2 To last_row_ticker
        If ws.Cells(i, 12).Value > stock_max_volume Then
            stock_max_volume = ws.Cells(i, 12).Value
            stock_max_volumeTicker = ws.Cells(i, 9).Value
        End If
    Next i
    max = FormatPercent(max)
    min = FormatPercent(min)

    ' Final challenge summary table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P2").Value = stock_max
    ws.Range("P3").Value = stock_min
    ws.Range("P4").Value = stock_max_volumeTicker
    ws.Range("Q2").Value = max
    ws.Range("Q3").Value = min
    ws.Range("Q4").Value = stock_max_volume
   ws.Cells.Columns.AutoFit

    Next ws
End Sub

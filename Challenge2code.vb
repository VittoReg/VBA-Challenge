Sub CalculateYearPercentChangeAndTotalVolumeForAllSheets2()
    ' Loop through each worksheet in the workbook
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim i As Long
        Dim start As Long
        Dim lastRow As Long
        Dim change As Double
        Dim percentChange As Double
        Dim finalPrice As Double
        Dim initialPrice As Double
        Dim totalVolume As Double
        Dim highestIncrease As Double
        Dim highestDecrease As Double
        Dim highestVolume As Double
        Dim highestIncreaseTicker As String
        Dim highestDecreaseTicker As String
        Dim highestVolumeTicker As String
        Dim j As Integer
        ' Activate the current worksheet
        ws.Activate
        ' Set title rows
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Year Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ' Initialize variables
        start = 2
        j = 0 ' To keep track of output row
        highestIncrease = 0
        highestDecrease = 0
        highestVolume = 0
        ' Get the row number of the last row with data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' Loop through each row
        For i = 2 To lastRow
            ' Accumulate total volume for the current ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            ' If ticker changes or it's the last row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the end price for the current ticker
                finalPrice = ws.Cells(i, 6).Value
                ' Get start price for the ticker
                initialPrice = ws.Cells(start, 3).Value
                ' Calculate change from start to end
                change = finalPrice - initialPrice
                ' Calculate percent change
                If initialPrice <> 0 Then
                    percentChange = change / initialPrice
                Else
                    percentChange = 0
                End If
                ' Output ticker, year change, percent change, and total volume
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value ' Ticker
                ws.Range("J" & 2 + j).Value = change
                ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).Value = percentChange
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                ws.Range("L" & 2 + j).Value = totalVolume
                ' Conditional formatting for year change
                If change > 0 Then
                    ws.Range("J" & 2 + j).Interior.Color = vbGreen
                ElseIf change < 0 Then
                    ws.Range("J" & 2 + j).Interior.Color = vbRed
                Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = xlNone
                End If
                ' Conditional formatting for percent change
                If percentChange > 0 Then
                    ws.Range("K" & 2 + j).Interior.Color = vbGreen
                ElseIf percentChange < 0 Then
                    ws.Range("K" & 2 + j).Interior.Color = vbRed
                Else
                    ws.Range("K" & 2 + j).Interior.ColorIndex = xlNone
                End If
                ' Update greatest increase, decrease, and total volume
                If percentChange > highestIncrease Then
                    highestIncrease = percentChange
                    highestIncreaseTicker = ws.Cells(i, 1).Value
                ElseIf percentChange < highestDecrease Then
                    highestDecrease = percentChange
                    highestDecreaseTicker = ws.Cells(i, 1).Value
                End If
                If totalVolume > highestVolume Then
                    highestVolume = totalVolume
                    highestVolumeTicker = ws.Cells(i, 1).Value
                End If
                ' Prepare for next ticker
                start = i + 1
                j = j + 1 ' Move to the next row for the next output
                totalVolume = 0 ' Reset total volume for next ticker
            End If
        Next i
        ' After loop, output greatest increase, decrease, and volume
        ws.Range("P2").Value = highestIncreaseTicker
        ws.Range("Q2").Value = highestIncrease
        ws.Range("Q2").NumberFormat = "0.00" & "%"
        ws.Range("P3").Value = highestDecreaseTicker
        ws.Range("Q3").Value = highestDecrease
        ws.Range("Q3").NumberFormat = "0.00" & "%"
        ws.Range("P4").Value = highestVolumeTicker
        ws.Range("Q4").Value = highestVolume
        ws.Range("Q4").NumberFormat = "0"
        ' Set additional titles for clarity
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    Next ws
End Sub


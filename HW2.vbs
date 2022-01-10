Sub HW2()

    For Each ws In Worksheets

        'Create Variables
        Dim ticker As String
        
        Dim totalStockVolume As LongLong
        totalStockVolume = 0

        Dim numOfRows As LongLong 'Used to capture total number of rows in a column
        
        Dim i As LongLong 'Counter

        Dim TickerReportingRow As LongLong 'Used to move Reporting Row
        TickerReportingRow = 2

        Dim openingPrice As Double
        
        
        Dim closingPrice As Double


        Dim yearlyChange As Double
        
        Dim percentChange As Double
        
        'Set Additional Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Value"
        'Get number of rows and set variable
        numOfRows = ws.Cells(Rows.Count, 1).End(xlUp).Row

        

        For i = 2 To numOfRows

            If (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) Then
                openingPrice = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value

                ws.Range("I" & TickerReportingRow).Value = ticker

                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
                ws.Range("L" & TickerReportingRow).Value = totalStockVolume

                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                ws.Range("J" & TickerReportingRow).Value = yearlyChange
                
                    If openingPrice <> 0 Then
                        percentChange = yearlyChange / openingPrice
                        ws.Range("K" & TickerReportingRow).Value = percentChange
                    Else
                        ws.Range("K" & TickerReportingRow).Value = "Cannot Compute"
                    End If

                
                TickerReportingRow = TickerReportingRow + 1
                totalStockVolume = 0


            Else
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i


    Next ws
    
End Sub

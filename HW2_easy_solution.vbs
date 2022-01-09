Sub HW2_easy_solution()

    For Each ws In Worksheets

        'Sort the Worksheet
        With ActiveSheet.Sort
            .SortFields.Add Key:=Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=Range("B1"), Order:=xlAscending
            .SetRange Range("A:G")
            .Header = xlYes
            .Apply
        End With

        'Create Variables

        Dim ticker As String
        
        Dim totalStockVolume As LongLong
        totalStockVolume = 0

        Dim numOfRows As LongLong 'Used to capture total number of rows in a column
        
        Dim i As LongLong 'Counter

        Dim TickerReportingRow As LongLong 'Used to move Reporting Row
        TickerReportingRow = 2


        'Set Additional Headers on each sheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Value"


        'Get number of rows for range counting and set variable
        numOfRows = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        For i = 2 To numOfRows
 

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value

                ws.Range("I" & TickerReportingRow).Value = ticker


                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
                ws.Range("J" & TickerReportingRow).Value = totalStockVolume

                TickerReportingRow = TickerReportingRow + 1
                totalStockVolume = 0
            Else
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i


    Next ws
    
End Sub





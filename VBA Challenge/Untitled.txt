Attribute VB_Name = "Module1"
Sub Stocks()

    ' Looping through all sheets
    For Each ws In ThisWorkbook.Worksheets
    
        ' set variables
        Dim Ticker As String
        Dim quarterlychange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim LastRow As Long
        Dim SummaryRow As Long
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        
        ' name columns
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Total Volume"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        
        ' set up loop logic - looping through each stock on each sheet
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize SummaryRow to start filling from row 2 (below headers)
        SummaryRow = 2
        TotalVolume = 0 ' Reset Total Volume before starting
        
        ' Get the first opening price
        OpeningPrice = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
            
            ' Add to Total Volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Check if the ticker changes (i.e., new stock starts)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = LastRow Then
            
                ' Assign the ticker
                Ticker = ws.Cells(i, 1).Value
                
                ' Get the closing price
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the quarterly change
                quarterlychange = ClosingPrice - OpeningPrice
                
                ' Calculate the percent change
                If OpeningPrice <> 0 Then
                    PercentChange = (quarterlychange / OpeningPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ' Insert data into summary table
                ws.Cells(SummaryRow, 8).Value = Ticker ' Ticker in column H
                ws.Cells(SummaryRow, 9).Value = TotalVolume ' Total Volume in column I
                ws.Cells(SummaryRow, 10).Value = quarterlychange ' Quarterly Change in column J
                ws.Cells(SummaryRow, 11).Value = PercentChange ' Percent Change in column K
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%" ' Format Percent Change as percentage
                
                ' Color the Quarterly Change cell based on value
                If quarterlychange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                ElseIf quarterlychange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If
                
                ' Move to the next summary row
                SummaryRow = SummaryRow + 1
                
                ' Reset total volume and opening price for the next ticker
                TotalVolume = 0
                If i < LastRow Then
                    OpeningPrice = ws.Cells(i + 1, 3).Value ' Next opening price for the new stock
                End If
                
            End If
        Next i
    Next ws

End Sub


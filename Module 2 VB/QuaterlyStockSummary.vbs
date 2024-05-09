Sub Ticker()

Dim TickerSymbol As String
Dim QuarterOpen As Double
Dim QarterClose As Double
Dim StartPrice As Double
Dim ClosePrice As Double
Dim RowCount As Double
Dim FindStickerRowCount As Double
Dim PercentChange As Double
Dim StartTotalVolume As Double
Dim EndTotalVolume As Double
Dim TotalStockVolume As Double
Dim StoreGreatPercent As Double
Dim PreGreatPercent As Double
Dim StoreGreatSticker As String
Dim PreDecreasePercent As Double
Dim GreatTotalStockVolume As Double
Dim StoreTotalStockVolume As Double


Dim stockws As Worksheet

For Each qaurterws In Worksheets
    qaurterws.Range("I1").Value = "Ticker"
    qaurterws.Range("J1").Value = "Quaterly Change"
    qaurterws.Range("K1").Value = "Percent Change"
    qaurterws.Range("L1").Value = "Total Stock Volume"
    qaurterws.Range("P1").Value = "Ticker"
    qaurterws.Range("Q1").Value = "Value"
    qaurterws.Range("L1").EntireColumn.AutoFit
    qaurterws.Range("P1").EntireColumn.AutoFit
    qaurterws.Range("Q1").EntireColumn.AutoFit
    qaurterws.Cells(2, 15).Value = "Greatest % Increase"
    qaurterws.Cells(3, 15).Value = "Greatest % Decrease"
    qaurterws.Cells(4, 15).Value = "Greatest Total Volume"
    RowCount = 2
    FindStickerRowCount = 2
    StoreGreatPercent = 0
    PreGreatPercent = 0
    PreDecreasePercent = 0
    GreatTotalStockVolume = 0
    
    AllRowsCounts = qaurterws.Cells(Rows.Count, "A").End(xlUp).Row
        Debug.Print AllRowsCounts
        For i = 2 To AllRowsCounts
        
            If qaurterws.Cells(i + 1, 1).Value <> qaurterws.Cells(i, 1).Value Then
                TickerSymbol = qaurterws.Cells(i, 1).Value
                StartTotalVolume = i
                ClosePrice = qaurterws.Cells(i, 6).Value
                StartPrice = qaurterws.Cells(FindStickerRowCount, 3).Value
                qaurterws.Cells(RowCount, 9).Value = TickerSymbol
                qaurterws.Cells(RowCount, 10).Value = ClosePrice - StartPrice
                StoreGreatPercent = (ClosePrice - StartPrice) / StartPrice
                StoreTotalStockVolume = Application.Sum(Range(qaurterws.Cells(FindStickerRowCount, 7), qaurterws.Cells(i, 7)))
                If StoreGreatPercent > PreGreatPercent Then
                   PreGreatPercent = StoreGreatPercent
                   StoreGreatSticker = TickerSymbol
                   qaurterws.Cells(2, 16).Value = TickerSymbol
                   qaurterws.Cells(2, 17).Value = PreGreatPercent
                   qaurterws.Cells(2, 17).NumberFormat = "0.00%"
                End If
                If StoreGreatPercent < PreDecreasePercent Then
                   PreDecreasePercent = StoreGreatPercent
                   StoreGreatSticker = TickerSymbol
                   qaurterws.Cells(3, 16).Value = TickerSymbol
                   qaurterws.Cells(3, 17).Value = PreDecreasePercent
                   qaurterws.Cells(3, 17).NumberFormat = "0.00%"
                End If
                If StoreTotalStockVolume > GreatTotalStockVolume Then
                   GreatTotalStockVolume = StoreTotalStockVolume
                   StoreGreatSticker = TickerSymbol
                   qaurterws.Cells(4, 16).Value = TickerSymbol
                   qaurterws.Cells(4, 17).Value = StoreTotalStockVolume
                End If
                
                qaurterws.Cells(RowCount, 11).Value = StoreGreatPercent
                qaurterws.Cells(RowCount, 11).NumberFormat = "0.00%"
                qaurterws.Cells(RowCount, 12).Value = StoreTotalStockVolume
                FindStickerRowCount = i + 1
                Debug.Print FindStickerRowCount
                If qaurterws.Cells(RowCount, 10) > 0 Then
                  qaurterws.Cells(RowCount, 10).Interior.ColorIndex = 4
                ElseIf Cells(RowCount, 10) < 0 Then
                  qaurterws.Cells(RowCount, 10).Interior.ColorIndex = 3
                End If
                RowCount = RowCount + 1
             End If
             
            
            Next i
Next qaurterws
        
        



    


End Sub






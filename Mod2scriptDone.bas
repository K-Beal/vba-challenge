Attribute VB_Name = "Module1"
Sub Stocks()

  Dim totalStockVolume As Double
  Dim openingPrice As Double
  Dim closingPrice As Double
  Dim percentChange As Double
  Dim greatestPercentIncreaseValue As Double
  Dim greatestPercentIncreaseTicker As String
  Dim greatestPercentDecreaseValue As Double
  Dim greatestPercentDecreaseTicker As String
  Dim greatestTotalVolumeValue As Double
  Dim greatestTotalVolumeTicker As String
  Dim ws As Worksheet
  Dim rowCount As Long
  Dim nextRow As Long
  
  
  For Each ws In ThisWorkbook.Sheets
  
  nextRow = 2
  totalStockVolume = 0
  openingPrice = 0
  closingPrice = 0
  greatestTotalVolumeValue = 0
  greatestTotalVolumeTicker = "<NONE>"
  greatestPercentIncreaseValue = 0
  greatesPercentIncreaseTicker = "<NONE>"
  greatestPercentDecreaseValue = 0
  greatestPercentDecreaseTicker = "<NONE>"
  
  
  rowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  
  For Row = 2 To rowCount
    If ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
        startingRow = Row
        totalStockVolume = 0
        openingPrice = ws.Cells(Row, 3).Value
        
    End If
    
    
    totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
    
    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
        endingRow = Row
        closingPrice = ws.Cells(Row, 6).Value
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Cells(nextRow, 9).Value = ws.Cells(Row, 1).Value
        ws.Cells(nextRow, 10).Value = closingPrice - openingPrice
        If closingPrice > openingPrice Then
            ws.Cells(nextRow, 10).Interior.Color = vbGreen
        ElseIf closingPrice < openingPrice Then
            ws.Cells(nextRow, 10).Interior.Color = vbRed
        Else
            ws.Cells(nextRow, 10).Interior.Color = vbYellow
        End If
        
        
        ws.Cells(nextRow, 10).NumberFormat = "$#,##0.00"
        ws.Cells(nextRow, 11).NumberFormat = "0.00%"
        percentChange = (closingPrice - openingPrice) / openingPrice
        
        ws.Cells(nextRow, 11).Value = (closingPrice - openingPrice) / openingPrice
        If percentChange > greatestPercentIncreaseValue Then
            greatestPercentIncreaseValue = percentChange
            greatestPercentIncreaseTicker = ws.Cells(Row, 1).Value
            
            
        End If
        
        If closingPrice > openingPrice Then
            ws.Cells(nextRow, 11).Interior.Color = vbGreen
        ElseIf closingPrice < openingPrice Then
            ws.Cells(nextRow, 11).Interior.Color = vbRed
        Else
            ws.Cells(nextRow, 11).Interior.Color = vbYellow
        
        End If
        
        If percentChange < greatestPercentDecreaseValue Then
            greatestPercentDecreaseValue = percentChange
            greatestPercentDecreaseTicker = ws.Cells(Row, 1).Value
            
        End If
        
        
        If totalStockVolume > greatestTotalVolumeValue Then
            greatestTotalVolumeTicker = ws.Cells(Row, 1).Value
            greatestTotalVolumeValue = totalStockVolume
        
        End If
            
        ws.Cells(nextRow, 12).Value = totalStockVolume
        nextRow = nextRow + 1
        
      End If
    Next Row
      
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
        
    ws.Range("O2").Value = "Greatest Percent Increase"
    ws.Range("P2").Value = greatestPercentIncreaseTicker
    ws.Range("Q2").Value = greatestPercentIncreaseValue
        
    ws.Range("O3").Value = "Greatest Percent Decrease"
    ws.Range("P3").Value = greatestPercentDecreaseTicker
    ws.Range("Q3").Value = greatestPercentDecreaseValue
        
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
        
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = greatestTotalVolumeTicker
    ws.Range("Q4").Value = greatestTotalVolumeValue
    
  
    MsgBox (greatestPercentIncreaseValue)
    MsgBox (greatestPercentIncreaseTicker)
Next ws

End Sub

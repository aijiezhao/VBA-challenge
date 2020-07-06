Sub worksheet_loop()

For Each ws In Worksheets

  Dim Ticker As String
  Dim Yearly_Change As Double
  Yearly_Change = 0
  Dim Percent_change As Double
  Percent_change = 0
  Dim total_stock_volume As Double
  total_stock_volume = 0
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Dim open_price As Double
  open_price = ws.Cells(2, 3).Value
  Dim Lastrow As Long

  Dim max_ticker As String
  max_ticker = " "
  Dim max_percent As Double
  max_percent = 0
  Dim max_volume_ticker As String
  max_volume_ticker = " "
  Dim max_volume As Double
  max_volume = 0
  Dim min_ticker As String
  min_ticker = " "
  Dim min_percent As Double
  min_percent = 0



  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"

  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"

  For i = 2 To Lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Ticker = ws.Cells(i, 1).Value
      Yearly_Change = ws.Cells(i, 6).Value - open_price
      Percent_change = Yearly_Change / open_price
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      ws.Range("K" & Summary_Table_Row).Value = Percent_change
      ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
      Summary_Table_Row = Summary_Table_Row + 1
      open_price = ws.Cells(i + 1, 3).Value
      total_stock_volume = 0

    Else
      
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
     
    End If

  Next i

      ws.Range("K2:K" & Summary_Table_Row).NumberFormat = "0.00%"

  For i = 2 To Summary_Table_Row


    If ws.Cells(i, 10).Value <= 0 Then
      ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)

    Else

      ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)

    End If

    If ws.Cells(i, 11).Value > max_percent Then
      max_percent = ws.Cells(i, 11).Value
      max_ticker = ws.Cells(i, 9).Value
    Else
      max_percent = max_percent
      max_ticker = max_ticker
    End If

    If ws.Cells(i, 11).Value < min_percent Then
      min_percent = ws.Cells(i, 11).Value
      min_ticker = ws.Cells(i, 9).Value
    Else
      min_percent = min_percent
      min_ticker = min_ticker
    End If

    If ws.Cells(i, 12).Value > max_volume Then
      max_volume = ws.Cells(i, 12).Value
      max_volume_ticker = ws.Cells(i, 9).Value
    Else
      max_volume = max_volume
      max_volume_ticker = max_volume_ticker
    End If
      
  Next i
  ws.Range("Q2").Value = max_percent
  ws.Range("Q3").Value = min_percent
  ws.Range("P2").Value = max_ticker
  ws.Range("P3").Value = min_ticker
  ws.Range("Q4").Value = max_volume
  ws.Range("P4").Value = max_volume_ticker


Next ws

End Sub


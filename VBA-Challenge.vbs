Sub Stock()

For Each ws In Worksheets
    
    ws.Range("I1") = "Stock Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Volume"
    
    ws.Columns("I:L").HorizontalAlignment = xlCenter
    
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 12

    Dim Stock_Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Stock_Volume As Double
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Last_Row As Double
    
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Stock_Volume = 0

    Dim Summary_Row As Double
    Summary_Row = 2

    For i = 2 To Last_Row

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        Stock_Ticker = ws.Cells(i, 1).Value
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

        ws.Range("I" & Summary_Row).Value = Stock_Ticker
        ws.Range("L" & Summary_Row).Value = Stock_Volume

        Stock_Volume = 0
        Stock_Close = ws.Cells(i, 6)

        If Stock_Open = 0 Then
            Yearly_Change = 0
            Percent_Change = 0
        Else:
            Yearly_Change = Stock_Close - Stock_Open
            Percent_Change = (Stock_Close - Stock_Open) / Stock_Open
        End If
    
        ws.Range("J" & Summary_Row).Value = Yearly_Change
        ws.Range("K" & Summary_Row).Value = Percent_Change
        ws.Range("K" & Summary_Row).Style = "Percent"
        ws.Range("K" & Summary_Row).NumberFormat = "0.00%"

        Summary_Row = Summary_Row + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
        Stock_Open = ws.Cells(i, 3)
    Else: Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    End If

    Next i

    For i = 2 To Last_Row
        If ws.Range("J" & i).Value > 0 Then
            ws.Range("J" & i).Interior.ColorIndex = 43
        ElseIf ws.Range("J" & i).Value < 0 Then
            ws.Range("J" & i).Interior.ColorIndex = 46
        End If

    Next i
    Next ws
End Sub

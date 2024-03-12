VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Challenge2FINAL()

    ' Set initial variables
    Dim Ticker_Unique As String
    Dim Total_volume As LongLong
    Dim RowCount As Long
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Start As Long
    Dim ws As Worksheet
    Dim Summary_volume_row As Long
    
    'Set ws and connect to the correct worksheet,set Summary_volume_row and set Rowcount
    Set ws = ThisWorkbook.Worksheets("2018")
    Summary_volume_row = 2
    RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Set names for 4 new coloms
    ws.Cells(1, 14).Value = "Ticker_Unique"
    ws.Cells(1, 15).Value = "Yearly Change"
    ws.Cells(1, 16).Value = "Percent Change"
    ws.Cells(1, 17).Value = "Total_volume"

    ' Create a starting point for total_volume and start
    Total_volume = 0
    Start = 2
    

    ' Loop from row 2 to last row through tickers
    For i = 2 To RowCount
        ' Test if next ticker is same with previous one
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_Unique = ws.Cells(i, 1).Value
            Total_volume = Total_volume + ws.Cells(i, 7).Value

            ' Do the loop for YearlyChange and PercentChange calculation
            If ws.Cells(Start, 3).Value <> 0 Then
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
                PercentChange = (YearlyChange / ws.Cells(Start, 3).Value) * 100
            Else
                YearlyChange = 0
                PercentChange = 0
            End If

            ' Print our results in 4 new coloms N,O,P & Q
            ws.Cells(Summary_volume_row, 15).Value = YearlyChange
            ws.Cells(Summary_volume_row, 16).Value = PercentChange
            ws.Cells(Summary_volume_row, 14).Value = Ticker_Unique
            ws.Cells(Summary_volume_row, 17).Value = Total_volume

            ' Reset Total_volume and move to the next summary row
            Total_volume = 0
            Summary_volume_row = Summary_volume_row + 1
            ' Set start to the next ticker's first row
            Start = i + 1
        Else
            Total_volume = Total_volume + ws.Cells(i, 7).Value
        End If
    Next i
    

    ' set new values to for greatest % increase, % decrease, and total volume
    Dim increase_number As Long
    Dim decrease_number As Long
    Dim volume_number As Long
    Dim maxPercentChange As Double
    Dim minPercentChange As Double
    Dim maxTotalVolume As Double
    Dim LastRow As Long
    
    LastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row

    maxPercentChange = WorksheetFunction.Max(ws.Range("P2:P" & LastRow))
    minPercentChange = WorksheetFunction.Min(ws.Range("P2:P" & LastRow))
    maxTotalVolume = WorksheetFunction.Max(ws.Range("Q2:Q" & LastRow))
    
    ws.Range("T2").Value = maxPercentChange
    ws.Range("T3").Value = minPercentChange
    ws.Range("T4").Value = maxTotalVolume
    
    ws.Range("S2:S3").NumberFormat = "0.00%"
    

    ' Find the matching row numbers for the ticker symbols
    increase_number = WorksheetFunction.Match(maxPercentChange, ws.Range("P2:P" & LastRow), 0) + 1
    decrease_number = WorksheetFunction.Match(minPercentChange, ws.Range("P2:P" & LastRow), 0) + 1
    volume_number = WorksheetFunction.Match(maxTotalVolume, ws.Range("Q2:Q" & LastRow), 0) + 1
    
    ' Print ticker symbols that repond with gratest%increase,gratest%decrease and gratest total volume
    ws.Range("S2").Value = ws.Cells(increase_number, "N").Value
    ws.Range("S3").Value = ws.Cells(decrease_number, "N").Value
    ws.Range("S4").Value = ws.Cells(volume_number, "N").Value
    
    'Set Colomn U that labels our results
    ws.Cells(2, 21).Value = "Greatest%increase"
    ws.Cells(3, 21).Value = "Greatest%decrease"
    ws.Cells(4, 21).Value = "GreatestTotalVolume"
    
    'Do coloring(red and green) in colomn O based if the YearlyChange is positive or negative
    For i = 2 To LastRow
    If IsEmpty(Cells(i, 15).Value) Then Exit For
        If Cells(i, 15).Value > 0 Then
            Cells(i, 15).Interior.ColorIndex = 4
        Else
            Cells(i, 15).Interior.ColorIndex = 3
    End If
    Next i
    
    
End Sub

Sub stock_data()
For Each ws In Worksheets
'Populate the Header Values on each sheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "YearlyChange"
    ws.Range("K1").Value = "PercentChange"
    ws.Range("L1").Value = "TotalStockVolume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
'Declare Variables and their Data Types
Dim i As Long
Dim LastRow As Long
Dim TickerCounter As Integer
Dim TotalCounter As Long
Dim YearChange As Double
Dim YearOpen As Double
Dim PercentChange As Double
Dim TSV As Double
Dim YearOpenRow As Long

'Set initial values for Variables
 YearChange = 0
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 TickerCounter = 2
 TSV = 0
 YearOpenRow = 2
 'Set the K column and Q2 and Q3 to Percentage
 ws.Range("Q2","Q3").Style = "Percent"
    'Compare ticker values in column A and print Unique values in Column I
    For i = 2 To LastRow
        YearOpen = ws.Cells(YearOpenRow, 3).Value
        If ws.Cells(i+1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(TickerCounter, 9).Value = ws.Cells(i, 1).Value
            'Calculate the value change from close in column F to open in column C, print the value in column J based on uniqe ticker value and set color format
            YearChange = ws.Cells(i,6).value - YearOpen
            ws.Cells(TickerCounter, 10).Value = YearChange
                If ws.Cells(TickerCounter, 10).Value >= 0 Then 
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
                End If
            'Calculate the percent of change for each unique ticker value and print in column K
            PercentChange = (YearChange / YearOpen)
            ws.Cells(TickerCounter, 11).Value = PercentChange
            ws.Cells(TickerCounter, 11).NumberFormat = "0.00%"
            'Print the total stock volume in column L
            TSV = TSV + ws.Cells(i, 7).Value
            ws.Cells(TickerCounter, 12).Value = TSV
            'Reset values for variables 
            YearOpen = 0
            YearChange = 0
            TickerCounter = TickerCounter + 1
            TSV = 0
            YearOpenRow = i + 1
            'Add up the volume values for each unique ticker value
        Else
            TSV = TSV + ws.Cells(i, 7).Value
        End If
    Next i 
'Declare variables for Greatest Section of the worksheet
Dim PercentLastRow As Integer
Dim TSVLastRow As Integer
Dim GrtIncrease As Double
Dim GrtDecrease As Double
Dim GrtTSV As Double
Dim j As Long
Dim k As Long
'Set initial values for variables
 PercentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
 TSVLastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
 GrtIncrease = 0
 GrtDecrease = 0
 GrtTSV = 0
    'Compare the Values in all the rows for column L and assign the largest number to the variable
    For j = 2 To TSVLastRow
        If GrtTSV < ws.Cells(j, 12).Value Then
            GrtTSV = ws.Cells(j, 12).Value
            ws.Range("Q4").Value = GrtTSV
            ws.Range("P4").Value = ws.Cells(j, 9).Value
        End If
    Next j 
    'Compare the Values in column J and determine the highest and lowest percentages Print those values with the ticker value in columns P and Q
    For k = 2 To PercentLastRow
        If GrtIncrease < ws.Cells(k, 11).Value Then 
            GrtIncrease = ws.Cells(k, 11).Value
            ws.Range("Q2").Value = GrtIncrease
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("P2").Value = ws.Cells(k, 9).Value
        ElseIf GrtDecrease > ws.Cells(k, 11).Value Then
            GrtDecrease = ws.Cells(k, 11).Value
            ws.Range("Q3").Value = GrtDecrease
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("P3").Value = ws.Cells(k, 9)
        End If
    Next K
    Next ws 


End Sub

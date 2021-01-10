Attribute VB_Name = "Module1"
Sub Stocks():

For Each ws In Worksheets

' declare variables

Dim row As Integer
Dim ticker As String
Dim totalVolume As Double
Dim yearlyChange As Double
Dim percentChange As Double

'initializing variables

totalVolume = 0
row = 2
openPrice = ws.Cells(2, 3).Value

'creating column headers for the output table

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

' starting loop to go through all rows
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row


For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        yearlyChange = ws.Cells(i, 6).Value - openPrice
        If openPrice = 0 Then
            percentChange = 0
        Else
            percentChange = yearlyChange / openPrice
        End If
        
' print the output on the summary table

ws.Cells(row, 9).Value = ticker
ws.Cells(row, 10).Value = yearlyChange
ws.Cells(row, 11).Value = percentChange
ws.Cells(row, 12).Value = totalVolume

' conditional color formatting

If yearlyChange > 0 Then
    ws.Cells(row, 10).Interior.ColorIndex = 4
Else
    ws.Cells(row, 10).Interior.ColorIndex = 3
End If

' formatting percent change

ws.Cells(row, 11).NumberFormat = "0.00%"

' increment row

row = row + 1

'reset totalVolume to 0 so that accumulation will not occur

totalVolume = 0

' find open price for each ticker

openPrice = ws.Cells(i + 1, 3).Value

' if ticker is still the same, just add result

        Else
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        yearlyChange = ws.Cells(i, 6).Value - openPrice
        
        End If
        
        
Next i

Next ws


End Sub


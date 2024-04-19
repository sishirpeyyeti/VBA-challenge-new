Sub vbachallenge()

Dim i As Long
Dim rowCount As Long
Dim Location As Long
Dim total_stock_volume As Double
Dim x As Double
Dim y As Double

x = 20180102
y = 20181231

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


total_stock_volume = 0
Location = 2
rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

'The ticker symbol
For i = 2 To rowCount
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ws.Cells(Location, 9).Value = ws.Cells(i, 1).Value
    Location = Location + 1
    End If
Next i

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of the year.
Location = 2

For i = 2 To rowCount
    If ws.Cells(i, 2).Value = x Then
    startprice = ws.Cells(i, 3).Value
    End If
    If ws.Cells(i, 2).Value = y Then
    endprice = ws.Cells(i, 6).Value
    ws.Cells(Location, 10).Value = (startprice - endprice) * -1
    Location = Location + 1
    End If
Next i

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year
Location = 2

For i = 2 To rowCount
    If ws.Cells(i, 2).Value = x Then
    startprice = ws.Cells(i, 3).Value
    End If
    If ws.Cells(i, 2).Value = y Then
    endprice = ws.Cells(i, 6).Value
    ws.Cells(Location, 11).Value = (((startprice - endprice) / startprice) * 100) * -1
    Location = Location + 1
    End If
Next i

'Total Stock Volume
Location = 2

For i = 2 To rowCount
    If ws.Cells(i, 2).Value = x Then
    total_stock_volume = ws.Cells(i, 7).Value
    End If
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
    total_stock_volume = total_stock_volume + ws.Cells(i + 1, 7).Value
    End If
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    ws.Cells(Location, 12).Value = total_stock_volume
    Location = Location + 1
    total_stock_volume = 0
    End If
Next i

volume = ws.Cells(Rows.Count, "L").End(xlUp).Row
gincrease = ws.Range("K2").Value
gvolume = ws.Range("L2").Value
gdecrease = ws.Range("K2").Value

For i = 2 To volume
    If gvolume < ws.Cells(i + 1, 12).Value Then
    gvolume = ws.Cells(i + 1, 12).Value
    End If
Next i

ws.Range("Q4").Value = gvolume

For i = 2 To volume
    If ws.Cells(i, 12).Value = gvolume Then
    ws.Range("P4").Value = ws.Cells(i, 9).Value
    End If
Next i

For i = 2 To volume
    If gincrease < ws.Cells(i + 1, 11).Value Then
    gincrease = ws.Cells(i + 1, 11).Value
    End If
Next i

ws.Range("Q2").Value = gincrease

For i = 2 To volume
    If ws.Cells(i, 11).Value = gincrease Then
    ws.Range("P2").Value = ws.Cells(i, 9).Value
    End If
Next i

For i = 2 To volume
    If gdecrease > ws.Cells(i + 1, 11).Value Then
    gdecrease = ws.Cells(i + 1, 11).Value
    End If
Next i

ws.Range("Q3").Value = gdecrease

For i = 2 To volume
    If ws.Cells(i, 11).Value = gdecrease Then
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    End If
Next i


x = x + 10000
y = y + 10000
Next ws

End Sub
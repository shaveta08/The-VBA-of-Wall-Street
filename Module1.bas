Attribute VB_Name = "Module1"
Sub Calculate():
Dim rowCount As Long
Dim i As Long
Dim ansRow As Long
Dim totalStock As Double
Dim startPrice As Double
Dim endPrice As Double
Dim maxPercent As Double
Dim minPercent As Double
Dim maxTotalStock As Double
Dim tickerMaxPer As String
Dim tickerMinPer As String
Dim tickerMaxTotal As String


For Each ws In Worksheets

ansRow = 2
maxPercent = 0
minPercent = 0
maxTotalStock = 0
totalStock = 0
rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
endPrice = 0
startPrice = 0


ws.Range("M:M").Style = "Percent"
ws.Columns("K:N").AutoFit
ws.Range("K1").Value = "Ticker"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percentage Change"
ws.Range("N1").Value = "Total Stock Volume"

startPrice = ws.Range("C2").Value

For i = 2 To rowCount


If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then

    endPrice = ws.Cells(i, 6).Value
    
    
    If (startPrice <> 0 And endPrice <> 0) Then
        'Put ticker number change and Percentage Change in cell
        ws.Range("K" & ansRow) = ws.Cells(i, 1).Value
        ws.Range("L" & ansRow) = endPrice - startPrice
        ws.Range("M" & ansRow) = ws.Range("L" & ansRow) / startPrice
    
        ws.Range("N" & ansRow) = totalStock + ws.Cells(i, 7)
    
        'Color Formating of Change.
        If ws.Range("L" & ansRow) < 0 Then
            ws.Range("L" & ansRow).Interior.ColorIndex = 3
        ElseIf ws.Range("L" & ansRow) > 0 Then
            ws.Range("L" & ansRow).Interior.ColorIndex = 4
        End If
    
        'Check Maximum value of percent change
        If ws.Range("M" & ansRow) > maxPercent Then
            maxPercent = ws.Range("M" & ansRow)
            tickerMaxPer = ws.Range("K" & ansRow)
        End If
        'Check Minimum value of Percentage
        If ws.Range("M" & ansRow) < minPercent Then
            minPercent = ws.Range("M" & ansRow)
            tickerMinPer = ws.Range("K" & ansRow)
        End If
        'Check Maximum Value of total stock
        If ws.Range("N" & ansRow) > maxTotalStock Then
            maxTotalStock = ws.Range("N" & ansRow)
            tickerMaxTotal = ws.Range("K" & ansRow)
        End If
    End If
    
    totalStock = 0
    ansRow = ansRow + 1
    startPrice = ws.Cells(i + 1, 3).Value
Else
    totalStock = totalStock + ws.Cells(i, 7)
End If
Next i
'Fill Greatest values in the cells of the current sheet.
ws.Columns("P:R").AutoFit
ws.Range("P1").Value = ""
ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"
ws.Range("Q2").Value = tickerMaxPer
ws.Range("Q3").Value = tickerMinPer
ws.Range("Q4").Value = tickerMaxTotal
ws.Range("R2").Value = maxPercent
ws.Range("R3").Value = minPercent
ws.Range("R2").Style = "Percent"
ws.Range("R3").Style = "Percent"
ws.Range("R4").Value = maxTotalStock

Next ws

End Sub



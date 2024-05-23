Sub ExtractTickers()

'Declare variables
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")

'Create collection to store tickers
    Dim tickers As Collection
    Set tickers = New Collection

'Determine last row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

'Loop through Column A
    Dim r As Long
    For r = 1 To lastRow
        Dim ticker As Variant
        ticker = ws.Cells(r, 1).Value
        If Not IsInCollection(ticker, tickers) Then
            tickers.Add ticker
        End If
    Next r

    ' Write the tickers into column J
    Dim i As Long
    i = 1
    Dim item As Variant
    For Each item In tickers
        ws.Cells(i, 10).Value = item
        i = i + 1
    Next item
End Sub

Function IsInCollection(val As Variant, col As Collection) As Boolean
    Dim elem As Variant
    On Error Resume Next
    IsInCollection = False
    For Each elem In col
        If elem = val Then
            IsInCollection = True
            Exit For
        End If
    Next elem
    On Error GoTo 0
End Function

Sub QuarterlyChange()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim change As Double
    Dim uniqueTickers As Range
    Dim cell As Range
    Dim tickerRange As Range
    Dim firstRow As Long
    Dim lastTickerRow As Long
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ' Get the range of unique tickers in column J
    Set uniqueTickers = ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
    
    ' Loop through each unique ticker
    For Each cell In uniqueTickers
        ticker = cell.Value
        Set tickerRange = ws.Range("A2", ws.Cells(ws.Rows.Count, "A").End(xlUp))
        
        ' Find the first and last row of the current ticker symbol
        firstRow = 0
        lastTickerRow = 0
        For Each tickerCell In tickerRange
            If tickerCell.Value = ticker Then
                If firstRow = 0 Then firstRow = tickerCell.Row
                lastTickerRow = tickerCell.Row
            End If
        Next tickerCell
        
        ' Check if symbol has 62 rows
        If lastTickerRow - firstRow + 1 = 62 Then
            openingPrice = ws.Cells(firstRow, "C").Value
            closingPrice = ws.Cells(lastTickerRow, "F").Value
            change = closingPrice - openingPrice
            
            ' Place results in column 'K'
            cell.Offset(0, 1).Value = change
        Else
            cell.Offset(0, 1).Value = "Error: Not 62 rows"
        End If
    Next cell
End Sub


Sub CalculateQuarterlyPercentageChange()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim percentageChange As Double
    Dim uniqueTickers As Range
    Dim cell As Range
    Dim tickerRange As Range
    Dim firstRow As Long
    Dim lastTickerRow As Long
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ' Get the range of unique tickers in column J
    Set uniqueTickers = ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
    
    ' Loop through each unique ticker
    For Each cell In uniqueTickers
        ticker = cell.Value
        Set tickerRange = ws.Range("A2", ws.Cells(ws.Rows.Count, "A").End(xlUp))
        
        ' Find the first and last row of the current ticker symbol
        firstRow = 0
        lastTickerRow = 0
        For Each tickerCell In tickerRange
            If tickerCell.Value = ticker Then
                If firstRow = 0 Then firstRow = tickerCell.Row
                lastTickerRow = tickerCell.Row
            End If
        Next tickerCell
        
        ' Check if the ticker symbol has 62 rows
        If lastTickerRow - firstRow + 1 = 62 Then
            openingPrice = ws.Cells(firstRow, "C").Value
            closingPrice = ws.Cells(lastTickerRow, "F").Value
            If openingPrice <> 0 Then
                percentageChange = ((closingPrice - openingPrice) / openingPrice) * 100
            Else
                percentageChange = 0
            End If
            
            ' Place the percentage change in column 'L'
            cell.Offset(0, 2).Value = Format(percentageChange, "0.00") & "%"
        Else
            cell.Offset(0, 2).Value = "Error: Not 62 rows"
        End If
    Next cell
End Sub

Sub CalculateVolume()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim volumeSum As Double
    Dim uniqueTickers As Range
    Dim cell As Range
    Dim tickerRange As Range
    Dim firstRow As Long
    Dim lastTickerRow As Long
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ' Get the range of unique tickers in column J
    Set uniqueTickers = ws.Range("J2", ws.Cells(ws.Rows.Count, "J").End(xlUp))
    
    ' Loop through each unique ticker
    For Each cell In uniqueTickers
        ticker = cell.Value
        Set tickerRange = ws.Range("A2", ws.Cells(ws.Rows.Count, "A").End(xlUp))
        
        ' Find the first and last row of the current ticker symbol
        firstRow = 0
        lastTickerRow = 0
        volumeSum = 0
        For Each tickerCell In tickerRange
            If tickerCell.Value = ticker Then
                If firstRow = 0 Then firstRow = tickerCell.Row
                lastTickerRow = tickerCell.Row
                volumeSum = volumeSum + ws.Cells(tickerCell.Row, "G").Value
            End If
        Next tickerCell
        
        ' Check if the ticker symbol has 62 rows
        If lastTickerRow - firstRow + 1 = 62 Then
            ' Place the volume sum in column 'M'
            cell.Offset(0, 3).Value = volumeSum
        Else
            cell.Offset(0, 3).Value = "Error: Not 62 rows"
        End If
    Next cell
End Sub

Sub ConditionalFormatting()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Define the worksheet and range
    Set ws = ThisWorkbook.Sheets("Q1")
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    Set rng = ws.Range("K2:K" & lastRow)
    
    ' Clear any existing conditional formatting
    rng.FormatConditions.Delete
    
    ' Add conditional formatting for positive numbers
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0)
    End With
    
    ' Add conditional formatting for negative numbers
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0)
    End With
    
    ' Add conditional formatting for zero
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        .Interior.Color = RGB(255, 255, 255)
    End With
End Sub

Sub ConditionalFormatting_Percent()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Define the worksheet and range
    Set ws = ThisWorkbook.Sheets("Q1")
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    Set rng = ws.Range("L2:L" & lastRow)
    
    ' Clear any existing conditional formatting
    rng.FormatConditions.Delete
    
    ' Add conditional formatting for positive numbers
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0)
    End With
    
    ' Add conditional formatting for negative numbers
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0)
    End With
    
    ' Add conditional formatting for zero
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        .Interior.Color = RGB(255, 255, 255)
    End With
End Sub

Sub GreatestPercentageChanges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim maxVal As Double
    Dim minVal As Double
    Dim maxTicker As String
    Dim minTicker As String
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ' Get the last row in column J
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Define the range to search for percentages in column L
    Set rng = ws.Range("L2:L" & lastRow)
    
    ' max and min values
    maxVal = -15000 ' A very small number
    minVal = 15000  ' A very large number
    
    ' Loop through each cell in the range
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value > maxVal Then
                maxVal = cell.Value
                maxTicker = ws.Cells(cell.Row, "J").Value
            End If
            If cell.Value < minVal Then
                minVal = cell.Value
                minTicker = ws.Cells(cell.Row, "J").Value
            End If
        End If
    Next cell
    
    ' Place the results in the  cells
    ws.Range("P2").Value = maxTicker
    ws.Range("Q2").Value = maxVal
    ws.Range("P3").Value = minTicker
    ws.Range("Q3").Value = minVal
End Sub

Sub greatest_volume()

Dim ws As Worksheet
Dim lastRow As Long
Dim rng As Range
Dim cell As Range
Dim maxVal As Double
Dim maxTicker As String

    
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ' Get the last row in column J
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Define the range to search for total stock volumes in column M
    Set rng = ws.Range("M2:M" & lastRow)
    
    ' max volume
    maxVolume = 0
    
    ' Loop through each cell in the range
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value > maxVolume Then
                maxVolume = cell.Value
                maxTicker = ws.Cells(cell.Row, "J").Value
            End If
        End If
    Next cell
    
    ' Place the results in the  cells
    ws.Range("P4").Value = maxTicker
    ws.Range("Q4").Value = maxVolume
End Sub


'Create a script that loops through the rows to determine what the difference in the
'yearly opening and closing values for each stock as well as the percent change in value
'and the total volume of the stock for the entire year


Sub StockLoop()

'Creates a loop that will loop through all the worksheets
For Each ws In Worksheets

'Sets the last row where the loop will stop
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Defines the Variables
Dim Ticker As String
Dim OpenValue, CloseValue, Dif, PercentChange As Double
Dim Volume, Start, j As Integer
Dim MaximumVolume As LongLong
Dim MaximumPercentChange, MinimumPercentChange As Double, MaxCell As Range
Dim MaxValue As String
Dim rowNum As Integer

'Sets the headers for columns and rows
ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N1").Value = "Ticker Symbol"
ws.Range("O1").Value = "Value"
ws.Range("M2").Value = "Greatest % Increase"
ws.Range("M3").Value = "Greatest % Decrease"
ws.Range("M4").Value = "Greatest Volume"

'Sets the starting values for the variables that need to be counted
Volume = 0
Start = 2
j = 2

'Creates a For loop that will loop through all the rows
    For i = 2 To lastrow
    
    'Creates a conditional that activates when one row is not equal to the previous row
    'It will returen the total volume, stock ticker, yearly difference and percent change
    'It will also return a green color in the cell for a positive yearly change and a red for negative
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        'Assigns values to their appropriate cells
            Volume = Volume + ws.Range("G" & i).Value
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(j, 9).Value = Ticker
            ws.Cells(j, 12).Value = Volume
            OpenValue = ws.Cells(Start, 3).Value
            CloseValue = ws.Cells(i, 6).Value
            
        'Defines a new variable
            Dif = CloseValue - OpenValue
            ws.Range("J" & j).Value = Dif
        
        'Assigns colors to the approprate cells
            If ws.Range("J" & j).Value >= 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & j).Value < 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 3
            Else
                ws.Range("J" & j).Interior.ColorIndex = 2
            End If
            
        'Fixes the errors caused by the division of 0
            If OpenValue = 0 Then
                ws.Range("K" & j).Value = FormatPercent(0, 2, , , vbFalse)
            Else
                ws.Range("K" & j).Value = FormatPercent(Dif / OpenValue, 2, , , vbFalse)
            End If
        
            Start = i + 1
            Volume = 0
            j = j + 1
       
       'Adds up the volumes from the entire year of stocks
        Else
            Volume = Volume + ws.Range("G" & i).Value
     
        End If
        
    Next i
    
'Resizes the L column
    ws.Columns("L:L").AutoFit
    
'Assigns values to appropriate cells
    ws.Range("N1").Value = "Ticker Symbol"
    ws.Range("O1").Value = "Value"
    ws.Range("M2").Value = "Greatest % Increase"
    ws.Range("M3").Value = "Greatest % Decrease"
    ws.Range("M4").Value = "Greatest Volume"
    
'Creates a loop that pulls the greatest value for the Percent Change column
'Determines the row with the greatest value and assigns the row number to formula to find ticker symbol
'Assigns the ticker symbol to N2 and the value of the Percent Change to O2
    MaximumPercentChange = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Range("O2").Value = FormatPercent(MaximumPercentChange, 2, , , vbFalse)
    MaxValue = FormatPercent(MaximumPercentChange, 2, , , vbFalse)
    Set MaxCell = ws.Range("K:K").Find(MaxValue, , LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    rowNum = MaxCell.Row
    ws.Range("N2").Value = ws.Cells(rowNum, 9).Value
    
'Creates a loop that pulls the minimum value for the Percent Change column
'Determines the row with the minimum value and assigns the row number to formula to find ticker symbol
'Assigns the ticker symbol to N3 and the value of the Percent Change to O3
    MinimumPercentChange = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Range("O3").Value = FormatPercent(MinimumPercentChange, 2, , , vbFalse)
    MaxValue = FormatPercent(MinimumPercentChange, 2, , , vbFalse)
    Set MaxCell = ws.Range("K:K").Find(MaxValue, , LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    rowNum = MaxCell.Row
    ws.Range("N3").Value = ws.Cells(rowNum, 9).Value
    
'Creates a loop that pulls the greatest value for the Total Stock Volume column
'Resizes the O column
'Determines the row with the greatest value and assigns the row number to formula to find ticker symbol
'Assigns the ticker symbol to N4 and the value of the Total Stock Volume to O4
    MaximumVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Range("O4").Value = MaximumVolume
    ws.Columns("O:O").AutoFit
    Set MaxCell = ws.Range("L:L").Find(MaximumVolume, , LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    rowNum = MaxCell.Row
    ws.Range("N4").Value = ws.Cells(rowNum, 9).Value

Next ws

End Sub


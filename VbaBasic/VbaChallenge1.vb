Sub StockRunner()

Dim WorksheetName As String

'declare vars
    'openPrice
    dim openPrice as double
    'closePrice
    dim closePrice as double
    'ticker
    dim ticker as string
    'tickerHigh & Low
    dim tickerHigh as string
    dim tickerLow as string
    'vol
    dim vol as double
    'volCapture
    dim volCapture as double
    dim tickerVol as string
    'percentChange
    dim percentChange as double
    'percentCapture
    dim percentHigh as double
    dim percentLow as double
    'summaryRow
    dim summaryRow as integer
    'change
    dim change as double
    dim LR as long 
    dim i as long

    LR = Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ""
    tickerHigh = ""
    tickerLow = ""
    tickerVol = ""
    closePrice = 0
    openPrice = 0
    summaryRow = 2
    percentChange = 0
    vol = 0
    percentHigh = 0
    volCapture = 0
  
        Cells(1, 9).Value = "ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Stock Volume"

'loop through rows and create summary table
For i = 2 to LR

'identify if the next ticker is different
If cells(i,1).Value <> cells(i+1,1).Value Then

    'capture closePrice
    closePrice = cells(i,6).Value
    change = cells(i,6).Value - openPrice
    'calculate percentChange between open and close
    if openPrice <> 0 then
    percentChange = (change/openPrice)*100 
    Range("K" & summaryRow).Value = percentChange & "%"
    else 
    Range("K" & summaryRow).Value = "-"
    end if

'push my data to the summary table
    'push ticker to summary
     Range("I" & summaryRow).Value = ticker
    'push change to summary
     Range("J" & summaryRow).Value = change

    
    'push vol to summary
    Range("L" & summaryRow).Value = vol
    'iterate summaryRow counter

        'Color Code According to Negative or positive Yearly Change
        if change > 0 then
        Range("J" & summaryRow).Interior.ColorIndex = 4

        elseif change < 0 then
        Range("J" & summaryRow).Interior.ColorIndex = 3

        elseif change = 0 then
        Range("J" & summaryRow).Interior.ColorIndex = 6
        
        end If

        'Capture Greatest % Increase
        if percentHigh < percentChange then
        percentHigh = percentChange
        tickerHigh = cells(i,1).Value
        
        'Capture Greatest % Decrease
        elseif percentLow > percentChange and openPrice <> 0 then
        percentLow = percentChange
        tickerLow = cells(i,1).Value
        end if
        'Greatest total Volume
        if vol > volCapture then
        volCapture = vol
        tickerVol = cells(i,1).Value
        end if

        'summaryRow = summaryRow + 1
        summaryRow = summaryRow + 1
        'set vol = 0
        vol = 0
        'set openPrice = 0
        openPrice = 0
        'set closePrice = 0
        closePrice = 0
        'set percentChange = 0
        percentChange = 0
        'set change = 0
        change = 0
        'set ticker = " "
        ticker = ""

 

'if ticker is the same   
ElseIf cells(i,1).Value = cells(i+1,1).Value Then

        If openPrice = 0 Then
            'capture openPrice
            openPrice = cells(i,3).Value
            'capture ticker
            ticker = cells(i,1).Value
            
        End If
        'define vol = vol + Cells(i,7).Value
        vol = vol + Cells(i,7).Value


End If

        
Next i

        'Greatest Total Volume, % Increase & % Decrease

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"

        Cells(1, 16).Value = "ticker"
        Cells(2, 16).Value = tickerHigh
        Cells(3, 16).Value = tickerLow
        Cells(4, 16).Value = tickerVol

        
        Cells(1, 17).Value = "Value"
        Cells(2, 17).Value = percentHigh & "%"
        Cells(3, 17).Value = percentLow & "%"
        Cells(4, 17).Value = volCapture

        tickerHigh = ""
        tickerLow = ""
        tickerVol = ""
        percentLow = 0
        percentHigh = 0 
        volCapture = 0

End Sub
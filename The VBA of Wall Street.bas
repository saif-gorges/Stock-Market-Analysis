Attribute VB_Name = "Module1"
Sub Analysis()

'Worksheet loop
For Each ws In Worksheets

'Variable declaration
Dim rows As Long
Dim tickerrows As Long
Dim yearopen As Double
Dim yearclose As Double
Dim totalvolume As Double
Dim summary As Range
Dim maxinc As Double
Dim mininc As Double
Dim maxtot As Double

'Cell titles and formatting
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("Q2:Q3").NumberFormat = "0.00%"

'Populate ticker column and total volume
rows = Application.CountA(ws.Range("A:A"))
tickerrows = 2

For i = 2 To rows

    If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
    
    ws.Cells(tickerrows, 9) = ws.Cells(i, 1)
    tickerrows = tickerrows + 1
    totalvolume = 0
    
    Else
    totalvolume = totalvolume + ws.Cells(i, 7)
    
    End If

ws.Cells(tickerrows - 1, 12) = totalvolume

Next i

'Find first and last occurence of each ticker symbol
tickerrows = tickerrows - 1
For i = 2 To tickerrows

    yearopen = ws.Columns(1).Find(What:=ws.Cells(i, 9), LookAt:=xlWhole).Row
    yearclose = ws.Columns(1).Find(What:=ws.Cells(i, 9), LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
    
    'Calculate and display yearly change
    ws.Cells(i, 10) = ws.Cells(yearclose, 6) - ws.Cells(yearopen, 3)
    
        'Format cells based on yearly change
        If ws.Cells(i, 10) > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    'Calculate and display percent change
    If ws.Cells(yearopen, 3) <> 0 Then
        ws.Cells(i, 11) = (ws.Cells(yearclose, 6) - ws.Cells(yearopen, 3)) / ws.Cells(yearopen, 3)
    Else
        ws.Cells(i, 11) = 0
    End If
    
Next i


'Calculate greatest % increase and display with ticker symbol
Set summary = ws.Range("K:K")
maxinc = Application.WorksheetFunction.Max(summary)
ws.Cells(2, 17) = maxinc
ws.Cells(2, 16) = ws.Cells(Application.WorksheetFunction.Match(maxinc, summary, 0), 9)

'Calculate greatest % decrease and display with ticker symbol
mininc = Application.WorksheetFunction.Min(summary)
ws.Cells(3, 17) = mininc
ws.Cells(3, 16) = ws.Cells(Application.WorksheetFunction.Match(mininc, summary, 0), 9)

'Calculate greatest total volume and display with ticker symbol
Set summary = ws.Range("L:L")
maxtot = Application.WorksheetFunction.Max(summary)
ws.Cells(4, 17) = maxtot
ws.Cells(4, 16) = ws.Cells(Application.WorksheetFunction.Match(maxtot, summary, 0), 9)

Next ws

End Sub

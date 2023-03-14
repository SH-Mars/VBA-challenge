Attribute VB_Name = "Module1"
Sub loop_allws():
    
    Dim ws As Worksheet
    
    WkSheets = Array("2018", "2019", "2020")
    
    For Each ws In Sheets(Array("2018", "2019", "2020"))
        ws.Select
        
        Call get_ticker
        Call YearlyPrice
        Call matrix
    
    Next ws

End Sub

Sub get_ticker():

    Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True

End Sub

Sub YearlyPrice():

    Dim rowcount As Long
    Dim tickercount As Long

    rowcount = Cells(Rows.Count, 1).End(xlUp).Row
    tickercount = 2
    
    Range("J1").value = "Yearly Change"
    Range("K1").value = "Percent Change"
    Range("L1").value = "Total Stock Volume"
    
    For i = 2 To rowcount

        'If there is a new ticker then
        If Cells(i, 1).value <> Cells(i + 1, 1) Then
        'Take the (close price - open price) and put it in the appropriate cell
        Cells(tickercount, 10).value = Cells(i, 6).value - Cells(i - 250, 3).value
        'Take the ((close price/open price)-1) and put it in the appropriate cell
        Cells(tickercount, 11).value = (Cells(i, 6).value / Cells(i - 250, 3).value) - 1
        'Sum the volume
        Cells(tickercount, 12).value = WorksheetFunction.Sum(Range(Cells(i, 7), Cells(i - 250, 7)))
        
        CondColor = Cells(tickercount, 10).value
        Select Case CondColor
            Case Is > 0
                Cells(tickercount, 10).Interior.ColorIndex = 4
            Case Is < 0
                Cells(tickercount, 10).Interior.ColorIndex = 3
        End Select
        
        Cells(tickercount, 11).NumberFormat = "0.00%"
            
        
        tickercount = tickercount + 1

        End If
    Next i
    
End Sub

Sub matrix():

    Dim value As Double
    Dim ticker As String
    
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    For i = 2 To 3001
    
    ticker = Cells(i, 9).value
    
        value = WorksheetFunction.Max(Range("K2:K3001"))
        Range("Q2").value = value
        
        If Cells(i, 11).value = value Then
            Range("P2").value = ticker
        End If
    Next i
    Range("Q2").NumberFormat = "0.00%"
    
    For i = 2 To 3001
        
    ticker = Cells(i, 9).value
    
        value = WorksheetFunction.Min(Range("K2:K3001"))
        Range("Q3").value = value
        
        If Cells(i, 11).value = value Then
            Range("P3").value = ticker
        End If
    Next i
    Range("Q3").NumberFormat = "0.00%"
    
    x = WorksheetFunction.Max(Range("L2:L3001"))
    
    For i = 2 To 3001
    
    ticker = Cells(i, 9).value
    
        Range("Q4").value = x
        
        If Cells(i, 12).value = x Then
            Range("P4").value = ticker
        End If
    Next i
    
End Sub

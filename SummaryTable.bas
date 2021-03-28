Attribute VB_Name = "Module1"
Sub Summary():

' declare variables
    Dim LastRow As Long
    Dim Row As Long
    Dim Counter As Integer
    Dim FirstPrice As Single
    Dim LastPrice As Single
    
    
' Create Table headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Price Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
'Find Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Initialize variables
    Counter = 2
    Row = 2
' Store FirstPrice
    FirstPrice = Cells(Row, 3).Value
'Print First Ticker Symbol to Table
    Cells(Counter, 9).Value = Cells(2, 1).Value
       
'Loop to find next ticker symbol
    For Row = 2 To LastRow
       
       
        If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
        'Store Lastprice
        LastPrice = Cells(Row, 6).Value
        'Print Price Change to Table
        Cells(Counter, 10) = (LastPrice - FirstPrice)
        'Print Percent Change to table
        Cells(Counter, 11) = (LastPrice - FirstPrice) / FirstPrice
        
        'Store firstprice for new ticker
        FirstPrice = Cells(Row + 1, 3).Value
        
        
        
        'Step up Summary Counter
        Counter = Counter + 1
        'Print Ticker Symbol for next
        Cells(Counter, 9) = Cells(Row + 1, 1).Value
        
       
        

        
        End If
    
        Row = Row + 1
    Next Row

    
    
End Sub

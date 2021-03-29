Attribute VB_Name = "Module1"
Sub Summary():

' declare variables
    Dim LastRow As Long
    Dim Row As Long
    Dim Counter As Integer
    Dim FirstPrice As Double
    Dim Volume As Double
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets

    ws.Activate

' Create Table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Price Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"


'Find Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Initialize variables
    Counter = 2
    Row = 2
    Volume = 0
' Store FirstPrice
    FirstPrice = ws.Cells(Row, 3).Value
'Print First Ticker Symbol to Table
    ws.Cells(Counter, 9).Value = ws.Cells(Row, 1).Value
       
'Loop to find next ticker symbol
    For Row = 2 To LastRow
       
        'Add Volume to total
        
        Volume = (Volume + (ws.Cells(Row, 7).Value))
        
        If ws.Cells(Row, 1).Value <> ws.Cells((Row + 1), 1).Value Then

        'Print Price Change to Table
        ws.Cells(Counter, 10).Value = ws.Cells(Row, 6).Value - FirstPrice
        
        
        'Print Total Voloume to Table
        ws.Cells(Counter, 12).Value = Volume
        'Store firstprice for new ticker
        FirstPrice = ws.Cells(Row + 1, 3).Value
        'Reset Volume to 0
        Volume = 0
        'Step up Summary Counter
        Counter = Counter + 1
        'Print Ticker Symbol for next
        ws.Cells(Counter, 9) = ws.Cells(Row + 1, 1).Value

        End If
        If FirstPrice <> 0 Then 'Skip First Price if it equals 0
        ws.Cells(Counter, 11).Value = ((ws.Cells(Row, 6).Value - FirstPrice) / FirstPrice) 'Print Percent Change to table
        
        End If
    Next Row
    
    'Format Table
    ws.Columns("I:L").EntireColumn.AutoFit
    ws.Columns("K:K").NumberFormat = "0.00%"

Next
    
starting_ws.Activate 'activate the worksheet that was originally active


' Used code and coments from "http://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/" accessed 3/29/21
End Sub



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
            If ((ws.Cells(Row, 6).Value - FirstPrice) / FirstPrice) < 0 Then
                ws.Cells(Counter, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(Counter, 11).Interior.ColorIndex = 4
            End If
        End If
    Next Row
   
   
    'Declare Variable for callouts
    Dim Decrease As Double
    Dim Increase As Double
    Dim MaxVolume As Double
    Dim ChangeEval As Long
    
    'Draw table
    
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    
    ' Initialize Callout variables
    Decrease = 0
    Increase = 0
    MaxVolume = 0
    
    For ChangeEval = 2 To Counter
    
        'Find Greatest Increase
        If Cells(ChangeEval, 11) > Increase Then
            Increase = Cells(ChangeEval, 11)
            Cells(2, 15) = Cells(ChangeEval, 9) 'Print Ticker
            Cells(2, 16) = Cells(ChangeEval, 11) ' Print Value
        End If
        
        ' FindGreatest Decrease
        If Cells(ChangeEval, 11) < Decrease Then
            Decrease = Cells(ChangeEval, 11)
            Cells(3, 15) = Cells(ChangeEval, 9) 'Print Ticker
            Cells(3, 16) = Cells(ChangeEval, 11) ' Print Value
        End If
        
        ' Find Greatest Volume
        
        If Cells(ChangeEval, 12) > MaxVolume Then
            MaxVolume = Cells(ChangeEval, 12)
            Cells(4, 15) = Cells(ChangeEval, 9) 'Print Ticker
            Cells(4, 16) = Cells(ChangeEval, 12) ' Print Value
        End If
        
    Next ChangeEval
    
    'Format Table
    ws.Columns("I:P").EntireColumn.AutoFit
    ws.Columns("K:K").NumberFormat = "0.00%"
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"

Next
    
starting_ws.Activate 'activate the worksheet that was originally active


' Used code and coments from "http://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/" accessed 3/29/21
End Sub



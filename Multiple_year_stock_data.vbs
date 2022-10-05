Sub MultiYearStockData()



Dim Ticker As String
Dim i As Long
Dim j As Long
Dim LRow As Long
Dim LCol As Long
Dim OpenStart As Double
Dim CloseEnd As Double
Dim TickerEnd As String
Dim StockVolCount As Double
Dim SummaryCount As Long
Dim MaxChange As Double
Dim MaxChangeTicker As String
Dim MinChange As Double
Dim MinChangeTicker As String
Dim MaxVol As Double
Dim MaxVolTicker As String
Dim ws As Worksheet




'Application.Cursor = xlWait

Worksheets(1).Cells(10, "O").Value = "MACRO MAY TAKE UP TO 7 MINUTES TO COMPLETE"
Worksheets(1).Cells(10, "O").Font.ColorIndex = 3
Worksheets(1).Cells(10, "O").Font.Bold = True


Set ws = ActiveSheet


For Each ws In ActiveWorkbook.Worksheets

Application.ScreenUpdating = False
ws.Activate


    LCol = ws.Cells(100, Columns.Count).End(xlToLeft).Column
    LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OpenStart = ws.Cells(2, "C").Value
    StockVolCount = ws.Cells(2, "G").Value
    SummaryCount = 2
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    MaxChange = ws.Cells(2, "K").Value
    MinChange = ws.Cells(2, "K").Value
    MaxVol = ws.Cells(2, "L").Value
    MaxVolTicker = ws.Cells(2, "I").Value
    
    For i = 2 To LRow
        If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
            'define openend for ticker
            CloseEnd = ws.Cells(i, "F").Value
            ws.Cells(SummaryCount, "I").Value = ws.Cells(i, "A").Value
            'calculate yearly change
            ws.Cells(SummaryCount, "J").Value = CloseEnd - OpenStart
            ws.Cells(SummaryCount, "J").NumberFormat = "0.00"

            'calculate percent change
            ws.Cells(SummaryCount, "K").Value = (CloseEnd - OpenStart) / OpenStart
            ws.Cells(SummaryCount, "K").NumberFormat = "0.00%"
            ws.Cells(SummaryCount, "L").Value = StockVolCount
    
            'reset openstart and totalvol for next ticker
            OpenStart = ws.Cells(i + 1, "C")
            StockVolCount = ws.Cells(i + 1, "G")
            'increase summarycount
            SummaryCount = SummaryCount + 1
            'MsgBox (Cells(i, "A").Value)
        Else
            StockVolCount = StockVolCount + ws.Cells(i + 1, "G").Value
        End If
        
    Next i
    
    'determine min and max percent changes
    
    For i = 2 To LRow
        If ws.Cells(i, "K").Value > MaxChange Then
            MaxChange = ws.Cells(i, "K").Value
            MaxChangeTicker = ws.Cells(i, "I").Value
        ElseIf ws.Cells(i, "K").Value < MinChange Then
            MinChange = ws.Cells(i, "K").Value
            MinChangeTicker = ws.Cells(i, "I").Value
        End If
    Next i
    
    ws.Cells(2, "P").Value = MaxChangeTicker
    ws.Cells(2, "Q").Value = MaxChange
    ws.Cells(2, "Q").NumberFormat = "0.00%"
    ws.Cells(3, "P").Value = MinChangeTicker
    ws.Cells(3, "Q").Value = MinChange
    ws.Cells(3, "Q").NumberFormat = "0.00%"
    
    
    'determine greatest total volume
    For i = 2 To LRow
        If ws.Cells(i, "L").Value > MaxVol Then
            MaxVol = ws.Cells(i, "L").Value
            MaxVolTicker = ws.Cells(i, "I").Value
        End If
    Next i
        
    ws.Cells(4, "P").Value = MaxVolTicker
    ws.Cells(4, "Q").Value = MaxVol
    
    ws.Columns("I:Q").AutoFit
    'MsgBox ws.Name
    
'Run all conditionals at the end

   Application.CutCopyMode = False
   
    ws.Range("J2").Select
    ws.Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65280
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
Next ws


Worksheets(1).Cells(10, "O").Value = Null
Worksheets(1).Cells(10, "O").Font.ColorIndex = Null
Worksheets(1).Cells(10, "O").Font.Bold = False


'Application.Cursor = xlDefault
End Sub





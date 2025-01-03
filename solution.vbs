Attribute VB_Name = "Module11"
' Part 1:

' Create a script that loops through all the stocks for each quarter and outputs the following information:
' The ticker symbol
' Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
' The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
' The total stock volume of the stock. The result should match the following image:
' Add functionality to return stock with
'   Greatest % increase
'   Greatest % decrease
'   Greatest total volume
Sub Stock_summary()

    Dim i As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim LastCol As Long
    Dim PartOneCol As Long
    Dim PartTwoCol As Long
    
    For Each ws In Worksheets
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        PartOneCol = 9
        'MsgBox ("PartOneCol: " & PartOneCol)
        ws.Cells(1, PartOneCol).Value = "Ticker"
        ws.Cells(1, PartOneCol + 1).Value = "Quarterly Change"
        ws.Cells(1, PartOneCol + 2).Value = "Percent Change"
        ws.Cells(1, PartOneCol + 3).Value = "Total Stock Volume"
        PartTwoCol = 15
        ws.Cells(1, PartTwoCol + 1).Value = "Ticker"
        ws.Cells(1, PartTwoCol + 2).Value = "Value"
        ws.Cells(2, PartTwoCol).Value = "Greatest % Increase"
        ws.Cells(3, PartTwoCol).Value = "Greatest % Decrease"
        ws.Cells(4, PartTwoCol).Value = "Greatest Total Volume"
        'MsgBox ("PartTwoCol: " & PartTwoCol)
        Dim CurTicker As String
        Dim NextTicker As String
        CurTicker = ""
        Dim CurRow As Integer
        CurRow = 2
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim TotalVolume As LongLong
        ' MsgBox ("OpenPrice: " & OpenPrice & "CLosePrice: " & ClosePrice)
        For i = 2 To LastRow
            NextTicker = ws.Cells(i, 1).Value
            If CurTicker = "" Then
                ws.Cells(2, PartOneCol).Value = NextTicker
                CurTicker = NextTicker
                OpenPrice = ws.Cells(i, 3).Value
                TotalVolume = ws.Cells(i, 7)
                ws.Cells(CurRow, PartOneCol + 3).Value = TotalVolume
            ElseIf CurTicker = NextTicker Then
                ClosePrice = ws.Cells(i, 6).Value
                ws.Cells(CurRow, PartOneCol + 1).Value = ClosePrice - OpenPrice
                ws.Cells(CurRow, PartOneCol + 2).Value = ws.Cells(CurRow, PartOneCol + 1).Value / OpenPrice
                TotalVolume = TotalVolume + ws.Cells(i, 7)
                ws.Cells(CurRow, PartOneCol + 3).Value = TotalVolume
            ElseIf CurTicker <> NextTicker Then
                CurRow = CurRow + 1
                ws.Cells(CurRow, PartOneCol).Value = NextTicker
                CurTicker = NextTicker
                OpenPrice = ws.Cells(i, 3).Value
                TotalVolume = ws.Cells(i, 7)
                ws.Cells(CurRow, PartOneCol + 3).Value = TotalVolume
            End If
        Next i
        
        LastRowPartTwo = ws.Cells(Rows.Count, PartOneCol).End(xlUp).Row
        For i = 2 To LastRowPartTwo
            If ws.Cells(i, PartOneCol + 1).Value > 0 Then
                ws.Cells(i, PartOneCol + 1).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, PartOneCol + 1).Value < 0 Then
                ws.Cells(i, PartOneCol + 1).Interior.ColorIndex = 3
            End If
        Next i
        ws.Range("K2:K" & LastRowPartTwo).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Cells(2, PartTwoCol + 2).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowPartTwo))
        ws.Cells(3, PartTwoCol + 2).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowPartTwo))
        ws.Cells(4, PartTwoCol + 2).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowPartTwo))
        
        ws.Cells(2, PartTwoCol + 1).Value = Application.WorksheetFunction.Index(ws.Range("I2:I" & LastRowPartTwo), Application.WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & LastRowPartTwo), 0))
        ws.Cells(3, PartTwoCol + 1).Value = Application.WorksheetFunction.Index(ws.Range("I2:I" & LastRowPartTwo), Application.WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & LastRowPartTwo), 0))
        ws.Cells(4, PartTwoCol + 1).Value = Application.WorksheetFunction.Index(ws.Range("I2:I" & LastRowPartTwo), Application.WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & LastRowPartTwo), 0))
    Next ws
    
    MsgBox ("Summary Complete")
End Sub

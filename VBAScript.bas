Attribute VB_Name = "Module1"

Sub change()
Dim ws As Worksheet
Dim total As Double
Dim i As Long
Dim change As Single
Dim j As Integer
Dim start As Long
Dim rowCount As Long
Dim percentchange As Single
Dim percentincrease As Integer
Dim percentdecrease As Integer
Dim highestvolume As Integer

For Each ws In Worksheets
    total = 0
    j = 0
    change = 0
    start = 2
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    For i = 2 To rowCount
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            total = total + ws.Cells(i, 7).Value
            If total = 0 Then
                ws.Range("I" & j + 2).Value = ws.Cells(i, 1).Value
                ws.Range("J" & j + 2).Value = 0
                ws.Range("K" & j + 2).Value = "%" & 0
                ws.Range("L" & j + 2).Value = 0
            Else
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                        Exit For
                    End If
                Next find_value
            End If
        change = (ws.Cells(i, 6) - ws.Cells(start, 3))
        percentchange = Round((change / ws.Cells(start, 3) * 100), 2)
    
        start = i + 1
        
        ws.Range("I" & j + 2).Value = ws.Cells(i, 1).Value
        ws.Range("J" & j + 2).Value = Round(change, 2)
        ws.Range("K" & j + 2).Value = "%" & percentchange
        ws.Range("L" & j + 2).Value = total
        
        Select Case change
            Case Is > 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            Case Else
                ws.Range("J" & 2 + j).Interior.ColorIndex = 0
        End Select
    End If
    
        change = 0
        total = 0
        j = j + 1
    Else
        total = total + ws.Cells(i, 7).Value

End If
Next i

        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        percentincrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        percentdecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        highestvolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        ws.Range("P2") = ws.Cells(percentincrease + 1, 9)
        ws.Range("P3") = ws.Cells(percentdecrease + 1, 9)
        ws.Range("P4") = ws.Cells(highestvolume + 1, 9)

Next ws
End Sub

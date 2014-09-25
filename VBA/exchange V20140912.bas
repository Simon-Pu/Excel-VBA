Attribute VB_Name = "exchange rate"
Sub exchangeCalculate()

    Dim DaysofCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim cnt As Integer
    Dim 程 As Integer
    Dim TotalRunTime As String
    Dim TotalRunDate As Date
    
    Worksheets("exchange").Visible = xlSheetVisible
    Worksheets("Temp").Visible = xlSheetVisible
    Worksheets("List").Visible = xlSheetVisible
    TotalRunDate = Now
    
    程 = Sheets("exchange").Range("A" & Rows.Count).End(xlUp).Row
    Sheets("exchange").Range("C10:I" & 程).ClearContents
    Sheets("exchange").Range("A10:B65536").ClearContents
    
    Sheets("Result").Range("E8") = Sheets("List").Range("B" & (Sheets("List").Range("C1") + 1))
    DaysofCount = CDbl(Sheets("Result").Range("C4") - Sheets("Result").Range("C2"))
    '
    'Perpared of dynamic info for iqy needed
    Sheets("Temp").Range("A1") = Format((Sheets("Result").Range("C2")), "yyyymmdd")
    Sheets("Temp").Range("B1") = Sheets("Result").Range("E8")
    
    Sheets("Temp").Select
    Application.ScreenUpdating = False
    
    'query date from Web
    For i = 0 To DaysofCount
        Sheets("exchange").Range("A" & i + 8) = Sheets("Result").Range("C2") + i
        Sheets("Temp").Range("A1") = Format((Sheets("exchange").Range("A" & i + 8)), "yyyymmdd")
        
        'Sheets("Temp").Select
        Sheets("Temp").Range("A8:B10").ClearContents
        ActiveWorkbook.RefreshAll
        Application.Wait (Now + TimeValue("0:00:01"))
        
        Sheets("exchange").Range("B" & i + 8) = Sheets("Temp").Range("A9")
        
        'delete Connections iqy if count >1
        cnt = ActiveWorkbook.Connections.Count
        For j = 1 To cnt
            If j + 1 <= cnt Then
                ActiveWorkbook.Connections.Item(j + 1).Delete
            End If
        Next j
        Application.StatusBar = "Processing done about " & Format(i / DaysofCount, "0%")
        
    Next i
    
    'delete no result of raw
    For i = 0 To DaysofCount
        If IsEmpty(Sheets("exchange").Cells((DaysofCount + 8 - i), 2)) Then
            Sheets("exchange").Cells((DaysofCount + 8 - i), 2).EntireRow.Delete
        End If
    Next i
    
    'Force set first item no to 1 if the first raw was empty
    If Sheets("exchange").Range("C8") <> 1 Then
        Sheets("exchange").Range("C8") = 1
    End If
    If IsEmpty(Sheets("exchange").Range("C9")) Then
        Sheets("exchange").Range("C9").Formula = "=IF(B9<>"", C8+1)"
        Sheets("exchange").Range("D9").Formula = "=C9*$D$3+$D$4"
        Sheets("exchange").Range("E9").Formula = "=B9-D9"
        Sheets("exchange").Range("F9").Formula = "=D9+(2*$D$5)"
        Sheets("exchange").Range("G9").Formula = "=D9+$D$5"
        Sheets("exchange").Range("H9").Formula = "=D9-$D$5"
        Sheets("exchange").Range("I9").Formula = "=D9-(2*$D$5)"
    End If
    
    
    '干贾硄笵璸衡そΑ戈
    Sheets("exchange").Select
    'Application.ScreenUpdating = False
    程 = Sheets("exchange").Range("A" & Rows.Count).End(xlUp).Row
    If 程 > 9 Then
        Application.Calculation = xlCalculationManual
        Sheets("exchange").Range("C9:I9").AutoFill Destination:=Sheets("exchange").Range("C9:I" & 程)
        Application.Calculation = xlCalculationAutomatic
    End If

    
    'Re-Draw Chart
    UpdateDrawChart
    
    Application.ScreenUpdating = True
    
    Sheets("Result").Select
    TotalRunTime = Format((Now - TotalRunDate), "h:mm:ss")
    Application.StatusBar = "Finished and Total run about " & TotalRunTime
    
    Worksheets("exchange").Visible = xlSheetHidden
    Worksheets("Temp").Visible = xlSheetHidden
    Worksheets("List").Visible = xlSheetHidden

End Sub



Sub UpdateDrawChart()

    Dim ChartTitle As String
    Dim rng As Range
    Dim min As Double, min2 As Double, max As Double, max2 As Double
    Dim O_Index As Double, R_Line As Double, U95_trend As Double, U75_trend As Double
    Dim D95_trend As Double, D75_trend As Double
    Dim 穝程 As Integer
    Dim ChartTitleIndex As Integer
    Dim FinalRow As Integer
    
    Worksheets("exchange").Visible = xlSheetVisible
    穝程 = Sheets("exchange").Range("A" & Rows.Count).End(xlUp).Row
    Sheets("exchange").Select
    'Set range from which to determine smallest value
    Set rng = Sheets("exchange").Range("B8:B" & 穝程)

    'Worksheet function MIN returns the smallest value in a range
    If Not IsEmpty(Sheets("exchange").Range("B8")) Then
        min = Application.WorksheetFunction.min(rng)
        max = Application.WorksheetFunction.max(rng)
    End If
    
    If Not IsError(Sheets("exchange").Range("I8")) Then
        Set rng = Sheets("exchange").Range("I8:I" & 穝程)
        min2 = Application.WorksheetFunction.min(rng)
    End If
    
    If Not IsError(Sheets("exchange").Range("F8")) Then
        Set rng = Sheets("exchange").Range("F8:F" & 穝程)
        max2 = Application.WorksheetFunction.max(rng)
    End If
    
    Set rng = Nothing
    
    If max < max2 Then
        max = max2
    End If
    
    If min2 < min Then
        min = min2
    End If
    
    
    Sheets("Result").Select
    ActiveSheet.ChartObjects("Chart 1").Activate
    If Err Then
        ActiveSheet.ChartObjects("瓜 1").Select
        On Error GoTo 0
    End If
    With ActiveChart
        ActiveChart.HasTitle = True
        ChartTitleIndex = Sheets("list").Cells(1, 3) + 1
        ChartTitle = Sheets("list").Cells(ChartTitleIndex, 1)
        ActiveChart.ChartTitle.Text = ChartTitle & " 蹲瞯-贾き絬眯"
        
        i = .SeriesCollection.Count
        'MsgBox i
        
        For i = i + 1 To 6
            ActiveChart.SeriesCollection.NewSeries
            If i = 1 Then
                ActiveChart.SeriesCollection(1).Name = "=exchange!$B$7"
            ElseIf i = 2 Then
                ActiveChart.SeriesCollection(2).Name = "=exchange!$D$7"
            ElseIf i = 3 Then
                ActiveChart.SeriesCollection(3).Name = "=exchange!$F$7"
            ElseIf i = 4 Then
                ActiveChart.SeriesCollection(4).Name = "=exchange!$G$7"
            ElseIf i = 5 Then
                ActiveChart.SeriesCollection(5).Name = "=exchange!$H$7"
            ElseIf i = 6 Then
                ActiveChart.SeriesCollection(6).Name = "=exchange!$I$7"
            End If
        Next
        
        FinalRow = 穝程
        'MsgBox FinalRow
        If i = .SeriesCollection.Count <= 6 Then
            ActiveChart.SeriesCollection(1).XValues = Sheets("exchange").Range("A8:A" & FinalRow)
            ActiveChart.SeriesCollection(1).Values = Sheets("exchange").Range("B8:B" & FinalRow)
            ActiveChart.SeriesCollection(2).XValues = Sheets("exchange").Range("A8:A" & FinalRow)
            ActiveChart.SeriesCollection(2).Values = Sheets("exchange").Range("D8:D" & FinalRow)
            ActiveChart.SeriesCollection(3).XValues = Sheets("exchange").Range("A8:A" & FinalRow)
            ActiveChart.SeriesCollection(3).Values = Sheets("exchange").Range("F8:F" & FinalRow)
            ActiveChart.SeriesCollection(4).XValues = Sheets("exchange").Range("A8:A" & FinalRow)
            ActiveChart.SeriesCollection(4).Values = Sheets("exchange").Range("G8:G" & FinalRow)
            ActiveChart.SeriesCollection(5).XValues = Sheets("exchange").Range("A8:A" & FinalRow)
            ActiveChart.SeriesCollection(5).Values = Sheets("exchange").Range("H8:H" & FinalRow)
            ActiveChart.SeriesCollection(6).XValues = Sheets("exchange").Range("A8:A" & FinalRow)
            ActiveChart.SeriesCollection(6).Values = Sheets("exchange").Range("I8:I" & FinalRow)
        Else
                MsgBox "Active Chart Series Count >6 "
        End If
        
    End With
    
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = min * 0.98
        .MaximumScale = max * 1.02
    End With
    ActiveChart.ChartArea.Select
    
    'Show Last result
    Sheets("exchange").Select
    'Date
    Sheets("exchange").Cells(2, 7) = Sheets("exchange").Cells(FinalRow, 1)
    '夹
    Sheets("exchange").Cells(3, 7) = Sheets("exchange").Cells(FinalRow, 2)
    O_Index = Sheets("exchange").Cells(3, 7)
    '+2SD
    Sheets("exchange").Cells(5, 7) = Sheets("exchange").Cells(FinalRow, 6)
    U95_trend = Sheets("exchange").Cells(5, 7)
    '+SD
    Sheets("exchange").Cells(3, 9) = Sheets("exchange").Cells(FinalRow, 7)
    U75_trend = Sheets("exchange").Cells(3, 9)
    '-SD
    Sheets("exchange").Cells(4, 9) = Sheets("exchange").Cells(FinalRow, 8)
    D75_trend = Sheets("exchange").Cells(4, 9)
    '-2SD
    Sheets("exchange").Cells(5, 9) = Sheets("exchange").Cells(FinalRow, 9)
    D95_trend = Sheets("exchange").Cells(5, 9)
    'Regression Line
    Sheets("exchange").Cells(4, 7) = Sheets("exchange").Cells(FinalRow, 4)
    R_Line = Sheets("exchange").Cells(FinalRow, 4)
    Sheets("exchange").Cells(2, 9).Interior.ColorIndex = 2
    
    If O_Index <= D95_trend Then
        Sheets("exchange").Cells(2, 9).Interior.ColorIndex = 3
        Sheets("exchange").Cells(2, 9) = 1
    End If
    If ((O_Index > D95_trend) And (O_Index <= D75_trend)) Then
        Sheets("exchange").Cells(2, 9).Interior.ColorIndex = 3
        Sheets("exchange").Cells(2, 9) = 2
    End If
    If ((O_Index > D75_trend) And (O_Index <= R_Line)) Then
        Sheets("exchange").Cells(2, 9) = 3
    End If
    If ((O_Index > R_Line) And (O_Index <= U75_trend)) Then
        Sheets("exchange").Cells(2, 9) = 4
    End If
    If ((O_Index > U75_trend) And (O_Index <= U95_trend)) Then
        Sheets("exchange").Cells(2, 9).Interior.ColorIndex = 4
        Sheets("exchange").Cells(2, 9) = 5
    End If
    If O_Index > U95_trend Then
        Sheets("exchange").Cells(2, 9).Interior.ColorIndex = 4
        Sheets("exchange").Cells(2, 9) = 6
    End If
    
    Sheets("exchange").Cells(2, 11) = (((CDate(Sheets("exchange").Cells(FinalRow, 1))) - (CDate(Sheets("exchange").Cells(8, 1)))) / 365)
    Worksheets("exchange").Visible = xlSheetHidden

End Sub


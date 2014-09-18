Attribute VB_Name = "Module4"
'# MSCIIndex-曾氏通道計算
'# CopyRight@小工友, simon.pu@gmail.com

Sub CalculateMSCI_GPS()
    Dim basebook As Workbook
    Dim mybook As Workbook
    Dim sourceRange As Range
    Dim destrange As Range
    Dim SourceRcount As Long
    Dim n As Long
    Dim rnum As Long
    Dim MyPath As String
    Dim SaveDriveDir As String
    Dim FName(2) As Variant
    Dim ChartTitle As String
    Dim rng As Range
    Dim min As Double, min2 As Double, max As Double, max2 As Double
    Dim O_Index As Double, R_Line As Double, U95_trend As Double, U75_trend As Double
    Dim D95_trend As Double, D75_trend As Double
    Dim NumofGPS As Integer
    Dim x As Integer
    Dim TotalofTime As Date

    
    Application.ScreenUpdating = False
    Sheets("MSCI").Visible = xlSheetVisible
    Sheets("MSCI_Index_List").Visible = xlSheetVisible
    Sheets("MSCI_GPS").Select
       
    SaveDriveDir = CurDir
    MyPath = "C:\Test"
    If Len(Dir(MyPath, vbDirectory)) = 0 Then
        MkDir (MyPath)
        Sheets("MSCI").Range("L3") = MyPath
    End If
    
    ChDrive MyPath
    ChDir MyPath
    
    Sheets("MSCI_GPS").Range("B3:K57").Cells.ClearContents
    
    'FolderSelection
    
    'If Sheets("MSCI").Range("L3").Value = "-" Then
    '    GoTo EndAndExit
    'End If
    Sheets("MSCI_GPS").Range("I1") = Date
    
    TotalofTime = Now
    
    SetDefaultDate
    Sheets("MSCI_GPS").Range("G1") = Sheets("Result").Range("C6")
    
    NumofGPS = Sheets("MSCI_Index_List").Range("A" & Rows.Count).End(xlUp).Row
    
    For x = 2 To NumofGPS
        Sheets("MSCI").Range("K3") = x - 1
    
        DownloadFiles
    
        最後一列 = Sheets("MSCI").Range("A" & Rows.Count).End(xlUp).Row
        Sheets("MSCI").Range("C10:J" & 最後一列).Cells.ClearContents
        Sheets("MSCI").Range("A10:B" & 最後一列).Cells.ClearContents
    
        Sheets("MSCI").Activate
    
        FName(0) = Sheets("MSCI").Range("L3").Value & "\" & Sheets("MSCI").Range("M3").Value
    
        If IsArray(FName) Then
            Application.ScreenUpdating = False
            Set basebook = ThisWorkbook

            For n = LBound(FName) To UBound(FName)
                Set mybook = Workbooks.Open(FName(n))
            
                ' Add 50 more for extra buffer
                rnum = LastRow(basebook.Worksheets(1)) + 50
            
                For i = 1 To rnum
                    If mybook.Worksheets(1).Range("A" & i).Value <> "Date" Then
                    'j = i
                    'mybook.Worksheets(1).Rows(i).Delete
                    'mybook.Worksheets(1).Range("A" & i).Delete
                    ElseIf mybook.Worksheets(1).Range("A" & i).Value = "Date" Then
                        j = i
                        GoTo FindDate
                    End If
                Next i
FindDate:
                For i = 1 To j - 1
                    mybook.Worksheets(1).Rows(j - i).Delete
                Next i
            
                'rnum = mybook.Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
                j = 0
                For i = 1 To rnum
                    If mybook.Worksheets(1).Range("A" & i).Value = "" Then
                        j = i
                        GoTo FindDate2
                    End If
                Next i
FindDate2:
                j = rnum - j
                n = 0
                For i = 0 To j
                    'N = 0
                    mybook.Worksheets(1).Rows(rnum - n).Delete
                    n = n + 1
                Next i
            
                Set sourceRange = mybook.Worksheets(1).Range("A1:B" & rnum)
                SourceRcount = sourceRange.Rows.Count
                Set destrange = basebook.Worksheets("MSCI").Cells(7, "A")

                'basebook.Worksheets(1).Cells(rnum, "D").Value = mybook.Name
                ' This will add the workbook name in column D if you want

                sourceRange.Copy destrange
                ' Instead of this line you can use the code below to copy only the values

                ' With sourceRange
                    ' Set destrange = basebook.Worksheets(1).Cells(rnum, "A"). _
                    ' Resize(.Rows.Count, .Columns.Count)
                ' End With
                ' destrange.Value = sourceRange.Value

                mybook.Close False

            Next n
        End If
        ChDrive SaveDriveDir
        ChDir SaveDriveDir
        Set basebook = Nothing  '釋放物件變數
        Set mybook = Nothing
        Set sourceRange = Nothing
        Set destrange = Nothing
    
        'Application.ScreenUpdating = True
        '補上曾氏通道計算公式資料
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        新最後一列 = Sheets("MSCI").Range("A" & Rows.Count).End(xlUp).Row
        Sheets("MSCI").Range("C9:J9").AutoFill Destination:=Sheets("MSCI").Range("C9:J" & 新最後一列)
        Application.Calculation = xlCalculationAutomatic
        'Application.ScreenUpdating = True
    
        'Set range from which to determine smallest value
        Set rng = Sheets("MSCI").Range("D8:D" & 新最後一列)

        'Worksheet function MIN returns the smallest value in a range
        min = Application.WorksheetFunction.min(rng)
        max = Application.WorksheetFunction.max(rng)
    
        Set rng = Sheets("MSCI").Range("J8:J" & 新最後一列)
        min2 = Application.WorksheetFunction.min(rng)
    
        Set rng = Sheets("MSCI").Range("G8:G" & 新最後一列)
        max2 = Application.WorksheetFunction.max(rng)
    
        Set rng = Nothing
    
        If max < max2 Then
            max = max2
        End If
    
        If min2 < min Then
            min = min2
        End If
    
        Sheets("Result").Activate
    
        ActiveSheet.ChartObjects("Chart 4").Activate
        If Err Then
            ActiveSheet.ChartObjects("圖表 4").Select
            On Error GoTo 0
        End If
        With ActiveChart
            ActiveChart.HasTitle = True
            ChartTitle = Sheets("MSCI").Cells(7, 2)
            ActiveChart.ChartTitle.Text = "MSCI " & ChartTitle & " Index-曾氏通道"
        
            i = .SeriesCollection.Count
            'MsgBox i
        
            For i = i + 1 To 6
                ActiveChart.SeriesCollection.NewSeries
                If i = 1 Then
                    ActiveChart.SeriesCollection(1).Name = "=MSCI!$B$7"
                ElseIf i = 2 Then
                    ActiveChart.SeriesCollection(2).Name = "=MSCI!$E$7"
                ElseIf i = 3 Then
                    ActiveChart.SeriesCollection(3).Name = "=MSCI!$G$7"
                ElseIf i = 4 Then
                    ActiveChart.SeriesCollection(4).Name = "=MSCI!$H$7"
                ElseIf i = 5 Then
                    ActiveChart.SeriesCollection(5).Name = "=MSCI!$I$7"
                ElseIf i = 6 Then
                    ActiveChart.SeriesCollection(6).Name = "=MSCI!$J$7"
                End If
            Next
        
            FinalRow = 新最後一列
            'MsgBox FinalRow
            If i = .SeriesCollection.Count <= 6 Then
                ActiveChart.SeriesCollection(1).XValues = Sheets("MSCI").Range("A8:A" & FinalRow)
                ActiveChart.SeriesCollection(1).Values = Sheets("MSCI").Range("D8:D" & FinalRow)
                ActiveChart.SeriesCollection(2).XValues = Sheets("MSCI").Range("A8:A" & FinalRow)
                ActiveChart.SeriesCollection(2).Values = Sheets("MSCI").Range("E8:E" & FinalRow)
                ActiveChart.SeriesCollection(3).XValues = Sheets("MSCI").Range("A8:A" & FinalRow)
                ActiveChart.SeriesCollection(3).Values = Sheets("MSCI").Range("G8:G" & FinalRow)
                ActiveChart.SeriesCollection(4).XValues = Sheets("MSCI").Range("A8:A" & FinalRow)
                ActiveChart.SeriesCollection(4).Values = Sheets("MSCI").Range("H8:H" & FinalRow)
                ActiveChart.SeriesCollection(5).XValues = Sheets("MSCI").Range("A8:A" & FinalRow)
                ActiveChart.SeriesCollection(5).Values = Sheets("MSCI").Range("I8:I" & FinalRow)
                ActiveChart.SeriesCollection(6).XValues = Sheets("MSCI").Range("A8:A" & FinalRow)
                ActiveChart.SeriesCollection(6).Values = Sheets("MSCI").Range("J8:J" & FinalRow)
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
        'Date
        Sheets("MSCI").Cells(1, 6) = Sheets("MSCI").Cells(FinalRow, 1)
        '標的
        Sheets("MSCI").Cells(2, 6) = Sheets("MSCI").Cells(FinalRow, 2)
        O_Index = Sheets("MSCI").Cells(2, 6)
        '95%樂觀
        Sheets("MSCI").Cells(3, 6) = Exp(Sheets("MSCI").Cells(FinalRow, 7).Value)
        U95_trend = Sheets("MSCI").Cells(3, 6)
        '75%樂觀
        Sheets("MSCI").Cells(4, 6) = Exp(Sheets("MSCI").Cells(FinalRow, 8).Value)
        U75_trend = Sheets("MSCI").Cells(4, 6)
        '75%悲觀
        Sheets("MSCI").Cells(3, 8) = Exp(Sheets("MSCI").Cells(FinalRow, 9).Value)
        D75_trend = Sheets("MSCI").Cells(3, 8)
        '95%悲觀
        Sheets("MSCI").Cells(4, 8) = Exp(Sheets("MSCI").Cells(FinalRow, 10).Value)
        D95_trend = Sheets("MSCI").Cells(4, 8)
        'Regression Line
        R_Line = Exp(Sheets("MSCI").Cells(FinalRow, 5).Value)
        Sheets("Result").Range("I10") = Exp(Sheets("MSCI").Cells(FinalRow, 5).Value)
        Sheets("MSCI").Cells(2, 8).Interior.ColorIndex = 2
        '實際統計(年)
        Sheets("Result").Range("I11") = ((Sheets("MSCI").Cells(FinalRow, 1)) - (Sheets("MSCI").Cells(8, 1))) / 365
    
        If O_Index <= D95_trend Then
            Sheets("MSCI").Cells(2, 8).Interior.ColorIndex = 4
            Sheets("MSCI").Cells(2, 8) = 1
        End If
        If ((O_Index > D95_trend) And (O_Index <= D75_trend)) Then
            Sheets("MSCI").Cells(2, 8).Interior.ColorIndex = 4
            Sheets("MSCI").Cells(2, 8) = 2
        End If
        If ((O_Index > D75_trend) And (O_Index <= R_Line)) Then
            Sheets("MSCI").Cells(2, 8) = 3
        End If
        If ((O_Index > R_Line) And (O_Index <= U75_trend)) Then
            Sheets("MSCI").Cells(2, 8) = 4
        End If
        If ((O_Index > U75_trend) And (O_Index <= U95_trend)) Then
            Sheets("MSCI").Cells(2, 8).Interior.ColorIndex = 3
            Sheets("MSCI").Cells(2, 8) = 5
        End If
        If O_Index > U95_trend Then
            Sheets("MSCI").Cells(2, 8).Interior.ColorIndex = 3
            Sheets("MSCI").Cells(2, 8) = 6
        End If
    
EndAndExit:
        Sheets("MSCI_GPS").Activate
        '最新日期
        Sheets("MSCI_GPS").Range("B" & x + 1) = Sheets("Result").Range("I2")
        '目前位階
        Sheets("MSCI_GPS").Range("C" & x + 1) = Sheets("Result").Range("I8")
        Sheets("MSCI_GPS").Range("C" & x + 1).Interior.ColorIndex = 2
        If Sheets("MSCI_GPS").Range("C" & x + 1) >= 5 Then
            Sheets("MSCI_GPS").Range("C" & x + 1).Interior.Color = RGB(255, 102, 102)
        ElseIf Sheets("MSCI_GPS").Range("C" & x + 1) <= 2 Then
            Sheets("MSCI_GPS").Range("C" & x + 1).Interior.ColorIndex = 4
        End If
        '目前資料
        Sheets("MSCI_GPS").Range("D" & x + 1) = Sheets("Result").Range("I3")
        '斜率
        Sheets("MSCI_GPS").Range("E" & x + 1) = Sheets("Result").Range("I9")
        Sheets("MSCI_GPS").Range("E" & x + 1).Interior.ColorIndex = 2
        If Sheets("MSCI_GPS").Range("E" & x + 1) < 0 Then
            Sheets("MSCI_GPS").Range("E" & x + 1).Interior.Color = RGB(255, 102, 102)
        End If
        '95%樂觀
        Sheets("MSCI_GPS").Range("F" & x + 1) = Sheets("Result").Range("I4")
        '75%樂觀
        Sheets("MSCI_GPS").Range("G" & x + 1) = Sheets("Result").Range("I5")
        '趨勢線
        Sheets("MSCI_GPS").Range("H" & x + 1) = Sheets("Result").Range("I10")
        '75%悲觀
        Sheets("MSCI_GPS").Range("I" & x + 1) = Sheets("Result").Range("I6")
        '95悲觀
        Sheets("MSCI_GPS").Range("J" & x + 1) = Sheets("Result").Range("I7")
        '實際統計(年)
        Sheets("MSCI_GPS").Range("K" & x + 1) = Sheets("Result").Range("I11")
        If Sheets("MSCI_GPS").Range("K" & x + 1) <= 3.4 Then
            Sheets("MSCI_GPS").Range("K" & x + 1).Interior.Color = RGB(255, 102, 102)
        Else
            Sheets("MSCI_GPS").Range("K" & x + 1).Interior.ColorIndex = 2
        End If
        
        Application.StatusBar = "Done by " & Format(x / NumofGPS, "0%")
           
    Next x
    
    On Error Resume Next
    Kill MyPath & "\*.*"    ' delete all files in the folder
    RmDir MyPath            ' delete folder
    On Error GoTo 0
    TotalofTime = Now - TotalofTime
    Application.StatusBar = "Done by 100%，TotalRunTime: " & Format(TotalofTime, "hh:mm:ss")
    Sheets("MSCI").Visible = xlSheetVeryHidden
    Sheets("MSCI_Index_List").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True

End Sub

Function LastRow(sh As Worksheet)

    On Error Resume Next
    'LastRow = sh.Cells.Find(What:="*", _
    'After:=sh.Range("A1"), _
    'Lookat:=xlPart, _
    'LookIn:=xlFormulas, _
    'SearchOrder:=xlByRows, _
    'SearchDirection:=xlPrevious, _
    'MatchCase:=False).Row
    'LastRow = sh.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    On Error GoTo 0
    
End Function

Sub SetDefaultDate()

    Dim exampleDate As Date
    Dim DateofRanger As Integer
    
    
    exampleDate = Date
    Sheets("Result").Range("C4").Value = exampleDate
    If Sheets("MSCI").Range("P3") = 1 Then
        Sheets("Result").Range("C2").Value = exampleDate - (3.5 * 365)
    Else
        Sheets("Result").Range("C2").Value = exampleDate - (10 * 365)
    End If

End Sub



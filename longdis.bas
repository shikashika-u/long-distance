Attribute VB_Name = "Module1"
Sub Longdistance_GCaMP()
Attribute Longdistance_GCaMP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
    Dim i As Integer
    Dim j As Integer
    ROI = InputBox("ROI数を入力")
    interval = InputBox("撮影間隔(秒)を入力")
    If IsNumeric(ROI) = False Or IsNumeric(interval) = False Then
        MsgBox ("入力失敗")
        Exit Sub
    End If
    Application.Calculation = xlManual '自動計算をオフ
    Application.ScreenUpdating = False '画面表示の更新を停止
    sheetName = ActiveSheet.Name
    ActiveSheet.Copy Before:=Worksheets(1)
    ActiveSheet.Name = sheetName + "_calulated"
    With ActiveSheet.UsedRange
        MaxRow = .Rows.Count + 1
        MaxCol = .Columns.Count
    End With
    
    Path = ThisWorkbook.Path 'ワークブックがいるフォルダのパス
    Num_of_file = MaxCol \ (ROI + 2)
    Dim Filelist() As String
    ReDim Filelist(1 To Num_of_file) As String
    Filelist(1) = Dir(Path & "\ROIimage\*.jpg")
    For i = 2 To Num_of_file Step 1
        Filelist(i) = Dir()
    Next i
    file = 1
    Range(Cells(1, 1), Cells(MaxRow - 1, MaxCol)).Cut Destination:=Range("A2")
    
' ------------------------------------------------------------------------------------------------------------------------------
    For i = 2 To MaxCol - 1 Step ROI + 2
        Dim c As Range
        For Each c In Range(Cells(3, i + 1), Cells(MaxRow, i + ROI + 1))
            If c.Value = 0 Then c.Value = ""
        Next c
        Columns(i + ROI + 1).ColumnWidth = 50
        Dim Title As String
        Title = Replace(Cells(2, i).Text, ":time", "")
        Cells(1, i) = Title
        Cells(1, i).Font.Bold = True
        For j = 0 To ROI + 1 Step 1
            Cells(2, i + j) = Replace(Cells(2, i + j), Title + ":", "")
        Next j
        'グラフ作成
        Cells(ROI + 2, i + ROI + 1) = "Velocity"
        Cells(2 * ROI + 2, i + ROI + 1) = "Amplitude"
        Cells(3 * ROI + 3, i + ROI + 1) = "Duration"
        Cells(4 * ROI + 4, i + ROI + 1) = "average"
        Cells(5 * ROI + 5, i + ROI + 1) = "3xSD"
        Cells(6 * ROI + 6, i + ROI + 1) = "Significantly increased"
        Cells(7 * ROI + 7, i + ROI + 1) = "Half of amplitude"
        For j = 1 To ROI Step 1
        '計算(Amplitude, Average, 3xSD, Significantly increased, Half of amplitude------------------------------
            Cells(4 * ROI + 4 + j, i + ROI + 1).FormulaR1C1 = _
                "=AVERAGE(R3C" & i + j & ": R22C" & i + j & ")" 'Average
            Cells(2 * ROI + 2 + j, i + ROI + 1).FormulaR1C1 = _
                "=-R[" & 2 * ROI + 2 & "]C[0]+MAX(R3C" & i + j & ":R" & MaxRow & "C" & i + j & " )" 'Amplitude
            Cells(5 * ROI + 5 + j, i + ROI + 1).FormulaR1C1 = _
                "=R[" & -ROI - 1 & "]C[0]+3*STDEVP(R3C" & i + j & ": R22C" & i + j & ")" '3xSD
            Cells(7 * ROI + 7 + j, i + ROI + 1) = Cells(2 * ROI + 2 + j, i + ROI + 1).Value / 2 'Half of Amplitude
            Dim counter As Integer
            '有意に上昇するタイミングをfor構文使ってカウント
            For counter = 0 To MaxRow - 2 Step 1
                If Cells(counter + 3, i + j).Value >= Cells(5 * ROI + 5 + j, i + ROI + 1).Value Then
                    Exit For
                End If
            Next counter
            If counter = MaxRow - 1 Then
                Cells(6 * ROI + 6 + j, i + ROI + 1) = "False"
            Else
                Cells(6 * ROI + 6 + j, i + ROI + 1) = counter * interval
            End If
            Dim rRange As Range
            Set rRange = Range(Cells(3, i + j), Cells(MaxRow, i + j))
            Cells(3 * ROI + 3 + j, i + ROI + 1) = _
                interval * WorksheetFunction.CountIf(rRange, ">=" & Cells(7 * ROI + 7 + j, i + ROI + 1).Value & "") 'Duration
            '条件付き書式の追加
            Dim halfamp As FormatCondition
            Set halfamp = rRange.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=R[0]C[0]>=R" & 7 * ROI + 7 + j & "C" & i + ROI + 1 & "")
            '太字・背景色をグレーに変更
            halfamp.Font.Bold = True
            halfamp.Font.ColorIndex = 3
            halfamp.Interior.Color = RGB(204, 204, 204)
            Dim over_SD As FormatCondition
            Set over_SD = rRange.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=R[0]C[0]>=R" & 5 * ROI + 5 + j & "C" & i + ROI + 1 & "")
            over_SD.Interior.Color = RGB(204, 204, 204)
            Next j
        For j = 1 To ROI - 1 Step 1
                Cells(ROI + 2 + j, i + ROI + 1).FormulaR1C1 = _
                    "=R[" & -ROI & "]C[0]/(R[" & 5 * ROI + 5 & "]C[0]-R[" & 5 * ROI + 4 & "]C[0])"
        Next j
        
        With ActiveSheet.Pictures.Insert(Path & "\ROIimage\" & Filelist(file))
        .Top = Cells(4 * ROI + 4, i + ROI + 1).Top
        .Left = Cells(4 * ROI + 4, i + ROI + 1).Left
        .Width = Cells(4 * ROI + 4, i + ROI + 1).Width
        End With
        file = file + 1
    Next i
    Application.Calculation = xlAutomatic  '画面の表示更新をオン
    Application.ScreenUpdating = True  '自動計算をオン
    
End Sub

Sub graph()
    
    With ActiveSheet.Shapes.AddChart.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .SetSourceData Range(Cells(2, 2), Cells(274, 5))
        .HasTitle = True
        .ChartTitle.Text = Cells(1, 2).Text
        .Legend.Position = xlLegendPositionBottom
        .ChartArea.Format.Line.Visible = msoFalse
        With .Axes(xlValue)
            With .TickLabels
                .Font.Size = 14
                .Font.Bold = True
                .Font.Name = "Times New Roman"
            End With
            .Format.Line.ForeColor.SchemeColor = 8
            .Format.Line.Weight = 1.5
            .HasMajorGridlines = False
            .HasTitle = False
            .CrossesAt = -1
        End With
        With .Axes(xlCategory)
            With .TickLabels
                .Font.Size = 15
                .Font.Bold = True
                .Font.Name = "Times New Roman"
            End With
            .Format.Line.ForeColor.SchemeColor = 8
            .Format.Line.Weight = 1.5
            .HasMajorGridlines = False
            .HasTitle = True
            .AxisTitle.Text = "Time(s)"
            .AxisTitle.Font.Size = 14
            .AxisTitle.Font.Name = "Times New Roman"
        End With
        For i = 1 To 3 Step 1
                With .FullSeriesCollection(i)
                    .Format.Line.ForeColor.SchemeColor = 8 + 2 * (i - 1)
                    .Format.Line.Weight = 2
                End With
        Next i
    End With
    With ActiveSheet.ChartObjects(1)
        .Top = Cells(44, 6).Top
        .Left = Cells(44, 6).Left
        .Width = Cells(44, 6).Width
        .Height = Range(Cells(44, 6), Cells(59, 6)).Height
    End With
End Sub




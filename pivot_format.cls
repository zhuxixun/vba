Option Explicit

Sub FormatCurrentPivot()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim rngData As Range
    Dim rngLabel As Range
    
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo SafeExit
    
    '1. 判断当前工作表是否包含透视表
    If ws.PivotTables.Count = 0 Then
        MsgBox "当前工作表不包含透视表，程序结束。", vbInformation, "提示"
        Exit Sub
    End If
    
    '3. 遍历工作表里的所有透视表（可能不止一个）
    For Each pt In ws.PivotTables

        '1. 数据值区域
        Set rngData = pt.DataBodyRange        '所有数值，整个“数值区域”——就是你拖进“值”里的那些数字所在的矩形方块
        If Not rngData Is Nothing Then
            With rngData
                '.Font.Name = "Calibri"
                '.Font.Size = 11
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
        
        '2. 数据列标题区域（Σ 值行标题）
        Set rngLabel = pt.DataLabelRange      '只有“Σ 值”那一行的标题（“求和项:金额”、“计数项:数量”等）
        If Not rngLabel Is Nothing Then
            With rngLabel
                '.Font.Name = "Calibri"
                '.Font.Size = 11
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
        End If
        
        '--- 分别设置数据字段的数字格式 ---
        For Each pf In pt.DataFields
            If InStr(1, pf.Caption, "%", vbTextCompare) > 0 Then
                '含 % 的字段
                pf.NumberFormat = "0.0%"
            Else
                '不含 % 的字段
                pf.NumberFormat = "#,##0,,;[Red](#,##0,,)"
            End If
        Next pf
        
    Next pt

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    If Err.Number <> 0 Then
        MsgBox "运行过程出现错误：" & Err.Description, vbCritical, "提示"
    Else
        MsgBox "已完成格式化！", vbInformation, "提示"
    End If

End Sub
' ******************************************************************************
' 作者: Tobin
' 版本: 260605
' 功能：删除重复行并验证关键列存在性
' 操作说明：
' 1. 根据全部5列的组合判断重复数据（预算合同编号、合同预算项目、累计已用预算、调整后预算金额.、最终预算）
' 2. 自动查找表头所在行
' 3. 验证以下列是否存在：
'    - 预算合同编号
'    - 合同预算项目
'    - 累计已用预算
'    - 调整后预算金额.
'    - 最终预算（累计已用与调整后预算取大值）
' ******************************************************************************

Sub DeleteDuplicateRows()
    Const VERSION As String = "260605"
    Dim ws As Worksheet
    Dim lastRow As Long, headerRow As Long
    Dim i As Long
    Dim col1 As Integer, col2 As Integer, col3 As Integer, col4 As Integer, col5 As Integer
    Dim dict As Object
    Dim key As String
    Dim deleteCount As Long
    Dim missingColumns As Collection
    Dim startTime As Double, endTime As Double, elapsedTime As Double
    
    Set missingColumns = New Collection
    
    Set ws = ActiveSheet
    
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then
        MsgBox "未找到表头行！" & vbCrLf & vbCrLf & "程序已在第 1-100 行范围内搜索，但未找到包含""预算合同编号""或""合同预算项目""的表头行。" & vbCrLf & vbCrLf & "请检查工作表中是否包含这两个关键列名！", vbCritical, VERSION & "版本自动化程序提醒您"
        Exit Sub
    End If
    
    col1 = GetColumnNumber(ws, "预算合同编号", headerRow)
    If col1 = 0 Then missingColumns.Add "预算合同编号"
    
    col2 = GetColumnNumber(ws, "合同预算项目", headerRow)
    If col2 = 0 Then missingColumns.Add "合同预算项目"
    
    col3 = GetColumnNumber(ws, "累计已用预算", headerRow)
    If col3 = 0 Then missingColumns.Add "累计已用预算"
    
    col4 = GetColumnNumber(ws, "调整后预算金额.", headerRow)
    If col4 = 0 Then missingColumns.Add "调整后预算金额."
    
    col5 = GetColumnNumber(ws, "最终预算（累计已用与调整后预算取大值）", headerRow)
    If col5 = 0 Then missingColumns.Add "最终预算（累计已用与调整后预算取大值）"
    
    If missingColumns.Count > 0 Then
        Dim errMsg As String
        errMsg = "列验证失败！" & vbCrLf & vbCrLf & "在表头行（第 " & headerRow & " 行）未找到以下必要列：" & vbCrLf & vbCrLf
        Dim colName As Variant
        For Each colName In missingColumns
            errMsg = errMsg & "  ● " & colName & vbCrLf
        Next colName
        errMsg = errMsg & vbCrLf & "请确保所有必要列与""预算合同编号""和""合同预算项目""在同一行。" & vbCrLf & vbCrLf & "必要列清单：" & vbCrLf & "  ● 预算合同编号" & vbCrLf & "  ● 合同预算项目" & vbCrLf & "  ● 累计已用预算" & vbCrLf & "  ● 调整后预算金额." & vbCrLf & "  ● 最终预算（累计已用与调整后预算取大值）"
        MsgBox errMsg, vbCritical, VERSION & "版本自动化程序提醒您"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    deleteCount = 0
    
    startTime = Timer
    
    For i = lastRow To headerRow + 1 Step -1
        key = ws.Cells(i, col1).Value & "|" & ws.Cells(i, col2).Value & "|" & ws.Cells(i, col3).Value & "|" & ws.Cells(i, col4).Value & "|" & ws.Cells(i, col5).Value
        
        If dict.Exists(key) Then
            ws.Rows(i).Delete
            deleteCount = deleteCount + 1
        Else
            dict.Add key, 1
        End If
    Next i
    
    endTime = Timer
    elapsedTime = (endTime - startTime) * 1000
    
    If deleteCount > 0 Then
        MsgBox "操作完成！" & vbCrLf & vbCrLf & "已成功删除 " & deleteCount & " 行重复数据。" & vbCrLf & vbCrLf & "程序耗时：" & Format(elapsedTime, "0.00") & " 毫秒" & vbCrLf & vbCrLf & "数据去重任务已完成！", vbInformation, VERSION & "版本自动化程序提醒您"
    Else
        MsgBox "操作完成！" & vbCrLf & vbCrLf & "未找到重复行，所有数据唯一。" & vbCrLf & vbCrLf & "数据验证任务已完成！", vbInformation, VERSION & "版本自动化程序提醒您"
    End If
    
    Set dict = Nothing
    Set missingColumns = Nothing
End Sub

Function FindHeaderRow(ws As Worksheet) As Long
    Dim row As Long
    Dim col As Integer
    
    For row = 1 To 100
        col = 1
        Do While ws.Cells(row, col).Value <> ""
            If ws.Cells(row, col).Value = "预算合同编号" Or ws.Cells(row, col).Value = "合同预算项目" Then
                FindHeaderRow = row
                Exit Function
            End If
            col = col + 1
        Loop
    Next row
    
    FindHeaderRow = 0
End Function

' ******************************************************************************
' 辅助函数：根据列标题获取列号
' 参数：
'   ws - 需要查找的工作表
'   columnName - 要查找的列标题名称
'   headerRow - 表头所在行号
' 返回值：
'   Integer - 列号（从1开始），未找到返回0
' ******************************************************************************
Function GetColumnNumber(ws As Worksheet, columnName As String, headerRow As Long) As Integer
    Dim col As Integer
    Dim lastCol As Integer
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If Trim(ws.Cells(headerRow, col).Value) = Trim(columnName) Then
            GetColumnNumber = col
            Exit Function
        End If
    Next col
    
    GetColumnNumber = 0
End Function

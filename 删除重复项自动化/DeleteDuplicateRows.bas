' ******************************************************************************
' 作者: Tobin
' 版本: 250604
' 功能：删除重复行并验证关键列存在性
' 操作说明：
' 1. 根据"预算合同编号"和"合同预算项目"两列判断重复数据
' 2. 验证以下列是否存在：
'    - 预算合同编号
'    - 合同预算项目
'    - 累计已用预算
'    - 调整后预算金额.
'    - 最终预算（累计已用与调整后预算取大值）
' ******************************************************************************

Sub DeleteDuplicateRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim col1 As Integer, col2 As Integer, col3 As Integer, col4 As Integer, col5 As Integer
    Dim dict As Object
    Dim key As String
    Dim deleteCount As Long
    Dim missingColumns As Collection
    
    ' 初始化缺失列集合
    Set missingColumns = New Collection
    
    ' 设置工作表（当前活动工作表）
    Set ws = ActiveSheet
    
    ' 查找各列列号并验证存在性
    col1 = GetColumnNumber(ws, "预算合同编号")
    If col1 = 0 Then missingColumns.Add "预算合同编号"
    
    col2 = GetColumnNumber(ws, "合同预算项目")
    If col2 = 0 Then missingColumns.Add "合同预算项目"
    
    col3 = GetColumnNumber(ws, "累计已用预算")
    If col3 = 0 Then missingColumns.Add "累计已用预算"
    
    col4 = GetColumnNumber(ws, "调整后预算金额.")
    If col4 = 0 Then missingColumns.Add "调整后预算金额."
    
    col5 = GetColumnNumber(ws, "最终预算（累计已用与调整后预算取大值）")
    If col5 = 0 Then missingColumns.Add "最终预算（累计已用与调整后预算取大值）"
    
    ' 处理缺失列情况
    If missingColumns.Count > 0 Then
        Dim errMsg As String
        errMsg = "未找到以下必要列：" & vbCrLf & vbCrLf
        Dim colName As Variant
        For Each colName In missingColumns
            errMsg = errMsg & "- " & colName & vbCrLf
        Next colName
        errMsg = errMsg & vbCrLf & "请检查工作表列名是否完整正确！"
        MsgBox errMsg, vbCritical, "列缺失错误"
        Exit Sub
    End If
    
    ' 获取数据的最后一行
    lastRow = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
    
    ' 创建字典用于存储已检查的行内容（复合主键）
    Set dict = CreateObject("Scripting.Dictionary")
    
    deleteCount = 0
    
    ' 从最后一行开始向上遍历（避免删除行后影响行号）
    For i = lastRow To 2 Step -1
        ' 使用"|"作为分隔符构建复合键
        key = ws.Cells(i, col1).Value & "|" & ws.Cells(i, col2).Value
        
        ' 如果字典中已经存在该键，说明是重复行，标记删除
        If dict.Exists(key) Then
            ws.Rows(i).Delete
            deleteCount = deleteCount + 1
        Else
            ' 如果不存在，将键添加到字典中
            dict.Add key, 1
        End If
    Next i
    
    ' 提示结果
    If deleteCount > 0 Then
        MsgBox "已删除 " & deleteCount & " 行重复数据！", vbInformation, "操作完成"
    Else
        MsgBox "未找到重复行，所有数据唯一！", vbInformation, "操作完成"
    End If
    
    ' 清理对象
    Set dict = Nothing
    Set missingColumns = Nothing
End Sub

' ******************************************************************************
' 辅助函数：根据列标题获取列号
' 参数：
'   ws - 需要查找的工作表
'   columnName - 要查找的列标题名称
' 返回值：
'   Integer - 列号（从1开始），未找到返回0
' ******************************************************************************
Function GetColumnNumber(ws As Worksheet, columnName As String) As Integer
    Dim col As Integer
    col = 1
    ' 遍历第一行所有列直到找到目标列或遇到空列
    Do While ws.Cells(1, col).Value <> ""
        If ws.Cells(1, col).Value = columnName Then
            GetColumnNumber = col
            Exit Function
        End If
        col = col + 1
    Loop
    GetColumnNumber = 0
End Function

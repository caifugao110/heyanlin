' ******************************************************************************
' 作者: Tobin
' 版本: 260616
' 功能：提取合同号和税率信息并生成税率一览表
' 操作说明：
' 1. 从[销售台账]工作表中提取合同号和税率列数据
' 2. 自动查找表头所在行
' 3. 验证以下列是否存在：
'    - 合同号
'    - 税率
' 4. 将提取的数据去重后保存到[税率一览表]工作表
' ******************************************************************************

Sub ExtractTaxRate()
    Const VERSION As String = "260616"
    Dim startTime As Double
    startTime = Timer
    
    ' ★性能优化：关闭屏幕更新、自动计算和事件
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long, headerRow As Long
    Dim i As Long, targetRow As Long
    Dim colContractNo As Integer, colTaxRate As Integer
    Dim dict As Object
    Dim key As String
    Dim missingColumns As Collection
    Dim sheetExists As Boolean
    
    Set missingColumns = New Collection
    Set dict = CreateObject("Scripting.Dictionary")
    
    Set wsSource = FindWorksheet("销售台账")
    
    If wsSource Is Nothing Then
        MsgBox "未找到[销售台账]工作表！", vbCritical, VERSION & "版本自动化程序提醒您"
        GoTo CleanUp
    End If
    
    headerRow = FindHeaderRow(wsSource)
    If headerRow = 0 Then
        MsgBox "未找到表头行！" & vbCrLf & vbCrLf & "程序已在第 1-100 行范围内搜索，但未找到包含""合同号""或""税率""的表头行。" & vbCrLf & vbCrLf & "请检查工作表中是否包含这两个关键列名！", vbCritical, VERSION & "版本自动化程序提醒您"
        GoTo CleanUp
    End If
    
    colContractNo = GetColumnNumber(wsSource, "合同号", headerRow)
    If colContractNo = 0 Then missingColumns.Add "合同号"
    
    colTaxRate = GetColumnNumber(wsSource, "税率", headerRow)
    If colTaxRate = 0 Then missingColumns.Add "税率"
    
    If missingColumns.Count > 0 Then
        Dim errMsg As String
        errMsg = "列验证失败！" & vbCrLf & vbCrLf & "在表头行（第 " & headerRow & " 行）未找到以下必要列：" & vbCrLf & vbCrLf
        Dim colName As Variant
        For Each colName In missingColumns
            errMsg = errMsg & "  ● " & colName & vbCrLf
        Next colName
        errMsg = errMsg & vbCrLf & "请确保所有必要列与""合同号""和""税率""在同一行。" & vbCrLf & vbCrLf & "必要列清单：" & vbCrLf & "  ● 合同号" & vbCrLf & "  ● 税率"
        MsgBox errMsg, vbCritical, VERSION & "版本自动化程序提醒您"
        GoTo CleanUp
    End If
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, colContractNo).End(xlUp).Row
    Dim dataRowCount As Long
    dataRowCount = lastRow - headerRow
    If dataRowCount < 1 Then
        MsgBox "数据区域为空，无数据可提取！", vbExclamation, VERSION & "版本自动化程序提醒您"
        GoTo CleanUp
    End If
    
    ' ★性能优化：一次性读取整列数据到数组（避免逐格读取）
    Dim arrContract() As Variant
    Dim arrTaxRate() As Variant
    arrContract = wsSource.Range(wsSource.Cells(headerRow + 1, colContractNo), wsSource.Cells(lastRow, colContractNo)).Value
    arrTaxRate = wsSource.Range(wsSource.Cells(headerRow + 1, colTaxRate), wsSource.Cells(lastRow, colTaxRate)).Value
    
    Dim cellValue As Variant
    Dim taxRateValue As Double
    For i = 1 To dataRowCount
        If Trim(CStr(arrContract(i, 1) & "")) <> "" Then
            cellValue = arrTaxRate(i, 1)
            
            If Not IsNumeric(cellValue) Then
                GoTo NextRow
            End If
            
            taxRateValue = Round(CDbl(cellValue), 5)
            key = Trim(CStr(arrContract(i, 1) & "")) & "|" & CStr(taxRateValue)
            If Not dict.Exists(key) Then
                dict.Add key, Array(arrContract(i, 1), taxRateValue)
            End If
        End If
NextRow:
    Next i
    
    Dim wbSource As Workbook
    Set wbSource = wsSource.Parent
    
    sheetExists = False
    On Error Resume Next
    Set wsTarget = wbSource.Sheets("税率一览表")
    sheetExists = (Err.Number = 0)
    On Error GoTo 0
    
    If sheetExists Then
        wsTarget.Cells.Clear
    Else
        Set wsTarget = wbSource.Sheets.Add(After:=wsSource)
        wsTarget.Name = "税率一览表"
    End If
    
    wsTarget.Cells(1, 1).Value = "合同号"
    wsTarget.Cells(1, 2).Value = "税率"
    
    ' ★性能优化：用数组批量写入输出（避免逐格写入）
    Dim outputArr() As Variant
    ReDim outputArr(1 To dict.Count, 1 To 2)
    
    Dim idx As Long
    idx = 1
    Dim item As Variant
    For Each item In dict.Items
        outputArr(idx, 1) = item(0)
        outputArr(idx, 2) = item(1)
        idx = idx + 1
    Next item
    
    If dict.Count > 0 Then
        wsTarget.Range("A2").Resize(dict.Count, 2).Value = outputArr
    End If
    targetRow = dict.Count + 2
    
    ' 格式化
    With wsTarget
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 12
        .Columns("A").HorizontalAlignment = xlCenter
        .Columns("B").HorizontalAlignment = xlRight
        .Rows(1).Font.Bold = True
        
        If dict.Count > 0 Then
            Dim dataRange As Range
            Set dataRange = .Range(.Cells(1, 1), .Cells(targetRow - 1, 2))
            dataRange.Borders.LineStyle = xlContinuous
        End If
        
        .Range("A1:B1").AutoFilter
    End With
    
    Dim elapsed As Long
    elapsed = CLng((Timer - startTime) * 1000)
    
    wsTarget.Activate
    
    MsgBox "操作完成！" & vbCrLf & vbCrLf & "已成功提取 " & dict.Count & " 条合同号-税率记录。" & vbCrLf & vbCrLf & "数据已保存到[税率一览表]工作表中。" & vbCrLf & vbCrLf & "耗时：" & elapsed & " ms", vbInformation, VERSION & "版本自动化程序提醒您"
    
CleanUp:
    ' ★恢复Excel设置
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Set dict = Nothing
    Set missingColumns = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
End Sub

Function FindHeaderRow(ws As Worksheet) As Long
    Dim row As Long
    Dim col As Integer
    
    For row = 1 To 100
        col = 1
        Do While ws.Cells(row, col).Value <> ""
            If ws.Cells(row, col).Value = "合同号" Or ws.Cells(row, col).Value = "税率" Then
                FindHeaderRow = row
                Exit Function
            End If
            col = col + 1
        Loop
    Next row
    
    FindHeaderRow = 0
End Function

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

Function FindWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim targetName As String
    
    targetName = Trim(sheetName)
    
    For Each ws In ActiveWorkbook.Sheets
        If Trim(ws.Name) = targetName Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next ws
    
    Set FindWorksheet = Nothing
End Function
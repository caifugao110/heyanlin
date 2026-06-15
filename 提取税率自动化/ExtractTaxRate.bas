' ******************************************************************************
' 作者: Tobin
' 版本: 260615
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
    Const VERSION As String = "260615"
    Dim startTime As Double
    startTime = Timer
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
        Exit Sub
    End If
    
    headerRow = FindHeaderRow(wsSource)
    If headerRow = 0 Then
        MsgBox "未找到表头行！" & vbCrLf & vbCrLf & "程序已在第 1-100 行范围内搜索，但未找到包含""合同号""或""税率""的表头行。" & vbCrLf & vbCrLf & "请检查工作表中是否包含这两个关键列名！", vbCritical, VERSION & "版本自动化程序提醒您"
        Exit Sub
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
        Exit Sub
    End If
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, colContractNo).End(xlUp).Row
    
    For i = headerRow + 1 To lastRow
        If Trim(wsSource.Cells(i, colContractNo).Value) <> "" Then
            Dim cellValue As Variant
            Dim taxRateValue As Double
            cellValue = wsSource.Cells(i, colTaxRate).Value
            
            If Not IsNumeric(cellValue) Then
                GoTo NextRow
            End If
            
            taxRateValue = Round(CDbl(cellValue), 5)
            key = Trim(CStr(wsSource.Cells(i, colContractNo).Value)) & "|" & CStr(taxRateValue)
            If Not dict.Exists(key) Then
                dict.Add key, Array(wsSource.Cells(i, colContractNo).Value, taxRateValue)
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
    
    targetRow = 2
    Dim item As Variant
    For Each item In dict.Items
        wsTarget.Cells(targetRow, 1).Value = item(0)
        wsTarget.Cells(targetRow, 2).Value = item(1)
        targetRow = targetRow + 1
    Next item
    
    ' 格式化
    With wsTarget
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 12
        .Columns("A").HorizontalAlignment = xlCenter
        .Columns("B").HorizontalAlignment = xlRight
        .Rows(1).Font.Bold = True
        
        Dim dataRange As Range
        Set dataRange = .Range(.Cells(1, 1), .Cells(targetRow - 1, 2))
        dataRange.Borders.LineStyle = xlContinuous
        
        .Range("A1:B1").AutoFilter
    End With
    
    Dim elapsed As Long
    elapsed = CLng((Timer - startTime) * 1000)
    
    wsTarget.Activate
    
    MsgBox "操作完成！" & vbCrLf & vbCrLf & "已成功提取 " & dict.Count & " 条合同号-税率记录。" & vbCrLf & vbCrLf & "数据已保存到[税率一览表]工作表中。" & vbCrLf & vbCrLf & "耗时：" & elapsed & " ms", vbInformation, VERSION & "版本自动化程序提醒您"
    
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
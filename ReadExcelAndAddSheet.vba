Sub ReadExcelAndAddSheet()
    Dim filePath As String
    filePath = "/Users/luo/Documents/testFiles/testExcel.xlsx"

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath)

    ' 获取工作簿中的所有Sheet数量
    Dim sheetCount As Integer
    sheetCount = wb.Sheets.Count

    ' 遍历所有Sheet，删除除了Sheet1和Sheet2之外的所有sheet
    Dim i As Integer
    For i = sheetCount To 3 Step -1
        If i <> 1 And i <> 2 Then
            wb.Sheets(i).Delete
        End If
    Next i

    ' 获取第一个Sheet
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    ' 从第33行开始读取每一行的前5个字段
    Dim startRow As Integer
    startRow = 33
    ' 获取实际显示的最后一行的行数
    Dim endRow As Integer
    endRow = WorksheetFunction.Min(ws.Cells(Rows.Count, 1).End(xlUp).Row, 66)

    ' 从第2个Sheet开始创建新的Sheet
    Dim j As Integer
    For j = startRow To endRow
        Dim row As Range
        Set row = ws.Rows(j)

        ' 在获取第一个单元格之前，再次检查行是否为空
        Dim firstCell As Range
        Set firstCell = row.Cells(1, 1)
        If firstCell Is Nothing Or firstCell.Value = "" Then
            Exit For ' 如果该行的第一个格子没有值，中断循环
        End If

        Dim functionNameLogical As String
        functionNameLogical = GetCellValueAsString(row.Cells(1, 4))

        ' 复制第二个Sheet作为模板
        Dim templateSheet As Worksheet
        Set templateSheet = wb.Sheets(2)
        Dim newSheet As Worksheet
        Set newSheet = templateSheet.Copy(, wb.Sheets(wb.Sheets.Count))
        newSheet.Name = "functionName_" & functionNameLogical & j

        ' 复制模板Sheet的内容到新Sheet
        templateSheet.UsedRange.Copy newSheet.Range("A1")

        ' 在新Sheet的第6行创建单元格，并在6C写入字段值
        Dim row6 As Range
        Set row6 = newSheet.Rows(6)
        row6.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 3))
        ' 在新Sheet的第7行创建单元格，并在7C，7F，7M写入字段值
        Dim row7 As Range
        Set row7 = newSheet.Rows(7)
        row7.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 2))
        row7.Cells(1, 6).Value = GetCellValueAsString(row.Cells(1, 5))
        row7.Cells(1, 13).Value = GetCellValueAsString(row.Cells(1, 4))

        ' 输出字段值
        Debug.Print "FunctionId: " & GetCellValueAsString(row.Cells(1, 2))
        Debug.Print "Modifier: " & GetCellValueAsString(row.Cells(1, 3))
        Debug.Print "FunctionName (Logical): " & GetCellValueAsString(row.Cells(1, 5))
        Debug.Print "FunctionName (Physical): " & GetCellValueAsString(row.Cells(1, 4))
    Next j

    ' 删除原始的Sheet2
    ' wb.Sheets(2).Delete

    ' 保存修改后的Excel文件
    wb.Save
    wb.Close

    Set wb = Nothing
End Sub

Function GetCellValueAsString(cell As Range) As String
    If cell Is Nothing Then
        GetCellValueAsString = ""
    Else
        GetCellValueAsString = CStr(cell.Value)
    End If
End Function

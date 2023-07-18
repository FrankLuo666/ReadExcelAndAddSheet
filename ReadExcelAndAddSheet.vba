Sub ReadExcelAndAddSheet()
    Dim filePath As String
    filePath = "/Users/luo/Documents/testFiles/testExcel.xlsx"

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath)

    ' ��ȡ�������е�����Sheet����
    Dim sheetCount As Integer
    sheetCount = wb.Sheets.Count

    ' ��������Sheet��ɾ������Sheet1��Sheet2֮�������sheet
    Dim i As Integer
    For i = sheetCount To 3 Step -1
        If i <> 1 And i <> 2 Then
            wb.Sheets(i).Delete
        End If
    Next i

    ' ��ȡ��һ��Sheet
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    ' �ӵ�33�п�ʼ��ȡÿһ�е�ǰ5���ֶ�
    Dim startRow As Integer
    startRow = 33
    ' ��ȡʵ����ʾ�����һ�е�����
    Dim endRow As Integer
    endRow = WorksheetFunction.Min(ws.Cells(Rows.Count, 1).End(xlUp).Row, 66)

    ' �ӵ�2��Sheet��ʼ�����µ�Sheet
    Dim j As Integer
    For j = startRow To endRow
        Dim row As Range
        Set row = ws.Rows(j)

        ' �ڻ�ȡ��һ����Ԫ��֮ǰ���ٴμ�����Ƿ�Ϊ��
        Dim firstCell As Range
        Set firstCell = row.Cells(1, 1)
        If firstCell Is Nothing Or firstCell.Value = "" Then
            Exit For ' ������еĵ�һ������û��ֵ���ж�ѭ��
        End If

        Dim functionNameLogical As String
        functionNameLogical = GetCellValueAsString(row.Cells(1, 4))

        ' ���Ƶڶ���Sheet��Ϊģ��
        Dim templateSheet As Worksheet
        Set templateSheet = wb.Sheets(2)
        Dim newSheet As Worksheet
        Set newSheet = templateSheet.Copy(, wb.Sheets(wb.Sheets.Count))
        newSheet.Name = "functionName_" & functionNameLogical & j

        ' ����ģ��Sheet�����ݵ���Sheet
        templateSheet.UsedRange.Copy newSheet.Range("A1")

        ' ����Sheet�ĵ�6�д�����Ԫ�񣬲���6Cд���ֶ�ֵ
        Dim row6 As Range
        Set row6 = newSheet.Rows(6)
        row6.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 3))
        ' ����Sheet�ĵ�7�д�����Ԫ�񣬲���7C��7F��7Mд���ֶ�ֵ
        Dim row7 As Range
        Set row7 = newSheet.Rows(7)
        row7.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 2))
        row7.Cells(1, 6).Value = GetCellValueAsString(row.Cells(1, 5))
        row7.Cells(1, 13).Value = GetCellValueAsString(row.Cells(1, 4))

        ' ����ֶ�ֵ
        Debug.Print "FunctionId: " & GetCellValueAsString(row.Cells(1, 2))
        Debug.Print "Modifier: " & GetCellValueAsString(row.Cells(1, 3))
        Debug.Print "FunctionName (Logical): " & GetCellValueAsString(row.Cells(1, 5))
        Debug.Print "FunctionName (Physical): " & GetCellValueAsString(row.Cells(1, 4))
    Next j

    ' ɾ��ԭʼ��Sheet2
    ' wb.Sheets(2).Delete

    ' �����޸ĺ��Excel�ļ�
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

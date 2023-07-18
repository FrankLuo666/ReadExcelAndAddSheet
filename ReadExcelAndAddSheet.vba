Sub ReadExcelAndAddSheet()
    Dim filePath As String
    filePath = "/Users/luo/Documents/testFiles/testExcel.xlsx"

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath)

    ' ��`���֥å��ڤΤ��٤ƤΥ��`������ȡ��
    Dim sheetCount As Integer
    sheetCount = wb.Sheets.Count

    ' ���٤ƤΥ��`�Ȥ��ߖˤ���Sheet1��Sheet2����Τ��٤ƤΥ��`�Ȥ�����
    Dim i As Integer
    For i = sheetCount To 3 Step -1
        If i <> 1 And i <> 2 Then
            wb.Sheets(i).Delete
        End If
    Next i

    ' ����Υ��`�Ȥ�ȡ��
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    ' 33��Ŀ������Ф������5�ĤΥե��`��ɤ��i��ȡ��
    Dim startRow As Integer
    startRow = 33
    ' �g�H�˱�ʾ�����������з��Ť�ȡ��
    Dim endRow As Integer
    endRow = WorksheetFunction.Min(ws.Cells(Rows.Count, 1).End(xlUp).Row, 66)

    ' 2��Ŀ�Υ��`�Ȥ����¤������`�Ȥ�����
    Dim j As Integer
    For j = startRow To endRow
        Dim row As Range
        Set row = ws.Rows(j)

        ' ����Υ����ȡ�ä���ǰ�ˡ��Ф��դ��ɤ������ٶȴ_�J
        Dim firstCell As Range
        Set firstCell = row.Cells(1, 1)
        If firstCell Is Nothing Or firstCell.Value = "" Then
            Exit For ' �����Ф�����Υ��뤬����֤äƤ��ʤ����ϡ���`�פ��ж�
        End If

        Dim functionNameLogical As String
        functionNameLogical = GetCellValueAsString(row.Cells(1, 4))

        ' 2��Ŀ�Υ��`�Ȥ�ƥ�ץ�`�ȤȤ��ƥ��ԩ`
        Dim templateSheet As Worksheet
        Set templateSheet = wb.Sheets(2)
        Dim newSheet As Worksheet
        Set newSheet = templateSheet.Copy(, wb.Sheets(wb.Sheets.Count))
        newSheet.Name = "functionName_" & functionNameLogical & j

        ' �ƥ�ץ�`�ȤΥ��`�����ݤ��¤������`�Ȥ˥��ԩ`
        templateSheet.UsedRange.Copy newSheet.Range("A1")

        ' �¤������`�Ȥ�6��Ŀ�˥�������ɤ���6C�˥ե��`��ɂ���ӛ��
        Dim row6 As Range
        Set row6 = newSheet.Rows(6)
        row6.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 3))
        ' �¤������`�Ȥ�7��Ŀ�˥�������ɤ���7C��7F��7M�˥ե��`��ɂ���ӛ��
        Dim row7 As Range
        Set row7 = newSheet.Rows(7)
        row7.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 2))
        row7.Cells(1, 6).Value = GetCellValueAsString(row.Cells(1, 5))
        row7.Cells(1, 13).Value = GetCellValueAsString(row.Cells(1, 4))

        ' �ե��`��ɂ������
        Debug.Print "FunctionId: " & GetCellValueAsString(row.Cells(1, 2))
        Debug.Print "Modifier: " & GetCellValueAsString(row.Cells(1, 3))
        Debug.Print "FunctionName (Logical): " & GetCellValueAsString(row.Cells(1, 5))
        Debug.Print "FunctionName (Physical): " & GetCellValueAsString(row.Cells(1, 4))
    Next j

    ' ���ꥸ�ʥ��2��Ŀ�Υ��`�Ȥ�����
    ' wb.Sheets(2).Delete

    ' ������Excel�ե�����򱣴�
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

Sub ReadExcelAndAddSheet()
    Dim filePath As String
    filePath = "/Users/luo/Documents/testFiles/testExcel.xlsx"

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath)

    ' ワークブック内のすべてのシート数を取得
    Dim sheetCount As Integer
    sheetCount = wb.Sheets.Count

    ' すべてのシートを走査し、Sheet1とSheet2以外のすべてのシートを削除
    Dim i As Integer
    For i = sheetCount To 3 Step -1
        If i <> 1 And i <> 2 Then
            wb.Sheets(i).Delete
        End If
    Next i

    ' 最初のシートを取得
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    ' 33行目から各行の最初の5つのフィールドを読み取る
    Dim startRow As Integer
    startRow = 33
    ' 実際に表示される最後の行番号を取得
    Dim endRow As Integer
    endRow = WorksheetFunction.Min(ws.Cells(Rows.Count, 1).End(xlUp).Row, 66)

    ' 2番目のシートから新しいシートを作成
    Dim j As Integer
    For j = startRow To endRow
        Dim row As Range
        Set row = ws.Rows(j)

        ' 最初のセルを取得する前に、行が空かどうかを再度確認
        Dim firstCell As Range
        Set firstCell = row.Cells(1, 1)
        If firstCell Is Nothing Or firstCell.Value = "" Then
            Exit For ' その行の最初のセルが値を持っていない場合、ループを中断
        End If

        Dim functionNameLogical As String
        functionNameLogical = GetCellValueAsString(row.Cells(1, 4))

        ' 2番目のシートをテンプレートとしてコピー
        Dim templateSheet As Worksheet
        Set templateSheet = wb.Sheets(2)
        Dim newSheet As Worksheet
        Set newSheet = templateSheet.Copy(, wb.Sheets(wb.Sheets.Count))
        newSheet.Name = "functionName_" & functionNameLogical & j

        ' テンプレートのシート内容を新しいシートにコピー
        templateSheet.UsedRange.Copy newSheet.Range("A1")

        ' 新しいシートの6行目にセルを作成し、6Cにフィールド値を記入
        Dim row6 As Range
        Set row6 = newSheet.Rows(6)
        row6.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 3))
        ' 新しいシートの7行目にセルを作成し、7C、7F、7Mにフィールド値を記入
        Dim row7 As Range
        Set row7 = newSheet.Rows(7)
        row7.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 2))
        row7.Cells(1, 6).Value = GetCellValueAsString(row.Cells(1, 5))
        row7.Cells(1, 13).Value = GetCellValueAsString(row.Cells(1, 4))

        ' フィールド値を出力
        Debug.Print "FunctionId: " & GetCellValueAsString(row.Cells(1, 2))
        Debug.Print "Modifier: " & GetCellValueAsString(row.Cells(1, 3))
        Debug.Print "FunctionName (Logical): " & GetCellValueAsString(row.Cells(1, 5))
        Debug.Print "FunctionName (Physical): " & GetCellValueAsString(row.Cells(1, 4))
    Next j

    ' オリジナルの2番目のシートを削除
    ' wb.Sheets(2).Delete

    ' 変更後のExcelファイルを保存
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

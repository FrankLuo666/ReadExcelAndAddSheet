Sub ReadExcelAndAddSheet()
    Dim filePath As String
    filePath = "/Users/luo/Documents/testFiles/testExcel.xlsx"

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath)

    ' Get the number of sheets in the workbook
    Dim sheetCount As Integer
    sheetCount = wb.Sheets.Count

    ' Iterate through all sheets and delete all sheets except Sheet1 and Sheet2
    Dim i As Integer
    For i = sheetCount To 3 Step -1
        If i <> 1 And i <> 2 Then
            wb.Sheets(i).Delete
        End If
    Next i

    ' Get the first sheet
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    ' Read the first 5 fields from each row starting from row 33
    Dim startRow As Integer
    startRow = 33
    ' Get the actual last row number to be displayed
    Dim endRow As Integer
    endRow = WorksheetFunction.Min(ws.Cells(Rows.Count, 1).End(xlUp).Row, 66)

    ' Create new sheets starting from the second sheet
    Dim j As Integer
    For j = startRow To endRow
        Dim row As Range
        Set row = ws.Rows(j)

        ' Before getting the first cell, check again if the row is empty
        Dim firstCell As Range
        Set firstCell = row.Cells(1, 1)
        If firstCell Is Nothing Or firstCell.Value = "" Then
            Exit For ' If the first cell in the row has no value, exit the loop
        End If

        Dim functionNameLogical As String
        functionNameLogical = GetCellValueAsString(row.Cells(1, 4))

        ' Copy the second sheet as a template
        Dim templateSheet As Worksheet
        Set templateSheet = wb.Sheets(2)
        Dim newSheet As Worksheet
        Set newSheet = templateSheet.Copy(, wb.Sheets(wb.Sheets.Count))
        newSheet.Name = "functionName_" & functionNameLogical & j

        ' Copy the content of the template sheet to the new sheet
        templateSheet.UsedRange.Copy newSheet.Range("A1")

        ' Create cells in the 6th row of the new sheet and fill in the values in column 6C
        Dim row6 As Range
        Set row6 = newSheet.Rows(6)
        row6.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 3))
        ' Create cells in the 7th row of the new sheet and fill in the values in columns 7C, 7F, 7M
        Dim row7 As Range
        Set row7 = newSheet.Rows(7)
        row7.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 2))
        row7.Cells(1, 6).Value = GetCellValueAsString(row.Cells(1, 5))
        row7.Cells(1, 13).Value = GetCellValueAsString(row.Cells(1, 4))

        ' Output the field values
        Debug.Print "FunctionId: " & GetCellValueAsString(row.Cells(1, 2))
        Debug.Print "Modifier: " & GetCellValueAsString(row.Cells(1, 3))
        Debug.Print "FunctionName (Logical): " & GetCellValueAsString(row.Cells(1, 5))
        Debug.Print "FunctionName (Physical): " & GetCellValueAsString(row.Cells(1, 4))
    Next j

    ' Delete the original Sheet2
    ' wb.Sheets(2).Delete

    ' Save the modified Excel file
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

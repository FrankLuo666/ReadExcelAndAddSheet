Sub ReadExcelAndAddSheet()
    Dim wb As Workbook
    Set wb = ThisWorkbook ' Use the active workbook as the source

    ' Get the first sheet
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    ' Read each row's first 5 fields starting from row 33
    Dim startRow As Integer
    startRow = 33
    ' Get the actual last displayed row number
    Dim endRow As Integer
    endRow = WorksheetFunction.Min(ws.Cells(Rows.Count, 1).End(xlUp).Row, 66)

    ' Create new sheets starting from the second sheet
    Dim j As Integer
    For j = startRow To endRow
        Dim row As Range
        Set row = ws.Rows(j)

        ' Check again if the row is empty before getting the first cell
        Dim firstCell As Range
        Set firstCell = row.Cells(1, 1)
        If firstCell Is Nothing Or firstCell.Value = "" Then
            Exit For ' Exit the loop if the first cell in the row is empty
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

        ' Create cells in the 6th row of the new sheet and fill in field values in column C
        Dim row6 As Range
        Set row6 = newSheet.Rows(6)
        row6.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 3))
        ' Create cells in the 7th row of the new sheet and fill in field values in columns C, F, and M
        Dim row7 As Range
        Set row7 = newSheet.Rows(7)
        row7.Cells(1, 3).Value = GetCellValueAsString(row.Cells(1, 2))
        row7.Cells(1, 6).Value = GetCellValueAsString(row.Cells(1, 5))
        row7.Cells(1, 13).Value = GetCellValueAsString(row.Cells(1, 4))

        ' Output field values
        Debug.Print "FunctionId: " & GetCellValueAsString(row.Cells(1, 2))
        Debug.Print "Modifier: " & GetCellValueAsString(row.Cells(1, 3))
        Debug.Print "FunctionName (Logical): " & GetCellValueAsString(row.Cells(1, 5))
        Debug.Print "FunctionName (Physical): " & GetCellValueAsString(row.Cells(1, 4))
    Next j

    ' Save the modified Excel file
    wb.Save
    'wb.Close ' You may or may not want to close the workbook depending on your requirements

    'Set wb = Nothing ' No need to set to Nothing since it's the active workbook
End Sub

Function GetCellValueAsString(cell As Range) As String
    If cell Is Nothing Then
        GetCellValueAsString = ""
    Else
        GetCellValueAsString = CStr(cell.Value)
    End If
End Function

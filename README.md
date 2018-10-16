# Excel-Text-Finder
A vb script code to calculate count of a text in an excel file
Private Sub CommandButton1_Click()
string_to_search = Sheets("Sheet1").Range("K10").Value
count_of_string = 0
Dim my_FileName As Variant
my_FileName = Application.GetOpenFilename
Dim wb As Excel.Workbook
Set wb = Workbooks.Open(my_FileName)
Dim sinput
sinput = InputBox("Enter Sheet Name")
ActiveWorkbook.Sheets(sinput).Activate
For Row = 1 To ActiveSheet.UsedRange.Rows.Count
    For Column = 1 To ActiveSheet.UsedRange.Columns.Count
        CurrentCellText = ActiveSheet.Cells(Row, Column).Value
        startIndex = 1
        Position = InStr(startIndex, CurrentCellText, string_to_search)
        Do While Position > 0
            Position = InStr(startIndex, CurrentCellText, string_to_search)
            startIndex = Position + Len(string_to_search)
        If Position > 0 Then
             count_of_string = count_of_string + 1
            End If
            Loop
            Next Column
        Next Row
        MsgBox "The given text is found " & count_of_string & " times."
End Sub


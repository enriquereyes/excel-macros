' Get customer workbook...
Dim customerBook As Workbook
Dim filter As String
Dim caption As String
Dim customerFilename As String
Dim customerWorkbook As Workbook
Dim targetWorkbook As Workbook

' make weak assumption that active workbook is the target
Set targetWorkbook = Application.ActiveWorkbook

' get the customer workbook
filter = "Text files (*.xlsx),*.xlsx"
caption = "Please Select an input file "
customerFilename = Application.GetOpenFilename(filter, , caption)

Set customerWorkbook = Application.Workbooks.Open(customerFilename)

' assume range is A1 - C10 in sheet1
' copy data from customer to target workbook
Dim targetSheet As Worksheet
Set targetSheet = targetWorkbook.Worksheets(1)
Dim sourceSheet As Worksheet
Set sourceSheet = customerWorkbook.Worksheets(1)

targetSheet.Range("A1", "C10").Value = sourceSheet.Range("A1", "C10").Value

' Close customer workbook
customerWorkbook.Close
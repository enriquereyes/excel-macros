Sub PegarMonto()
'
' PegarMonto Macro
'

Dim book As Workbook
Dim hoja1 As Worksheet
Dim hoja2 As Worksheet
Dim i As Integer
Dim j As Integer
Dim x As String
Dim y As String

Set book = Application.ActiveWorkbook
Set hoja1 = book.Worksheets(1)
Set hoja2 = book.Worksheets(2)

For i = 4 To 1000
    x = CStr(hoja1.Cells(i, "E").Value)
    For j = 2 To 1000
        y = CStr(hoja2.Cells(j, "C").Value)
        If x = y Then
            hoja1.Cells(i, "K").Value = hoja2.Cells(j, "E").Value
        End If
    Next j
    If hoja1.Cells(i, "E").Value = vbNullString Then
        End
    End If
Next i
'
End Sub
Sub Hipervinculos()
'
' Hipervinculos Macro
'

'Crea hipervínculos al documento llamado igual al texto de la celda "iA" con extensión .docx

    Dim n As Long, i As Long
    Dim cell As String
    n = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To n
        Cells(i, "A").Select
        cell = CStr(Cells(i, "A").Value) + ".docx"
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=cell
    Next i
End Sub

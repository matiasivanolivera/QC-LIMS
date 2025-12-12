Attribute VB_Name = "modFuncionesAutocompletar"
Public Function ObtenerValoresDeHoja(ws As Worksheet, Optional col As Long = 1) As Collection
    Dim colVals As New Collection
    Dim i As Long, ult As Long
    ult = ws.Cells(ws.Rows.Count, col).End(xlUp).row

    For i = 2 To ult
        If Trim(ws.Cells(i, col).Value) <> "" Then
            colVals.Add CStr(ws.Cells(i, col).Value)
        End If
    Next i

    Set ObtenerValoresDeHoja = colVals
End Function


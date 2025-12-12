Attribute VB_Name = "modExtractorIndustrial"
Option Explicit

Public Function ExtraerActividadesIndustrialesDesdeHoja(ByVal wsh As Worksheet) As Collection

    Dim colRes As New Collection          ' siempre existe, aunque quede vacía
    Dim ultimaFila As Long
    Dim ultimaCol As Long
    Dim r As Long, c As Long
    Dim txt As String
    Dim colLinea As Collection
    Dim i As Long
    Dim a As cActividadOT

    '------------------------------------
    ' Delimitamos el área de trabajo
    '------------------------------------
    ' Fila: hasta la última fila con algo en la columna C
    ultimaFila = wsh.Cells(wsh.Rows.Count, "C").End(xlUp).row
    ' Columna: desde C hasta la última usada en la fila 3 (cabeceras)
    ultimaCol = wsh.Cells(3, wsh.Columns.Count).End(xlToLeft).Column

    ' Cuerpo de la semana: filas 4..ultimaFila, columnas 3..ultimaCol
    For r = 4 To ultimaFila
        For c = 3 To ultimaCol

            txt = CStr(wsh.Cells(r, c).Value)

            If Trim$(txt) <> "" Then

                ' Parsear SOLO esa celda
                Set colLinea = ParsearLinea(txt)

                If Not colLinea Is Nothing Then
                    If colLinea.Count > 0 Then

                        For i = 1 To colLinea.Count

                            Set a = colLinea(i)

                            ' completar hoja y celda origen
                            a.Hoja = wsh.Name
                            a.Celda = wsh.Cells(r, c).Address(False, False)

                            colRes.Add a

                        Next i

                    End If
                End If

            End If

        Next c
    Next r

    ' Devolvemos SIEMPRE una colección (aunque vacía)
    Set ExtraerActividadesIndustrialesDesdeHoja = colRes

End Function



Private Function ContieneLote(ByVal txt As String) As Boolean
    Dim rgx As Object
    Dim m As Object
    
    Set rgx = CreateObject("VBScript.RegExp")
    With rgx
        .Pattern = "l\d{6}"
        .IgnoreCase = True
        .Global = False
    End With
    
    ContieneLote = rgx.Test(LCase$(txt))
End Function


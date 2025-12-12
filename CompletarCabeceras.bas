Attribute VB_Name = "CompletarCabeceras"
Sub CompletarCabecerasDesdeMarzo_Seguro()

    Dim wsCal As Worksheet
    Dim Hoja As Worksheet
    Dim idx As Long
    Dim filaCal As Long
    Dim ultimaFila As Long
    Dim col As Long

    Set wsCal = ThisWorkbook.Worksheets("CALENDARIO_2026")

    ' Buscar primer día hábil de marzo (lunes a sábado)
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, "A").End(xlUp).row
    
    filaCal = 0
    For i = 2 To ultimaFila
        If Month(wsCal.Cells(i, "A").Value) = 3 Then               ' mes marzo
            If LCase(wsCal.Cells(i, "B").Value) <> "domingo" Then  ' excluir domingo
                filaCal = i
                Exit For
            End If
        End If
    Next i
    
    If filaCal = 0 Then
        MsgBox "No se encontró un día válido para iniciar marzo.", vbCritical
        Exit Sub
    End If

    ' Ubicar índice de la hoja MAR(1)
    On Error Resume Next
    idx = ThisWorkbook.Worksheets("MAR(1)").Index
    On Error GoTo 0

    If idx = 0 Then
        MsgBox "No existe la hoja MAR(1).", vbCritical
        Exit Sub
    End If

    ' Recorrer hojas desde MAR(1) hacia adelante
    For i = idx To ThisWorkbook.Worksheets.Count
        
        Set Hoja = ThisWorkbook.Worksheets(i)

        ' Solo procesar hojas cuyo nombre contenga "(" y ")"
        If InStr(Hoja.Name, "(") > 0 And InStr(Hoja.Name, ")") > 0 Then
            
            ' Limpiar cabeceras
            Hoja.Range("B1:G2").ClearContents

            ' Colocar 6 fechas consecutivas (lunes a sábado)
            For col = 2 To 7
                
                Hoja.Cells(1, col).Value = wsCal.Cells(filaCal, "A").Value   ' Fecha
                Hoja.Cells(2, col).Value = wsCal.Cells(filaCal, "B").Value   ' Día

                ' Avanzar una fila en el calendario
                filaCal = filaCal + 1

                ' Saltar domingos
                Do While LCase(wsCal.Cells(filaCal, "B").Value) = "domingo"
                    filaCal = filaCal + 1
                Loop

            Next col

        End If

    Next i

    MsgBox "Cabeceras actualizadas desde MAR(1).", vbInformation

End Sub



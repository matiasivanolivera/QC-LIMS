Attribute VB_Name = "modOT"
Public Function ParsearActividad(texto As String) As Collection
    Dim resultado As New Collection
    Dim lineas() As String
    Dim linea As Variant
    Dim actsLinea As Collection
    Dim a As cActividadOT
    Dim i As Long

    If Trim(texto) = "" Then
        Set ParsearActividad = resultado
        Exit Function
    End If

    lineas = Split(texto, vbLf)

    For Each linea In lineas
        If Trim(linea) <> "" Then
            Set actsLinea = ParsearLinea(CStr(linea))

            If Not actsLinea Is Nothing Then
                For i = 1 To actsLinea.Count
                    resultado.Add actsLinea(i)
                Next i
            End If
        End If
    Next linea

    Set ParsearActividad = resultado
End Function


Private Function ParsearLinea(linea As String) As Collection
    Dim partes1() As String
    Dim partes2() As String
    Dim Especialidad As String, Variante As String
    Dim tecnicaCompleta As String
    Dim lotesTexto As String
    Dim lotes() As String
    Dim lote As Variant
    Dim subt As Collection
    Dim acts As New Collection
    Dim subTec As Variant
    Dim cleaned As String
    Dim act As cActividadOT      ' ? solo una variable objeto

    linea = Trim(linea)
    If linea = "" Then Exit Function

    partes1 = Split(linea, ":", 2)
    If UBound(partes1) < 1 Then Exit Function

    cleaned = Trim(partes1(0))
    Call SepararEspecialidadVariante(cleaned, Especialidad, Variante)

    partes2 = Split(partes1(1), "-", 2)
    If UBound(partes2) < 1 Then Exit Function

    tecnicaCompleta = Trim(partes2(0))
    lotesTexto = Trim(partes2(1))

    lotes = Split(lotesTexto, ",")

    Set subt = ExpandirTecnica(tecnicaCompleta)

    For Each lote In lotes
        lote = Trim(CStr(lote))
        If lote <> "" Then
            For Each subTec In subt
                Set act = New cActividadOT                  ' ? crear instancia
                act.Especialidad = Especialidad
                act.Variante = Variante
                act.Ensayo = Especialidad
                act.tecnica = CStr(subTec)
                act.NPLote = lote
                act.TextoCrudo = linea

                acts.Add act                                ' ? objeto a la colección
            Next subTec
        End If
    Next lote

    Set ParsearLinea = acts
End Function


Private Sub SepararEspecialidadVariante(texto As String, _
                                        ByRef Especialidad As String, _
                                        ByRef Variante As String)

    Dim partes() As String
    partes = Split(texto, " ")

    Dim last As String
    last = partes(UBound(partes))

    If InStr(1, last, "(", vbTextCompare) > 0 And _
       Right(last, 1) = ")" Then

        Variante = last
        Especialidad = Trim(Left(texto, Len(texto) - Len(last)))
    Else
        Variante = ""
        Especialidad = texto
    End If
End Sub

Public Function ExpandirTecnica(tec As String) As Collection
    Dim col As New Collection
    Dim tmp As String
    Dim partes() As String
    Dim raw As Variant
    Dim cleaned As String
    Dim baseEsTest As Boolean

    tmp = Replace(Replace(tec, "/", "+"), "  ", " ")
    partes = Split(tmp, "+")

    tmp = UCase$(Trim$(tec))
    If Left$(tmp, 4) = "TEST" Then baseEsTest = True

    For Each raw In partes
        cleaned = UCase$(Trim(CStr(raw)))

        If cleaned = "" Then GoTo NextRaw

        If cleaned = "S1" Or cleaned = "S2" Then
            If baseEsTest Then
                col.Add "TEST " & cleaned
            End If
            GoTo NextRaw
        End If

        col.Add cleaned

NextRaw:
    Next raw

    Set ExpandirTecnica = col
End Function

Public Function ExtraerActividadesPendientes(wsh As Worksheet) As Collection
    
    Dim col As New Collection
    Dim ultimaFila As Long, ultimaCol As Long
    Dim r As Long, c As Long
    Dim texto As String
    Dim acts As Collection
    Dim act As cActividadOT
    
    ultimaFila = wsh.UsedRange.Rows.Count
    ultimaCol = wsh.UsedRange.Columns.Count
    
    For r = 1 To ultimaFila
        For c = 1 To ultimaCol
            
            texto = Trim(wsh.Cells(r, c).Value)
            If texto <> "" Then
                
                ' --- llamar al parser ---
                Set acts = ParsearLinea(texto)
                
                If Not acts Is Nothing Then
                    If acts.Count > 0 Then
                        
                        Dim i As Long
                        For i = 1 To acts.Count
                            
                            act = acts(i)
                            
                            ' completar datos de ubicación
                            act.Hoja = wsh.Name
                            act.Celda = wsh.Cells(r, c).Address
                            
                            ' agregar al resultado
                            col.Add act
                            
                        Next i
                        
                    End If
                End If
            
            End If
        
        Next c
    Next r

    Set ExtraerActividadesPendientes = col
End Function

Public Function GenerarOT_ID(ByVal fechaOT As Date, ByVal Analista As String) As String
    Dim ws As Worksheet
    Dim last As Long
    Dim n As Long
    Dim partes() As String
    Dim sufijo As Long
    
    Set ws = ThisWorkbook.Worksheets("ORDENES_TRABAJO")
    
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Por defecto, primer número de secuencia
    sufijo = 1
    
    If last >= 2 Then
        partes = Split(CStr(ws.Cells(last, 1).Value), "-")
        If UBound(partes) >= 3 Then
            On Error Resume Next
            n = CLng(partes(3))
            On Error GoTo 0
            If n > 0 Then sufijo = n + 1
        End If
    End If
    
    GenerarOT_ID = "OT-" & Format(fechaOT, "yyyymmdd") & "-" & Analista & "-" & Format(sufijo, "000")
End Function

Public Function OTExiste(ByVal OT_ID As String) As Boolean
    Dim ws As Worksheet
    Dim f As Range
    
    Set ws = ThisWorkbook.Worksheets("ORDENES_TRABAJO")
    
    On Error Resume Next
    Set f = ws.Columns(1).Find(What:=OT_ID, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    OTExiste = Not f Is Nothing
End Function

Public Sub RegistrarOT(ByVal OT_ID As String, ByVal colSel As Collection)
    Dim ws As Worksheet
    Dim wsLog As Worksheet
    Dim i As Long
    Dim r As Long
    Dim a As cActividadOT
    Dim rLog As Long
    
    Set ws = ThisWorkbook.Worksheets("ORDENES_TRABAJO")
    
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("LOG_OT")
    On Error GoTo 0
    
    For i = 1 To colSel.Count
        
        Set a = colSel(i)
        
        r = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
        
        ws.Cells(r, 1).Value = OT_ID               ' OT_ID
        ws.Cells(r, 2).Value = a.Fecha             ' Fecha
        ws.Cells(r, 3).Value = a.Analista          ' Analista
        ws.Cells(r, 4).Value = a.Ensayo            ' Actividad / Ensayo
        ws.Cells(r, 5).Value = a.NPLote            ' NP / Lote
        ws.Cells(r, 6).Value = a.tecnica           ' Técnica
        ws.Cells(r, 7).Value = "PENDIENTE"         ' Estado inicial
        ws.Cells(r, 8).Value = Now                 ' Timestamp
        
        If Not wsLog Is Nothing Then
            rLog = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
            wsLog.Cells(rLog, 1).Value = Now
            wsLog.Cells(rLog, 2).Value = OT_ID
            wsLog.Cells(rLog, 3).Value = a.Hoja
            wsLog.Cells(rLog, 4).Value = a.Celda
            wsLog.Cells(rLog, 5).Value = a.TextoCrudo
        End If
        
    Next i
End Sub

Public Sub ColorearCeldas(ByVal colSel As Collection)
    Dim i As Long
    Dim a As cActividadOT
    Dim ws As Worksheet
    
    For i = 1 To colSel.Count
        Set a = colSel(i)
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(a.Hoja)
        If Not ws Is Nothing Then
            ws.Range(a.Celda).Interior.Color = vbYellow
        End If
        On Error GoTo 0
        
        Set ws = Nothing
    Next i
End Sub

Public Sub RegistrarLOG_OT(ByVal mensaje As String)
    Dim ws As Worksheet
    Dim r As Long
    
    ' Intentar usar la hoja LOG_OT si existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOG_OT")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Si no existe, simplemente mostrar en la ventana Inmediato
        Debug.Print Now & " - " & mensaje
        Exit Sub
    End If
    
    ' Escribir una línea de log resumen
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = mensaje
End Sub
Public Function ObtenerFechaDesdeColumna(ByVal wsh As Worksheet, ByVal col As Long) As Date
    Dim v As Variant
    
    ' Las fechas están en la fila 1, desde la columna B en adelante
    If col < 2 Then
        ' Columna A no tiene fecha de cabecera, usar hoy como salvaguarda
        ObtenerFechaDesdeColumna = Date
        Exit Function
    End If
    
    v = wsh.Cells(1, col).Value
    
    If IsDate(v) Then
        ObtenerFechaDesdeColumna = CDate(v)
    Else
        ' Si por algún motivo no hay fecha válida, usar hoy para no romper el flujo
        ObtenerFechaDesdeColumna = Date
    End If
End Function








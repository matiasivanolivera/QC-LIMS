Attribute VB_Name = "modParserIndustrial"
'==========================================
' Módulo : modParserIndustrial
' Proyecto: CMDT '25 – mini LIMS
' Rol    : Parsing de celdas del cronograma
'          ? genera cActividadOT a partir
'          de texto de cronograma + maestros.
'
' Depende de:
'   - cActividadOT
'   - modDiccionarios (maestros)
'
' Usado por:
'   - modExtractorIndustrial
'   - frmOTReview (a través de ExtraerActividadesIndustrialesDesdeHoja)
'==========================================

Option Explicit

Public Function ParsearActividadIndustrial(txt As String) As Collection
    Dim col As New Collection
    Dim lineas() As String
    Dim i As Long

    lineas = Split(txt, vbLf)

    For i = LBound(lineas) To UBound(lineas)
        Dim linea As String
        linea = Trim(lineas(i))

        If linea <> "" Then
            Dim acts As Collection
            Set acts = ParsearLinea(linea)

            Dim j As Long
            For j = 1 To acts.Count
                col.Add acts(j)
            Next j
        End If
    Next i

    Set ParsearActividadIndustrial = col
End Function

Public Function ParsearLinea(ByVal t As String) As Collection
    Dim col As New Collection
    Dim texto As String
    Dim partesGuion() As String
    Dim parteIzq As String, parteDer As String
    Dim esp As String, pres As String, ff As String, aliasAuto As String, tipo As String
    Dim tecnicas As Collection
    Dim lots() As String
    Dim lote As Variant, subTec As Variant
    Dim a As cActividadOT

    ' Normalizar texto
    texto = Trim$(t)
    If texto = "" Then Exit Function
    
    ' Debe tener al menos un ":" y un "-"
    If InStr(texto, ":") = 0 Then Exit Function
    If InStr(texto, "-") = 0 Then Exit Function
    
    ' Separar en parteIzq (antes de "-") y parteDer (después de "-")
    partesGuion = Split(texto, "-", 2)
    If UBound(partesGuion) < 1 Then Exit Function
    
    parteIzq = Trim$(partesGuion(0))  ' especialidad + técnicas
    parteDer = Trim$(partesGuion(1))  ' lotes
    
    '===============================
    ' 1) ESPECIALIDAD + PRESENTACIÓN
    '===============================
    Call DetectarEspecialidad(LCase$(parteIzq), esp, pres, ff, tipo, aliasAuto)
    
    '===============================
    ' 2) TÉCNICAS (lista de ensayos)
    '===============================
    Set tecnicas = ExpandirTecnicaDesdeTexto(parteIzq)
    If tecnicas Is Nothing Then Exit Function
    If tecnicas.Count = 0 Then Exit Function
    
    '===============================
    ' 3) LOTES (solo en la parte derecha)
    '===============================
    lots = ExtraerLotes(parteDer)
    If Not TieneElementos(lots) Then Exit Function
    
    '===============================
    ' 4) GENERAR COMBINACIONES
    '===============================
    For Each lote In lots
        lote = Trim$(CStr(lote))
        If lote <> "" Then
            For Each subTec In tecnicas
                
                Set a = New cActividadOT

                ' Técnicas
                a.Ensayo = CStr(subTec)
                a.tecnica = CStr(subTec)

                ' Producto (lo nuevo)
                a.Especialidad = esp
                a.Presentacion = pres
                a.FormaFF = ff
                a.TipoProducto = tipo
                a.aliasAuto = aliasAuto

                ' Lo que se ve en la ListView
                a.Muestra = Trim$(esp & IIf(pres <> "", " " & pres, ""))
                a.NPLote = lote

                ' Metadatos
                a.TextoCrudo = texto
                a.Hoja = ""
                a.Celda = ""
                
                col.Add a
            Next subTec
        End If
    Next lote

    Set ParsearLinea = col
End Function




Private Sub DetectarEspecialidad(t As String, _
    ByRef esp As String, ByRef pres As String, _
    ByRef ff As String, ByRef tipo As String, _
    ByRef aliasAuto As String)

    Dim k As Variant
    For Each k In DicEsp.Keys
        If InStr(t, LCase(Split(k, " ")(0))) > 0 Then
            esp = DicEsp(k)(0)
            pres = DicEsp(k)(1)
            ff = DicEsp(k)(2)
            tipo = DicEsp(k)(3)
            aliasAuto = DicEsp(k)(4)
            Exit Sub
        End If
    Next k
End Sub

Private Function ExtraerLotes(ByVal t As String) As Variant
    Dim rgx As Object
    Dim matches As Object
    Dim i As Long
    Dim arr() As String

    Set rgx = CreateObject("VBScript.RegExp")
    With rgx
        .Global = True
        .IgnoreCase = True
        ' Captura:
        '   - L + 5 ó 6 dígitos  (ej: L049924, L12345)
        '   - NP + 4 a 6 dígitos (ej: NP0123, NP012345)
        .Pattern = "\b(L\d{5,6}|NP\d{4,6})\b"
    End With

    Set matches = rgx.Execute(t)

    If matches.Count = 0 Then
        ' Devolvemos un array vacío compatible
        ExtraerLotes = Split("", ";")
        Exit Function
    End If

    ReDim arr(matches.Count - 1)

    For i = 0 To matches.Count - 1
        arr(i) = UCase$(matches(i).Value)
    Next i

    ExtraerLotes = arr
End Function

Private Function ExtraerTecnica(t As String) As String
    Dim tecnicas As Variant
    tecnicas = Array("dosis", "id", "io", "ph", "ts", "test", "visc")

    Dim i As Long
    For i = LBound(tecnicas) To UBound(tecnicas)
        If InStr(t, tecnicas(i)) > 0 Then
            ExtraerTecnica = UCase(tecnicas(i))
            Exit Function
        End If
    Next i

    ExtraerTecnica = ""
End Function

Private Function ExtraerSubtecnicas(t As String) As Variant
    ExtraerSubtecnicas = Split("", ";")
End Function

Private Function TieneElementos(arr As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsArray(arr) Then
        If UBound(arr) >= LBound(arr) Then
            TieneElementos = True
        End If
    End If
    Exit Function
ErrHandler:
    TieneElementos = False
End Function

Public Function ExpandirTecnicaDesdeTexto(ByVal parteIzq As String) As Collection
    Dim c As New Collection
    Dim txt As String
    Dim trozos() As String
    Dim i As Long
    Dim token As String
    
    ' parteIzq viene con algo como:
    ' "CLARITROMICINA 500mg PI(GCR): DOSIS / ID / PH"
    ' Nos quedamos con lo posterior a los ":" si los hay
    If InStr(parteIzq, ":") > 0 Then
        txt = Split(parteIzq, ":", 2)(1)
    Else
        txt = parteIzq
    End If
    
    txt = UCase$(txt)
    
    ' Separar por "/" y espacios
    trozos = Split(txt, "/")
    
    For i = LBound(trozos) To UBound(trozos)
        token = Trim$(trozos(i))
        If token <> "" Then
            ' Opcional: limpiar cosas tipo "DOSIS S1" ? "DOSIS"
            token = Split(token, " ")(0)
            
            ' Podrías validar contra una lista de técnicas válidas si querés
            On Error Resume Next
            c.Add token, token   ' evita duplicados usando la clave = valor
            On Error GoTo 0
        End If
    Next i
    
    Set ExpandirTecnicaDesdeTexto = c
End Function






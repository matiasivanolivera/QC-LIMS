Attribute VB_Name = "modAnalistas"
Option Explicit

'==========================================
' Módulo : modAnalistas
' Rol    : Resolver el ID de analista de
'          una actividad del cronograma.
'
' Estrategia:
'   - Los IDs válidos (MF, PMV, LNM, MSG, ...)
'     se cargan desde MAESTRO_ANALISTAS (col A).
'   - Para una celda de actividad (fila, col):
'       * se sube POR LA COLUMNA B
'         hasta encontrar una cabecera de bloque
'         (texto no vacío).
'       * se normaliza ese texto (MGA TARDE ? MGA)
'       * se acepta solo si el ID existe
'         en MAESTRO_ANALISTAS.
'==========================================

' IDs válidos según MAESTRO_ANALISTAS
Private DicAnalistasMaestro As Object

'---------------------------------------
' Cargar IDs válidos desde MAESTRO_ANALISTAS
'---------------------------------------
Private Sub AsegurarMaestroAnalistas()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim id As String

    If Not DicAnalistasMaestro Is Nothing Then Exit Sub

    Set DicAnalistasMaestro = CreateObject("Scripting.Dictionary")
    DicAnalistasMaestro.CompareMode = vbTextCompare

    Set ws = ThisWorkbook.Worksheets("MAESTRO_ANALISTAS")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 2 To lastRow
        id = UCase$(Trim$(CStr(ws.Cells(r, "A").Value)))
        If id <> "" Then
            If Not DicAnalistasMaestro.Exists(id) Then
                DicAnalistasMaestro.Add id, True
            End If
        End If
    Next r

End Sub

'---------------------------------------
' Normalizar texto crudo a ID:
'   "MGA TARDE"  ? "MGA"
'   "PMV - MAÑANA" ? "PMV"
'---------------------------------------
Public Function ExtraerIdAnalista(ByVal raw As String) As String

    Dim i As Long
    Dim ch As String

    raw = UCase$(Trim$(raw))
    If raw = "" Then Exit Function

    ' Cortar por espacio y por guión
    raw = Split(raw, " ")(0)
    raw = Split(raw, "-")(0)

    ' Dejar solo letras A..Z al principio
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch < "A" Or ch > "Z" Then Exit For
    Next i

    ExtraerIdAnalista = Left$(raw, i - 1)

End Function

'---------------------------------------
' Buscar hacia ARRIBA en la COLUMNA B
' hasta encontrar un ID válido
'---------------------------------------
Public Function ObtenerAnalistaDesdeBloque( _
                ByVal wsh As Worksheet, _
                ByVal fila As Long, _
                ByVal col As Long) As String

    Dim r As Long
    Dim raw As String
    Dim id As String

    AsegurarMaestroAnalistas

    ' Cabeceras de bloque están en columna B (col = 2)
    For r = fila To 3 Step -1

        raw = CStr(wsh.Cells(r, 2).Value)   ' columna B fija

        If Trim$(raw) <> "" Then
            id = ExtraerIdAnalista(raw)
            If id <> "" Then
                If DicAnalistasMaestro.Exists(id) Then
                    ObtenerAnalistaDesdeBloque = id
                    Exit Function
                End If
            End If
        End If

    Next r

    ObtenerAnalistaDesdeBloque = ""

End Function

'---------------------------------------
' Por si cambiás MAESTRO_ANALISTAS
'---------------------------------------
Public Sub InvalidarCacheAnalistas()
    Set DicAnalistasMaestro = Nothing
End Sub



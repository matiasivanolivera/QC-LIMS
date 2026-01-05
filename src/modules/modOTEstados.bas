Attribute VB_Name = "modOTEstados"
Option Explicit

' Devuelve el número de columna donde está el encabezado (fila 1)
' Ej: ColIndex(ws, "OT_ID") -> 1
Private Function ColIndex(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerText Then
            ColIndex = c
            Exit Function
        End If
    Next c
    
    ColIndex = 0 ' no encontrado
End Function

' Devuelve la última fila usada basada en una columna (colIdx)
Private Function LastRow(ByVal ws As Worksheet, ByVal colIdx As Long) As Long
    If colIdx <= 0 Then
        LastRow = 1
        Exit Function
    End If
    LastRow = ws.Cells(ws.Rows.Count, colIdx).End(xlUp).row
End Function
'========================
' OT - CAMBIO DE ESTADO
'========================

Public Function CambiarEstadoOT(ByVal otId As String, _
                                ByVal nuevoEstado As String, _
                                Optional ByVal motivo As String = "", _
                                Optional ByVal usuario As String = "SISTEMA") As Boolean
    On Error GoTo EH
    
    Dim wsOT As Worksheet, wsLog As Worksheet
    Set wsOT = ThisWorkbook.Worksheets("ORDENES_TRABAJO")
    Set wsLog = ThisWorkbook.Worksheets("LOG_OT")
    
    ' Validar estado
    nuevoEstado = UCase$(Trim$(nuevoEstado))
    If Not EstadoValido(nuevoEstado) Then
        MsgBox "Estado inválido: " & nuevoEstado, vbCritical
        CambiarEstadoOT = False
        Exit Function
    End If
    
    ' Columnas en ORDENES_TRABAJO
    Dim cOtId As Long, cFecha As Long, cAnalista As Long
    Dim cEstadoOld As Long, cTsOld As Long
    Dim cEstadoNew As Long, cTsNew As Long, cUserNew As Long, cMotNew As Long
    
    cOtId = ColIndex(wsOT, "OT_ID")
    cFecha = ColIndex(wsOT, "Fecha")
    cAnalista = ColIndex(wsOT, "Analista")
    cEstadoOld = ColIndex(wsOT, "Estado")
    cTsOld = ColIndex(wsOT, "Timestamp")
    
    cEstadoNew = ColIndex(wsOT, "ESTADO_OT")
    cTsNew = ColIndex(wsOT, "ESTADO_TS")
    cUserNew = ColIndex(wsOT, "ESTADO_USUARIO")
    cMotNew = ColIndex(wsOT, "ESTADO_MOTIVO")
    
    If cOtId = 0 Then Err.Raise vbObjectError + 100, , "No existe encabezado OT_ID en ORDENES_TRABAJO."
    If cEstadoNew = 0 Or cTsNew = 0 Or cUserNew = 0 Or cMotNew = 0 Then
        Err.Raise vbObjectError + 101, , "Faltan columnas nuevas de estado (ESTADO_OT/TS/USUARIO/MOTIVO)."
    End If
    
    Dim lastR As Long, r As Long
    lastR = LastRow(wsOT, cOtId)
    If lastR < 2 Then
        MsgBox "ORDENES_TRABAJO no tiene datos.", vbInformation
        CambiarEstadoOT = False
        Exit Function
    End If
    
    ' Encontrar filas del OT y capturar datos base
    Dim encontrado As Boolean
    Dim estadoAnterior As String, fechaAct As Variant, analistaAct As String
    
    encontrado = False
    For r = 2 To lastR
        If Trim$(CStr(wsOT.Cells(r, cOtId).Value)) = otId Then
            If Not encontrado Then
                ' Estado anterior: preferir ESTADO_OT si ya tenía algo
                estadoAnterior = Trim$(CStr(wsOT.Cells(r, cEstadoNew).Value))
                If Len(estadoAnterior) = 0 And cEstadoOld > 0 Then
                    estadoAnterior = Trim$(CStr(wsOT.Cells(r, cEstadoOld).Value))
                End If
                If cFecha > 0 Then fechaAct = wsOT.Cells(r, cFecha).Value
                If cAnalista > 0 Then analistaAct = Trim$(CStr(wsOT.Cells(r, cAnalista).Value))
                encontrado = True
            End If
        End If
    Next r
    
    If Not encontrado Then
        MsgBox "No se encontró OT_ID=" & otId & " en ORDENES_TRABAJO.", vbExclamation
        CambiarEstadoOT = False
        Exit Function
    End If
    
    ' Actualizar todas las filas del OT
    Dim nowTs As Date
    nowTs = Now
    
    For r = 2 To lastR
        If Trim$(CStr(wsOT.Cells(r, cOtId).Value)) = otId Then
            wsOT.Cells(r, cEstadoNew).Value = nuevoEstado
            wsOT.Cells(r, cTsNew).Value = nowTs
            wsOT.Cells(r, cUserNew).Value = usuario
            wsOT.Cells(r, cMotNew).Value = motivo
            
            ' Compatibilidad con columnas viejas si existen
            If cEstadoOld > 0 Then wsOT.Cells(r, cEstadoOld).Value = nuevoEstado
            If cTsOld > 0 Then wsOT.Cells(r, cTsOld).Value = nowTs
        End If
    Next r
    
    ' Registrar log (1 fila, scope OT)
    RegistrarLogCambioEstado wsLog, otId, nowTs, usuario, fechaAct, analistaAct, estadoAnterior, nuevoEstado, motivo
    
    CambiarEstadoOT = True
    Exit Function
    
EH:
    MsgBox "CambiarEstadoOT error: " & Err.Description, vbCritical
    CambiarEstadoOT = False
End Function

Private Function EstadoValido(ByVal e As String) As Boolean
    Select Case e
        Case "PENDIENTE", "EN_PROCESO", "FINALIZADA", "ANULADA", "CANCELADA"
            EstadoValido = True
        Case Else
            EstadoValido = False
    End Select
End Function

Private Sub RegistrarLogCambioEstado(ByVal wsLog As Worksheet, _
                                     ByVal otId As String, _
                                     ByVal ts As Date, _
                                     ByVal usuario As String, _
                                     ByVal fechaAct As Variant, _
                                     ByVal analistaAct As String, _
                                     ByVal estadoAnt As String, _
                                     ByVal estadoNuevo As String, _
                                     ByVal motivo As String)
    Dim cTs As Long, cUser As Long, cFecha As Long, cAnalista As Long, cOtId As Long, cAccion As Long, cDetalle As Long
    Dim cLogId As Long, cEvento As Long, cEA As Long, cEN As Long, cMot As Long, cScope As Long, cClave As Long, cHoja As Long, cCelda As Long
    
    cTs = ColIndex(wsLog, "Timestamp")
    cUser = ColIndex(wsLog, "Usuario")
    cFecha = ColIndex(wsLog, "Fecha")
    cAnalista = ColIndex(wsLog, "Analista")
    cOtId = ColIndex(wsLog, "OT_ID")
    cAccion = ColIndex(wsLog, "Acción")
    cDetalle = ColIndex(wsLog, "Detalle")
    
    cLogId = ColIndex(wsLog, "LOG_ID")
    cEvento = ColIndex(wsLog, "EVENTO_TIPO")
    cEA = ColIndex(wsLog, "ESTADO_ANT")
    cEN = ColIndex(wsLog, "ESTADO_NUEVO")
    cMot = ColIndex(wsLog, "MOTIVO")
    cScope = ColIndex(wsLog, "SCOPE")
    cClave = ColIndex(wsLog, "CLAVE_ACTIVIDAD")
    cHoja = ColIndex(wsLog, "HOJA")
    cCelda = ColIndex(wsLog, "CELDA")
    
    Dim lastR As Long, newR As Long
    lastR = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row
    If lastR < 1 Then lastR = 1
    newR = lastR + 1
    
    ' Base
    If cTs > 0 Then wsLog.Cells(newR, cTs).Value = ts
    If cUser > 0 Then wsLog.Cells(newR, cUser).Value = usuario
    If cFecha > 0 Then wsLog.Cells(newR, cFecha).Value = fechaAct
    If cAnalista > 0 Then wsLog.Cells(newR, cAnalista).Value = analistaAct
    If cOtId > 0 Then wsLog.Cells(newR, cOtId).Value = otId
    If cAccion > 0 Then wsLog.Cells(newR, cAccion).Value = "CAMBIAR_ESTADO"
    If cDetalle > 0 Then wsLog.Cells(newR, cDetalle).Value = "Estado: " & estadoAnt & " -> " & estadoNuevo & ". Motivo: " & motivo
    
    ' v2
    If cLogId > 0 Then wsLog.Cells(newR, cLogId).Value = Format$(ts, "yyyymmdd-hhnnss") & "-" & otId
    If cEvento > 0 Then wsLog.Cells(newR, cEvento).Value = "CAMBIAR_ESTADO"
    If cEA > 0 Then wsLog.Cells(newR, cEA).Value = estadoAnt
    If cEN > 0 Then wsLog.Cells(newR, cEN).Value = estadoNuevo
    If cMot > 0 Then wsLog.Cells(newR, cMot).Value = motivo
    If cScope > 0 Then wsLog.Cells(newR, cScope).Value = "OT"
    
    ' En este paso (scope OT) estos quedan vacíos
    If cClave > 0 Then wsLog.Cells(newR, cClave).Value = vbNullString
    If cHoja > 0 Then wsLog.Cells(newR, cHoja).Value = vbNullString
    If cCelda > 0 Then wsLog.Cells(newR, cCelda).Value = vbNullString
End Sub



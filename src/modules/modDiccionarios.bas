Attribute VB_Name = "modDiccionarios"
Option Explicit

Public DicEsp As Object
Public DicFF As Object
Public DicAnalistasMaestro As Object

Public Sub CargarAnalistasDesdeMaestro()

    Dim ws As Worksheet
    Dim last As Long, r As Long
    Dim id As String, nombre As String

    ' Si ya está cargado, no repetir
    If Not DicAnalistasMaestro Is Nothing Then Exit Sub

    Set DicAnalistasMaestro = CreateObject("Scripting.Dictionary")
    DicAnalistasMaestro.CompareMode = vbTextCompare

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MAESTRO_ANALISTAS")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No se encontró la hoja MAESTRO_ANALISTAS.", vbExclamation
        Exit Sub
    End If

    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 2 To last
        id = Trim$(ws.Cells(r, 1).Value)       ' Columna A = ID (MF, PMV, PAG, ...)
        nombre = Trim$(ws.Cells(r, 2).Value)   ' Columna B = nombre (opcional)
        
        If id <> "" Then
            If Not DicAnalistasMaestro.Exists(id) Then
                DicAnalistasMaestro.Add id, nombre
            End If
        End If
    Next r

End Sub


Public Sub CargarDiccionariosDesdeMaestros()

    Dim wsEsp As Worksheet, wsFF As Worksheet
    Dim last As Long, r As Long
    Dim esp As String, pres As String, ff As String, tipo As String, aliasAuto As String

    Set DicEsp = CreateObject("Scripting.Dictionary")
    Set DicFF = CreateObject("Scripting.Dictionary")

    '====================
    ' ESPECIALIDADES
    '====================
    Set wsEsp = ThisWorkbook.Worksheets("MAESTRO_ESPECIALIDADES")
    last = wsEsp.Cells(wsEsp.Rows.Count, "A").End(xlUp).row

    For r = 2 To last
        esp = Trim(wsEsp.Cells(r, 1).Value)
        pres = Trim(wsEsp.Cells(r, 2).Value)
        ff = Trim(wsEsp.Cells(r, 3).Value)
        tipo = Trim(wsEsp.Cells(r, 4).Value)
        aliasAuto = GenerarAliasAutomatico(esp, pres, ff)

        Dim key As String
        key = LCase(esp & " " & pres)

        If Not DicEsp.Exists(key) Then
            DicEsp.Add key, Array(esp, pres, ff, tipo, aliasAuto)
        End If
    Next r

    '====================
    ' FORMAS FARMACÉUTICAS
    '====================
    Set wsFF = ThisWorkbook.Worksheets("MAESTRO_FF")
    last = wsFF.Cells(wsFF.Rows.Count, "A").End(xlUp).row

    For r = 2 To last
        Dim forma As String, cod As String, tipoRes As String
        forma = Trim(wsFF.Cells(r, 1).Value)
        cod = Trim(wsFF.Cells(r, 2).Value)
        tipoRes = Trim(wsFF.Cells(r, 3).Value)

        DicFF(cod) = tipoRes
    Next r

End Sub



Public Function GenerarAliasAutomatico(esp As String, pres As String, ffBase As String) As String

    Dim p As String: p = Replace(pres, " ", "")
    p = Replace(p, "%", "")
    p = Replace(p, "mg", "")
    p = Replace(p, "mcg", "")

    Dim siglas As String
    siglas = UCase(Left(esp, 4))

    GenerarAliasAutomatico = siglas & p & ffBase
End Function


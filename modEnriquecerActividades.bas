Attribute VB_Name = "modEnriquecerActividades"
Option Explicit

' ============================
' ESTRUCTURA GLOBAL DEL MÓDULO
' ============================
Type ParsedInfo
    Alias As String
    tipo As String
    Categoria As String
    tecnica As String
    Equipo As String
    StdPrim As String
    StdSec As String
    lote As String
End Type

Sub EnriquecerActividades()

    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim resp As VbMsgBoxResult
    Dim i As Long
    Dim info As ParsedInfo
    Dim Fecha As Date

    Set ws = ThisWorkbook.Worksheets("BASE_ACTIVIDADES")

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If ultimaFila < 2 Then
        MsgBox "No hay datos para enriquecer.", vbExclamation
        Exit Sub
    End If

    ' CONFIRMACIÓN PARA REGENERAR
    resp = MsgBox("¿Desea regenerar los datos de enriquecimiento?" & vbCrLf & _
                  "Esto sobrescribirá Alias, Tipo, Categoría, Técnica," & vbCrLf & _
                  "Instrumento, Estándares, Lote, Trimestre e ID_Extendido.", _
                  vbYesNo + vbQuestion)

    If resp = vbNo Then Exit Sub

    ' LIMPIAR COLUMNAS 9 A 18
    ws.Range("I2:R" & ultimaFila).ClearContents

    ' CARGAR MAESTROS EN MEMORIA
    Dim dicAlias As Object, dicVar As Object
    Dim dicTipo As Object, dicCat As Object, dicTec As Object
    Dim dicEq As Object, dicPrim As Object, dicSec As Object

    Set dicAlias = Load_DiccionarioAliasExactos()
    Set dicVar = Load_DiccionarioAliasVariantes()
    Set dicTipo = Load_DiccionarioTipos()
    Set dicCat = Load_DiccionarioCategorias()
    Set dicTec = Load_DiccionarioTecnicas()
    Set dicEq = Load_DiccionarioEquipos()
    Set dicPrim = Load_DiccionarioEstandaresPrimarios()
    Set dicSec = Load_DiccionarioEstandaresSecundarios()

    ' ENRIQUECER FILA POR FILA
    For i = 2 To ultimaFila
        
        Dim texto As String
        Dim Analista As String

        texto = CStr(ws.Cells(i, 7).Value)
        Analista = CStr(ws.Cells(i, 6).Value)
        Fecha = ws.Cells(i, 2).Value

        info = ParsearActividad(texto, dicAlias, dicVar, dicTipo, dicCat, dicTec, dicEq, dicPrim, dicSec)

        ' GUARDAR RESULTADOS
        ws.Cells(i, 9).Value = info.Alias
        ws.Cells(i, 10).Value = info.tipo
        ws.Cells(i, 11).Value = info.Categoria
        ws.Cells(i, 12).Value = info.tecnica
        ws.Cells(i, 13).Value = info.Equipo
        ws.Cells(i, 14).Value = info.StdPrim
        ws.Cells(i, 15).Value = info.StdSec
        ws.Cells(i, 16).Value = info.lote
        ws.Cells(i, 17).Value = TrimestreDesdeFecha(Fecha)
        ws.Cells(i, 18).Value = GenerarIDExtendido(Fecha, Analista, i - 1)
    Next i

    MsgBox "Enriquecimiento completado.", vbInformation

End Sub

  

    Function ParsearActividad(texto As String, _
                          dicAlias As Object, dicVar As Object, _
                          dicTipo As Object, dicCat As Object, dicTec As Object, _
                          dicEq As Object, dicPrim As Object, dicSec As Object) As ParsedInfo
                          
    Dim info As ParsedInfo
    Dim aliasDet As String
    Dim palabra As Variant
    Dim clave As Variant

    Dim t As String
    t = LCase(texto)

    ' 1) MATCH EXACTO
    For Each clave In dicAlias.Keys
        If InStr(t, LCase(clave)) > 0 Then
            aliasDet = clave
            Exit For
        End If
    Next clave

    ' 2) MATCH FLEXIBLE SI NO ENCONTRÓ EXACTO
    If aliasDet = "" Then
        For Each palabra In dicVar.Keys
            If InStr(t, LCase(palabra)) > 0 Then
                aliasDet = dicVar(palabra)
                Exit For
            End If
        Next palabra
    End If

    info.Alias = aliasDet

    ' 3) TIPO / CATEGORIA / TECNICA
    If aliasDet <> "" Then
        info.tipo = dicTipo(aliasDet)
        info.Categoria = dicCat(aliasDet)
        info.tecnica = dicTec(aliasDet)
    End If

    ' 4) EQUIPO (si aplica técnica instrumental)
    Dim tec As String
    tec = info.tecnica
    If tec <> "" Then
        For Each clave In dicEq.Keys
            If InStr(LCase(clave), LCase(tec)) > 0 Then
                info.Equipo = dicEq(clave)
                Exit For
            End If
        Next clave
    End If

    ' 5) DETECTAR LOTE
    info.lote = DetectarLote(t)

    ' 6) DETECTAR ESTÁNDARES
    info.StdPrim = DetectarEstandar(t, dicPrim)
    info.StdSec = DetectarEstandar(t, dicSec)

    ParsearActividad = info

End Function

Function DetectarLote(t As String) As String
    Dim rgx As Object, mc As Object
    Set rgx = CreateObject("VBScript.RegExp")
    
    rgx.Pattern = "(lote|lt|l)[ -:]?(\d+)"
    rgx.IgnoreCase = True
    rgx.Global = False

    If rgx.Test(t) Then
        Set mc = rgx.Execute(t)
        DetectarLote = mc(0).SubMatches(1)
    Else
        DetectarLote = ""
    End If
End Function


Function DetectarEstandar(t As String, dic As Object) As String
    Dim clave As Variant
    For Each clave In dic.Keys
        If InStr(t, LCase(clave)) > 0 Then
            DetectarEstandar = dic(clave)
            Exit Function
        End If
    Next clave
    DetectarEstandar = ""
End Function

Function TrimestreDesdeFecha(f As Date) As Integer
    TrimestreDesdeFecha = WorksheetFunction.RoundUp(Month(f) / 3, 0)
End Function


Function GenerarIDExtendido(f As Date, Analista As String, nro As Long) As String
    GenerarIDExtendido = Format(f, "yyyymmdd") & "-" & UCase(Analista) & "-ACT" & Format(nro, "0000")
End Function

Function Load_DiccionarioAliasExactos() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "Visc", "Visc"
    d.Add "pH", "pH"
    d.Add "PI", "PI"
    d.Add "IO", "IO"
    d.Add "KF", "KF"
    d.Add "TS", "TS"
    d.Add "ID", "ID"
    d.Add "UDS", "UDS"
    d.Add "UUD", "UUD"
    d.Add "Valor", "Valor"
    d.Add "Rev", "Rev"
    d.Add "Doc", "Doc"
    d.Add "AA", "AA"
    d.Add "Planta", "Planta"
    d.Add "Equipo", "Equipo"
    d.Add "Valid", "Valid"
    d.Add "CP", "CP"
    d.Add "CI", "CI"
    d.Add "OBS", "OBS"
    d.Add "TA", "TA"

    Set Load_DiccionarioAliasExactos = d
End Function

Function Load_DiccionarioAliasVariantes() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "visco", "Visc"
    d.Add "viscosidad", "Visc"
    d.Add "disol", "TS"
    d.Add "disolución", "TS"
    d.Add "karl", "KF"
    d.Add "fischer", "KF"
    d.Add "ph", "pH"
    d.Add "imp", "IO"
    d.Add "impureza", "IO"
    d.Add "valorac", "Valor"
    d.Add "titul", "Valor"
    d.Add "ident", "ID"
    d.Add "agua", "AA"
    d.Add "document", "Doc"
    d.Add "rev", "Rev"
    
    Set Load_DiccionarioAliasVariantes = d
End Function

Function Load_DiccionarioCategorias() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "Visc", "Reología"
    d.Add "pH", "Fisicoquímico"
    d.Add "PI", "Fisicoquímico"
    d.Add "IO", "Impurezas"
    d.Add "KF", "Humedad"
    d.Add "TS", "Fisicoquímico"
    d.Add "ID", "Identidad"
    d.Add "Valor", "Volumetría"
    d.Add "UUD", "Uniformidad"
    d.Add "UDS", "Uniformidad"
    d.Add "Rev", "Gestión"
    d.Add "Doc", "Gestión"
    d.Add "AA", "Humedad"
    d.Add "Planta", "Servicios"
    d.Add "Equipo", "Mantenimiento"
    d.Add "Valid", "Validación"
    d.Add "CP", "Control Interno"
    d.Add "CI", "Control Interno"
    d.Add "OBS", "Observaciones"
    d.Add "TA", "Administrativo"

    Set Load_DiccionarioCategorias = d
End Function



Function Load_DiccionarioTipos() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "Visc", "Fisicoquímico"
    d.Add "pH", "Fisicoquímico"
    d.Add "PI", "Fisicoquímico"
    d.Add "Valor", "Fisicoquímico"
    d.Add "KF", "Instrumental"
    d.Add "TS", "Instrumental"
    d.Add "IO", "Instrumental"
    d.Add "ID", "Fisicoquímico"
    d.Add "UUD", "Fisicoquímico"
    d.Add "UDS", "Fisicoquímico"
    d.Add "Rev", "Documental"
    d.Add "Doc", "Documental"
    d.Add "AA", "Instrumental"
    d.Add "Planta", "Servicios"
    d.Add "Equipo", "Soporte"
    d.Add "Valid", "Documental"
    d.Add "CP", "Control Interno"
    d.Add "CI", "Control Interno"
    d.Add "OBS", "Documental"
    d.Add "TA", "Documental"
    
    Set Load_DiccionarioTipos = d
End Function

Function Load_DiccionarioEquipos() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "hplc", "HPLC"
    d.Add "uv", "Espectrofotómetro UV"
    d.Add "brookfield", "Viscosímetro Brookfield"
    d.Add "kf", "Karl Fischer"
    d.Add "ph", "pHmetro"
    
    Set Load_DiccionarioEquipos = d
End Function

Function Load_DiccionarioEstandaresPrimarios() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "amoxicilina", "Amoxicilina"
    d.Add "cefale", "Cefalexina"
    d.Add "metformina", "Metformina"
    d.Add "atorvastatina", "Atorvastatina"
    d.Add "ibuprofeno", "Ibuprofeno"
    d.Add "azitro", "Azitromicina"
    d.Add "rifampi", "Rifampicina"
    d.Add "gliclazida", "Gliclazida"
    
    Set Load_DiccionarioEstandaresPrimarios = d
End Function


Function Load_DiccionarioEstandaresSecundarios() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "amoxicilina", "Amoxicilina (Sec)"
    d.Add "metformina", "Metformina (Sec)"
    d.Add "sulfamet", "Sulfametoxazol (Sec)"
    d.Add "morfina", "Morfina (Sec)"
    d.Add "diazepam", "Diazepam (Sec)"
    
    Set Load_DiccionarioEstandaresSecundarios = d
End Function

Function Load_DiccionarioTecnicas() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    d.Add "Visc", "Viscosimetría"
    d.Add "pH", "Potenciometría"
    d.Add "PI", "Punto de Ignición"
    d.Add "IO", "Cromatografía"
    d.Add "KF", "Karl Fischer"
    d.Add "TS", "Ensayo de Disolución"
    d.Add "ID", "Identificación"
    d.Add "Valor", "Valoración"
    d.Add "UUD", "Uniformidad"
    d.Add "UDS", "Uniformidad"
    d.Add "Rev", "Revisión Documental"
    d.Add "Doc", "Documentación"
    d.Add "AA", "Actividad de Agua"
    d.Add "Planta", "Control de Agua"
    d.Add "Equipo", "Mantenimiento"
    d.Add "Valid", "Validación"
    d.Add "CP", "Control Interno"
    d.Add "CI", "Control Interno"
    d.Add "OBS", "Observación"
    d.Add "TA", "Tarea Administrativa"

    Set Load_DiccionarioTecnicas = d
End Function



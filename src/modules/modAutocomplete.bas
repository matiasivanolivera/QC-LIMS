Attribute VB_Name = "modAutocomplete"
Option Explicit

' Variable pública para transferencia de datos entre formulario y módulo
Public ValorSeleccionado As String
Public AutocompleteActivo As Boolean

' ---------------------------------------------------------------
'  Detectar ingreso de "$" en celdas del cronograma
' ---------------------------------------------------------------
Public Sub DetectarAutocompletado(ByVal Target As Range)

    On Error GoTo Salir

    ' Solo una celda
    If Target.CountLarge > 1 Then Exit Sub

    ' Solo texto
    If Not VarType(Target.Value) = vbString Then Exit Sub

    ' Solo si empieza con "$"
    If Left$(Trim$(Target.Value), 1) <> "$" Then Exit Sub

    ' Evitar reentradas
    If AutocompleteActivo Then Exit Sub
    AutocompleteActivo = True

    ' ==========================
    ' CARGAR LISTA DESDE HOJA
    ' ==========================
    Dim colVals As Collection
    Set colVals = ObtenerValoresDeHoja(ThisWorkbook.Sheets("LISTA_$"))

    ' Mostrar formulario
    frmAutocomplete.CargarLista colVals
    frmAutocomplete.Show

    ' ==========================
    ' SI EL USUARIO AGREGA UN VALOR NUEVO
    ' ==========================
    If frmAutocomplete.MaestroSeleccionado = "NUEVO" Then
        Dim ws As Worksheet, lf As Long
        Set ws = ThisWorkbook.Sheets("LISTA_$")

        lf = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
        ws.Cells(lf, 1).Value = frmAutocomplete.ValorSeleccionado
    End If

    ' ==========================
    ' ASIGNAR RESULTADO A LA CELDA
    ' ==========================
    If frmAutocomplete.ValorSeleccionado = "" Then
        Target.Value = ""
    Else
        Target.Value = frmAutocomplete.ValorSeleccionado
    End If

Salir:
    AutocompleteActivo = False

End Sub



' ---------------------------------------------------------------
'  Función para saber si la hoja es semanal (ENE(1), FEB(2), etc.)
' ---------------------------------------------------------------
Public Function EsHojaCronograma(wsName As String) As Boolean
    On Error Resume Next

    ' Debe contener "(" y ")"
    If InStr(wsName, "(") > 0 And InStr(wsName, ")") > 0 Then
        EsHojaCronograma = True
    Else
        EsHojaCronograma = False
    End If

End Function



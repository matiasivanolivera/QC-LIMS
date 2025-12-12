VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOTReview 
   Caption         =   "Revisión de actividades pendientes"
   ClientHeight    =   8520
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "frmOTReview.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmOTReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================
' UserForm : frmOTReview
' Proyecto : CMDT '25 – mini LIMS
' Rol      : Revisar actividades pendientes
'            en hoja semanal y generar OT.
'
' Interfaz:
'   - cmbAnalista    : filtro por analista
'   - lvActividades  : lista de cActividadOT
'   - cmdConfirmar   : genera OT + LOG
'   - cmdCancelar    : cierra el formulario
'
' Depende de:
'   - cActividadOT
'   - modParserIndustrial (ExtraerActividades..., ObtenerAnalistaDesdeBloque)
'   - modDiccionarios
'   - modOT (GenerarOT_ID, RegistrarOT, ColorearCeldas)
'   - modValidadoresOT (OTExiste, etc.)
'==========================================
Option Explicit

'===============================
' VARIABLES DE MÓDULO
'===============================
Private colActsFull As Collection        ' todas las actividades detectadas
Private bloqueandoCombo As Boolean       ' para no disparar eventos mientras cargamos

'===============================
' CONFIGURAR LISTVIEW
'===============================
Private Sub ConfigurarListView()

    With lvActividades
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .HideSelection = False

        ' Si no hay columnas definidas, las creamos ahora
        If .ColumnHeaders.Count = 0 Then
            With .ColumnHeaders
                .Add , , "Ensayo", 80
                .Add , , "Muestra", 120
                .Add , , "NP/Lote", 70
                .Add , , "Técnica", 70
                .Add , , "Analista", 80
                .Add , , "Fecha", 70
                .Add , , "Hoja", 60
                .Add , , "Celda", 60
            End With
        End If

        .ListItems.Clear
    End With

End Sub


'===============================
' CARGAR ACTIVIDADES DESDE HOJA
'===============================
Public Sub CargarActividades()

    Dim wsh As Worksheet
    Dim i As Long
    Dim cel As Range
    Dim a As cActividadOT
    Dim dicAnalistas As Object
    Dim k As Variant
    Dim idAnalista As String

    Set wsh = ActiveSheet

    '--- ListView listo ---
    ConfigurarListView

    '--- Maestros (NP, FF, etc.) ---
    CargarDiccionariosDesdeMaestros

    '--- Parser industrial sobre la hoja activa ---
    Set colActsFull = ExtraerActividadesIndustrialesDesdeHoja(wsh)

    If colActsFull Is Nothing Or colActsFull.Count = 0 Then
        MsgBox "No se detectaron actividades industriales válidas en " & wsh.Name, _
               vbInformation, "Revisión de actividades"
        Exit Sub
    End If

    '--- Diccionario local de analistas con actividad ---
    Set dicAnalistas = CreateObject("Scripting.Dictionary")
    dicAnalistas.CompareMode = vbTextCompare

    '--- Completar Fecha + Analista en cada actividad ---
    For i = 1 To colActsFull.Count

        Set a = colActsFull(i)
        Set cel = wsh.Range(a.Celda)

        ' Fecha por columna (ya la tenías implementada)
        a.Fecha = ObtenerFechaDesdeColumna(wsh, cel.Column)

        ' Analista según MAESTRO_ANALISTAS (rangos de fila)
        idAnalista = ObtenerAnalistaDesdeBloque(wsh, cel.row, cel.Column)
        a.Analista = idAnalista

        ' Solo agregamos IDs válidos
        If idAnalista <> "" Then
            If Not dicAnalistas.Exists(idAnalista) Then
                dicAnalistas.Add idAnalista, True
            End If
        End If

    Next i

    '--- Cargar combo de analistas ---
    bloqueandoCombo = True

    With cmbAnalista
        .Clear
        .AddItem "(Todos)"
        For Each k In dicAnalistas.Keys
            .AddItem CStr(k)
        Next k
        .ListIndex = 0
    End With

    bloqueandoCombo = False

    '--- Mostrar todas las actividades al inicio ---
    RefrescarListaActividades ""

End Sub



'===============================
' REFRESCAR LISTVIEW SEGÚN FILTRO
'===============================
Private Sub RefrescarListaActividades(ByVal filtroAnalista As String)

    Dim i As Long
    Dim a As cActividadOT
    Dim it As ListItem
    Dim filtro As String

    lvActividades.ListItems.Clear

    If colActsFull Is Nothing Then Exit Sub
    If colActsFull.Count = 0 Then Exit Sub

    filtro = UCase$(Trim$(filtroAnalista))

    For i = 1 To colActsFull.Count

        Set a = colActsFull(i)

        If filtro = "" Or UCase$(a.Analista) = filtro Then

            Set it = lvActividades.ListItems.Add(, , a.Ensayo)  ' col 0 = Ensayo
            it.SubItems(1) = a.Muestra
            it.SubItems(2) = a.NPLote
            it.SubItems(3) = a.tecnica
            it.SubItems(4) = a.Analista
            it.SubItems(5) = Format$(a.Fecha, "dd/mm/yyyy")
            it.SubItems(6) = a.Hoja
            it.SubItems(7) = a.Celda

        End If

    Next i

End Sub

'===============================
' CAMBIO EN COMBO DE ANALISTAS
'===============================
Private Sub cmbAnalista_Change()

    Dim filtro As String

    If bloqueandoCombo Then Exit Sub

    If cmbAnalista.ListIndex <= 0 Then
        filtro = ""
    Else
        filtro = cmbAnalista.Value
    End If

    RefrescarListaActividades filtro

End Sub


'===============================
' BOTÓN CANCELAR
'===============================
Private Sub cmdCancelar_Click()
    Unload Me
End Sub


'===============================
' BOTÓN CONFIRMAR = GENERAR OT
'===============================
Private Sub cmdConfirmar_Click()

    Dim i As Long
    Dim it As ListItem
    Dim colSel As New Collection
    Dim a As cActividadOT
    Dim analistaRef As String
    Dim fechaMin As Date
    Dim primer As Boolean
    Dim OT_ID As String

    '-------------------------------
    ' REUNIR ITEMS SELECCIONADOS
    '-------------------------------
    For i = 1 To lvActividades.ListItems.Count

        Set it = lvActividades.ListItems(i)

        If it.Checked Then

            Set a = New cActividadOT

            ' Mapeo según columnas del ListView
            a.Ensayo = it.Text          ' Columna 0 (texto principal)
            a.Muestra = it.SubItems(1)  ' Columna 1
            a.NPLote = it.SubItems(2)   ' Columna 2
            a.tecnica = it.SubItems(3)  ' Columna 3
            a.Analista = it.SubItems(4) ' Columna 4

            If IsDate(it.SubItems(5)) Then
                a.Fecha = CDate(it.SubItems(5))  ' Columna 5 = Fecha
            Else
                a.Fecha = 0
            End If

            a.Hoja = it.SubItems(6)      ' Columna 6
            a.Celda = it.SubItems(7)     ' Columna 7

            colSel.Add a

        End If
    Next i

    If colSel.Count = 0 Then
        MsgBox "Debe seleccionar al menos una actividad.", vbExclamation
        Exit Sub
    End If

    '-------------------------------
    ' VALIDAR QUE SEA UN SOLO ANALISTA
    ' Y CALCULAR FECHA MÍNIMA
    '-------------------------------
    primer = True

    For i = 1 To colSel.Count

        Set a = colSel(i)

        If primer Then
            analistaRef = a.Analista
            fechaMin = a.Fecha
            primer = False
        Else
            If UCase$(a.Analista) <> UCase$(analistaRef) Then
                MsgBox "Las actividades seleccionadas pertenecen a más de un analista." & vbCrLf & _
                       "Genere OTs separadas por analista.", vbCritical
                Exit Sub
            End If
            If a.Fecha <> 0 And a.Fecha < fechaMin Then
                fechaMin = a.Fecha
            End If
        End If

    Next i

    '-------------------------------
    ' GENERAR OT_ID ÚNICO
    '-------------------------------
    OT_ID = GenerarOT_ID(fechaMin, analistaRef)

    '-------------------------------
    ' EVITAR DUPLICADOS
    '-------------------------------
    If OTExiste(OT_ID) Then
        MsgBox "La OT " & OT_ID & " ya existe. No se puede duplicar.", vbCritical
        Exit Sub
    End If

    '-------------------------------
    ' REGISTRAR OT
    '-------------------------------
    RegistrarOT OT_ID, colSel

    '-------------------------------
    ' COLOREAR CELDAS PROCESADAS
    '-------------------------------
    ColorearCeldas colSel

    '-------------------------------
    ' LOG
    '-------------------------------
    RegistrarLOG_OT "Creada " & OT_ID & " con " & colSel.Count & " actividades."

    MsgBox "OT generada correctamente: " & OT_ID, vbInformation
    Unload Me

End Sub

Private Sub UserForm_Click()
    ' reservado
End Sub



VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCerrarJornada_OLD 
   Caption         =   "Cerrar jornada del analista"
   ClientHeight    =   10632
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15384
   OleObjectBlob   =   "frmCerrarJornada_OLD.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCerrarJornada_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConfigurarListView()

    Dim col As ColumnHeader
    
    With lvActividades
        
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .CheckBoxes = True
        .HideColumnHeaders = False

        ' Limpiar columnas previas
        .ColumnHeaders.Clear

        ' Crear columnas
        .ColumnHeaders.Add , , "Tipo", 80
        .ColumnHeaders.Add , , "Producto", 160
        .ColumnHeaders.Add , , "Muestra", 80
        .ColumnHeaders.Add , , "Ensayo", 140
        .ColumnHeaders.Add , , "Forma", 100
        .ColumnHeaders.Add , , "Analista", 80
        .ColumnHeaders.Add , , "Descripción", 220
        
    End With

End Sub

Private Sub UserForm_Initialize()

    ' Configurar columnas del ListView
    ConfigurarListView

    ' Cargar analistas en ComboBox
    CargarAnalistas

    ' Preparar fecha
    txtFecha.Value = Date
    
End Sub

Private Sub CargarAnalistas()

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("MAESTRO_ANALISTAS")

    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    cboAnalista.Clear

    For i = 2 To ultFila
        cboAnalista.AddItem ws.Cells(i, 1).Value   ' Alias
    Next i

End Sub

Private Sub CargarActividadesPorFechaYAnalista()

    Dim ws As Worksheet
    Dim i As Long, ultFila As Long
    Dim fechaSel As Date, analistaSel As String
    Dim item As ListItem
    
    fechaSel = txtFecha.Value
    analistaSel = cboAnalista.Value
    
    lvActividades.ListItems.Clear
    
    For Each ws In ThisWorkbook.Worksheets
        
        If EsHojaCronograma(ws.Name) Then
            
            ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

            For i = 2 To ultFila

                If ws.Cells(i, 2).Value = fechaSel _
                And ws.Cells(i, 6).Value = analistaSel _
                And ws.Cells(i, 7).Value <> "" Then
                    
                    Set item = lvActividades.ListItems.Add(, , ws.Cells(i, 11).Value)  ' Tipo
                    
                    item.SubItems(1) = ws.Cells(i, 7).Value   ' Producto / Actividad original
                    item.SubItems(2) = ws.Cells(i, 15).Value  ' Muestra / NP / Lote
                    item.SubItems(3) = ws.Cells(i, 13).Value  ' Ensayo
                    item.SubItems(4) = ws.Cells(i, 12).Value  ' Forma farmacéutica
                    item.SubItems(5) = ws.Cells(i, 6).Value   ' Analista
                    item.SubItems(6) = ws.Cells(i, 7).Value   ' Descripción completa

                End If

            Next i

        End If
    
    Next ws

End Sub




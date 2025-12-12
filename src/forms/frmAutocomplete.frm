VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutocomplete 
   Caption         =   "Autocompletar"
   ClientHeight    =   5772
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5856
   OleObjectBlob   =   "frmAutocomplete.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAutocomplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ValorSeleccionado As String
Public MaestroSeleccionado As String
Public ListaValores As Collection

' Inicializa el formulario
Private Sub UserForm_Initialize()
    ValorSeleccionado = ""
    MaestroSeleccionado = ""
    txtBuscar.Text = ""
    lstResultados.Clear
End Sub

' Cargar lista inicial
Public Sub CargarLista(colValores As Collection)
    Set ListaValores = colValores
    Call FiltrarLista("")
End Sub

' Filtro dinámico al escribir
Private Sub txtBuscar_Change()
    Call FiltrarLista(txtBuscar.Text)
End Sub

' Aplica filtro
Private Sub FiltrarLista(filtro As String)
    Dim v As Variant
    Dim item As Variant
    
    lstResultados.Clear
    
    filtro = LCase(filtro)
    
    For Each item In ListaValores
        If filtro = "" Or InStr(LCase(item), filtro) > 0 Then
            lstResultados.AddItem item
        End If
    Next item
End Sub

' Doble clic = seleccionar directo
Private Sub lstResultados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstResultados.ListIndex >= 0 Then
        ValorSeleccionado = lstResultados.Value
        Me.Hide
    End If
End Sub

' Botón Seleccionar
Private Sub cmdSeleccionar_Click()
    If lstResultados.ListIndex < 0 Then
        MsgBox "Seleccione un valor de la lista.", vbInformation
        Exit Sub
    End If
    
    ValorSeleccionado = lstResultados.Value
    Me.Hide
End Sub

' Botón Cancelar
Private Sub cmdCancelar_Click()
    ValorSeleccionado = ""
    Me.Hide
End Sub

' Botón Agregar Nuevo
Private Sub cmdNuevo_Click()
    Dim nuevo As String
    nuevo = InputBox("Ingrese el nuevo valor para el maestro:", "Agregar nuevo")

    If Trim(nuevo) = "" Then Exit Sub

    ValorSeleccionado = nuevo
    MaestroSeleccionado = "NUEVO"  ' marca especial para el módulo principal
    Me.Hide
End Sub

